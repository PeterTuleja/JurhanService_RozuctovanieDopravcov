using JurhanLib.Import;
using JurhanLib.Import.Dopravcovia;
using JurhanModels.Import;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using OmegaLib.Enums;
using OmegaLib.Repository;
using OmegaLib.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JurhanService_RozuctovanieDopravcov
{
    /// <summary>
    /// Prejde IMAP schranku platby@jurhan.com: priecinky dopravcov -> prilohy emailov -> rozuctovanie.
    /// Uspesne spracovane emaily presuva do podpriecinka "Zaúčtované".
    /// </summary>
    internal class RozuctovanieEmailov
    {
        private const string NazovPodpriecinkaZauctovane = "Zaúčtované";
        private static readonly string[] _povolenePripony = { ".csv", ".xlsx", ".xls" };

        private readonly PripojeneFirmy _pripojeneFirmy;
        private readonly RozuctovanieLogger _logger;
        private readonly string _workDir;

        internal RozuctovanieEmailov(PripojeneFirmy pripojeneFirmy, RozuctovanieLogger logger)
        {
            _pripojeneFirmy = pripojeneFirmy;
            _logger = logger;
            _workDir = Path.Combine(pripojeneFirmy._omegaPath, "Import", "RozuctovanieDopravcov");
        }

        internal void Execute()
        {
            Directory.CreateDirectory(_workDir);
            try
            {
                foreach (string f in Directory.GetFiles(_workDir))
                {
                    File.Delete(f);
                }
            }
            catch (Exception ex)
            {
                _logger.Loguj($"Nepodarilo sa vyčistiť pracovný adresár '{_workDir}': {ex}", true);
            }

            using (ImapClient client = new ImapClient())
            {
                client.Connect(Constants.ClientHostImapJurhan, 993, true);
                client.Authenticate(Constants.MessageToPlatbyJurhan, Constants.ClientPasswordJurhan);

                var folders = client.GetFolders(client.PersonalNamespaces[0]);

                foreach (IMailFolder folder in folders)
                {
                    //if (!string.IsNullOrEmpty(folder.ParentFolder?.FullName))
                    //{
                    //    // spracuvame len priecinky priamo pod korenom schranky (podpriecinky = "Zaúčtované" a pod.)
                    //    continue;
                    //}

                    if (folder.Name == NazovPodpriecinkaZauctovane)
                    {
                        continue;
                    }

                    eTypSuboru typSuboru = FolderMapping.DajTypSuboru(folder.Name);
                    if (typSuboru == eTypSuboru.Undefined)
                    {
                        _logger.Loguj($"Priečinok '{folder.FullName}' nie je namapovaný na typ dopravcu - preskakujem.", true);
                        continue;
                    }

                    try
                    {
                        SpracujPriecinok(folder, typSuboru);
                    }
                    catch (Exception ex)
                    {
                        _logger.Loguj($"Chyba pri spracovaní priečinka '{folder.FullName}': {ex}", true);
                    }
                }

                client.Disconnect(true);
            }
        }

        private void SpracujPriecinok(IMailFolder folder, eTypSuboru typSuboru)
        {
            folder.Open(FolderAccess.ReadWrite);

            IList<UniqueId> uids = folder.Search(SearchQuery.NotDeleted);
            if (!uids.Any())
            {
                return;
            }

            _logger.Loguj($"Priečinok '{folder.FullName}' ({typSuboru}): {uids.Count} emailov.", true);

            IMailFolder zauctovane = null;
            foreach (UniqueId uid in uids)
            {
                try
                {
                    MimeMessage message = folder.GetMessage(uid);
                    if (SpracujEmail(message, typSuboru))
                    {
                        if (zauctovane == null)
                        {
                            zauctovane = DajAleboVytvorZauctovane(folder);
                        }

                        // presuvame hned po zauctovani - pad medzi zauctovanim a davkovym presunom
                        // by nechal zauctovane emaily navzdy v priecinku (pri retry vratia ZiadneUhrady)
                        folder.MoveTo(uid, zauctovane);
                        _logger.Loguj($"Email '{message.Subject}' presunutý do '{zauctovane.FullName}'.", true);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri spracovaní emailu {uid} v priečinku '{folder.FullName}': {ex}", true);
                }
            }
        }

        /// <returns>true, ak sa ma email presunut do podpriecinka "Zaúčtované"</returns>
        private bool SpracujEmail(MimeMessage message, eTypSuboru typSuboru)
        {
            bool asponJedenSubor = false;
            bool vsetkoZauctovane = true;

            // BodyParts namiesto Attachments - dopravcovia posielaju subory aj ako inline prilohy (bez Content-Disposition: attachment)
            foreach (MimePart attachment in message.BodyParts.OfType<MimePart>().Where(p => !string.IsNullOrEmpty(p.FileName)))
            {
                string filePath = UlozPrilohu(attachment);
                if (filePath == null)
                {
                    continue;
                }

                asponJedenSubor = true;
                _logger.Loguj($"Email '{message.Subject}': spracúvam súbor {Path.GetFileName(filePath)}.", true);

                eVysledokRozuctovania vysledok = SpracujSubor(filePath, typSuboru);
                _logger.Loguj($"Súbor {Path.GetFileName(filePath)}: {vysledok}.", true);

                // duplicita = subor uz bol zauctovany skor -> email tiez patri do "Zaúčtované"
                if (vysledok != eVysledokRozuctovania.Rozuctovane && vysledok != eVysledokRozuctovania.Duplicita)
                {
                    vsetkoZauctovane = false;
                }
            }

            return asponJedenSubor && vsetkoZauctovane;
        }

        private eVysledokRozuctovania SpracujSubor(string filePath, eTypSuboru typSuboru)
        {
            short mesiac = NazovSuboru.DajMesiac(Path.GetFileNameWithoutExtension(filePath));
            if (Lib.NastavTypRozuctovania(typSuboru) == eTypRozuctovania.BezZapoctuBanky && mesiac == 0)
            {
                _logger.Loguj($"V názve súboru {Path.GetFileName(filePath)} sa nenašiel mesiac (MM.RRRR) - súbor preskakujem.", true);
                return eVysledokRozuctovania.Chyba;
            }
            if (mesiac == 0)
            {
                // pre typy so zapoctom banky sa datum berie z bankoveho dokladu, mesiac je nepodstatny
                mesiac = (short)DateTime.Now.Month;
            }

            RozuctovanieContext ctx = new RozuctovanieContext
            {
                pripojeneFirmy = _pripojeneFirmy,
                typSuboru = typSuboru,
                fileName = filePath,
                mesiac = mesiac,
                interneCislo = null, // bankovy doklad sa hlada podla poznamky = nazov suboru bez pripony
                zobrazenieChyby = eZobrazenieChyby.ZapisDoSuboru,
                typSpustenia = Program.typSpustenia,
            };

            RozuctovanieCore core = new RozuctovanieCore(ctx);
            ServicesFile.ZalohujSubor(filePath); // zaloha do <exe>\Log s casovou peciatkou (spec: subor sa po spracovani zmaze, zaloha ostava)
            eVysledokRozuctovania vysledok;
            try
            {
                vysledok = core.Execute();
            }
            finally
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            return vysledok;
        }

        private string UlozPrilohu(MimePart attachment)
        {
            string fileName = attachment.FileName;
            if (!_povolenePripony.Contains(Path.GetExtension(fileName).ToLowerInvariant()))
            {
                return null;
            }

            string filePath = Path.Combine(_workDir, ToSafeFileName(fileName));
            using (FileStream stream = File.Create(filePath))
            {
                attachment.Content.DecodeTo(stream);
            }
            return filePath;
        }

        private IMailFolder DajAleboVytvorZauctovane(IMailFolder folder)
        {
            IMailFolder zauctovane = folder.GetSubfolders(false).FirstOrDefault(f => f.Name == NazovPodpriecinkaZauctovane);
            if (zauctovane == null)
            {
                zauctovane = folder.Create(NazovPodpriecinkaZauctovane, true);
            }
            return zauctovane;
        }

        private string ToSafeFileName(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                s = s.Replace(c, '_');
            }
            return s;
        }
    }
}
