using JurhanLib.Email;
using JurhanLib.Import;
using JurhanLib.Import.Dopravcovia;
using JurhanModels.Enum;
using JurhanModels.Import;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
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

        // PILOT (skusobne nasadenie): v tychto priecinkoch sa robi ostre rozuctovanie (import do Omegy),
        // vo vsetkych ostatnych sa rozuctovanie iba simuluje a loguje, co by sa v ostrej prevadzke stalo.
        private static readonly string[] _pilotnePriecinky = {
            "INBOX.Dopravcovia.DPD CZ",
            "INBOX.Dopravcovia.DPD HR",
            "INBOX.Dopravcovia.DPD PL",
            "INBOX.Dopravcovia.DPD SI",
            "INBOX.Dopravcovia.DPD HU",
            "INBOX.Dopravcovia.DPD RO",
            "INBOX.Dopravcovia.DPD SK",
            "INBOX.Dopravcovia.GLS CZ",
            "INBOX.Dopravcovia.GLS HU",
            "INBOX.Dopravcovia.GLS PLN",
            "INBOX.Dopravcovia.GLS RO",
            "INBOX.Dopravcovia.PACKETA",
            "INBOX.Dopravcovia.SPS",
        };
        // PILOT krok b (zapnut az pred nasadenim na server): v pilotnych priecinkoch sa emaily aj presuvaju
        // do podpriecinka "Zaúčtované"; kym je false, presun sa iba loguje.
        private const bool PresuvatVPilotnychPriecinkoch = true;

        private readonly PripojeneFirmy _pripojeneFirmy;
        private readonly string _workDir;
        private readonly RozuctovanieLogger _logger;
        // suhrnne zoznamy z rozuctovania za cely beh - na konci sa z nich vytvoria max 4 subory
        // (Rozuctovane / Nesparovane / Uzuhradene / Nerozuctovane) a poslu jednym emailom
        private readonly SuhrnneZoznamyRozuctovania _suhrnneZoznamy = new SuhrnneZoznamyRozuctovania();
        // vynimky (s call stackom) zachytene pocas behu - na konci sa z nich sklada celkovy vysledok do .err suboru
        private readonly List<string> _chybyBehu = new List<string>();
        // nazvy priecinkov, ktore sme v schranke naozaj videli - na konci z nich zistime, ci sa niektory
        // pilotny priecinok netrafil (preklep v nazve tichu vyradil DPD RO z pilotu a iba sa simulovalo)
        private readonly HashSet<string> _videnePriecinky = new HashSet<string>();
        // Omega nedokoncila import v limite - dalsie subory by na nu cakali rovnako dlho (pri desiatkach
        // emailov aj hodinu), preto sa beh prerusi a nespracovane emaily ostanu na dalsi beh
        private bool _importNedostupny;

        internal RozuctovanieEmailov(PripojeneFirmy pripojeneFirmy, RozuctovanieLogger logger)
        {
            _pripojeneFirmy = pripojeneFirmy;
            _logger = logger;
            _workDir = Path.Combine(pripojeneFirmy._omegaPath, "Import", "RozuctovanieDopravcov");
        }

        internal void Execute()
        {
            Directory.CreateDirectory(_workDir);
            ServicesLog.VytvorLogovaciAdresar();
            _logger.NacitajDataZoSuboru();
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
                _chybyBehu.Add($"Pracovný adresár '{_workDir}': {ex}");
            }

            using (ImapClient client = new ImapClient())
            {
                client.Connect(Constants.ClientHostImapJurhan, 993, true);
                client.Authenticate(Constants.MessageToPlatbyJurhan, Constants.ClientPasswordJurhan);

                var folders = client.GetFolders(client.PersonalNamespaces[0]);

                foreach (IMailFolder folder in folders)
                {
                    if (_importNedostupny)
                    {
                        break;
                    }

                    //if (!string.IsNullOrEmpty(folder.ParentFolder?.FullName))
                    //{
                    //    // spracuvame len priecinky priamo pod korenom schranky (podpriecinky = "Zaúčtované" a pod.)
                    //    continue;
                    //}

                    if (folder.Name == NazovPodpriecinkaZauctovane)
                    {
                        continue;
                    }

                    _videnePriecinky.Add(folder.FullName);

                    eTypSuboru typSuboru = FolderMapping.DajTypSuboru(folder.Name);
                    if (typSuboru == eTypSuboru.Undefined)
                    {
                        _logger.PrazdnyRiadok(1);
                        _logger.Loguj($"Priečinok '{folder.FullName}' nie je namapovaný na typ dopravcu - preskakujem.", true);
                        _logger.PrazdnyRiadok(1);
                        continue;
                    }

                    try
                    {
                        SpracujPriecinok(folder, typSuboru);
                    }
                    catch (Exception ex)
                    {
                        _logger.Loguj($"Chyba pri spracovaní priečinka '{folder.FullName}': {ex}", true);
                        _chybyBehu.Add($"Priečinok '{folder.FullName}': {ex}");
                    }
                }

                client.Disconnect(true);
            }

            SkontrolujNazvyPilotnychPriecinkov();
            PosliLogSuboryRozuctovania();
            ZapisCelkovyVysledokDoErrSuboru();

            _logger.ZapisDataDoSuboru();
            _logger.PosliLogSuborEmailom(AdresatiEmailov());
        }

        /// <summary>
        /// Adresati oboch emailov servisy (subory z rozuctovania aj logy sluzby):
        /// v mode servisa obchod aj Tuleja, v mode program iba Tuleja.
        /// </summary>
        private static List<string> AdresatiEmailov()
        {
            return Program.typSpustenia == eTypSpustenia.Servica
                ? new List<string> { Constants.MessageToPlatbyJurhan, Constants.MessageToObchodJurhan,
                    Constants.MessageToTulejaX }
                : new List<string> { Constants.MessageToTulejaX };
        }

        /// <summary>
        /// Pilotný priečinok, ktorý sa v schránke nenašiel, znamená preklep v <see cref="_pilotnePriecinky"/> -
        /// taký priečinok sa potichu rozúčtováva iba nasucho. Radšej to nahlásime, ako aby to zas niekto objavil
        /// až z toho, že doklady v Omege chýbajú.
        /// </summary>
        private void SkontrolujNazvyPilotnychPriecinkov()
        {
            if (_importNedostupny)
            {
                // beh sme prerušili, časť priečinkov sme ani nevideli - hlásenie by bolo falošné
                return;
            }

            List<string> nenajdene = _pilotnePriecinky.Where(p => !_videnePriecinky.Contains(p)).ToList();
            if (nenajdene.Any())
            {
                _logger.Loguj($"[PILOT] Tieto pilotné priečinky v schránke neexistujú - preklep v názve? " +
                    $"{string.Join(", ", nenajdene)}", true);
                _chybyBehu.Add($"Pilotné priečinky, ktoré sa v schránke nenašli: {string.Join(", ", nenajdene)}");
            }
        }

        /// <summary>
        /// .err subor ma odrazat cely beh, nie poslednu hlasku z rozuctovania jedneho suboru (hlasky typu
        /// "Nenašiel sa bankový výpis..." nie su chyby behu). Ak pocas behu nastali vynimky, zapisu sa
        /// vsetky s call stackom; inak sa zapise, ze beh prebehol v poriadku.
        /// </summary>
        private void ZapisCelkovyVysledokDoErrSuboru()
        {
            if (_chybyBehu.Any())
            {
                ServicesError.ErrorEnd($"Beh rozúčtovania {DateTime.Now:dd.MM.yyyy HH:mm:ss} skončil s chybami ({_chybyBehu.Count}):" +
                    Environment.NewLine + Environment.NewLine +
                    string.Join(Environment.NewLine + Environment.NewLine, _chybyBehu));
            }
            else
            {
                ServicesError.ErrorEnd($"Beh rozúčtovania {DateTime.Now:dd.MM.yyyy HH:mm:ss} prebehol v poriadku.");
            }
        }

        /// <summary>
        /// Suhrnne log subory za cely beh v jednom emaili: Rozuctovane (iba nazvy bezchybnych dokladov),
        /// Nesparovane/Uzuhradene/Nerozuctovane (polozky zoskupene pod nazvom dokladu). Prazdne sa nevytvoria.
        /// </summary>
        private void PosliLogSuboryRozuctovania()
        {
            List<string> prilohy = _suhrnneZoznamy.ZapisSubory(ServicesLog.NameApplicationLogPath());
            if (!prilohy.Any())
            {
                return;
            }

            bool odoslane = EmailService.PosliEmail(
                AdresatiEmailov(),
                null,
                null,
                "Log súbory z rozúčtovania dopravcov",
                "V prílohe posielam log súbory zo spustenia programu pre rozúčtovanie dopravcov",
                prilohy);
            if (!odoslane)
            {
                _logger.Loguj($"Nepodarilo sa odoslať email so súbormi z rozúčtovania " +
                    $"({string.Join(", ", prilohy.Select(Path.GetFileName))}).", true);
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

            bool pilotny = _pilotnePriecinky.Contains(folder.FullName);

            _logger.PrazdnyRiadok(1);
            _logger.Loguj($"Priečinok '{folder.FullName}' ({typSuboru}): {uids.Count} emailov.", true);
            if (!pilotny)
            {
                _logger.Loguj($"[PILOT] Priečinok je mimo pilotu - rozúčtovanie sa iba simuluje (bez importu do Omegy a bez presunov).", true);
            }
            _logger.PrazdnyRiadok(1);

            IMailFolder zauctovane = null;
            foreach (UniqueId uid in uids)
            {
                if (_importNedostupny)
                {
                    break;
                }

                try
                {
                    MimeMessage message = folder.GetMessage(uid);
                    if (SpracujEmail(message, typSuboru, folder.Name, ibaSimulacia: !pilotny))
                    {
                        if (pilotny && PresuvatVPilotnychPriecinkoch)
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
                        else
                        {
                            _logger.Loguj($"[PILOT] Email '{message.Subject}' by bol presunutý do " +
                                $"'{folder.FullName}.{NazovPodpriecinkaZauctovane}'" +
                                (pilotny ? " (presun v pilote zatiaľ vypnutý)." : " (priečinok mimo pilotu)."), true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri spracovaní emailu {uid} v priečinku '{folder.FullName}': {ex}", true);
                    _chybyBehu.Add($"Email {uid} v priečinku '{folder.FullName}': {ex}");
                }
            }
        }

        /// <returns>true, ak sa ma email presunut do podpriecinka "Zaúčtované"</returns>
        private bool SpracujEmail(MimeMessage message, eTypSuboru typSuboru, string nazovPriecinka, bool ibaSimulacia)
        {
            bool asponJedenSubor = false;
            bool vsetkoZauctovane = true;

            // BodyParts namiesto Attachments - dopravcovia posielaju subory aj ako inline prilohy (bez Content-Disposition: attachment)
            List<MimePart> prilohy = message.BodyParts.OfType<MimePart>()
                .Where(p => !string.IsNullOrEmpty(p.FileName))
                .ToList();

            // niektori dopravcovia (napr. GoPay) posielaju ten isty vypis ako .csv aj .xlsx.
            // .xlsx sa nekonvertuje (nie je v PreConvertCsv) a CSV parser ho nespracuje (Chyba/MalformedLineException),
            // preto ak k rovnakemu nazvu existuje .csv, tabulkoveho dvojnika (.xlsx/.xls) preskakujeme.
            HashSet<string> zakladySCsv = new HashSet<string>(
                prilohy.Where(p => string.Equals(Path.GetExtension(p.FileName), ".csv", StringComparison.OrdinalIgnoreCase))
                       .Select(p => Path.GetFileNameWithoutExtension(p.FileName)),
                StringComparer.OrdinalIgnoreCase);

            foreach (MimePart attachment in prilohy)
            {
                if (_importNedostupny)
                {
                    // zvyšné prílohy tohto emailu sa už nespracovali - email nie je vybavený
                    vsetkoZauctovane = false;
                    break;
                }

                string pripona = Path.GetExtension(attachment.FileName);
                if ((string.Equals(pripona, ".xlsx", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(pripona, ".xls", StringComparison.OrdinalIgnoreCase))
                    && zakladySCsv.Contains(Path.GetFileNameWithoutExtension(attachment.FileName)))
                {
                    _logger.Loguj($"Email '{message.Subject}': prílohu {attachment.FileName} preskakujem (existuje .csv verzia).", true);
                    continue;
                }

                // PRIPRAVENÉ NA OSTRÚ PREVÁDZKU (zatiaľ zakomentované): samostatný email s duplikátom -
                // súbor s rovnakým názvom (inou príponou) už bol rozúčtovaný z iného emailu, počíta sa ako
                // zaúčtovaný, aby sa aj tento email presunul do podpriečinka "Zaúčtované".
                //if (_povolenePripony.Contains(Path.GetExtension(attachment.FileName).ToLowerInvariant())
                //    && SuborUzBolRozuctovany(attachment.FileName))
                //{
                //    _logger.Loguj($"Email '{message.Subject}': súbor {attachment.FileName} už bol rozúčtovaný z iného emailu - počítam ako zaúčtovaný.", true);
                //    asponJedenSubor = true;
                //    continue;
                //}

                string filePath = UlozPrilohu(attachment);
                if (filePath == null)
                {
                    continue;
                }

                asponJedenSubor = true;
                _logger.Loguj($"Email '{message.Subject}': spracúvam súbor {Path.GetFileName(filePath)}.", true);

                eVysledokRozuctovania vysledok = SpracujSubor(filePath, typSuboru, nazovPriecinka, ibaSimulacia);
                _logger.Loguj($"Súbor {Path.GetFileName(filePath)}: {vysledok}.", true);

                // duplicita = subor uz bol zauctovany skor; vsetko uz uhradene = najdene faktury su uz
                // zaplatene (Emag) -> email v oboch pripadoch patri do "Zaúčtované"
                if (vysledok != eVysledokRozuctovania.Rozuctovane && vysledok != eVysledokRozuctovania.Duplicita
                    && vysledok != eVysledokRozuctovania.VsetkoUzUhradene)
                {
                    vsetkoZauctovane = false;
                }
            }

            return asponJedenSubor && vsetkoZauctovane;
        }

        private eVysledokRozuctovania SpracujSubor(string filePath, eTypSuboru typSuboru, string nazovPriecinka, bool ibaSimulacia)
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
                interneCislo = null, // sluzba: doklad sa hlada podla textu hlavicky (C099) a datumu vypisu
                nazovPriecinka = nazovPriecinka,
                ibaSimulacia = ibaSimulacia, // pilot: mimo pilotnych priecinkov sa nic nezapisuje do Omegy
                Loguj = s => _logger.Loguj(s, true), // kritéria hľadania dokladu do logu služby
                PrazdnyRiadok = n => _logger.PrazdnyRiadok(n),
                zobrazenieChyby = eZobrazenieChyby.ZapisDoSuboru,
                typSpustenia = Program.typSpustenia,
            };

            RozuctovanieCore core = new RozuctovanieCore(ctx);
            // zaloha do <exe>\Log s povodnym nazvom prilohy (bez casovej peciatky);
            // subor sa po spracovani z pracovneho adresara zmaze, zaloha ostava (pri rovnakom nazve sa prepise)
            string zalohaDir = ServicesLog.NameApplicationLogPath();
            Directory.CreateDirectory(zalohaDir);
            File.Copy(filePath, Path.Combine(zalohaDir, Path.GetFileName(filePath)), true);
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
                _suhrnneZoznamy.PridajZKontextu(ctx);
                if (ctx.importNedostupny)
                {
                    PrerusBeh();
                }
            }
            return vysledok;
        }

        /// <summary>
        /// Omega nedokončila import v limite. Každý ďalší súbor by na ňu čakal rovnako dlho, preto sa beh
        /// ukončí; nespracované emaily ostávajú v priečinkoch a spracujú sa pri ďalšom behu.
        /// </summary>
        private void PrerusBeh()
        {
            if (_importNedostupny)
            {
                return;
            }
            _importNedostupny = true;
            _logger.Loguj($"Import do Omegy nedobehol v limite - prerušujem beh. Zvyšné emaily ostávajú " +
                $"nespracované v priečinkoch a spracujú sa pri ďalšom behu.", true);
            _chybyBehu.Add("Import do Omegy nedobehol v limite - beh bol prerušený, " +
                "zvyšné emaily ostali nespracované.");
        }

        /// <summary>
        /// Súbor s týmto názvom (bez prípony) už bol rozúčtovaný - existuje doklad s C149_ImportText.
        /// Kontrolujú sa oba vzory: holý názov aj "dopravca: názov" (zapisuje ho rozúčtovanie).
        /// </summary>
        private bool SuborUzBolRozuctovany(string fileName)
        {
            string nazov = Path.GetFileNameWithoutExtension(fileName);
            return new EudHlavickaRepository(_pripojeneFirmy.dataProvider)
                .DajDoklady("C149_ImportText = @1 OR C149_ImportText = @2", nazov, "dopravca: " + nazov)
                .Any();
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
