using JurhanLib.Import;
using JurhanLib.Logger;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JurhanService_RozuctovanieDopravcov
{
    /// <summary>
    /// JEDNORAZOVÁ obnova. Vráti konkrétne emaily, ktoré beh 2026-07-24 22:21 presunul po úspešnom
    /// rozúčtovaní do podpriečinkov „Zaúčtované“, späť do ich rodičovských priečinkov.
    /// Po použití môže byť trieda odstránená.
    ///
    /// Pozn.: vracia VŠETKY emaily s presne zhodným predmetom nájdené v danom „Zaúčtované“ podpriečinku.
    /// Ak by tam bol z predchádzajúceho behu ďalší email s rovnakým predmetom, vráti sa tiež.
    /// </summary>
    internal class ObnovaPresunutychEmailov
    {
        private const string NazovPodpriecinkaZauctovane = "Zaúčtované";

        private readonly RozuctovanieLogger _logger;

        // rodičovský priečinok (IMAP full name) -> presné predmety emailov, ktoré sa majú vrátiť
        private static readonly Dictionary<string, string[]> _naVratenie = new Dictionary<string, string[]>
        {
            ["INBOX.Dopravcovia.DPD RO"] = new[]
            {
                "DPD COD2026.07.20 (contr. 2038813)",
                "DPD COD2026.07.21 (contr. 2038813)",
                "DPD COD2026.07.22 (contr. 2038813)",
                "DPD COD2026.07.23 (contr. 2038813)",
            },
            ["INBOX.Dopravcovia.GLS RO"] = new[]
            {
                "Lista Colete cu Ramburs COD list – 21.07.2026",
                "Lista Colete cu Ramburs COD list – 22.07.2026",
            },
        };

        internal ObnovaPresunutychEmailov(RozuctovanieLogger logger)
        {
            _logger = logger;
        }

        internal void Execute()
        {
            using (ImapClient client = new ImapClient())
            {
                client.Connect(Constants.ClientHostImapJurhan, 993, true);
                client.Authenticate(Constants.MessageToPlatbyJurhan, Constants.ClientPasswordJurhan);

                foreach (KeyValuePair<string, string[]> par in _naVratenie)
                {
                    try
                    {
                        VratEmaily(client, par.Key, par.Value);
                    }
                    catch (Exception ex)
                    {
                        _logger.Loguj($"Obnova: chyba pri priečinku '{par.Key}': {ex}", true, FarbyLogu.Chyba);
                    }
                }

                client.Disconnect(true);
            }
        }

        private void VratEmaily(ImapClient client, string rodicFullName, string[] predmety)
        {
            IMailFolder rodic = client.GetFolder(rodicFullName);
            IMailFolder zauctovane = rodic.GetSubfolders(false)
                .FirstOrDefault(f => f.Name == NazovPodpriecinkaZauctovane);
            if (zauctovane == null)
            {
                _logger.Loguj($"Obnova: podpriečinok '{NazovPodpriecinkaZauctovane}' v '{rodicFullName}' neexistuje - preskakujem.", true, FarbyLogu.Chyba);
                return;
            }

            zauctovane.Open(FolderAccess.ReadWrite);

            HashSet<string> hladanePredmety = new HashSet<string>(predmety, StringComparer.Ordinal);
            HashSet<string> vratene = new HashSet<string>(StringComparer.Ordinal);

            // jeden prechod: každé UID sa dotkne raz (presun neinvaliduje ostatné UID)
            foreach (UniqueId uid in zauctovane.Search(SearchQuery.All))
            {
                MimeMessage sprava = zauctovane.GetMessage(uid);
                if (hladanePredmety.Contains(sprava.Subject))
                {
                    zauctovane.MoveTo(uid, rodic);
                    vratene.Add(sprava.Subject);
                    _logger.Loguj($"Obnova: email '{sprava.Subject}' vrátený z '{zauctovane.FullName}' do '{rodicFullName}'.", true, FarbyLogu.Uspech);
                }
            }

            foreach (string predmet in predmety)
            {
                if (!vratene.Contains(predmet))
                {
                    _logger.Loguj($"Obnova: email s predmetom '{predmet}' sa v '{zauctovane.FullName}' nenašiel.", true, FarbyLogu.Chyba);
                }
            }
        }
    }
}
