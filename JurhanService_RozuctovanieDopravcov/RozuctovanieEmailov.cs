using JurhanLib.Email;
using JurhanLib.Import;
using JurhanLib.Import.Dopravcovia;
using JurhanLib.Logger;
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
    ///
    /// Vynimka je Packeta: jej vypisy sa beru z API (PacketaZApi), nie z emailov, takze sa v jej
    /// priecinku emaily necitaju ani nepresuvaju. Emailom totiz chodia len faktury v EUR - faktury
    /// v CZK/HUF/PLN/RON nechodia vobec a z emailov by sa nikdy nerozuctovali.
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
            "INBOX.Dopravcovia.GLS HR",
            "INBOX.Dopravcovia.GLS HU",
            "INBOX.Dopravcovia.GLS PLN",
            "INBOX.Dopravcovia.GLS RO",
            "INBOX.Dopravcovia.GLS SI",
            "INBOX.Dopravcovia.GLS SK",
            "INBOX.Dopravcovia.PACKETA",
            "INBOX.Dopravcovia.SPS",
            // GoPay zapnuty spat 22.08.2026 - server uz preukazatelne bezi s parovanim N+1
            // (v logu 19.08. je riadok "GoPay: pozbieranych X prevodov z vypisov v priecinku")
            "INBOX.Ostatné .Platobné brány, Účty.GoPay",
        };
        // Testovaci rezim (lokalne spustenie nad kopiou databazy). Ked je true, VSETKO sa iba
        // simuluje - nepresuvaju sa emaily a nevola sa autoimport do Omegy (kvoli rychlosti
        // testovania). Na serveri musi byt false: pilotne priecinky vtedy ostro importuju
        // a presuvaju emaily, priecinky mimo pilotu simuluju vzdy.
        private const bool IbaSimulacia = false;

        // DOCASNE (test formatu v9 od Packety nad emailmi, ktore uz boli spracovane a presunute):
        // spracuje sa IBA podpriecinok "Zaúčtované" pod priecinkom PACKETA. Nic sa nepresuva a nic
        // sa neimportuje do Omegy - priecinok nie je v _pilotnePriecinky, takze rozuctovanie ide
        // v simulacii, a schranka sa otvara len na citanie.
        // PO DOTESTOVANI: DocasnyTestPacketaZauctovane = false (alebo cely blok odstranit).
        private const bool DocasnyTestPacketaZauctovane = false;
        private const string TestovaciNazovPriecinka = "PACKETA";
        private const eTypSuboru TestovaciTypSuboru = eTypSuboru.Dopravca_Packeta;

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
        // dopravcovia (najma GLS CZ) posielaju ten isty report aj dvoma emailami; kluc je kontrolny sucet
        // obsahu, aby sa duplicita chytila aj ked pride pod inym nazvom
        private readonly Dictionary<string, SpracovanaPriloha> _spracovanePrilohy =
            new Dictionary<string, SpracovanaPriloha>();
        // bankove vypisy obsadene v tomto behu - na jeden vypis smie byt naparovany len jeden subor
        private readonly HashSet<string> _pouziteBankoveDoklady = new HashSet<string>();
        // GoPay: prevod vyrovnavajuci vypis N je az vo vypise N+1 - prevody z celeho priecinka sa
        // pozbieraju vopred (NacitajPrevodyGopay), aby bol pri spracovani suboru N jeho prevod znamy
        private readonly List<PrevodGopayVypisu> _prevodyGopay = new List<PrevodGopayVypisu>();

        private class SpracovanaPriloha
        {
            internal string nazov;
            internal eVysledokRozuctovania vysledok;
        }

        private class PrevodGopayVypisu
        {
            internal string ucet;
            internal DateTime obdobieOd;
            internal GopayCsv.Prevod prevod;
        }

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

            // vypisy, ktore uz maju rozuctovanie z minulych behov - na ten isty vypis sa druhy subor
            // naparovat nesmie; nacitava sa raz za beh, dotaz ide cez celu evidenciu
            _pouziteBankoveDoklady.UnionWith(RozuctovanieCore.NacitajPouziteBankoveDoklady(_pripojeneFirmy.dataProvider));
            _logger.Loguj($"Bankových výpisov, ktoré už majú rozúčtovanie: {_pouziteBankoveDoklady.Count}.", true);

            try
            {
                foreach (string f in Directory.GetFiles(_workDir))
                {
                    File.Delete(f);
                }
            }
            catch (Exception ex)
            {
                _logger.Loguj($"Nepodarilo sa vyčistiť pracovný adresár '{_workDir}': {ex}", true, FarbyLogu.Chyba);
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

                    eTypSuboru typSuboru;

                    if (DocasnyTestPacketaZauctovane)
                    {
                        // DOCASNE: iba PACKETA z podpriecinka "Zaúčtované"; typ sa nastavuje rucne,
                        // FolderMapping by z nazvu "Zaúčtované" vratil Undefined.
                        // Hlada sa podla nazvu a nadradeneho priecinka, nie podla FullName - oddelovac
                        // v ceste urcuje server a preklep v celej ceste by test potichu vypol
                        if (folder.Name != NazovPodpriecinkaZauctovane
                            || folder.ParentFolder?.Name != TestovaciNazovPriecinka)
                        {
                            continue;
                        }

                        _videnePriecinky.Add(folder.FullName);
                        typSuboru = TestovaciTypSuboru;
                        _logger.PrazdnyRiadok(1);
                        _logger.Loguj($"[TEST v9] Spracúvam iba priečinok '{folder.FullName}' - bez presunov " +
                            $"emailov a bez importu do Omegy.", true, FarbyLogu.Priecinok);
                        _logger.PrazdnyRiadok(1);
                    }
                    else
                    {
                        if (folder.Name == NazovPodpriecinkaZauctovane)
                        {
                            continue;
                        }

                        _videnePriecinky.Add(folder.FullName);

                        typSuboru = FolderMapping.DajTypSuboru(folder.Name);
                        if (typSuboru == eTypSuboru.Undefined)
                        {
                            _logger.PrazdnyRiadok(1);
                            _logger.Loguj($"Priečinok '{folder.FullName}' nie je namapovaný na typ dopravcu - preskakujem.", true, FarbyLogu.Priecinok);
                            _logger.PrazdnyRiadok(1);
                            continue;
                        }
                    }

                    try
                    {
                       SpracujPriecinok(folder, typSuboru);
                       
                    }
                    catch (Exception ex)
                    {
                        _logger.Loguj($"Chyba pri spracovaní priečinka '{folder.FullName}': {ex}", true, FarbyLogu.Chyba);
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
        /// v mode servisa platby, obchod aj Tuleja; v mode program (rucne testovacie behy)
        /// iba Tuleja - firma testovacie emaily dostavat nema.
        /// Pozor: ServicesLogger posiela User log na adresy BEZ tulejax, takze v mode program
        /// sa User log neposle nikam - to je tu zamer a denny log to hlasi riadkom
        /// "User log neposielam...". Developer log ide na tulejax v oboch modoch.
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
            if (DocasnyTestPacketaZauctovane)
            {
                // DOCASNE (test v9): prechádza sa jediný priečinok, kontrola pilotných by hlásila všetky.
                // Namiesto toho hlásime, keď sa testovací priečinok v schránke vôbec nenašiel - inak by
                // beh dopadol ako "nič na spracovanie" a nebolo by vidieť, že sa test nespustil.
                if (!_importNedostupny && !_videnePriecinky.Any())
                {
                    _logger.Loguj($"[TEST v9] Podpriečinok '{NazovPodpriecinkaZauctovane}' pod " +
                        $"'{TestovaciNazovPriecinka}' sa v schránke nenašiel - nespracovalo sa nič.",
                        true, FarbyLogu.Chyba);
                }
                return;
            }

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
                    $"({string.Join(", ", prilohy.Select(Path.GetFileName))}).", true, FarbyLogu.Chyba);
            }
        }


        private void SpracujPriecinok(IMailFolder folder, eTypSuboru typSuboru)
        {
            // DOCASNE (test v9): schranka sa otvara len na citanie, aby sa emailom v "Zaúčtované"
            // nemohlo stat nic ani pri chybe v kode.
            // Packeta ide z API a emaily sa pri nej nepresuvaju, takze tam citanie staci tiez.
            bool ibaCitanie = DocasnyTestPacketaZauctovane || Lib.JeTypSuboruZApi(typSuboru);
            folder.Open(ibaCitanie ? FolderAccess.ReadOnly : FolderAccess.ReadWrite);

            IList<UniqueId> uids = folder.Search(SearchQuery.NotDeleted);

            // DOCASNE (test v9): testovaci priecinok v _pilotnePriecinky nie je, takze pilotny = false
            // a rozuctovanie ide v simulacii; poistka pre pripad, ze by sa do zoznamu dostal
            bool pilotny = !DocasnyTestPacketaZauctovane && _pilotnePriecinky.Contains(folder.FullName);

            _logger.PrazdnyRiadok(1);
            _logger.Loguj($"Priečinok '{folder.FullName}' ({typSuboru}): {uids.Count} emailov.", true, FarbyLogu.Priecinok);
            if (!pilotny)
            {
                _logger.Loguj($"[PILOT] Priečinok je mimo pilotu - rozúčtovanie sa iba simuluje (bez importu do Omegy a bez presunov).", true);
            }
            _logger.PrazdnyRiadok(1);

            if (Lib.JeTypSuboruZApi(typSuboru))
            {
                // Packeta: dáta sa berú z jej API, nie z emailov - preto sa spracúva aj vtedy, keď
                // v priečinku žiadny email nie je. Emailom chodia LEN faktúry v EUR (jedna za týždeň);
                // faktúry v CZK/HUF/PLN/RON emailom nechodia vôbec, tie by sa z emailov nikdy
                // nerozúčtovali. Naopak výpis EUR faktúry obsahuje zásielky celého týždňa vo všetkých
                // menách, takže rozdelené po menách nesedeli so sumami faktúr tých mien.
                SpracujPacketaZApi(folder, typSuboru, uids.Count, pilotny);
                return;
            }

            if (!uids.Any())
            {
                return;
            }

            if (typSuboru == eTypSuboru.PlatobnaBrana_Gopay)
            {
                // prevod k výpisu N je až vo výpise N+1 - najprv sa z celého priečinka pozbierajú
                // prevody, aby bol pri spracovaní súboru N jeho prevod už známy
                NacitajPrevodyGopay(folder, uids);
            }

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
                    // simuluje sa, ked priecinok nie je v pilote ALEBO bezi testovaci rezim.
                    // Podmienka MUSI obsahovat !pilotny - bez nej by na serveri ostro importovali
                    // aj priecinky mimo pilotu (21.08. tak GLS SK realne importoval a GoPay by sa
                    // zauctoval napriek vypnutiu na ziadost zakaznika).
                    // DOCASNE (test v9): nazov priecinka je len do logov a nazvov log suborov -
                    // "Zaúčtované" by tam nic nepovedalo, posielame nadradenu "PACKETA"
                    if (SpracujEmail(message, typSuboru,
                        DocasnyTestPacketaZauctovane ? TestovaciNazovPriecinka : folder.Name,
                        ibaSimulacia: !pilotny || IbaSimulacia))
                    {
                        if (pilotny && !IbaSimulacia)
                        {
                            if (zauctovane == null)
                            {
                                zauctovane = DajAleboVytvorZauctovane(folder);
                            }

                            // presuvame hned po zauctovani - pad medzi zauctovanim a davkovym presunom
                            // by nechal zauctovane emaily navzdy v priecinku (pri retry vratia ZiadneUhrady)
                            folder.MoveTo(uid, zauctovane);
                            _logger.Loguj($"Email '{message.Subject}' presunutý do '{zauctovane.FullName}'.", true, FarbyLogu.Uspech);
                        }
                        else
                        {
                            _logger.Loguj($"[PILOT] Email '{message.Subject}' by bol presunutý do " +
                                $"'{folder.FullName}.{NazovPodpriecinkaZauctovane}'" +
                                (pilotny ? " (testovací režim - nepresúvam)." : " (priečinok mimo pilotu)."), true, FarbyLogu.Uspech);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri spracovaní emailu {uid} v priečinku '{folder.FullName}': {ex}", true, FarbyLogu.Chyba);
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

            if (!prilohy.Any())
            {
                _logger.Loguj($"Email '{message.Subject}': neobsahuje žiadnu prílohu s názvom súboru - preskakujem.", true);
                return false;
            }

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

                string filePath = UlozPrilohu(attachment, DajNazovPrilohy(attachment, message, typSuboru));
                if (filePath == null)
                {
                    // nepodporovaná prípona - treba to vidieť v logu, inak sa nedá zistiť, v akej podobe
                    // dopravca výpis posiela (napr. .zip alebo odkaz na stiahnutie namiesto prílohy)
                    _logger.Loguj($"Email '{message.Subject}': prílohu {attachment.FileName} neviem spracovať " +
                        $"(prípona {(string.IsNullOrEmpty(pripona) ? "žiadna" : pripona)}, podporované sú " +
                        $"{string.Join(", ", _povolenePripony)}).", true);
                    continue;
                }

                asponJedenSubor = true;
                if (!SpracujJedenSubor(filePath, message, typSuboru, nazovPriecinka, ibaSimulacia))
                {
                    vsetkoZauctovane = false;
                }
            }

            return asponJedenSubor && vsetkoZauctovane;
        }

        /// <summary>
        /// Pozbiera prevody "GOPAY-vyuctovani" zo všetkých CSV výpisov v priečinku. Prílohy sa kvôli tomu
        /// sťahujú dvakrát (raz sem, raz pri spracovaní) - pri desiatkach emailov je to prijateľná cena
        /// za to, že spracovanie súborov ostáva jednoduché a nezávislé od poradia emailov.
        /// </summary>
        private void NacitajPrevodyGopay(IMailFolder folder, IList<UniqueId> uids)
        {
            _prevodyGopay.Clear();
            foreach (UniqueId uid in uids)
            {
                try
                {
                    MimeMessage message = folder.GetMessage(uid);
                    foreach (MimePart priloha in message.BodyParts.OfType<MimePart>()
                        .Where(p => string.Equals(Path.GetExtension(p.FileName), ".csv", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (!GopayCsv.SkusRozobratNazov(priloha.FileName, out string ucet, out DateTime od, out DateTime _))
                        {
                            continue;
                        }

                        string subor = UlozPrilohu(priloha);
                        if (subor == null)
                        {
                            continue;
                        }
                        GopayCsv.Prevod prevod = GopayCsv.DajPrevod(subor);
                        File.Delete(subor);
                        if (prevod != null
                            && !_prevodyGopay.Any(p => p.ucet == ucet && p.obdobieOd == od))
                        {
                            _prevodyGopay.Add(new PrevodGopayVypisu { ucet = ucet, obdobieOd = od, prevod = prevod });
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri zbieraní prevodov GoPay z emailu {uid}: {ex}", true, FarbyLogu.Chyba);
                    _chybyBehu.Add($"Zbieranie prevodov GoPay, email {uid}: {ex}");
                }
            }
            _logger.Loguj($"GoPay: pozbieraných {_prevodyGopay.Count} prevodov z výpisov v priečinku " +
                $"(prevod k výpisu N je až vo výpise N+1).", true);
        }

        /// <summary>
        /// Prevod vyrovnávajúci platby zadaného výpisu - z výpisu, ktorého obdobie začína deň po konci
        /// tohto. Null, keď nasledujúci výpis ešte neprišiel (najnovší týždeň).
        /// </summary>
        private GopayCsv.Prevod DajPrevodZNasledujucehoVypisu(string fileName)
        {
            if (!GopayCsv.SkusRozobratNazov(fileName, out string ucet, out DateTime _, out DateTime obdobieDo))
            {
                return null;
            }
            return _prevodyGopay
                .FirstOrDefault(p => p.ucet == ucet && GopayCsv.JeNasledujuceObdobie(obdobieDo, p.obdobieOd))
                ?.prevod;
        }

        /// <summary>
        /// Packeta: výpisy sa stiahnu z jej API (zoznam faktúr + výpis zásielok k faktúre), rozdelia sa
        /// na vlastnú menu faktúry a rozúčtujú jeden po druhom. Faktúry, ktoré už rozúčtované sú, vráti
        /// jadro ako duplicitu - názvy súborov sú zhodné s tými, aké robí formulár, takže sa faktúra
        /// nerozúčtuje dvakrát ani keď ju medzitým niekto pustil ručne.
        ///
        /// Emaily sa nepresúvajú: dáta z nich nejdú, takže "Zaúčtované" by o nich nič nevypovedalo.
        /// </summary>
        private void SpracujPacketaZApi(IMailFolder folder, eTypSuboru typSuboru, int pocetEmailov, bool pilotny)
        {
            // simuluje sa, ked priecinok nie je v pilote ALEBO bezi testovaci rezim - rovnaka
            // podmienka ako pri emailovych prilohach
            bool ibaSimulacia = !pilotny || IbaSimulacia;
            string nazovPriecinka = DocasnyTestPacketaZauctovane ? TestovaciNazovPriecinka : folder.Name;

            List<string> subory = PacketaZApi.PripravSuboryNaRozuctovanie(
                _pripojeneFirmy._omegaPath, s => _logger.Loguj(s, true));

            if (!subory.Any())
            {
                _logger.Loguj($"Packeta: z API sa za posledných {PacketaZApi.DniDozadu} dní nenašlo " +
                    $"nič na rozúčtovanie.", true);
                return;
            }

            foreach (string subor in subory)
            {
                if (_importNedostupny)
                {
                    // zvysne subory sa spracuju pri dalsom behu - z API sa stiahnu znova
                    break;
                }

                try
                {
                    _logger.Loguj($"Packeta: spracúvam súbor {Path.GetFileName(subor)}.", true);
                    eVysledokRozuctovania vysledok = SpracujSubor(subor, typSuboru, nazovPriecinka, ibaSimulacia);
                    _logger.Loguj($"Súbor {Path.GetFileName(subor)}: {vysledok}.", true);
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri rozúčtovaní súboru {Path.GetFileName(subor)} z API Packety: {ex}",
                        true, FarbyLogu.Chyba);
                    _chybyBehu.Add($"Súbor {Path.GetFileName(subor)} z API Packety: {ex}");
                }
            }

            if (pocetEmailov > 0)
            {
                _logger.Loguj($"Packeta: {pocetEmailov} emailov v priečinku nepresúvam - dáta idú z API " +
                    $"a emaily sú len oznámenie.", true);
            }
        }

        /// <summary>
        /// Rozúčtuje jeden súbor z prílohy emailu.
        /// </summary>
        /// <returns>true, ak sa súbor počíta ako vybavený (email môže ísť do "Zaúčtované")</returns>
        private bool SpracujJedenSubor(string filePath, MimeMessage message, eTypSuboru typSuboru,
            string nazovPriecinka, bool ibaSimulacia)
        {
            eVysledokRozuctovania vysledok;
            string odtlacok = DajOdtlacokSuboru(filePath);
            if (_spracovanePrilohy.TryGetValue(odtlacok, out SpracovanaPriloha prva))
            {
                // ten istý report druhý raz - rozúčtovanie by vytvorilo tie isté doklady ešte raz;
                // email dostane rovnaký osud ako prvý, aby sa presunuli obidva
                vysledok = prva.vysledok;
                _logger.Loguj($"Email '{message.Subject}': súbor {Path.GetFileName(filePath)} je totožný so súborom " +
                    $"{prva.nazov}, ktorý už bol v tomto behu spracovaný - druhý raz ho nerozúčtovávam " +
                    $"(výsledok preberám: {vysledok}).", true);
                _suhrnneZoznamy.PridajDuplicitnyReport(nazovPriecinka, message.Date.LocalDateTime,
                    prva.nazov, Path.GetFileName(filePath), message.Subject);
                File.Delete(filePath);
            }
            else
            {
                _logger.Loguj($"Email '{message.Subject}': spracúvam súbor {Path.GetFileName(filePath)}.", true);
                vysledok = SpracujSubor(filePath, typSuboru, nazovPriecinka, ibaSimulacia);
                _logger.Loguj($"Súbor {Path.GetFileName(filePath)}: {vysledok}.", true);
                _spracovanePrilohy[odtlacok] = new SpracovanaPriloha
                {
                    nazov = Path.GetFileName(filePath),
                    vysledok = vysledok,
                };
            }

            // duplicita = subor uz bol zauctovany skor; vsetko uz uhradene = najdene faktury su uz
            // zaplatene (Emag) -> email v oboch pripadoch patri do "Zaúčtované"
            return vysledok == eVysledokRozuctovania.Rozuctovane || vysledok == eVysledokRozuctovania.Duplicita
                || vysledok == eVysledokRozuctovania.VsetkoUzUhradene;
        }

        private eVysledokRozuctovania SpracujSubor(string filePath, eTypSuboru typSuboru, string nazovPriecinka, bool ibaSimulacia)
        {
            short mesiac = NazovSuboru.DajMesiac(Path.GetFileNameWithoutExtension(filePath));
            if (Lib.NastavTypRozuctovania(typSuboru) == eTypRozuctovania.BezZapoctuBanky && mesiac == 0)
            {
                _logger.Loguj($"V názve súboru {Path.GetFileName(filePath)} sa nenašiel mesiac (MM.RRRR) - súbor preskakujem.", true, FarbyLogu.Chyba);
                return eVysledokRozuctovania.Chyba;
            }
            if (mesiac == 0)
            {
                // pre typy so zapoctom banky sa datum berie z bankoveho dokladu, mesiac je nepodstatny
                mesiac = (short)DateTime.Now.Month;
            }

            // GoPay: prevod vyrovnavajuci tento vypis je az v NASLEDUJUCOM tyzdennom vypise -
            // datum a suma sa dosadia z neho; ak este nepriisiel, jadro subor odlozi na dalsi beh
            GopayCsv.Prevod prevodGopay = typSuboru == eTypSuboru.PlatobnaBrana_Gopay
                ? DajPrevodZNasledujucehoVypisu(filePath)
                : null;
            if (prevodGopay != null)
            {
                _logger.Loguj($"Súbor {Path.GetFileName(filePath)}: prevod z nasledujúceho výpisu " +
                    $"{prevodGopay.suma:0.00} z {prevodGopay.datum:dd.MM.yyyy}.", true);
            }

            RozuctovanieContext ctx = new RozuctovanieContext
            {
                pripojeneFirmy = _pripojeneFirmy,
                typSuboru = typSuboru,
                fileName = filePath,
                mesiac = mesiac,
                datumVypisu = prevodGopay?.datum,
                sumaPrevodu = prevodGopay?.suma,
                interneCislo = null, // sluzba: doklad sa hlada podla textu hlavicky (C099) a datumu vypisu
                nazovPriecinka = nazovPriecinka,
                ibaSimulacia = ibaSimulacia, // pilot: mimo pilotnych priecinkov sa nic nezapisuje do Omegy
                pouziteBankoveDoklady = _pouziteBankoveDoklady, // jeden bankovy vypis = jedno rozuctovanie
                logger = _logger, // riadky rozúčtovania do logu služby, jadro pri nich určuje farby
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

        /// <summary>Kontrolný súčet obsahu prílohy - dva emaily s tým istým reportom dajú rovnaký odtlačok.</summary>
        private static string DajOdtlacokSuboru(string filePath)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                return BitConverter.ToString(sha.ComputeHash(stream));
            }
        }

        private string UlozPrilohu(MimePart attachment, string nazovSuboru = null)
        {
            string fileName = nazovSuboru ?? attachment.FileName;
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

        /// <summary>
        /// Názov, pod ktorým sa príloha uloží a rozúčtuje. DPD SK posiela KAŽDÝ report s tým istým
        /// názvom (sales_company_..._LWD/LWW) - a kľúč "už zaúčtované" je názov súboru, takže po prvom
        /// zaúčtovaní sa každý ďalší report hlásil ako duplicita a email sa presunul BEZ zaúčtovania
        /// (26.08.2026). Dátum emailu robí názov jedinečným a je stabilný naprieč behmi, takže detekcia
        /// skutočných duplicít (ten istý email pri ďalšom behu) funguje ďalej. Ostatným dopravcom sa
        /// názov nemení - nesú dátum či číslo faktúry v názve a napr. GoPay/DPD RO si z názvu čítajú údaje.
        /// </summary>
        private static string DajNazovPrilohy(MimePart attachment, MimeMessage message, eTypSuboru typSuboru)
        {
            if (typSuboru != eTypSuboru.Dopravca_DPD_Slovensko)
            {
                return attachment.FileName;
            }
            return Path.GetFileNameWithoutExtension(attachment.FileName)
                + "_" + message.Date.LocalDateTime.ToString("yyyy-MM-dd")
                + Path.GetExtension(attachment.FileName);
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
