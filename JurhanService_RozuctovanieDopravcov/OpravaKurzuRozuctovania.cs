using DataProvider;
using JurhanLib.Logger;
using OmegaLib.Enums;
using OmegaLib.Models;
using OmegaLib.Repository;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace JurhanService_RozuctovanieDopravcov
{
    /// <summary>
    /// JEDNORAZOVÁ oprava. Dohľadá bankové výpisy (okruh BV/zBV) a k nim spárované interné doklady
    /// rozúčtovania (okruh ID/zID), ktoré majú:
    ///   - rovnakú cudzomennú sumu (C210 SumaSpoluZahranicnaMena) a
    ///   - rovnaký účet peniaze-na-ceste (syntetika 261): banka na strane DAL, interný doklad na strane MáDať.
    /// Na spárovanom internom doklade opraví kurz (C202 KurzNBS, C203 KurzBanka) a množstvo jednotky (C201)
    /// na hodnoty z bankového výpisu (rovnaká chyba ako v UhradyToTxt – kurz sa bral z lístka, nie z výpisu)
    /// a dotknutý doklad označí (C002_Oznaceny = -1).
    ///
    /// V režime dryRun=true sa NIČ nezapisuje, len sa zaloguje, čo by sa zmenilo. Po použití môže byť trieda odstránená.
    /// </summary>
    internal class OpravaKurzuRozuctovania
    {
        private const string SyntetikaPeniazeNaCeste = "261";

        private readonly PripojeneFirmy _pripojeneFirmy;
        private readonly RozuctovanieLogger _logger;
        private readonly bool _dryRun;

        internal OpravaKurzuRozuctovania(PripojeneFirmy pripojeneFirmy, RozuctovanieLogger logger, bool dryRun)
        {
            _pripojeneFirmy = pripojeneFirmy;
            _logger = logger;
            _dryRun = dryRun;
        }

        internal void Execute()
        {
            IDataProvider dp = _pripojeneFirmy.dataProvider;
            EudHlavickaRepository hlavickaRepo = new EudHlavickaRepository(dp);
            EudPolozkaRepository polozkaRepo = new EudPolozkaRepository(dp);

            _logger.Loguj($"Oprava kurzu rozúčtovania — {(_dryRun ? "DRY-RUN (bez zápisu do DB)" : "OSTRÝ ZÁPIS do DB")}.", true, FarbyLogu.Hlavicka);

            // položky s účtom 261: banka na strane DAL, interný doklad na strane MáDať
            List<EUDPolozka> polozkyVypis = polozkaRepo.GetItems("C108_DALSyntetickyUcet = @1", SyntetikaPeniazeNaCeste).ToList();
            List<EUDPolozka> polozkyInterny = polozkaRepo.GetItems("C106_MDSyntetickyUcet = @1", SyntetikaPeniazeNaCeste).ToList();
            _logger.Loguj($"Diagnostika: položiek s účtom {SyntetikaPeniazeNaCeste} na DAL = {polozkyVypis.Count}, na MáDať = {polozkyInterny.Count}.", true, FarbyLogu.Hlavicka);

            // všetky hlavičky s účtom 261 (bez filtra okruhu) — diagnostika, aké okruhy tie doklady majú
            Dictionary<int, EUDHlavicka> vsetkyVypis = NacitajHlavicky(hlavickaRepo, polozkyVypis.Select(p => p.IdEUD));
            Dictionary<int, EUDHlavicka> vsetkyInterny = NacitajHlavicky(hlavickaRepo, polozkyInterny.Select(p => p.IdEUD));
            _logger.Loguj($"Diagnostika: okruhy dokladov s 261 na DAL: [{OkruhyHistogram(vsetkyVypis.Values)}]; " +
                $"na MáDať: [{OkruhyHistogram(vsetkyInterny.Values)}].", true, FarbyLogu.Hlavicka);

            // hlavičky obmedzené na správne okruhy (výpisy BV/zBV, interné ID/zID)
            Dictionary<int, EUDHlavicka> vypisy = FiltrujOkruh(vsetkyVypis, (short)eTYP_OKRUH.BV, (short)eTYP_OKRUH.zBV);
            Dictionary<int, EUDHlavicka> interne = FiltrujOkruh(vsetkyInterny, (short)eTYP_OKRUH.ID, (short)eTYP_OKRUH.zID);
            _logger.Loguj($"Diagnostika: bankových výpisov (BV/zBV) = {vypisy.Count}, interných dokladov (ID/zID) = {interne.Count}.", true, FarbyLogu.Hlavicka);

            // analytika účtu 261 pre každý doklad (z príslušnej položky)
            Dictionary<int, string> vypisAnalytika = polozkyVypis
                .GroupBy(p => p.IdEUD).ToDictionary(g => g.Key, g => g.First().DalAnalytickyUcet);
            Dictionary<int, string> internyAnalytika = polozkyInterny
                .GroupBy(p => p.IdEUD).ToDictionary(g => g.Key, g => g.First().MDAnalytickyUcet);

            // výpisy podľa kľúča (analytika + cudzomenná suma); kolízie = nejednoznačné, preskočíme
            Dictionary<string, EUDHlavicka> vypisPodlaKluca = new Dictionary<string, EUDHlavicka>();
            HashSet<string> nejednoznacne = new HashSet<string>();
            foreach (EUDHlavicka v in vypisy.Values)
            {
                if (v.SumaSpoluZahranicnaMena == 0) continue; // len cudzomenné (EUR výpisy kurz nemajú)
                if (!vypisAnalytika.TryGetValue(v.Id, out string anal)) continue;
                string kluc = Kluc(anal, v.SumaSpoluZahranicnaMena);
                if (vypisPodlaKluca.ContainsKey(kluc)) { nejednoznacne.Add(kluc); continue; }
                vypisPodlaKluca[kluc] = v;
            }

            int opravene = 0, uzOk = 0, bezVypisu = 0, nejedn = 0;
            foreach (EUDHlavicka i in interne.Values)
            {
                if (i.SumaSpoluZahranicnaMena == 0) continue;
                if (!internyAnalytika.TryGetValue(i.Id, out string anal)) continue;
                string kluc = Kluc(anal, i.SumaSpoluZahranicnaMena);

                if (nejednoznacne.Contains(kluc))
                {
                    _logger.Loguj($"Oprava kurzu: interný doklad '{i.CisloInterne}' (id {i.Id}) — účet {SyntetikaPeniazeNaCeste}{anal}, " +
                        $"suma {i.SumaSpoluZahranicnaMena:0.00}: viac bankových výpisov s rovnakým kľúčom, preskakujem (nejednoznačné).", true, FarbyLogu.Hlavicka);
                    nejedn++;
                    continue;
                }
                if (!vypisPodlaKluca.TryGetValue(kluc, out EUDHlavicka v)) { bezVypisu++; continue; }

                bool trebaOpravit = i.KurzNBS != v.KurzNBS || i.KurzBanka != v.KurzBanka || i.MnozstvoJednotky != v.MnozstvoJednotky;
                if (!trebaOpravit) { uzOk++; continue; }

                _logger.Loguj($"{(_dryRun ? "[DRY-RUN] " : "")}Interný doklad '{i.CisloInterne}' (id {i.Id}) ↔ výpis '{v.CisloInterne}' (id {v.Id}), " +
                    $"účet {SyntetikaPeniazeNaCeste}{anal}, suma {v.SumaSpoluZahranicnaMena:0.00}: " +
                    $"KurzNBS {i.KurzNBS} → {v.KurzNBS}, KurzBanka {i.KurzBanka} → {v.KurzBanka}, " +
                    $"množstvo {i.MnozstvoJednotky} → {v.MnozstvoJednotky}; doklad sa označí (C002_Oznaceny = -1).", true);

                if (!_dryRun)
                {
                    dp.ExecuteCommand(
                        "UPDATE T040_EUD SET C202_KurzNBS = @1, C203_KurzBanka = @2, C201_MnozstvoJednotky = @3, " +
                        "C002_Oznaceny = -1 WHERE C000_ID = @4",
                        v.KurzNBS, v.KurzBanka, v.MnozstvoJednotky, i.Id);
                }
                opravene++;
            }

            _logger.Loguj($"Oprava kurzu — na opravu: {opravene}, už OK: {uzOk}, bez zhody výpisu: {bezVypisu}, nejednoznačných: {nejedn}. " +
                $"{(_dryRun ? "Nič sa nezapísalo (dry-run)." : "Zmeny zapísané do DB.")}", true, FarbyLogu.Hlavicka);

            _logger.ZapisDataDoLogSuboru(); // zapíš .log súbor (utilita beží mimo bežného behu služby)
        }

        private static string Kluc(string analytika, decimal suma) =>
            $"{analytika}|{System.Math.Round(suma, 2).ToString(CultureInfo.InvariantCulture)}";

        private Dictionary<int, EUDHlavicka> NacitajHlavicky(EudHlavickaRepository repo, IEnumerable<int> idcka)
        {
            List<int> ids = idcka.Distinct().ToList();
            if (ids.Count == 0)
            {
                return new Dictionary<int, EUDHlavicka>();
            }
            return repo.DajDoklady($"C000_ID IN ({string.Join(",", ids)})").ToDictionary(h => h.Id);
        }

        private static Dictionary<int, EUDHlavicka> FiltrujOkruh(Dictionary<int, EUDHlavicka> hlavicky, short okruh1, short okruh2)
        {
            return hlavicky.Values
                .Where(h => h.CisloInterneTypOkruh == okruh1 || h.CisloInterneTypOkruh == okruh2)
                .ToDictionary(h => h.Id);
        }

        private static string OkruhyHistogram(IEnumerable<EUDHlavicka> hlavicky)
        {
            return string.Join(", ", hlavicky
                .GroupBy(h => h.CisloInterneTypOkruh)
                .OrderBy(g => g.Key)
                .Select(g => $"{g.Key}×{g.Count()}"));
        }
    }
}
