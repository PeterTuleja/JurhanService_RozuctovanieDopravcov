# RozuctovanieDopravcov služba — implementačný plán

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Windows služba (net10 + net48), ktorá cez IMAP prejde schránku platby@jurhan.com, rozúčtuje súbory dopravcov z príloh a emaily presunie do podpriečinka „Zaúčtované“.

**Architecture:** Parsery a orchestrácia sa presunú z UI programu RozuctovanieDopravcov do JurhanLib (`JurhanLib.Import.Dopravcovia`) bez závislosti na statickej triede `Program` (nahradí ju `RozuctovanieContext`). Program aj nová služba volajú spoločné jadro `RozuctovanieCore`. Služba má štandardnú kostru `ServiceRunner → ServiceJurhan → IExecutable → SpustenieServisyBase`.

**Tech Stack:** C# (net10.0-windows SDK-style / net48 classic), MailKit/MimeKit 4.13, Kros.KORM cez `IDataProvider`, DevExpress.Spreadsheet, Omega TXT autoimport.

**Spec:** `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\docs\specs\2026-07-14-rozuctovanie-dopravcov-service-design.md`

**Poznámka k testom:** v celom ekosystéme (29 solutionov) neexistuje žiadny testovací projekt ani framework. Plán preto namiesto unit testov používa: (a) build po každom tasku, (b) regresný princíp „program sa správa rovnako“ (zachované presné texty chýb a poradie krokov), (c) manuálny smoke beh v režime Program. Nezavádzaj testovací framework.

**Poznámka ku commitom:** commituj v repozitári, do ktorého súbor patrí (každý projekt má vlastný `.git`). Pred commitom over `git -C <adresár> rev-parse --is-inside-work-tree`; ak adresár nie je git repo, krok commitu preskoč (plati pre staré repá, ak nie sú pod gitom). Pre **nové** service adresáre git repo inicializujeme (Task 12/16).

**Kľúčové existujúce súbory (na čítanie/vzory):**

| Účel | Cesta |
|---|---|
| Pôvodná orchestrácia (nová vetva) | `C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\Rozuctovanie.cs` |
| Pôvodné parsery (nová vetva) | tamtiež: `DopravcaToUhrady.cs`, `BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`, `EmagParser.cs` |
| Pôvodný program – statiky | tamtiež: `Program.cs` |
| Staré originály parserov | `C:\Users\tuleja\source\repos\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\*.cs` |
| Vzor služby net10 | `C:\Projekty\Private\JurhanService\ImportDobropisov\JurhanService_ImportDobropisov\*` |
| Vzor služby net48 | `C:\Users\tuleja\source\repos\JurhanService\ImportFaktur\JurhanService_ImportFaktur\*` |
| Vzor IMAP prístupu | `C:\Users\tuleja\source\repos\JurhanOdloz\ArchivaciaEmailov\ArchivaciaEmailov\ArchivaciaEmailov.cs` |
| Repository s hľadaním podľa poznámky | `C:\Projekty\Private\OmegaLib\OmegaLib\Repository\EudHlavickaRepository.cs:34` (`DajDokladPodlaPoznamky`) — existuje aj v starom IndiLib (`C:\Projekty\DevOps\Indivindi\IndiLib\IndiLib\Repository\EudHlavickaRepository.cs:36`) |

**Príkaz na build (net10):** `dotnet build <csproj> -v:m`
**Príkaz na build (net48):**
```powershell
$msbuild = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
& $msbuild <csproj> /p:Configuration=Debug /v:m
```

---

## ČASŤ I — nová (net10) strana

### Task 1: RozuctovanieContext + eVysledokRozuctovania (nový JurhanLib)

**Files:**
- Create: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\Dopravcovia\RozuctovanieContext.cs`

- [ ] **Step 1.1: Vytvor súbor s týmto obsahom**

```csharp
using JurhanModels.Enum;
using JurhanModels.Import;
using OmegaLib.Enums;
using OmegaLib.Models;
using OmegaLib.Repository;

namespace JurhanLib.Import.Dopravcovia
{
    public enum eVysledokRozuctovania
    {
        Rozuctovane,
        Duplicita,
        NenajdenyBankovyDoklad,
        ZiadneUhrady,
        Chyba
    }

    /// <summary>
    /// Zdieľaný stav rozúčtovania — náhrada statickej triedy Program z pôvodného UI programu,
    /// aby jadro vedelo bežať aj zo služby.
    /// </summary>
    public class RozuctovanieContext
    {
        public PripojeneFirmy pripojeneFirmy;
        public eTypRozuctovania typRozuctovania;
        public eTypSuboru typSuboru;
        public short mesiac;
        /// <summary>Interné číslo bankového dokladu zadané v UI; v službe null — doklad sa hľadá podľa poznámky.</summary>
        public string interneCislo;
        public EUDHlavicka dokladPreRozuctovanie;
        public string fileName;
        public eCurrency currency;
        public eCountry country;
        public eZobrazenieChyby zobrazenieChyby = eZobrazenieChyby.ZapisDoSuboru;
        public eTypSpustenia typSpustenia = eTypSpustenia.Servica;
    }
}
```

Pozn.: over, v ktorom namespace je `eCountry` — v `AllegroParser.cs` sa používa s usingom `JurhanModels.Enum`. Ak by kompilácia hlásila chybu, uprav using podľa skutočného namespace (rovnako `eZobrazenieChyby` je v `OmegaLib.Enums` — over podľa `Rozuctovanie.cs`, ktorý má using `OmegaLib.Services`; nechaj presne tie usingy, s ktorými sa build podarí).

- [ ] **Step 1.2: Build**

Run: `dotnet build C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\JurhanLib.csproj -v:m`
Expected: Build succeeded (SDK projekt zoberie nový .cs automaticky).

- [ ] **Step 1.3: Commit v repe JurhanLib**

```powershell
git -C C:\Projekty\Private\JurhanProgramy\JurhanLib add JurhanLib/Import/Dopravcovia/RozuctovanieContext.cs
git -C C:\Projekty\Private\JurhanProgramy\JurhanLib commit -m "feat: RozuctovanieContext a eVysledokRozuctovania pre zdielane jadro rozuctovania"
```

---

### Task 2: DopravcaToUhrady → JurhanLib (nový)

**Files:**
- Create: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\Dopravcovia\DopravcaToUhrady.cs`
- Zdroj: `C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\DopravcaToUhrady.cs` (v programe sa zmaže v Tasku 6)

- [ ] **Step 2.1: Vytvor súbor.** Obsah = pôvodný súbor s týmito zmenami: namespace `JurhanLib.Import.Dopravcovia`, `internal` → `public`, konštruktor berie `RozuctovanieContext`, všetky `Program.*` nahradené `_ctx.*`, chybová hláška používa `_ctx.zobrazenieChyby` (v službe nesmie vyskočiť MessageBox!). Kompletný výsledok:

```csharp
using OmegaLib.Repository;
using OmegaLib.Services;
using JurhanModels.Enum;
using JurhanModels.Import;
using OmegaLib.Enums;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JurhanLib.Import.Dopravcovia
{
    public class DopravcaToUhrady : CsvBase
    {
        readonly FakturaHlavickaRepository _fakturaHlavickaRepository;
        private readonly RozuctovanieContext _ctx;

        public DopravcaToUhrady(RozuctovanieContext ctx) : base(ctx.fileName)
        {
            _ctx = ctx;
            _fakturaHlavickaRepository = new FakturaHlavickaRepository(ctx.pripojeneFirmy.dataProvider);
        }

        public IEnumerable<Uhrada> NacitajUhrady(short indexVS, short indexSumaCM1, short indexSumaCM2, short indexMena,
            IEnumerable<string> nazvyPrvychStlpcov, System.Text.Encoding encoding, bool udajeSuVUvodzovkach)
        {
            return NacitajUhrady(indexVS, indexSumaCM1, indexSumaCM2, indexMena, nazvyPrvychStlpcov,
                encoding, udajeSuVUvodzovkach, ";");
        }

        public IEnumerable<Uhrada> NacitajUhrady(short indexVS, short indexSumaCM1, short indexSumaCM2, short indexMena,
            IEnumerable<string> nazvyPrvychStlpcov, System.Text.Encoding encoding, bool udajeSuVUvodzovkach,
            string oddelovac)
        {
            bool nasledujuData = false;
            List<Uhrada> uhrady = new List<Uhrada>();
            Nacitaj(oddelovac, encoding, udajeSuVUvodzovkach);

            if (!KontrolaNaSpravnyTypSuboru(records, nazvyPrvychStlpcov))
            {
                return null;
            }
            else
            {
                var maxIndex = Math.Max(indexVS, Math.Max(indexSumaCM1, Math.Max(indexSumaCM2, indexMena)));
                foreach (var record in records)
                {
                    if (nasledujuData)
                    {
                        if (record.Length > maxIndex)
                        {
                            decimal SumaCM1 = indexSumaCM1 != -1 ? ServicesNumeric.ParseDecimal(record[indexSumaCM1]) : 0;
                            decimal SumaCM2 = indexSumaCM2 != -1 ? ServicesNumeric.ParseDecimal(record[indexSumaCM2]) : 0;
                            Uhrada uhrada = new Uhrada
                            {
                                VS = record[indexVS],
                                SumaCM = SumaCM1 + SumaCM2,
                                Mena = indexMena != -1 ? record[indexMena] : eCurrency.EUR.ToString(),
                                MnozstvoJednotky = 1,
                                Kurz = 1,
                            };
                            if (uhrada.SumaCM != 0 && !string.IsNullOrEmpty(uhrada.VS))
                            {
                                switch (_ctx.typSuboru)
                                {
                                    case eTypSuboru.AmazonATDE:
                                    case eTypSuboru.AmazonES:
                                    case eTypSuboru.AmazonFR:
                                    case eTypSuboru.AmazonIT:
                                        {
                                            uhrada.VS = DajVSPreAmazon(uhrada.VS);
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Cesko:
                                        {
                                            uhrada.Mena = eCurrency.CZK.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Polsko:
                                        {
                                            uhrada.Mena = eCurrency.PLN.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Chorvatsko:
                                        {
                                            uhrada.Mena = eCurrency.EUR.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Madarsko:
                                        {
                                            uhrada.Mena = eCurrency.HUF.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Rumunsko:
                                        {
                                            uhrada.Mena = eCurrency.RON.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_DPD_Slovinsko:
                                        {
                                            uhrada.Mena = eCurrency.EUR.ToString();
                                            break;
                                        }
                                    case eTypSuboru.Dopravca_GLS_PLN:
                                        {
                                            uhrada.SumaCM = -uhrada.SumaCM;
                                            uhrada.Mena = eCurrency.PLN.ToString();
                                            break;
                                        }
                                    case eTypSuboru.PlatobnaBrana_ComgatePLN:
                                        {
                                            uhrada.VS = uhrada.VS.Replace("Order ", string.Empty);
                                            uhrada.Mena = eCurrency.PLN.ToString();
                                            break;
                                        }
                                    case eTypSuboru.PlatobnaBrana_Gopay:
                                        {
                                            if (uhrada.SumaCM < 0)
                                            {
                                                uhrada.VS = $"pri VS {uhrada.VS} je záporna suma {uhrada.SumaCM}";
                                            }
                                            break;
                                        }
                                }
                                if (!string.IsNullOrEmpty(uhrada.VS))
                                {
                                    uhrady.Add(uhrada);
                                }
                            }
                        }
                    }
                    if (record.Length > 0 && nazvyPrvychStlpcov.Contains(record[0]))
                    {
                        nasledujuData = true;
                    }
                }

                if (uhrady.Any())
                {
                    _ctx.currency = Lib.DajCurrency(uhrady.First().Mena);
                    _ctx.country = Lib.DajCountry(_ctx.typSuboru, uhrady.First().Mena);
                }
                return uhrady;
            }
        }

        private bool KontrolaNaSpravnyTypSuboru(IEnumerable<string[]> records, IEnumerable<string> nazvyPrvychStlpcov)
        {
            if (records.Any(f => f.Length > 0 && nazvyPrvychStlpcov.Contains(f[0])))
            {
                return true;
            }

            ServicesError.ErrorEnd($"Vybrali ste nesprávny typ súboru. " +
                $"Prvý stlpec v importnom csv súbore sa musí volať : " +
                $"{Environment.NewLine}{Environment.NewLine}{string.Join(",", nazvyPrvychStlpcov)} ",
                _ctx.zobrazenieChyby);
            return false;
        }

        private string DajVSPreAmazon(string id)
        {
            var faktura = _fakturaHlavickaRepository.DajDoklad("C100_TypDokladu = @1 AND C097_Poznamka LIKE @2",
                eTypFaktury.OdoslanaFaktura, "%" + id);
            if (faktura != null)
            {
                return faktura.Objednavka ?? string.Empty;
            }
            else
            {
                ServicesError.ErrorEnd($"Nenašla sa žiadna objednávka pre VS: {id}. " +
                    $"Skontrolujte, či je správne zadaný VS v súbore.", eZobrazenieChyby.ZapisDoSuboru);
                return null;
            }
        }
    }
}
```

Pozn.: pôvodný `enum eTypDatum` z hlavičky súboru NEPRENÁŠAJ (nikde sa nepoužíva — over grepom `eTypDatum` v oboch programoch; ak sa predsa používa, prenes ho tiež).
Pozn. 2: `CsvBase` je v `OmegaLib.Services` (súbor `C:\Projekty\Private\OmegaLib\OmegaLib\Services\CsvBase.cs`); ak má `internal` členy nedostupné z JurhanLib, over jej prístupnosť — je `public`, používa ju aj pôvodný program z iného assembly.

- [ ] **Step 2.2: Build** — rovnaký príkaz ako Step 1.2. Expected: Build succeeded.

- [ ] **Step 2.3: Commit** — `git -C C:\Projekty\Private\JurhanProgramy\JurhanLib add -A && git -C ... commit -m "feat: presun DopravcaToUhrady do JurhanLib.Import.Dopravcovia (ctx namiesto Program statics)"`

---

### Task 3: BaseParser + AllegroParser + KauflandParser → JurhanLib (nový)

**Files:**
- Create: `...\JurhanLib\JurhanLib\Import\Dopravcovia\BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`
- Zdroj: rovnomenné súbory v novom programe RozuctovanieDopravcov

- [ ] **Step 3.1: BaseParser.cs**

```csharp
using OmegaLib.Services;
using JurhanModels.Import;
using OmegaLib.Enums;
using System.Collections.Generic;
using System.Linq;

namespace JurhanLib.Import.Dopravcovia
{
    public abstract class BaseParser
    {
        protected readonly RozuctovanieContext _ctx;

        protected BaseParser(RozuctovanieContext ctx)
        {
            _ctx = ctx;
        }

        internal abstract string NatiahniKodCRACountry();

        public IEnumerable<Uhrada> NacitajUhrady()
        {
            string sql = "SELECT F.C023_KodCiselnaRada as KodCiselnaRada, F.C030_CisloFaktury as VS, F.C050_PartnerStat as PartnerStat, " +
                "F.C114_FormaUhrady as FormaUhrady, F.C115_SposobDopravy as SposobDopravy, F.C200_Mena as Mena," +
                "E.C211_SumaSpolu as SumaTM, E.C210_SumaSpoluZahranicnaMena as SumaCM, E.C220_Uhradene as Uhradene, " +
                "E.C201_MnozstvoJednotky as MnozstvoJednotky, E.C202_KurzNBS as Kurz " +
                "FROM T228_Faktury as F INNER JOIN T040_EUD as E ON F.C030_CisloFaktury = E.C030_CisloInterne " +
                $"WHERE F.C100_TypDokladu = {(short)eTypFaktury.OdoslanaFaktura} AND " +
                $"E.C003_Zamknuty = 0 AND  (E.C220_Uhradene IS NULL OR E.C220_Uhradene = 0) AND NOT EXISTS " +
                $"(SELECT 1 FROM T228_Faktury as D " +
                $"WHERE D.C100_TypDokladu = {(short)eTypFaktury.Dobropis} AND D.C030_CisloFaktury = F.C104_DodaciList) " +
                $"AND C061_MesVyst = {_ctx.mesiac} AND " +
                $"C023_KodCiselnaRada = {ServicesString.SqlString(NatiahniKodCRACountry(), _ctx.pripojeneFirmy.DbJeSQL())} " +
                $"ORDER BY F.C062_RokVyst, F.C061_MesVyst, F.C060_DenVyst";

            return _ctx.pripojeneFirmy.dataProvider.Query<Uhrada>()
                .Sql(sql)
                .ToList()
                .Where(f => !JurhanModels.Kurieri.Lib.JeToDobierka(f.FormaUhrady))
                .ToList();
        }
    }
}
```

- [ ] **Step 3.2: AllegroParser.cs** — pôvodný obsah, `Program.typSuboru` → `_ctx.typSuboru`, `Program.country` → `_ctx.country`, konštruktor:

```csharp
using JurhanModels.Enum;
using JurhanModels.Import;

namespace JurhanLib.Import.Dopravcovia
{
    public class AllegroParser : BaseParser
    {
        public AllegroParser(RozuctovanieContext ctx) : base(ctx)
        {
        }

        internal override string NatiahniKodCRACountry()
        {
            string kodCR;
            switch (_ctx.typSuboru)
            {
                case eTypSuboru.Allegro_CZ:
                    kodCR = Constants.KodCRzOFALCZ;
                    _ctx.country = eCountry.CZ;
                    break;
                case eTypSuboru.Allegro_HU:
                    kodCR = Constants.KodCRzOFALHU;
                    _ctx.country = eCountry.HU;
                    break;
                case eTypSuboru.Allegro_PL:
                    kodCR = Constants.KodCRzOFALPL;
                    _ctx.country = eCountry.PL;
                    break;
                default:
                    kodCR = Constants.KodCROFALSK;
                    _ctx.country = eCountry.SK;
                    break;
            }
            return kodCR;
        }
    }
}
```

- [ ] **Step 3.3: KauflandParser.cs** — rovnaká mechanika (ctor + `_ctx.typSuboru`/`_ctx.currency`/`_ctx.country`); vetvy prenes 1:1 z pôvodného súboru (`Kaufland_EUR_AT→KodCRzOFKLAT/EUR/AT`, `Kaufland_CZK_CZ→KodCRzOFKLCZ/CZK/CZ`, `Kaufland_EUR_DE→KodCRzOFKLDE/EUR/DE`, `Kaufland_EUR_FR→KodCRzOFKLFR/EUR/FR`, `Kaufland_EUR_IT→KodCRzOFKLIT/EUR/IT`, `Kaufland_PLN_PL→KodCRzOFKLPL/PLN/PL`, default→`KodCROFKLSK/EUR/SK`):

```csharp
using JurhanModels.Enum;
using JurhanModels.Import;

namespace JurhanLib.Import.Dopravcovia
{
    public class KauflandParser : BaseParser
    {
        public KauflandParser(RozuctovanieContext ctx) : base(ctx)
        {
        }

        internal override string NatiahniKodCRACountry()
        {
            string kodCR;
            switch (_ctx.typSuboru)
            {
                case eTypSuboru.Kaufland_EUR_AT:
                    kodCR = Constants.KodCRzOFKLAT;
                    _ctx.currency = eCurrency.EUR;
                    _ctx.country = eCountry.AT;
                    break;
                case eTypSuboru.Kaufland_CZK_CZ:
                    kodCR = Constants.KodCRzOFKLCZ;
                    _ctx.currency = eCurrency.CZK;
                    _ctx.country = eCountry.CZ;
                    break;
                case eTypSuboru.Kaufland_EUR_DE:
                    kodCR = Constants.KodCRzOFKLDE;
                    _ctx.currency = eCurrency.EUR;
                    _ctx.country = eCountry.DE;
                    break;
                case eTypSuboru.Kaufland_EUR_FR:
                    kodCR = Constants.KodCRzOFKLFR;
                    _ctx.currency = eCurrency.EUR;
                    _ctx.country = eCountry.FR;
                    break;
                case eTypSuboru.Kaufland_EUR_IT:
                    kodCR = Constants.KodCRzOFKLIT;
                    _ctx.currency = eCurrency.EUR;
                    _ctx.country = eCountry.IT;
                    break;
                case eTypSuboru.Kaufland_PLN_PL:
                    kodCR = Constants.KodCRzOFKLPL;
                    _ctx.currency = eCurrency.PLN;
                    _ctx.country = eCountry.PL;
                    break;
                default:
                    kodCR = Constants.KodCROFKLSK;
                    _ctx.currency = eCurrency.EUR;
                    _ctx.country = eCountry.SK;
                    break;
            }
            return kodCR;
        }
    }
}
```

- [ ] **Step 3.4: Build + commit** (ako v Tasku 2). Commit message: `feat: presun BaseParser/AllegroParser/KauflandParser do JurhanLib`.

---

### Task 4: EmagParser → JurhanLib (nový)

**Files:**
- Create: `...\JurhanLib\JurhanLib\Import\Dopravcovia\EmagParser.cs`

- [ ] **Step 4.1: Vytvor súbor** — pôvodný obsah s ctx (namespace, public, ctor, `Program.*`→`_ctx.*`; chybová hláška cez `_ctx.zobrazenieChyby`):

```csharp
using DevExpress.Spreadsheet;
using OmegaLib.Services;
using JurhanModels.Enum;
using JurhanModels.Import;
using OmegaLib.Enums;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace JurhanLib.Import.Dopravcovia
{
    public class EmagParser
    {
        private readonly RozuctovanieContext _ctx;

        public EmagParser(RozuctovanieContext ctx)
        {
            _ctx = ctx;
        }

        private class EmagRiadok
        {
            public string VS { get; set; }
            public decimal SumaCM { get; set; }
            public decimal SumaTM { get; set; }
            public string Mena { get; set; }
            public decimal Kurz { get; set; }
            public decimal MnozstvoJednotky { get; set; }
            public string Poznamka { get; set; }
        }

        public IEnumerable<Uhrada> NacitajUhrady()
        {
            string zdroj = _ctx.typSuboru == eTypSuboru.Emag_HUF_HU ? "Emag HU" : "Emag RO";

            HashSet<string> referenceIds = NacitajReferenceIdZExcelu(_ctx.fileName);
            if (referenceIds == null)
            {
                return null; // chyba uz bola oznamena
            }

            var regex = new Regex(Regex.Escape(zdroj) + @"\s+(\d+)");
            var uhrady = new List<Uhrada>();

            foreach (var r in NacitajFakturyZOmegy(zdroj))
            {
                if (string.IsNullOrEmpty(r.VS) || string.IsNullOrEmpty(r.Poznamka))
                {
                    continue;
                }
                Match m = regex.Match(r.Poznamka);
                if (!m.Success || !referenceIds.Contains(m.Groups[1].Value))
                {
                    continue;
                }
                uhrady.Add(new Uhrada
                {
                    VS = r.VS,
                    SumaCM = r.SumaCM,
                    SumaTM = r.SumaTM,
                    Mena = r.Mena,
                    Kurz = r.Kurz,
                    MnozstvoJednotky = r.MnozstvoJednotky
                });
            }

            if (uhrady.Any())
            {
                _ctx.currency = Lib.DajCurrency(uhrady.First().Mena);
                _ctx.country = Lib.DajCountry(_ctx.typSuboru, uhrady.First().Mena);
            }
            return uhrady;
        }

        private HashSet<string> NacitajReferenceIdZExcelu(string fileName)
        {
            var referenceIds = new HashSet<string>();
            Workbook workbook = new Workbook();
            workbook.LoadDocument(fileName);
            Worksheet sheet = workbook.Worksheets[0];
            CellRange used = sheet.GetUsedRange();

            int refCol = -1;
            for (int c = used.LeftColumnIndex; c <= used.RightColumnIndex; c++)
            {
                if (sheet.Cells[used.TopRowIndex, c].Value.TextValue == "Reference ID")
                {
                    refCol = c;
                    break;
                }
            }
            if (refCol == -1)
            {
                ServicesError.ErrorEnd(
                    "V súbore Emag sa nenašiel stĺpec 'Reference ID'. Skontrolujte, či ste vybrali správny súbor.",
                    _ctx.zobrazenieChyby);
                return null;
            }

            for (int r = used.TopRowIndex + 1; r <= used.BottomRowIndex; r++)
            {
                CellValue value = sheet.Cells[r, refCol].Value;
                if (value.IsNumeric)
                {
                    referenceIds.Add(((long)value.NumericValue).ToString());
                }
            }
            return referenceIds;
        }

        private IEnumerable<EmagRiadok> NacitajFakturyZOmegy(string zdroj)
        {
            string like = "%" + zdroj + " %";
            string sql =
                "SELECT E.C031_CisloExterne AS VS, " +
                "E.C210_SumaSpoluZahranicnaMena AS SumaCM, E.C211_SumaSpolu AS SumaTM, " +
                "F.C200_Mena AS Mena, E.C202_KurzNBS AS Kurz, E.C201_MnozstvoJednotky AS MnozstvoJednotky, " +
                "CAST(F.C097_Poznamka AS varchar(max)) AS Poznamka " +
                "FROM T228_Faktury AS F INNER JOIN T040_EUD AS E ON F.C030_CisloFaktury = E.C030_CisloInterne " +
                $"WHERE F.C100_TypDokladu = {(short)eTypFaktury.OdoslanaFaktura} AND " +
                $"CAST(F.C097_Poznamka AS varchar(max)) LIKE {ServicesString.SqlString(like, _ctx.pripojeneFirmy.DbJeSQL())}";

            return _ctx.pripojeneFirmy.dataProvider.Query<EmagRiadok>().Sql(sql).ToList();
        }
    }
}
```

- [ ] **Step 4.2: Over DevExpress referencie v JurhanLib.csproj.** `DevExpress.Spreadsheet` (Workbook) potrebuje `DevExpress.Docs.v24.2` + `DevExpress.Office.v24.2.Core` + `DevExpress.Spreadsheet.v24.2.Core`. JurhanLib má Docs; ak chýbajú Office.Core/Spreadsheet.Core, pridaj do existujúceho DevExpress `<ItemGroup>`:

```xml
<Reference Include="DevExpress.Office.v24.2.Core"><HintPath>C:/Program Files/DevExpress 24.2/Components/Bin/NetCore/DevExpress.Office.v24.2.Core.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
<Reference Include="DevExpress.Spreadsheet.v24.2.Core"><HintPath>C:/Program Files/DevExpress 24.2/Components/Bin/NetCore/DevExpress.Spreadsheet.v24.2.Core.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
```

- [ ] **Step 4.3: Build + commit** (`feat: presun EmagParser do JurhanLib`).

---

### Task 5: RozuctovanieCore → JurhanLib (nový)

**Files:**
- Create: `...\JurhanLib\JurhanLib\Import\Dopravcovia\RozuctovanieCore.cs`
- Zdroj: `Rozuctovanie.cs` z nového programu

- [ ] **Step 5.1: Vytvor súbor.** Zmeny oproti originálu: ctx namiesto `Program.*`, návratový `eVysledokRozuctovania`, hľadanie bankového dokladu podľa poznámky keď `interneCislo` je null, `zobrazenieChyby` parametrizované. Poradie krokov pre SoZapoctomBanky je: kontrola fileName → kontrola duplicity (C149_ImportText) → hľadanie bankového dokladu (v origináli bol doklad prvý; zámena je zámerná — duplicitný email sa v službe presunie do Zaúčtované aj keď bankový doklad chýba).

```csharp
using DevExpress.Spreadsheet;
using OmegaLib.ImportTxt;
using OmegaLib.Models;
using OmegaLib.Repository;
using OmegaLib.Services;
using JurhanLib.Omega;
using JurhanModels.Enum;
using JurhanModels.Import;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JurhanLib.Import.Dopravcovia
{
    /// <summary>
    /// Jadro rozúčtovania dopravcov — presunutá logika z programu RozuctovanieDopravcov,
    /// použiteľná z UI programu aj zo služby.
    /// </summary>
    public class RozuctovanieCore
    {
        private readonly RozuctovanieContext _ctx;
        private readonly EudHlavickaRepository _eudHlavickaRepository;

        public RozuctovanieCore(RozuctovanieContext ctx)
        {
            _ctx = ctx;
            _eudHlavickaRepository = new EudHlavickaRepository(ctx.pripojeneFirmy.dataProvider);
        }

        public eVysledokRozuctovania Execute()
        {
            ServicesBase.NastavCultureInfoSk();

            if (_ctx.pripojeneFirmy.dataProvider == null)
            {
                ServicesError.ErrorEnd($"Nepodarilo sa otvoriť databázu " +
                    $"{Path.Combine(_ctx.pripojeneFirmy._omegaPath, "omega.mdb")} pripojených firiem v Omege !", _ctx.zobrazenieChyby);
                return eVysledokRozuctovania.Chyba;
            }

            _ctx.typRozuctovania = Lib.NastavTypRozuctovania(_ctx.typSuboru);

            if (_ctx.typRozuctovania == eTypRozuctovania.SoZapoctomBanky)
            {
                if (string.IsNullOrEmpty(_ctx.fileName))
                {
                    ServicesError.ErrorEnd("Vyberte súbor s platbami dopravcov", _ctx.zobrazenieChyby);
                    return eVysledokRozuctovania.Chyba;
                }

                string fileNameInEud = Path.GetFileNameWithoutExtension(_ctx.fileName);
                var doklady = _eudHlavickaRepository.DajDoklady("C149_ImportText = @1", fileNameInEud);
                if (doklady.Any())
                {
                    ServicesError.ErrorEnd($"Doklad {Environment.NewLine}{Environment.NewLine}" +
                        $"{fileNameInEud}{Environment.NewLine}{Environment.NewLine}" +
                        $"už bol rozúčtovaný !{Environment.NewLine}{Environment.NewLine}" +
                        $"Najprv vymažte doklady, ktoré vznikli jeho rozúčtovaním, " +
                        $"alebo zvoľte iný doklad pre rozúčtovanie.", _ctx.zobrazenieChyby);
                    return eVysledokRozuctovania.Duplicita;
                }

                _ctx.dokladPreRozuctovanie = DajDokladPreRozuctovanie();
                if (_ctx.dokladPreRozuctovanie == null)
                {
                    // chybu uz oznamil DajDokladPreRozuctovanie
                    return eVysledokRozuctovania.NenajdenyBankovyDoklad;
                }
            }

            var uhrady = NacitajUhradyZoSuboru();
            if (uhrady == null)
            {
                // chyba uz bola oznamena pri nacitavani suboru
                return eVysledokRozuctovania.Chyba;
            }

            if (uhrady.Any())
            {
                return RozuctujSubor(uhrady);
            }

            if (Lib.JeTypSuboruDobropis(_ctx.typSuboru))
            {
                ServicesError.ErrorEnd($"V databáze sa nenašli žiadne dobropisy k faktúram s typom úhrady dobierka, " +
                        $"ktoré by bolo možné sparovať.", _ctx.zobrazenieChyby);
            }
            else if (Lib.JeTypSuboruAllegro(_ctx.typSuboru))
            {
                ServicesError.ErrorEnd($"V databáze sa nenašli žiadne neuhradené faktúry z Allegra s typom úhrady bankový prevod, " +
                    $"a na ktoré ešte nebol vystavený dobropis.", _ctx.zobrazenieChyby);
            }
            else if (Lib.JeTypSuboruKaufland(_ctx.typSuboru))
            {
                ServicesError.ErrorEnd($"V databáze sa nenašli žiadne neuhradené faktúry z Kaufland s typom úhrady bankový prevod, " +
                    $"a na ktoré ešte nebol vystavený dobropis.", _ctx.zobrazenieChyby);
            }
            else
            {
                ServicesError.ErrorEnd($"Zo súboru {_ctx.fileName} sa mi nepodarilo načítať ani jednu úhradu. " +
                    $"Asi je zlá štruktúra súboru alebo sa nepodarilo spárovať žiadnu faktúru.", _ctx.zobrazenieChyby);
            }
            return eVysledokRozuctovania.ZiadneUhrady;
        }

        private IEnumerable<Uhrada> NacitajUhradyZoSuboru()
        {
            if (Lib.JeTypSuboruKaufland(_ctx.typSuboru))
            {
                return new KauflandParser(_ctx).NacitajUhrady();
            }

            if (Lib.JeTypSuboruEmag(_ctx.typSuboru))
            {
                return new EmagParser(_ctx).NacitajUhrady();
            }

            if (Lib.JeTypSuboruAllegro(_ctx.typSuboru))
            {
                return new AllegroParser(_ctx).NacitajUhrady();
            }

            if (Lib.JeTypSuboruDobropis(_ctx.typSuboru))
            {
                _ctx.fileName = _ctx.typSuboru.ToString() + " " + _ctx.mesiac.ToString("00");
                DobropisyParser dobropisyParser = new DobropisyParser();
                return dobropisyParser.NacitajUhrady(_ctx.pripojeneFirmy.dataProvider,
                    new DateTime(_ctx.pripojeneFirmy.firma.UctovnyRokDoRok, _ctx.mesiac, 1),
                    new DateTime(_ctx.pripojeneFirmy.firma.UctovnyRokDoRok, _ctx.mesiac, DateTime.DaysInMonth(_ctx.pripojeneFirmy.firma.UctovnyRokDoRok, _ctx.mesiac)));
            }

            if (Lib.JeTypSuboruPreConvertCsv(_ctx.typSuboru))
            {
                _ctx.fileName = ConvertXlsxToCsv(_ctx.fileName);
            }

            DopravcaToUhrady dopravcaToUhrady = new DopravcaToUhrady(_ctx);

            switch (_ctx.typSuboru)
            {
                case eTypSuboru.AmazonATDE:
                case eTypSuboru.AmazonIT:
                case eTypSuboru.AmazonFR:
                case eTypSuboru.AmazonES:
                    return dopravcaToUhrady.NacitajUhrady(2, 4, 7, -1, new List<string> { "Date" }, System.Text.Encoding.UTF8, true, ",");
                case eTypSuboru.Dopravca_DPD_Cesko:
                    return dopravcaToUhrady.NacitajUhrady(6, 2, -1, -1, new List<string> { "pl_number" }, System.Text.Encoding.UTF8, false);
                case eTypSuboru.Dopravca_DPD_Chorvatsko:
                    return dopravcaToUhrady.NacitajUhrady(5, 2, -1, -1, new List<string> { "Parcel number" }, System.Text.Encoding.UTF8, false);
                // Encoding.Default bol na Net48 ANSI (cp1250), na .NET 10 je UTF-8 —
                // ANSI vypisy dopravcov drzime explicitne na Windows-1250
                case eTypSuboru.Dopravca_DPD_Madarsko:
                    return dopravcaToUhrady.NacitajUhrady(7, 2, -1, -1, new List<string> { "Parcel number" }, ServicesEncoding.Windows1250, false);
                case eTypSuboru.Dopravca_DPD_Polsko:
                    return dopravcaToUhrady.NacitajUhrady(10, 4, -1, -1, new List<string> { "Customer number" }, ServicesEncoding.Windows1250, false);
                case eTypSuboru.Dopravca_DPD_Rumunsko:
                    return dopravcaToUhrady.NacitajUhrady(7, 10, -1, 11, new List<string> { "Line" }, ServicesEncoding.Windows1250, false);
                case eTypSuboru.Dopravca_DPD_Slovensko:
                    return dopravcaToUhrady.NacitajUhrady(8, 2, -1, 4, new List<string> { "parcelno" }, System.Text.Encoding.UTF8, false);
                case eTypSuboru.Dopravca_DPD_Slovinsko:
                    return dopravcaToUhrady.NacitajUhrady(6, 2, -1, -1, new List<string> { "Št. paketa" }, System.Text.Encoding.UTF8, false);
                case eTypSuboru.Dopravca_Express_One:
                    return dopravcaToUhrady.NacitajUhrady(2, 3, -1, 4, new List<string> { "VS" }, System.Text.Encoding.UTF8, false);
                case eTypSuboru.Dopravca_GLS:
                    return dopravcaToUhrady.NacitajUhrady(2, 4, -1, 5, new List<string> { "Journal No." }, ServicesEncoding.Windows1250, false);
                case eTypSuboru.Dopravca_GLS_PLN:
                    return dopravcaToUhrady.NacitajUhrady(5, 3, -1, -1, new List<string> { "Nr paczki" }, ServicesEncoding.Windows1250, false);
                case eTypSuboru.Dopravca_Packeta:
                    return dopravcaToUhrady.NacitajUhrady(4, 12, -1, 13, new List<string> { "Váš e-shop" }, System.Text.Encoding.UTF8, true);
                case eTypSuboru.Dopravca_SPS:
                    return dopravcaToUhrady.NacitajUhrady(2, 3, -1, 4, new List<string> { "VS" }, System.Text.Encoding.UTF8, true);
                case eTypSuboru.PlatobnaBrana_ComgatePLN:
                    return dopravcaToUhrady.NacitajUhrady(8, 14, -1, 13, new List<string> { "Merchant" }, System.Text.Encoding.UTF8, true);
                case eTypSuboru.PlatobnaBrana_Gopay:
                    return dopravcaToUhrady.NacitajUhrady(4, 5, -1, 8, new List<string> { "ID pohybu" }, System.Text.Encoding.UTF8, true);
                case eTypSuboru.PlatobnaBrana_PayPal:
                case eTypSuboru.PlatobnaBrana_SkPay:
                    return dopravcaToUhrady.NacitajUhrady(27, 54, -1, 63, new List<string> { "ID platby" }, System.Text.Encoding.UTF8, false);
                default:
                    ServicesError.ErrorEnd($"Nepodporovaný typ súboru {_ctx.typSuboru} !", _ctx.zobrazenieChyby);
                    return null;
            }
        }

        private eVysledokRozuctovania RozuctujSubor(IEnumerable<Uhrada> uhrady)
        {
            string importPath = Path.Combine(_ctx.pripojeneFirmy._omegaPath, "Import");

            if (!Directory.Exists(importPath))
            {
                Directory.CreateDirectory(importPath);
            }

            _ctx.fileName = Path.Combine(importPath, Path.GetFileNameWithoutExtension(_ctx.fileName) + ".txt");

            if (!Lib.JeTypSuboruAllegroAleboKauflandAleboDobropis(_ctx.typSuboru) && SuborJeUzRozuctovany(_ctx.fileName))
            {
                return eVysledokRozuctovania.Duplicita;
            }

            UhradyToTxt uhradyToTxt = new UhradyToTxt(_ctx.pripojeneFirmy.dataProvider, _ctx.typRozuctovania, _ctx.typSuboru, _ctx.country,
                _ctx.fileName, _ctx.dokladPreRozuctovanie, _ctx.mesiac, _ctx.zobrazenieChyby, _ctx.typSpustenia);
            uhradyToTxt.Execute(uhrady);
            if (uhradyToTxt.sparovaneVS.Any())
            {
                NaimportujSubor(uhradyToTxt.doklady, importPath);
            }
            return eVysledokRozuctovania.Rozuctovane;
        }

        private void NaimportujSubor(IEnumerable<EUDHlavickaTxt> doklady, string importPath)
        {
            var autoimport = new MyAutoImport(_ctx.pripojeneFirmy.dataProvider, _ctx.pripojeneFirmy.firma.IdFirma,
                _ctx.pripojeneFirmy._omegaPath, importPath,
                false, eImportTxtPridatPrepisat.Pridat, _ctx.zobrazenieChyby);

            ImportTxtEUD importTxtEUD = new ImportTxtEUD(_ctx.fileName);
            importTxtEUD.Zapis(doklady, "T00");

            if (doklady.Any())
            {
                autoimport.Spusti();
                if (autoimport.BolUspesny())
                {
                    ServicesError.ErrorEnd("Program úspešne naimportoval doklady pre rozučtovanie !", _ctx.zobrazenieChyby);
                }
            }
            else
            {
                ServicesError.ErrorEnd("Nenaimportovali sa žiadne doklady !", _ctx.zobrazenieChyby);
            }
        }

        private string ConvertXlsxToCsv(string fileName)
        {
            string fileNameCsv = ServicesFile.VymenPriponuSuboru(fileName, "csv");
            Workbook workbook = new Workbook();
            workbook.LoadDocument(fileName);
            workbook.Options.Export.Csv.ValueSeparator = ';';
            workbook.SaveDocument(fileNameCsv, DocumentFormat.Csv);
            return fileNameCsv;
        }

        private EUDHlavicka DajDokladPreRozuctovanie()
        {
            EUDHlavicka doklad;
            if (!string.IsNullOrEmpty(_ctx.interneCislo))
            {
                doklad = _eudHlavickaRepository.DajDokladPodlaInternehoCisla(_ctx.interneCislo);
                if (doklad == null)
                {
                    ServicesError.ErrorEnd($"Doklad s interným číslom {_ctx.interneCislo} neexistuje v evidencii !",
                        _ctx.zobrazenieChyby);
                }
            }
            else
            {
                // sluzba: bankovy doklad sa hlada podla poznamky = nazov spracovaneho suboru bez pripony
                string poznamka = Path.GetFileNameWithoutExtension(_ctx.fileName);
                doklad = _eudHlavickaRepository.DajDokladPodlaPoznamky(poznamka, _ctx.pripojeneFirmy.DbJeSQL());
                if (doklad == null)
                {
                    ServicesError.ErrorEnd($"Nenašiel sa bankový doklad s poznámkou '{poznamka}' pre rozúčtovanie súboru {_ctx.fileName} !",
                        _ctx.zobrazenieChyby);
                }
            }
            return doklad;
        }

        private bool SuborJeUzRozuctovany(string fileName)
        {
            var doklady = _eudHlavickaRepository.DajDoklady("C149_ImportText = @1", "dopravca: " + Path.GetFileNameWithoutExtension(fileName));
            if (doklady.Any())
            {
                ServicesError.ErrorEnd($"Doklad {Environment.NewLine}{Environment.NewLine}" +
                    $"{Path.GetFileNameWithoutExtension(fileName)}{Environment.NewLine}{Environment.NewLine}" +
                    $"už bol rozúčtovaný !{Environment.NewLine}{Environment.NewLine}" +
                    $"Najprv vymažte doklady, ktoré vznikli jeho rozúčtovaním.",
                    _ctx.zobrazenieChyby);
                return true;
            }
            return false;
        }
    }
}
```

Pozn.: pôvodná podmienka v `DajDokladPreRozuctovanie` (`!JeTypSuboruAllegroAleboKauflandAleboDobropis && SoZapoctomBanky`) je teraz zaručená volajúcim (metóda sa volá len vo vetve SoZapoctomBanky a AKD/Emag typy sú vždy BezZapoctuBanky podľa `Lib.NastavTypRozuctovania`).

- [ ] **Step 5.2: Build + commit** (`feat: RozuctovanieCore - zdielane jadro rozuctovania s vysledkovym enumom a hladanim dokladu podla poznamky`).

---

### Task 6: Nový program RozuctovanieDopravcov volá RozuctovanieCore

**Files:**
- Modify: `C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\Rozuctovanie.cs` (nahradiť celý obsah)
- Delete: v tom istom adresári `DopravcaToUhrady.cs`, `BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`, `EmagParser.cs`

- [ ] **Step 6.1: Nahraď obsah Rozuctovanie.cs adaptérom**

```csharp
using JurhanLib.Import.Dopravcovia;
using JurhanModels.Enum;
using JurhanModels.Import;
using OmegaLib.Enums;

namespace RozuctovanieDopravcov
{
    /// <summary>
    /// Tenký adaptér — logika je v JurhanLib.Import.Dopravcovia.RozuctovanieCore,
    /// zdieľaná so službou JurhanService_RozuctovanieDopravcov.
    /// </summary>
    internal class Rozuctovanie
    {
        internal void Execute(eTypSuboru typSuboru)
        {
            RozuctovanieContext ctx = new RozuctovanieContext
            {
                pripojeneFirmy = Program.pripojeneFirmy,
                typSuboru = typSuboru,
                mesiac = Program.mesiac,
                interneCislo = Program.interneCislo,
                fileName = Program.fileName,
                zobrazenieChyby = eZobrazenieChyby.ZapisDoSuboruAOznamVoFormulari,
                typSpustenia = eTypSpustenia.Program,
            };

            RozuctovanieCore core = new RozuctovanieCore(ctx);
            core.Execute();

            Program.typRozuctovania = ctx.typRozuctovania;
            Program.dokladPreRozuctovanie = ctx.dokladPreRozuctovanie;
            Program.fileName = ctx.fileName;
            Program.currency = ctx.currency;
            Program.country = ctx.country;
        }
    }
}
```

Pozn.: over, či `eZobrazenieChyby` je v `OmegaLib.Enums` alebo `OmegaLib.Services` (podľa pôvodných usingov v tomto projekte) a či `frmVyber` volá `new Rozuctovanie().Execute(...)` s rovnakou signatúrou (áno — `Execute(eTypSuboru)`). `using JurhanModels.Import` je pre `eTypSuboru`.

- [ ] **Step 6.2: Zmaž presunuté súbory** `DopravcaToUhrady.cs`, `BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`, `EmagParser.cs` (SDK projekt — stačí zmazať súbory). Ak `frmVyber.cs` alebo iný súbor referencuje zmazané triedy priamo (over grepom `DopravcaToUhrady|BaseParser|AllegroParser|KauflandParser|EmagParser` v projekte), nahraď použitia — očakáva sa, že ich používal len `Rozuctovanie.cs`.

- [ ] **Step 6.3: Build programu**

Run: `dotnet build C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\RozuctovanieDopravcov.csproj -v:m`
Expected: Build succeeded.

- [ ] **Step 6.4: Commit v repe RozuctovanieDopravcov** (`refactor: rozuctovanie presunute do JurhanLib.Import.Dopravcovia, program vola RozuctovanieCore`).

---

### Task 7: Constants — IMAP prístup (nové JurhanModels)

**Files:**
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\Constants\Constants.cs` (za riadok `public const string ClientPasswordNechalova = "Arer951*";`, cca :802)

- [ ] **Step 7.1: Pridaj konštanty**

```csharp
    // IMAP pristup k schranke platby@jurhan.com (sluzba RozuctovanieDopravcov)
    // DOPLNIT (Tuleja): realny IMAP host a heslo
    public const string ClientHostImapJurhan = "imap.floxm.com";
    public const string ClientPasswordPlatby = "";
```

- [ ] **Step 7.2: Build JurhanModels + commit** (`feat: IMAP konstanty pre schranku platby@jurhan.com`).

---

### Task 8: Nová služba — projekt a kostra (net10)

**Files (Create), adresár `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\`:**
- `JurhanService_RozuctovanieDopravcov.csproj`
- `Program.cs`, `Service.cs`, `Execute.cs`, `SpustenieServicy.cs`, `RozuctovanieLogger.cs`
- `Properties\AssemblyInfo.cs` (skopíruj z `ImportDobropisov\JurhanService_ImportDobropisov\Properties\AssemblyInfo.cs` a premenuj názvy assembly na `JurhanService_RozuctovanieDopravcov`)
- `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.sln`

- [ ] **Step 8.1: csproj** (vzor ImportDobropisov + MailKit):

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <OutputType>WinExe</OutputType>
    <UseWindowsForms>true</UseWindowsForms>
    <AssemblyName>JurhanService_RozuctovanieDopravcov</AssemblyName>
    <RootNamespace>JurhanService_RozuctovanieDopravcov</RootNamespace>
    <GenerateAssemblyInfo>false</GenerateAssemblyInfo>
    <NoWarn>NU1701;CS0618;CA1416</NoWarn>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.Windows.Compatibility" Version="10.0.9" />
    <PackageReference Include="MailKit" Version="4.13.0" />
  </ItemGroup>

  <ItemGroup>
    <Reference Include="DevExpress.Docs.v24.2"><HintPath>C:/Program Files/DevExpress 24.2/Components/Bin/NetCore/DevExpress.Docs.v24.2.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
    <Reference Include="DevExpress.Office.v24.2.Core"><HintPath>C:/Program Files/DevExpress 24.2/Components/Bin/NetCore/DevExpress.Office.v24.2.Core.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
    <Reference Include="DevExpress.Spreadsheet.v24.2.Core"><HintPath>C:/Program Files/DevExpress 24.2/Components/Bin/NetCore/DevExpress.Spreadsheet.v24.2.Core.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\..\..\OmegaLib\OmegaLib\OmegaLib.csproj" />
    <ProjectReference Include="..\..\..\DataProvider\DataProvider\DataProvider.csproj" />
    <ProjectReference Include="..\..\..\JurhanProgramy\JurhanModels\JurhanModels\JurhanModels.csproj" />
    <ProjectReference Include="..\..\..\JurhanProgramy\JurhanLib\JurhanLib\JurhanLib.csproj" />
  </ItemGroup>

  <!-- Kros.KORM 3.8.0 + Kros.Utils z C:/Omega (jednotna verzia s DataProvider) -->
  <ItemGroup>
    <Reference Include="Kros.KORM"><HintPath>C:/Omega/Kros.KORM.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
    <Reference Include="Kros.Utils"><HintPath>C:/Omega/Kros.Utils.dll</HintPath><SpecificVersion>False</SpecificVersion></Reference>
  </ItemGroup>
</Project>
```

- [ ] **Step 8.2: Program.cs**

```csharp
using JurhanLib.Services;
using JurhanModels.Enum;

namespace JurhanService_RozuctovanieDopravcov
{
    internal static class Program
    {
        internal static eTypSpustenia typSpustenia = eTypSpustenia.Servica;
        internal static string serviceName = "JurhanServiceNew_RozuctovanieDopravcov";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            ServiceRunner.RunService<Service, Execute>(args, ref typSpustenia);
        }
    }
}
```

- [ ] **Step 8.3: Service.cs**

```csharp
using JurhanService;
using System;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class Service : ServiceJurhan
    {
        public override void ExecuteRun()
        {
            Execute execute = new Execute();
            execute.Run();
        }

        public override TimeSpan Interval()
        {
            return new TimeSpan(1, 0, 0); // kazdu hodinu
        }

        public override string Name()
        {
            return Program.serviceName;
        }
    }
}
```

- [ ] **Step 8.4: RozuctovanieLogger.cs**

```csharp
using JurhanLib.Logger;
using System;
using System.Collections.Generic;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class RozuctovanieLogger : Logger
    {
        public override void PosliLogSuborEmailom(IEnumerable<string> adresy, int hour, IEnumerable<DayOfWeek> days)
        {
        }
    }
}
```

- [ ] **Step 8.5: Execute.cs** — beh každý deň, celé časové okno (emaily chodia priebežne), bez emailu po každom behu (`IbaProgram`):

```csharp
using OmegaLib.Services;
using JurhanLib.Services;
using JurhanModels.Enum;
using System.Collections.Generic;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class Execute : IExecutable
    {
        private readonly RozuctovanieLogger _logger = new RozuctovanieLogger();

        public void Run()
        {
            ServicesLog.VytvorLogovaciAdresar();

            SpustenieServicy spustenieServicy = new SpustenieServicy(_logger);
            spustenieServicy.Execute(Program.typSpustenia, eTypServisy.IbaProgram, 0, new List<short> { 0, 1, 2, 3, 4, 5, 6 }, 0, 23,
                "Rozúčtovanie dopravcov",
                new List<string> { Constants.MessageToTulejaX });
        }
    }
}
```

- [ ] **Step 8.6: SpustenieServicy.cs** (zatiaľ len log — worker sa doplní v Tasku 10):

```csharp
using OmegaLib.Services;
using JurhanLib.Services;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class SpustenieServicy : SpustenieServisyBase
    {
        private readonly RozuctovanieLogger _logger;

        internal SpustenieServicy(RozuctovanieLogger logger) : base(logger)
        {
            _logger = logger;
        }

        public override void VykonajFunkciu()
        {
            RozuctovanieEmailov rozuctovanieEmailov = new RozuctovanieEmailov(pripojeneFirmy, _logger);
            rozuctovanieEmailov.Execute();
        }
    }
}
```

Pozn.: `: base(logger)` — `SpustenieServisyBase` má konštruktor s `Logger` (na logovanie štart/koniec cez logger; `CreditNotLogger` v ImportDobropisov ho nevyužíva, my áno).
Pozn. 2: tento súbor sa build-ne až so Step 10.1 (RozuctovanieEmailov ešte neexistuje) — Stepy 8.x a 9.x a 10.1 dokonč pred prvým buildom, alebo sem dočasne nedávaj telo `VykonajFunkciu`.

- [ ] **Step 8.7: sln**

```powershell
dotnet new sln -n JurhanService_RozuctovanieDopravcov -o C:\Projekty\Private\JurhanService\RozuctovanieDopravcov
dotnet sln C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.sln add C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.csproj
```

---

### Task 9: Nová služba — mapovanie priečinkov a mesiac z názvu súboru

**Files:**
- Create: `...\JurhanService_RozuctovanieDopravcov\FolderMapping.cs`
- Create: `...\JurhanService_RozuctovanieDopravcov\NazovSuboru.cs`

- [ ] **Step 9.1: FolderMapping.cs** — ľudské názvy sú prvotný odhad, používateľ si ich doplní/upraví:

```csharp
using JurhanModels.Import;

namespace JurhanService_RozuctovanieDopravcov
{
    /// <summary>
    /// Mapovanie nazvu IMAP priecinka schranky platby@jurhan.com na typ dopravcu.
    /// DOPLNIT (Tuleja): presne nazvy priecinkov podla realnej schranky.
    /// </summary>
    internal static class FolderMapping
    {
        internal static eTypSuboru DajTypSuboru(string nazovPriecinka)
        {
            switch (nazovPriecinka)
            {
                case "Amazon ATDE": return eTypSuboru.AmazonATDE;
                case "Amazon IT": return eTypSuboru.AmazonIT;
                case "Amazon FR": return eTypSuboru.AmazonFR;
                case "Amazon ES": return eTypSuboru.AmazonES;
                case "Allegro CZ": return eTypSuboru.Allegro_CZ;
                case "Allegro HU": return eTypSuboru.Allegro_HU;
                case "Allegro PL": return eTypSuboru.Allegro_PL;
                case "Allegro SK": return eTypSuboru.Allegro_SK;
                case "Dobropisy": return eTypSuboru.Dobropisy;
                case "Comgate PLN": return eTypSuboru.PlatobnaBrana_ComgatePLN;
                case "GoPay": return eTypSuboru.PlatobnaBrana_Gopay;
                case "SkPay": return eTypSuboru.PlatobnaBrana_SkPay;
                case "PayPal": return eTypSuboru.PlatobnaBrana_PayPal;
                case "DPD Cesko": return eTypSuboru.Dopravca_DPD_Cesko;
                case "DPD Chorvatsko": return eTypSuboru.Dopravca_DPD_Chorvatsko;
                case "DPD Madarsko": return eTypSuboru.Dopravca_DPD_Madarsko;
                case "DPD Polsko": return eTypSuboru.Dopravca_DPD_Polsko;
                case "DPD Rumunsko": return eTypSuboru.Dopravca_DPD_Rumunsko;
                case "DPD Slovensko": return eTypSuboru.Dopravca_DPD_Slovensko;
                case "DPD Slovinsko": return eTypSuboru.Dopravca_DPD_Slovinsko;
                case "Express One": return eTypSuboru.Dopravca_Express_One;
                case "GLS": return eTypSuboru.Dopravca_GLS;
                case "GLS PLN": return eTypSuboru.Dopravca_GLS_PLN;
                case "Packeta": return eTypSuboru.Dopravca_Packeta;
                case "SPS": return eTypSuboru.Dopravca_SPS;
                case "Kaufland AT": return eTypSuboru.Kaufland_EUR_AT;
                case "Kaufland DE": return eTypSuboru.Kaufland_EUR_DE;
                case "Kaufland FR": return eTypSuboru.Kaufland_EUR_FR;
                case "Kaufland IT": return eTypSuboru.Kaufland_EUR_IT;
                case "Kaufland SK": return eTypSuboru.Kaufland_EUR_SK;
                case "Kaufland CZ": return eTypSuboru.Kaufland_CZK_CZ;
                case "Kaufland PL": return eTypSuboru.Kaufland_PLN_PL;
                case "Emag RO": return eTypSuboru.Emag_RON_RO;
                case "Emag HU": return eTypSuboru.Emag_HUF_HU;
                default: return eTypSuboru.Undefined;
            }
        }
    }
}
```

- [ ] **Step 9.2: NazovSuboru.cs**

```csharp
using System.Text.RegularExpressions;

namespace JurhanService_RozuctovanieDopravcov
{
    internal static class NazovSuboru
    {
        /// <summary>
        /// Vytiahne mesiac z nazvu suboru. Podporovane vzory: MM.RRRR, MM-RRRR, MM_RRRR a RRRR-MM
        /// (aj s bodkou/podciarkovnikom). Vrati 0, ak sa mesiac nenasiel.
        /// </summary>
        internal static short DajMesiac(string nazovBezPripony)
        {
            Match m = Regex.Match(nazovBezPripony, @"(?<![0-9])(0[1-9]|1[0-2])[._-](20[0-9]{2})(?![0-9])");
            if (m.Success)
            {
                return short.Parse(m.Groups[1].Value);
            }

            m = Regex.Match(nazovBezPripony, @"(?<![0-9])(20[0-9]{2})[._-](0[1-9]|1[0-2])(?![0-9])");
            if (m.Success)
            {
                return short.Parse(m.Groups[2].Value);
            }

            return 0;
        }
    }
}
```

---

### Task 10: Nová služba — IMAP worker RozuctovanieEmailov

**Files:**
- Create: `...\JurhanService_RozuctovanieDopravcov\RozuctovanieEmailov.cs`

- [ ] **Step 10.1: Vytvor súbor**

```csharp
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

            using (ImapClient client = new ImapClient())
            {
                client.Connect(Constants.ClientHostImapJurhan, 993, true);
                client.Authenticate(Constants.MessageToPlatbyJurhan, Constants.ClientPasswordPlatby);

                foreach (IMailFolder folder in client.GetFolders(client.PersonalNamespaces[0]))
                {
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

            IList<UniqueId> uids = folder.Search(SearchQuery.All);
            if (!uids.Any())
            {
                return;
            }

            _logger.Loguj($"Priečinok '{folder.FullName}' ({typSuboru}): {uids.Count} emailov.", true);

            List<UniqueId> naPresun = new List<UniqueId>();
            foreach (UniqueId uid in uids)
            {
                MimeMessage message = folder.GetMessage(uid);
                try
                {
                    if (SpracujEmail(message, typSuboru))
                    {
                        naPresun.Add(uid);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Loguj($"Chyba pri spracovaní emailu '{message.Subject}': {ex}", true);
                }
            }

            if (naPresun.Any())
            {
                IMailFolder zauctovane = DajAleboVytvorZauctovane(folder);
                folder.MoveTo(naPresun, zauctovane);
                _logger.Loguj($"Presunutých {naPresun.Count} emailov do '{zauctovane.FullName}'.", true);
            }
        }

        /// <returns>true, ak sa ma email presunut do podpriecinka "Zaúčtované"</returns>
        private bool SpracujEmail(MimeMessage message, eTypSuboru typSuboru)
        {
            bool asponJedenSubor = false;
            bool vsetkoZauctovane = true;

            foreach (MimeEntity attachment in message.Attachments)
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

        private string UlozPrilohu(MimeEntity attachment)
        {
            string fileName = attachment.ContentDisposition?.FileName ?? attachment.ContentType.Name;
            if (fileName == null || !_povolenePripony.Contains(Path.GetExtension(fileName).ToLowerInvariant()))
            {
                return null;
            }

            string filePath = Path.Combine(_workDir, ToSafeFileName(fileName));
            using (FileStream stream = File.Create(filePath))
            {
                if (attachment is MessagePart messagePart)
                {
                    messagePart.Message.WriteTo(stream);
                }
                else
                {
                    ((MimePart)attachment).Content.DecodeTo(stream);
                }
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
```

Pozor na dve veci pri kompilácii:
1. `Lib` je `JurhanLib.Import.Lib` (using `JurhanLib.Import`) a `eTypRozuctovania` je tiež v `JurhanLib.Import` — ak by kolidoval `Lib` s `JurhanModels.Kurieri.Lib`, použi plnú kvalifikáciu `JurhanLib.Import.Lib.NastavTypRozuctovania`.
2. Pri XLSX→CSV konverzii jadro zmení `ctx.fileName` na `.csv`; pôvodný stiahnutý súbor mažeme cez lokálnu premennú `filePath` — vzniknutý `.csv` v `_workDir` ostáva; ak chceš čistiť aj ten, over `ctx.fileName` po behu (nie je nutné — adresár je pracovný).

- [ ] **Step 10.2: Build celej služby**

Run: `dotnet build C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.csproj -v:m`
Expected: Build succeeded.

---

### Task 11: Deploy skripty (nová strana)

**Files:**
- Modify: `C:\Projekty\Private\JurhanService\Deploy\Install-AllJurhanServices.ps1` (pole `$ExeBaseNames`, cca :54-71)
- Modify: Publish skript — over grepom `Publish-AllJurhanServices` v `C:\Projekty\Private\JurhanService\Deploy\` a pridaj analogickú položku, ak obsahuje zoznam služieb.

- [ ] **Step 11.1: Pridaj do `$ExeBaseNames` (abecedne za `'JurhanService_RecenzieEmailom'`):**

```powershell
    'JurhanService_RozuctovanieDopravcov'
```

- [ ] **Step 11.2:** Ak `Publish-AllJurhanServices.ps1` (alebo `.bat`) obsahuje zoznam projektov/ciest, pridaj rovnakým štýlom riadok pre `RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov`. Ak zoznam derivuje z adresárovej štruktúry automaticky, nič nemeň.

---

### Task 12: Git + smoke test (nová strana)

- [ ] **Step 12.1: Git init a commit novej služby**

```powershell
git -C C:\Projekty\Private\JurhanService\RozuctovanieDopravcov init
git -C C:\Projekty\Private\JurhanService\RozuctovanieDopravcov add -A
git -C C:\Projekty\Private\JurhanService\RozuctovanieDopravcov commit -m "feat: sluzba RozuctovanieDopravcov (net10) - IMAP schranka platby@jurhan.com, rozuctovanie priloh cez RozuctovanieCore"
```

(Ak treba, pridaj `.gitignore` s `bin/`, `obj/`, `Log/` pred `git add`.)

- [ ] **Step 12.2: Smoke beh v režime Program** — spustí sa len ak sú doplnené `ClientPasswordPlatby`/host (inak sa očakáva zalogovaná chyba pripojenia — to je akceptovateľný výsledok smoke testu):

```powershell
& C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\bin\Debug\net10.0-windows\JurhanService_RozuctovanieDopravcov.exe p
Get-Content (Get-ChildItem C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\bin\Debug\net10.0-windows\Log\*.log | Sort-Object LastWriteTime | Select-Object -Last 1).FullName -Tail 30
```

Expected: log obsahuje „Začiatok: Rozúčtovanie dopravcov“, otvorenie firmy; pri prázdnom hesle IMAP chybu (zachytenú, nie pád procesu).

---

## ADDENDUM po code review Batch 1 (záväzné pre Task 4, 5 a 13)

Code review odhalil diery pre servisné použitie; nasledovné zmeny sú súčasťou plánu (nová aj stará strana):

1. **RozuctovanieCore.Execute()** — hneď po `NastavCultureInfoSk()` pridaj `_ctx.dokladPreRozuctovanie = null;` (originál ho prepisoval pri každom behu; poistka proti znovupoužitiu kontextu).
2. **RozuctujSubor** — po `uhradyToTxt.Execute(uhrady)`: ak `!uhradyToTxt.sparovaneVS.Any()` → `return eVysledokRozuctovania.ZiadneUhrady;` inak `return NaimportujSubor(uhradyToTxt.doklady, importPath) ? eVysledokRozuctovania.Rozuctovane : eVysledokRozuctovania.Chyba;`
3. **NaimportujSubor** vracia `bool`: `true` len ak `doklady.Any() && autoimport.BolUspesny()` (úspešná hláška ostáva); inak `false` (ostatné hlášky ostávajú ako doteraz).
4. **Workbook dispose** — `ConvertXlsxToCsv` (RozuctovanieCore) aj `NacitajReferenceIdZExcelu` (EmagParser) obalia `Workbook` do `using` (dlho bežiaca služba).
5. **RozuctovanieContext** XML doc doplniť: „Jeden kontext = jedno spracovanie súboru; nepoužívaj opakovane.“
6. **DajDokladPreRozuctovanie** (po review Batch 2) — diskriminátor UI vs. služba je `_ctx.interneCislo != null` (NIE `IsNullOrEmpty`): UI posiela prázdny string z textboxu a musí dostať pôvodnú chybu „Doklad s interným číslom ... neexistuje v evidencii !“; len služba posiela `null` → hľadanie podľa poznámky.
7. **Záloha prílohy** (po review Batch 3, plní spec §4 bod 3a) — v `RozuctovanieEmailov.SpracujSubor` sa pred `core.Execute()` volá `ServicesFile.ZalohujSubor(filePath);` (záloha do `<exe>\Log` s časovou pečiatkou; súbor sa potom vo `finally` maže ako doteraz). Platí aj pre starú službu (Task 16).
8. **Perzistencia loggera** (po quality review Batch 3 — KRITICKÉ) — `JurhanLib.Logger.Logger.Loguj` zapisuje len do pamäte; bez flushu sa všetko stratí. `Execute.Run()` musí okolo `spustenieServicy.Execute(...)` volať `_logger.NacitajDataZoSuboru()` pred a vo `finally` perzistenciu (`ZapisDataDoSuboru` + zápis do log súboru) — presný API vzor prevziať z ImportObjednavok (`ImportOrders.cs` + `OrdersLogger.cs`).
9. **Okamžitý presun emailu** — email sa presúva do „Zaúčtované“ jednotlivo hneď po úspešnom `SpracujEmail` (per-uid `MoveTo`), nie dávkovo po celom priečinku — inak výnimka medzi zaúčtovaním a presunom nechá zaúčtované emaily v priečinku a Allegro/Kaufland/Dobropisy sa pri opakovaní nikdy nevrátia ako Duplicita (vracajú ZiadneUhrady) → večné opakovanie.
10. **Prílohy cez BodyParts** — `message.Attachments` nevidí inline prílohy; iterovať `message.BodyParts.OfType<MimePart>().Where(p => !string.IsNullOrEmpty(p.FileName))` (filter prípon ostáva).
11. **GetMessage v try + NotDeleted** — `folder.GetMessage(uid)` patrí dovnútra per-email try; `folder.Search(SearchQuery.NotDeleted)` namiesto `All`.
12. **Len top-level priečinky + čistenie workDir** — spracúvajú sa len priečinky priamo pod koreňom schránky (vnorený priečinok s náhodne zhodným menom sa nesmie mapovať); `_workDir` sa na začiatku behu vyčistí od zvyškov (skonvertované .csv).

13. **Logger — parametrizovateľný názov súborov** — `JurhanLib.Logger.Logger` má hardcoded názvy „ImportObjednavok*“. Pridá sa spätne kompatibilný overload `protected Logger(string nazovLogu)` (parameterless ctor deleguje `: this("ImportObjednavok")`), `RozuctovanieLogger` volá `: base("RozuctovanieDopravcov")`.
14. **RozuctovanieLogger = plný OrdersLogger vzor** (plní spec §6) — override `PosliLogSuborEmailom` podľa `ImportObjednavok\OrdersLogger.cs` (o danej hodine: flush, zálohy, email s logmi na `Constants.MessageToTulejaX`, po úspešnom odoslaní zmazanie 4 log súborov — inak retry). `Execute.Run` ho volá vo `finally` po perzistencii: hodina 22, všetky dni. Bez toho logy rastú neobmedzene a .log sa pri každom behu duplikuje (NacitajDataZoSuboru resetuje počítadlá).

Body 8–14 platia rovnako pre starú stranu (Task 13 — starý JurhanLib Logger; Task 16 — stará služba).

Dôsledok pre službu (Task 10): `ZiadneUhrady` aj `Chyba` nechávajú email na mieste (už platí — presúva sa len `Rozuctovane`/`Duplicita`). Task 13 preberá **aktuálne** súbory z nového JurhanLib (vrátane týchto opráv), nie pôvodné bloky z Taskov 1–5.

---

## ČASŤ II — stará (net48) strana

### Task 13: Parsery a jadro → starý JurhanLib

**Files:**
- Create v `C:\Users\tuleja\source\repos\JurhanProgramy\JurhanLib\JurhanLib\Import\Dopravcovia\`: `RozuctovanieContext.cs`, `DopravcaToUhrady.cs`, `BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`, `EmagParser.cs`, `RozuctovanieCore.cs`
- Modify: `C:\Users\tuleja\source\repos\JurhanProgramy\JurhanLib\JurhanLib\JurhanLib.csproj` (classic — pridať `<Compile Include>`)

- [ ] **Step 13.1: Vytvor súbory.** Postup pre každý súbor: zober **novú** verziu z Tasku 1–5 a nahraď blok usingov podľa starého sveta. Mapa náhrad (odvodená z pôvodných starých súborov v `C:\Users\tuleja\source\repos\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\` — pri každom súbore si otvor starý originál a použi presne jeho usingy + doplň `JurhanLib.Import`/nič podľa potreby):

| Nový using | Starý using |
|---|---|
| `OmegaLib.Repository` | `IndiLib.Repository` |
| `OmegaLib.Services` | `IndiLib.Services` |
| `OmegaLib.Models` | `IndiLib.Modely` |
| `OmegaLib.Enums` | `OmegaNeo`/`Omega.BusinessLayer.Common` — over v starom origináli daného súboru (napr. starý `DopravcaToUhrady.cs` a `EmagParser.cs` majú `using Omega.BusinessLayer.Common;` pre `eTypFaktury`) |
| `OmegaLib.ImportTxt` | over v starom `Rozuctovanie.cs` (`IndiLib` obsahuje `ImportTxtEUD`? — použi usingy starého originálu) |
| `ServicesEncoding.Windows1250` v `RozuctovanieCore` | v starom svete použi `System.Text.Encoding.Default` presne tam, kde ho používa starý `Rozuctovanie.cs` (DPD HU/PL/RO, GLS, GLS_PLN) — net48 správanie sa nemení! |

Telá metód (logika) musia byť identické s novými verziami; jediné rozdiely sú usingy a encoding poznámka vyššie.

- [ ] **Step 13.2: Pridaj Compile entries do starého JurhanLib.csproj** (do ItemGroup s existujúcimi Compile):

```xml
<Compile Include="Import\Dopravcovia\RozuctovanieContext.cs" />
<Compile Include="Import\Dopravcovia\BaseParser.cs" />
<Compile Include="Import\Dopravcovia\AllegroParser.cs" />
<Compile Include="Import\Dopravcovia\KauflandParser.cs" />
<Compile Include="Import\Dopravcovia\EmagParser.cs" />
<Compile Include="Import\Dopravcovia\DopravcaToUhrady.cs" />
<Compile Include="Import\Dopravcovia\RozuctovanieCore.cs" />
```

- [ ] **Step 13.3: Over DevExpress referencie v starom JurhanLib.csproj** — `DevExpress.Spreadsheet` potrebuje net48 DLL; skopíruj presné `<Reference>` riadky DevExpress zo starého `RozuctovanieDopravcov.csproj` (`C:\Users\tuleja\source\repos\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\RozuctovanieDopravcov.csproj`), ak v JurhanLib chýbajú (Docs/Office.Core/Spreadsheet.Core z `..\..\..\..\..\..\..\Omega\`).

- [ ] **Step 13.4: Build starého JurhanLib** (msbuild príkaz z hlavičky plánu, csproj `C:\Users\tuleja\source\repos\JurhanProgramy\JurhanLib\JurhanLib\JurhanLib.csproj`). Expected: Build succeeded.

- [ ] **Step 13.5: Commit** (ak je repo): `feat: presun parserov a RozuctovanieCore do JurhanLib.Import.Dopravcovia (net48)`.

---

### Task 14: Constants — IMAP (staré JurhanModels)

**Files:**
- Modify: `C:\Users\tuleja\source\repos\JurhanProgramy\JurhanModels\JurhanModels\Constants\Constants.cs`

- [ ] **Step 14.1:** Pridaj rovnaké 4 riadky ako v Tasku 7 (za `ClientPasswordNechalova`).
- [ ] **Step 14.2:** Build starých JurhanModels (msbuild) + commit.

---

### Task 15: Starý program volá RozuctovanieCore

**Files:**
- Modify: `C:\Users\tuleja\source\repos\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\Rozuctovanie.cs` (nahradiť adaptérom z Tasku 6 — usingy podľa starého sveta: `IndiLib.Services` pre `eZobrazenieChyby` — over podľa pôvodného súboru)
- Modify: `RozuctovanieDopravcov.csproj` — odstráň `<Compile Include="...">` pre `DopravcaToUhrady.cs`, `BaseParser.cs`, `AllegroParser.cs`, `KauflandParser.cs`, `EmagParser.cs` a zmaž tieto súbory. `KauflandPdfParser.cs` (mŕtvy kód, existuje len v starej vetve) nechaj tak.

- [ ] **Step 15.1:** Uprav Rozuctovanie.cs (adaptér), zmaž presunuté súbory + Compile entries.
- [ ] **Step 15.2:** Build starého programu (msbuild `RozuctovanieDopravcov.csproj`). Expected: Build succeeded. Pozn.: starý program referencuje JurhanLib cez HintPath `bin\Debug` — Task 13 musí byť zbuildovaný skôr.
- [ ] **Step 15.3:** Commit (ak repo): `refactor: program vola RozuctovanieCore z JurhanLib (net48)`.

---

### Task 16: Stará služba (net48)

**Files (Create), adresár `C:\Users\tuleja\source\repos\JurhanService\RozuctovanieDopravcov\`:**
- `JurhanService_RozuctovanieDopravcov.sln`
- `zaregistruj_RozuctovanieDopravcov.bat`, `odregistruj_RozuctovanieDopravcov.bat`
- `JurhanService_RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.csproj`
- `JurhanService_RozuctovanieDopravcov\App.config`
- `JurhanService_RozuctovanieDopravcov\Properties\AssemblyInfo.cs` (skopíruj z ImportFaktur a premenuj)
- `JurhanService_RozuctovanieDopravcov\Program.cs`, `Service.cs`, `Execute.cs`, `SpustenieServicy.cs`, `RozuctovanieLogger.cs`, `FolderMapping.cs`, `NazovSuboru.cs`, `RozuctovanieEmailov.cs`, `ProjectInstaller.cs`

- [ ] **Step 16.1: C# súbory** — `Program.cs` (serviceName = `"JurhanService_RozuctovanieDopravcov"`, bez prefixu New!), `Service.cs`, `Execute.cs`, `SpustenieServicy.cs`, `RozuctovanieLogger.cs`, `FolderMapping.cs`, `NazovSuboru.cs`, `RozuctovanieEmailov.cs` = identické s novými (Tasky 8–10) s náhradou usingov `OmegaLib.Services→IndiLib.Services`, `OmegaLib.Repository→IndiLib` (PripojeneFirmy je v namespace `IndiLib`), `OmegaLib.Enums→IndiLib.Services` pre `eZobrazenieChyby` (over podľa starého `ImportFaktur\Execute.cs`). Navyše:

`ProjectInstaller.cs`:
```csharp
using JurhanLib.Services;

namespace JurhanService_RozuctovanieDopravcov
{
    public class ProjectInstaller : ProjectInstallerBase
    {
        public override string ServiceName()
        {
            return "JurhanService_RozuctovanieDopravcov";
        }
    }
}
```

- [ ] **Step 16.2: csproj** (vzor ImportFaktur + MailKit z packages adresára ArchivaciaEmailov; DevExpress zo starého programu):

```xml
<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" Condition="Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">AnyCPU</Platform>
    <ProjectGuid>{3F6B2C1A-9D4E-4B7A-8C25-71E0A5D9B334}</ProjectGuid>
    <OutputType>WinExe</OutputType>
    <RootNamespace>JurhanService_RozuctovanieDopravcov</RootNamespace>
    <AssemblyName>JurhanService_RozuctovanieDopravcov</AssemblyName>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <FileAlignment>512</FileAlignment>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
    <Deterministic>true</Deterministic>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugSymbols>true</DebugSymbols>
    <DebugType>full</DebugType>
    <Optimize>false</Optimize>
    <OutputPath>bin\Debug\</OutputPath>
    <DefineConstants>DEBUG;TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugType>pdbonly</DebugType>
    <Optimize>true</Optimize>
    <OutputPath>bin\Release\</OutputPath>
    <DefineConstants>TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="IndiLib, Version=1.0.0.0, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\..\..\..\..\Projekty\DevOps\Indivindi\IndiLib\IndiLib\bin\Debug\IndiLib.dll</HintPath>
    </Reference>
    <Reference Include="JurhanLib, Version=1.0.0.0, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\JurhanProgramy\JurhanLib\JurhanLib\bin\Debug\JurhanLib.dll</HintPath>
    </Reference>
    <Reference Include="JurhanModels, Version=1.0.0.0, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\JurhanProgramy\JurhanModels\JurhanModels\bin\Debug\JurhanModels.dll</HintPath>
    </Reference>
    <Reference Include="Kros.KORM, Version=3.8.0.0, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\..\..\..\..\Omega\Kros.KORM.dll</HintPath>
    </Reference>
    <Reference Include="Omega.BusinessLayer, Version=1.0.9309.16746, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\..\..\..\..\Omega\Omega.BusinessLayer.dll</HintPath>
    </Reference>
    <Reference Include="Omega.Database">
      <HintPath>..\..\..\..\..\..\..\Omega\Omega.Database.dll</HintPath>
    </Reference>
    <Reference Include="OmegaNeo, Version=1.0.10.0, Culture=neutral, processorArchitecture=MSIL">
      <SpecificVersion>False</SpecificVersion>
      <HintPath>..\..\..\..\..\..\..\Omega\OmegaNeo.dll</HintPath>
    </Reference>
    <Reference Include="MailKit, Version=4.13.0.0, Culture=neutral, PublicKeyToken=4e064fe7c44a8f1b, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\MailKit.4.13.0\lib\net48\MailKit.dll</HintPath>
    </Reference>
    <Reference Include="MimeKit, Version=4.13.0.0, Culture=neutral, PublicKeyToken=bede1c8a46c66814, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\MimeKit.4.13.0\lib\net48\MimeKit.dll</HintPath>
    </Reference>
    <Reference Include="System.Buffers, Version=4.0.4.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Buffers.4.6.0\lib\net462\System.Buffers.dll</HintPath>
    </Reference>
    <Reference Include="System.Formats.Asn1, Version=8.0.0.1, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Formats.Asn1.8.0.1\lib\net462\System.Formats.Asn1.dll</HintPath>
    </Reference>
    <Reference Include="System.Memory, Version=4.0.2.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Memory.4.6.0\lib\net462\System.Memory.dll</HintPath>
    </Reference>
    <Reference Include="System.Numerics.Vectors, Version=4.1.5.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Numerics.Vectors.4.6.0\lib\net462\System.Numerics.Vectors.dll</HintPath>
    </Reference>
    <Reference Include="System.Runtime.CompilerServices.Unsafe, Version=6.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Runtime.CompilerServices.Unsafe.6.1.0\lib\net462\System.Runtime.CompilerServices.Unsafe.dll</HintPath>
    </Reference>
    <Reference Include="System.Threading.Tasks.Extensions, Version=4.2.1.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.Threading.Tasks.Extensions.4.6.0\lib\net462\System.Threading.Tasks.Extensions.dll</HintPath>
    </Reference>
    <Reference Include="System.ValueTuple, Version=4.0.3.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>..\..\..\JurhanOdloz\ArchivaciaEmailov\packages\System.ValueTuple.4.5.0\lib\net47\System.ValueTuple.dll</HintPath>
    </Reference>
    <Reference Include="System" />
    <Reference Include="System.Configuration.Install" />
    <Reference Include="System.Core" />
    <Reference Include="System.Xml.Linq" />
    <Reference Include="System.Data.DataSetExtensions" />
    <Reference Include="Microsoft.CSharp" />
    <Reference Include="System.Data" />
    <Reference Include="System.Net.Http" />
    <Reference Include="System.ServiceProcess" />
    <Reference Include="System.Xml" />
  </ItemGroup>
  <ItemGroup>
    <Compile Include="Execute.cs" />
    <Compile Include="FolderMapping.cs" />
    <Compile Include="NazovSuboru.cs" />
    <Compile Include="Program.cs" />
    <Compile Include="ProjectInstaller.cs">
      <SubType>Component</SubType>
    </Compile>
    <Compile Include="Properties\AssemblyInfo.cs" />
    <Compile Include="RozuctovanieEmailov.cs" />
    <Compile Include="RozuctovanieLogger.cs" />
    <Compile Include="Service.cs">
      <SubType>Component</SubType>
    </Compile>
    <Compile Include="SpustenieServicy.cs" />
  </ItemGroup>
  <ItemGroup>
    <None Include="App.config" />
  </ItemGroup>
  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
</Project>
```

Pozn.: DevExpress referencie pridaj len ak build zlyhá na `DevExpress.Spreadsheet` (typy sú použité v JurhanLib; pri Reference DLL sa transitívne nekopírujú — pre runtime pridaj rovnaké DevExpress `<Reference>` riadky ako v starom `RozuctovanieDopravcov.csproj`, nech sa DLL dostanú do výstupu).

- [ ] **Step 16.3: App.config** (binding redirecty pre MailKit — prevzaté z ArchivaciaEmailov):

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <startup>
        <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.8" />
    </startup>
  <runtime>
    <assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
      <dependentAssembly>
        <assemblyIdentity name="System.Memory" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.0.2.0" newVersion="4.0.2.0" />
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="System.Buffers" publicKeyToken="cc7b13ffcd2ddd51" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.0.4.0" newVersion="4.0.4.0" />
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="MailKit" publicKeyToken="4e064fe7c44a8f1b" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.13.0.0" newVersion="4.13.0.0" />
      </dependentAssembly>
      <dependentAssembly>
        <assemblyIdentity name="MimeKit" publicKeyToken="bede1c8a46c66814" culture="neutral" />
        <bindingRedirect oldVersion="0.0.0.0-4.13.0.0" newVersion="4.13.0.0" />
      </dependentAssembly>
    </assemblyBinding>
  </runtime>
</configuration>
```

- [ ] **Step 16.4: bat súbory**

`zaregistruj_RozuctovanieDopravcov.bat`:
```bat
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe c:\JurhanService\JurhanService_RozuctovanieDopravcov.exe
```

`odregistruj_RozuctovanieDopravcov.bat`:
```bat
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe -u c:\JurhanService\JurhanService_RozuctovanieDopravcov.exe
```

(Over formát v `ImportFaktur\odregistruj_ImportFaktur.bat` a použi rovnaký.)

- [ ] **Step 16.5: sln** — skopíruj `C:\Users\tuleja\source\repos\JurhanService\ImportFaktur\JurhanService_ImportFaktur.sln` a v texte nahraď `ImportFaktur` → `RozuctovanieDopravcov` a projektový GUID → `{3F6B2C1A-9D4E-4B7A-8C25-71E0A5D9B334}` (musí sedieť s csproj).

- [ ] **Step 16.6: Build** (msbuild na sln alebo csproj). Expected: Build succeeded.

- [ ] **Step 16.7: Git** — ak `C:\Users\tuleja\source\repos\JurhanService\RozuctovanieDopravcov` nie je repo, `git init` + commit ako v Task 12.1 (message `feat: sluzba RozuctovanieDopravcov (net48)`).

---

### Task 17: Záverečná verifikácia

- [ ] **Step 17.1: Buildy všetkých dotknutých projektov naraz:**

```powershell
dotnet build C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\JurhanLib.csproj -v:m
dotnet build C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\RozuctovanieDopravcov.csproj -v:m
dotnet build C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.csproj -v:m
# net48 (msbuild z hlavicky planu):
# JurhanLib, JurhanModels, RozuctovanieDopravcov (program), JurhanService_RozuctovanieDopravcov
```
Expected: všetko Build succeeded, žiadne nové warningy typu CS0103/CS0246.

- [ ] **Step 17.2: Kontrola úplnosti voči spec** — prejdi spec sekcie 2–6 a odškrtni: presun logiky ✓ (Tasky 1–5, 13), program volá core ✓ (6, 15), IMAP konštanty ✓ (7, 14), kostra služby ✓ (8, 16), tok spracovania vrátane duplicity→Zaúčtované a nenájdený doklad→email ostáva ✓ (10), mapovanie priečinkov ✓ (9), deploy ✓ (11, 16.4), logovanie ✓ (RozuctovanieLogger + Loguj volania).

- [ ] **Step 17.3: Odovzdanie používateľovi** — zoznam vecí, ktoré musí doplniť: IMAP host + heslo v oboch `Constants.cs`, názvy priečinkov v oboch `FolderMapping.cs`, poznámky bankových dokladov v Omege = názov súboru bez prípony, formát mesiaca v názvoch súborov (MM.RRRR / MM-RRRR / MM_RRRR / RRRR-MM).
