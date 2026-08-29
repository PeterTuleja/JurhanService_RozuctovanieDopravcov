# Rozúčtovanie kasy — implementačný plán

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rozúčtovať xlsx export dokladov z kasy: hotovosť → pokladničné doklady PD/P1/P (MD 211001), karta → interné doklady ID/IDPB/IDSKP (MD 315102), oboje ako inkaso faktúry s dátumom úhrady z riadku.

**Architecture:** Jeden nový `eTypSuboru.Kasa` + `KasaParser` (DevExpress Spreadsheet). `Uhrada` nesie dátum úhrady a spôsob platby; `UhradyToTxt` ich číta per úhrada, takže jeden beh zaúčtuje zmiešané PD aj ID doklady (žiadne delenie na dávky nie je potrebné). Faktúra sa hľadá podľa interného čísla (C030). Duplicity chráni existujúca kontrola `FakturaJeUzUhradena` + kľúč súboru v C112.

**Tech Stack:** C# net10.0-windows, DevExpress Spreadsheet, xUnit. Repá: JurhanModels, JurhanLib, služba, program (každé samostatný git na `master`). OmegaLib sa nemení.

**Spec:** `docs/superpowers/specs/2026-08-29-kasa-rozuctovanie-design.md` (v repe služby).

**Poznámky pre realizátora:**
- Buildy/testy sa robia v hlavných adresároch repozitárov (nie vo worktree — worktree nie je buildovateľný, viď pamäť projektu).
- Commity: každý repozitár zvlášť, priamo na `master`.
- Kasa je EUR-only tuzemsko. Známe obmedzenie: keby kasa uhradila faktúru v cudzej mene, 3M vetva (`ZapisUhradu3M`) by použila okruh bez spôsobu platby — v praxi nenastáva, neriešime.

---

### Task 1: JurhanModels — enum Kasa, spôsob platby, Uhrada, konštanty

**Files:**
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\Import\Enums.cs`
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\Import\Uhrada.cs`
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\Enum\Enum.cs`
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\Constants\Constants.cs`

- [ ] **Step 1.1: Pridať `Kasa` a `eSposobPlatbyKasa` do `Import\Enums.cs`**

Na koniec enumu `eTypSuboru` (za `Dopravca_GLS_RO`, pred zatváraciu `}` enumu) a za enum pridať nový enum:

```csharp
        Dopravca_GLS_SK,
        Dopravca_GLS_RO,
        // Kasa (registracna pokladnica): export dokladov "Doklady *.xlsx", kazdy riadok je uhrada
        // faktury. Hotovost -> PD/P1/P, karta -> ID/IDPB/IDSKP; sposob platby nesie
        // Uhrada.SposobPlatbyKasa, datum dokladu Uhrada.DatumUhrady.
        Kasa
    }

    /// <summary>Spôsob platby dokladu kasy - určuje okruh/číselný rad a protiúčet inkasa.</summary>
    public enum eSposobPlatbyKasa
    {
        Ziadna = 0,
        Hotovost,
        Karta
    }
}
```

(Pôvodná zatváracia `}` súboru ostáva — enum `eSposobPlatbyKasa` je vnútri namespace `JurhanModels.Import`.)

- [ ] **Step 1.2: Rozšíriť `Uhrada.cs`**

```csharp
namespace JurhanModels.Import
{
    public class Uhrada
    {
        public string VS { get; set; }
        public string CisloBalika { get; set; }
        public string Dobropis { get; set; }
        public decimal SumaTM { get; set; }
        public decimal SumaCM { get; set; }
        public string Mena { get; set; }        
        public decimal MnozstvoJednotky { get; set; }
        public decimal Kurz { get; set; }
        public string FormaUhrady { get; set; }
        /// <summary>Dátum úhrady z riadku exportu (kasa). Keď je vyplnený, doklad inkasa sa
        /// datuje ním - nie dátumom faktúry ani bankového výpisu.</summary>
        public System.DateTime? DatumUhrady { get; set; }
        /// <summary>Spôsob platby dokladu kasy; mimo kasy ostáva Ziadna.</summary>
        public eSposobPlatbyKasa SposobPlatbyKasa { get; set; }
    }
}
```

- [ ] **Step 1.3: Pridať kódy evidencie a radu PD do `Enum\Enum.cs`**

Za enum `eKod_zIDCR_EUD` (končí `zIDz` na riadku ~292) vložiť:

```csharp
    /// <summary>Evidencie pokladničných dokladov (okruh PD) - zatiaľ len pokladnica kasy.</summary>
    public enum eKod_PDEV_EUD
    {
        P1      // Pokladnica c. 1
    }

    /// <summary>Číselné rady pokladničných dokladov.</summary>
    public enum eKod_PDCR_EUD
    {
        P       // Prijmove pokladnicne doklady
    }
```

- [ ] **Step 1.4: Pridať účet pokladnice do `Constants\Constants.cs`**

Za `public const string UcetSyntetikaOdberatelia = "311";` (riadok ~198) vložiť:

```csharp
    /// <summary>Pokladnica č. 1 - protiúčet (MD) hotovostných inkás z kasy.</summary>
    public const string UcetPokladnica = "211001";
```

- [ ] **Step 1.5: Build JurhanModels**

Run: `dotnet build "C:\Projekty\Private\JurhanProgramy\JurhanModels\JurhanModels\JurhanModels.csproj"`
Expected: Build succeeded, 0 errors.

- [ ] **Step 1.6: Commit (repo JurhanModels)**

```bash
cd "C:\Projekty\Private\JurhanProgramy\JurhanModels"
git add JurhanModels/Import/Enums.cs JurhanModels/Import/Uhrada.cs JurhanModels/Enum/Enum.cs JurhanModels/Constants/Constants.cs
git commit -m "Kasa: typ suboru, sposob platby, datum uhrady na Uhrada, kody PD/P1/P a ucet pokladnice"
```

---

### Task 2: JurhanLib — smerovanie kasy v `Lib.cs` (TDD)

**Files:**
- Test: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.Tests\KasaLibTests.cs` (nový)
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.Tests\ImportLibTests.cs`
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\Lib.cs`

- [ ] **Step 2.1: Napísať zlyhávajúce testy `KasaLibTests.cs`**

```csharp
using JurhanLib.Import;
using JurhanModels.Enum;
using JurhanModels.Import;
using OmegaLib.Enums;
using Xunit;

namespace JurhanLib.Tests
{
    /// <summary>
    /// Smerovanie kasy: hotovosť ide do pokladničných dokladov (PD/P1/P, protiúčet 211001),
    /// karta do interných dokladov SK Pay (ID/IDPB/IDSKP, protiúčet 315102). Zámena okruhu
    /// alebo účtu by zaúčtovala doklady do nesprávnej evidencie.
    /// </summary>
    public class KasaLibTests
    {
        [Fact]
        public void HotovostIdeDoPokladnicnychDokladov()
        {
            var vysledok = Lib.DajKodEvidencieACiselnyRad(
                eTypSuboru.Kasa, eCountry.SK, eSposobPlatbyKasa.Hotovost);

            Assert.Equal(eTYP_OKRUH.PD, vysledok.Item1);
            Assert.Equal("P1", vysledok.Item2);
            Assert.Equal("P", vysledok.Item3);
        }

        [Fact]
        public void KartaIdeDoRaduSkPay()
        {
            var vysledok = Lib.DajKodEvidencieACiselnyRad(
                eTypSuboru.Kasa, eCountry.SK, eSposobPlatbyKasa.Karta);

            Assert.Equal(eTYP_OKRUH.ID, vysledok.Item1);
            Assert.Equal("IDPB", vysledok.Item2);
            Assert.Equal("IDSKP", vysledok.Item3);
        }

        [Fact]
        public void ProtiuctyInkasaKasy()
        {
            Assert.Equal("211001", Lib.DajUcetInkasaKasy(eSposobPlatbyKasa.Hotovost));
            Assert.Equal("315102", Lib.DajUcetInkasaKasy(eSposobPlatbyKasa.Karta));
        }

        [Fact]
        public void KasaJeBezZapoctuBankyXlsxATuzemsko()
        {
            Assert.Equal(eTypRozuctovania.BezZapoctuBanky, Lib.NastavTypRozuctovania(eTypSuboru.Kasa));
            Assert.True(Lib.JeTypSuboruXlsx(eTypSuboru.Kasa));
            Assert.True(Lib.JeTypSuboruKasa(eTypSuboru.Kasa));
            Assert.Equal(eCountry.SK, Lib.DajCountry(eTypSuboru.Kasa, "EUR"));
        }
    }
}
```

- [ ] **Step 2.2: Doplniť `Kasa` do výnimiek `ImportLibTests.cs`**

V `ImportLibTests.cs` do množiny `BezUctu` (riadok ~26) pridať:

```csharp
        private static readonly HashSet<eTypSuboru> BezUctu = new HashSet<eTypSuboru>
        {
            eTypSuboru.Undefined,
            eTypSuboru.Dobropisy,
            // Kasa nejde cez zápočet banky a protiúčty inkás má vlastné (Lib.DajUcetInkasaKasy)
            eTypSuboru.Kasa,
        };
```

- [ ] **Step 2.3: Spustiť testy — musia zlyhať (kompilácia)**

Run: `dotnet test "C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.sln" --filter KasaLibTests`
Expected: FAIL — `Lib` neobsahuje `DajUcetInkasaKasy` / `JeTypSuboruKasa`, overload s `eSposobPlatbyKasa` neexistuje.

- [ ] **Step 2.4: Implementovať zmeny v `Lib.cs`**

a) Signatúra `DajKodEvidencieACiselnyRad` (riadok ~27) dostane voliteľný parameter:

```csharp
        public static Tuple<eTYP_OKRUH, string, string> DajKodEvidencieACiselnyRad(eTypSuboru typSuboru, eCountry country,
            eSposobPlatbyKasa sposobPlatbyKasa = eSposobPlatbyKasa.Ziadna)
```

b) Do switchu (napr. za case `eTypSuboru.Kaufland_EUR_SK`, pred `default`) pridať:

```csharp
                case eTypSuboru.Kasa:
                    {
                        if (sposobPlatbyKasa == eSposobPlatbyKasa.Hotovost)
                        {
                            // hotovost z kasy: prijmovy pokladnicny doklad P1/P
                            typOkruh = eTYP_OKRUH.PD;
                            kodEV = eKod_PDEV_EUD.P1.ToString();
                            kodCR = eKod_PDCR_EUD.P.ToString();
                        }
                        else
                        {
                            // karta z kasy: terminal SK Pay - rovnaky rad ako PlatobnaBrana_SkPay
                            typOkruh = eTYP_OKRUH.ID;
                            kodEV = eKod_IDEV_EUD.IDPB.ToString();
                            kodCR = eKod_IDCR_EUD.IDSKP.ToString();
                        }
                        break;
                    }
```

c) Za metódu `DajUcetPeniazeNaCeste` (končí ~riadok 581) pridať:

```csharp
        /// <summary>
        /// Protiúčet (MD) inkasa dokladu kasy: hotovosť ide do pokladnice (211001), karta na
        /// pohľadávku voči SK Pay (315102) - peniaze prídu neskôr od acquirera na bankový účet.
        /// </summary>
        internal static string DajUcetInkasaKasy(eSposobPlatbyKasa sposobPlatby)
        {
            return sposobPlatby == eSposobPlatbyKasa.Hotovost
                ? Constants.UcetPokladnica
                : Constants.UcetOstatnePohladavkySyntetika + (short)eUcetOstatnePohladavkyAnalytika.SKPay;
        }
```

d) V `NastavTypRozuctovania` (riadok ~632) pridať case do skupiny `BezZapoctuBanky`:

```csharp
                case eTypSuboru.Emag_RON_RO:
                case eTypSuboru.Emag_HUF_HU:
                case eTypSuboru.Kasa:
                    return eTypRozuctovania.BezZapoctuBanky;
```

e) Za `JeTypSuboruDobropis` (riadok ~706) pridať:

```csharp
        public static bool JeTypSuboruKasa(eTypSuboru typSuboru)
        {
            return typSuboru == eTypSuboru.Kasa;
        }
```

f) V `JeTypSuboruXlsx` (riadok ~733) pridať `eTypSuboru.Kasa,` na koniec zoznamu (za `Emag_HUF_HU`). POZOR: do `JeTypSuboruPreConvertCsv` sa Kasa NEPRIDÁVA — parser číta xlsx priamo.

g) V `DajCountry` (riadok ~783) pridať case:

```csharp
                case eTypSuboru.Kasa: return eCountry.SK;
```

- [ ] **Step 2.5: Spustiť testy — musia prejsť (vrátane existujúcich)**

Run: `dotnet test "C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.sln"`
Expected: PASS všetko (KasaLibTests aj ImportLibTests — Kasa je v `BezUctu`, takže poistky na 315000/261000 neudrú).

- [ ] **Step 2.6: Commit (repo JurhanLib)**

```bash
cd "C:\Projekty\Private\JurhanProgramy\JurhanLib"
git add JurhanLib/Import/Lib.cs JurhanLib.Tests/KasaLibTests.cs JurhanLib.Tests/ImportLibTests.cs
git commit -m "Kasa: smerovanie okruhu/radu a protiuctov podla sposobu platby"
```

---

### Task 3: `KasaParser` (TDD)

**Files:**
- Test: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.Tests\KasaParserTests.cs` (nový)
- Create: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\Dopravcovia\KasaParser.cs`

- [ ] **Step 3.1: Napísať zlyhávajúce testy `KasaParserTests.cs`**

Fixture xlsx sa generuje DevExpressom priamo v teste (žiadny binárny súbor v repe):

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using JurhanLib.Import.Dopravcovia;
using JurhanModels.Import;
using OmegaLib.Enums;
using Xunit;

namespace JurhanLib.Tests
{
    /// <summary>
    /// Parser exportu kasy: riadok exportu -> Uhrada s interným číslom faktúry (Referencia),
    /// dátumom úhrady a spôsobom platby. Riadky, ktoré nie sú uhradenou úhradou faktúry alebo
    /// majú nepodporovaný spôsob platby (poukážky/QR/ostatné), sa preskočia a vykážu.
    /// </summary>
    public class KasaParserTests
    {
        private static readonly string[] Hlavicka =
        {
            KasaParser.StlpecDatum, KasaParser.StlpecTyp, KasaParser.StlpecReferencia,
            KasaParser.StlpecHotovost, KasaParser.StlpecKarta,
            "Poukážky", "QR platba", "Ostatné", KasaParser.StlpecStavUhrady
        };

        private static object[] Riadok(DateTime datum, string referencia, decimal hotovost, decimal karta,
            string typ = "10 - úhrada dokladu/faktúry", string stav = "uhradená",
            decimal poukazky = 0, decimal qr = 0, decimal ostatne = 0)
        {
            return new object[] { datum, typ, referencia, hotovost, karta, poukazky, qr, ostatne, stav };
        }

        private static string VytvorExport(string[] hlavicka, params object[][] riadky)
        {
            string subor = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
            using (var workbook = new DevExpress.Spreadsheet.Workbook())
            {
                var sheet = workbook.Worksheets[0];
                for (int c = 0; c < hlavicka.Length; c++)
                {
                    sheet.Cells[0, c].Value = hlavicka[c];
                }
                for (int r = 0; r < riadky.Length; r++)
                {
                    for (int c = 0; c < riadky[r].Length; c++)
                    {
                        object hodnota = riadky[r][c];
                        var bunka = sheet.Cells[r + 1, c];
                        if (hodnota is DateTime datum) bunka.Value = datum;
                        else if (hodnota is decimal suma) bunka.Value = (double)suma;
                        else if (hodnota is double cislo) bunka.Value = cislo;
                        else bunka.Value = (string)hodnota;
                    }
                }
                workbook.SaveDocument(subor, DevExpress.Spreadsheet.DocumentFormat.Xlsx);
            }
            return subor;
        }

        private static List<Uhrada> Parsuj(RozuctovanieContext ctx)
        {
            try
            {
                return new KasaParser(ctx).NacitajUhrady()?.ToList();
            }
            finally
            {
                File.Delete(ctx.fileName);
            }
        }

        private static RozuctovanieContext Kontext(string subor)
        {
            return new RozuctovanieContext { typSuboru = eTypSuboru.Kasa, fileName = subor };
        }

        [Fact]
        public void RozdeliHotovostAKartuDoUhrad()
        {
            var ctx = Kontext(VytvorExport(Hlavicka,
                Riadok(new DateTime(2026, 8, 3, 8, 7, 48), "1202616885", 64.8m, 0),
                Riadok(new DateTime(2026, 8, 4, 14, 35, 49), "1202617033", 0, 96.9m)));

            List<Uhrada> uhrady = Parsuj(ctx);

            Assert.Equal(2, uhrady.Count);
            Assert.Equal("1202616885", uhrady[0].VS);
            Assert.Equal(64.8m, uhrady[0].SumaCM);
            Assert.Equal(64.8m, uhrady[0].SumaTM);
            Assert.Equal("EUR", uhrady[0].Mena);
            Assert.Equal(eSposobPlatbyKasa.Hotovost, uhrady[0].SposobPlatbyKasa);
            Assert.Equal(new DateTime(2026, 8, 3), uhrady[0].DatumUhrady.Value.Date);
            Assert.Equal(eSposobPlatbyKasa.Karta, uhrady[1].SposobPlatbyKasa);
            Assert.Equal(96.9m, uhrady[1].SumaCM);
            Assert.Equal(eCurrency.EUR, ctx.currency);
            Assert.Equal(eCountry.SK, ctx.country);
        }

        [Fact]
        public void DelenaPlatbaVytvoriDveUhrady()
        {
            var ctx = Kontext(VytvorExport(Hlavicka,
                Riadok(new DateTime(2026, 8, 5, 10, 0, 0), "2026080010", 20m, 30m)));

            List<Uhrada> uhrady = Parsuj(ctx);

            Assert.Equal(2, uhrady.Count);
            Assert.Equal(eSposobPlatbyKasa.Hotovost, uhrady[0].SposobPlatbyKasa);
            Assert.Equal(20m, uhrady[0].SumaCM);
            Assert.Equal(eSposobPlatbyKasa.Karta, uhrady[1].SposobPlatbyKasa);
            Assert.Equal(30m, uhrady[1].SumaCM);
            Assert.Equal("2026080010", uhrady[1].VS);
        }

        [Fact]
        public void PreskociIneTypyANeuhradene()
        {
            var ctx = Kontext(VytvorExport(Hlavicka,
                Riadok(new DateTime(2026, 8, 5, 10, 0, 0), "111", 10m, 0, typ: "01 - predaj"),
                Riadok(new DateTime(2026, 8, 5, 11, 0, 0), "222", 10m, 0, stav: "neuhradená")));

            Assert.Empty(Parsuj(ctx));
        }

        [Fact]
        public void PreskociAVykazeNepodporovanySposobPlatby()
        {
            var ctx = Kontext(VytvorExport(Hlavicka,
                Riadok(new DateTime(2026, 8, 5, 10, 0, 0), "333", 0, 0, poukazky: 15m)));

            Assert.Empty(Parsuj(ctx));
            Assert.Single(ctx.nesparovaneVS);
            Assert.Contains("333", ctx.nesparovaneVS[0]);
        }

        [Fact]
        public void ChybajuciStlpecVratiNull()
        {
            // bez stĺpca Referencia - iný súbor, parser nesmie hádať
            string[] bezReferencie = Hlavicka.Where(h => h != KasaParser.StlpecReferencia).ToArray();
            var ctx = Kontext(VytvorExport(bezReferencie,
                new object[] { new DateTime(2026, 8, 5), "10 - úhrada dokladu/faktúry", 10.0, 0.0, 0.0, 0.0, 0.0, "uhradená" }));

            Assert.Null(Parsuj(ctx));
        }
    }
}
```

- [ ] **Step 3.2: Spustiť testy — musia zlyhať**

Run: `dotnet test "C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.sln" --filter KasaParserTests`
Expected: FAIL — `KasaParser` neexistuje.

- [ ] **Step 3.3: Implementovať `KasaParser.cs`**

```csharp
using DevExpress.Spreadsheet;
using JurhanModels.Import;
using OmegaLib.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace JurhanLib.Import.Dopravcovia
{
    /// <summary>
    /// Parser exportu dokladov z registračnej pokladnice ("Doklady *.xlsx"). Spracúva len riadky
    /// typu "10 - úhrada dokladu/faktúry" so stavom "uhradená". Stĺpec Referencia nesie INTERNÉ
    /// číslo faktúry v Omege (C030) - podľa neho sa faktúra páruje. Hotovosť a karta idú do
    /// samostatných úhrad (delená platba = dve úhrady); spôsob platby určuje okruh/rad a protiúčet,
    /// dátum úhrady datuje doklad. Stĺpce sa hľadajú podľa názvov v hlavičke, nie podľa indexov.
    /// </summary>
    public class KasaParser
    {
        internal const string StlpecDatum = "Dátum";
        internal const string StlpecTyp = "Typ";
        internal const string StlpecReferencia = "Referencia";
        internal const string StlpecHotovost = "Hotovosť";
        internal const string StlpecKarta = "Platobné karty";
        internal const string StlpecStavUhrady = "Stav úhrady";
        // nepodporované spôsoby platby - nenulová hodnota znamená riadok, ktorý nevieme zaúčtovať
        internal static readonly string[] NepodporovaneStlpce = { "Poukážky", "QR platba", "Ostatné" };

        private const string KodTypuUhradaFaktury = "10";
        private const string StavUhradena = "uhradená";

        private readonly RozuctovanieContext _ctx;

        public KasaParser(RozuctovanieContext ctx)
        {
            _ctx = ctx;
        }

        public IEnumerable<Uhrada> NacitajUhrady()
        {
            var uhrady = new List<Uhrada>();
            using (Workbook workbook = new Workbook())
            {
                workbook.LoadDocument(_ctx.fileName);
                Worksheet sheet = workbook.Worksheets[0];
                CellRange used = sheet.GetUsedRange();

                Dictionary<string, int> stlpce = NajdiStlpce(sheet, used);
                if (stlpce == null)
                {
                    return null; // chyba uz bola oznamena
                }

                for (int r = used.TopRowIndex + 1; r <= used.BottomRowIndex; r++)
                {
                    SpracujRiadok(sheet, r, stlpce, uhrady);
                }
            }

            _ctx.currency = Lib.DajCurrency("EUR");
            _ctx.country = Lib.DajCountry(_ctx.typSuboru, "EUR");
            _ctx.logger?.Loguj($"Kasa: načítaných {uhrady.Count} úhrad " +
                $"({uhrady.Count(u => u.SposobPlatbyKasa == eSposobPlatbyKasa.Hotovost)} hotovosť, " +
                $"{uhrady.Count(u => u.SposobPlatbyKasa == eSposobPlatbyKasa.Karta)} karta).", true);
            return uhrady;
        }

        private Dictionary<string, int> NajdiStlpce(Worksheet sheet, CellRange used)
        {
            var stlpce = new Dictionary<string, int>();
            for (int c = used.LeftColumnIndex; c <= used.RightColumnIndex; c++)
            {
                string nazov = sheet.Cells[used.TopRowIndex, c].Value.TextValue;
                if (!string.IsNullOrEmpty(nazov) && !stlpce.ContainsKey(nazov))
                {
                    stlpce[nazov] = c;
                }
            }

            IEnumerable<string> pozadovane = new[]
            {
                StlpecDatum, StlpecTyp, StlpecReferencia, StlpecHotovost, StlpecKarta, StlpecStavUhrady
            }.Concat(NepodporovaneStlpce);
            foreach (string nazov in pozadovane)
            {
                if (!stlpce.ContainsKey(nazov))
                {
                    ServicesError.ErrorEnd($"V exporte kasy sa nenašiel stĺpec '{nazov}'. " +
                        $"Skontrolujte, či ide o export dokladov z pokladnice.", _ctx.zobrazenieChyby);
                    return null;
                }
            }
            return stlpce;
        }

        private void SpracujRiadok(Worksheet sheet, int r, Dictionary<string, int> stlpce, List<Uhrada> uhrady)
        {
            string typ = DajText(sheet.Cells[r, stlpce[StlpecTyp]].Value);
            string referencia = DajText(sheet.Cells[r, stlpce[StlpecReferencia]].Value);

            if (string.IsNullOrEmpty(typ) && string.IsNullOrEmpty(referencia))
            {
                return; // prázdny riadok
            }

            // typ je "10 - úhrada dokladu/faktúry" - porovnáva sa kód pred medzerou, nie celý text
            if (typ == null || typ.Split(' ')[0] != KodTypuUhradaFaktury)
            {
                _ctx.logger?.Loguj($"Kasa: doklad '{referencia}' preskakujem - typ '{typ}' nie je úhrada faktúry.", true);
                return;
            }

            string stav = DajText(sheet.Cells[r, stlpce[StlpecStavUhrady]].Value);
            if (!string.Equals(stav, StavUhradena, StringComparison.OrdinalIgnoreCase))
            {
                _ctx.logger?.Loguj($"Kasa: doklad '{referencia}' preskakujem - stav úhrady '{stav}' nie je '{StavUhradena}'.", true);
                return;
            }

            decimal nepodporovane = NepodporovaneStlpce.Sum(s => DajSumu(sheet.Cells[r, stlpce[s]].Value));
            decimal hotovost = DajSumu(sheet.Cells[r, stlpce[StlpecHotovost]].Value);
            decimal karta = DajSumu(sheet.Cells[r, stlpce[StlpecKarta]].Value);
            if (nepodporovane != 0 || hotovost < 0 || karta < 0)
            {
                // poukážky/QR/ostatné nevieme zaúčtovať; záporná suma je storno - to tiež nie
                _ctx.nesparovaneVS.Add($"(kasa) {referencia}: nepodporovaný spôsob platby alebo storno " +
                    $"(hotovosť {hotovost:0.00}, karta {karta:0.00}, poukážky/QR/ostatné {nepodporovane:0.00})");
                _ctx.logger?.Loguj($"Kasa: doklad '{referencia}' preskakujem - nepodporovaný spôsob platby alebo storno.",
                    true, Logger.FarbyLogu.Upozornenie);
                return;
            }

            if (string.IsNullOrEmpty(referencia))
            {
                _ctx.nesparovaneVS.Add($"(kasa) doklad bez referencie - suma {hotovost + karta:0.00}, niet ho k čomu spárovať");
                return;
            }

            CellValue hodnotaDatumu = sheet.Cells[r, stlpce[StlpecDatum]].Value;
            if (!hodnotaDatumu.IsDateTime)
            {
                _ctx.nesparovaneVS.Add($"(kasa) {referencia}: riadok bez dátumu úhrady");
                return;
            }
            DateTime datum = hodnotaDatumu.DateTimeValue;

            PridajUhradu(uhrady, referencia, datum, hotovost, eSposobPlatbyKasa.Hotovost);
            PridajUhradu(uhrady, referencia, datum, karta, eSposobPlatbyKasa.Karta);
        }

        private static void PridajUhradu(List<Uhrada> uhrady, string referencia, DateTime datum,
            decimal suma, eSposobPlatbyKasa sposob)
        {
            if (suma == 0)
            {
                return;
            }
            uhrady.Add(new Uhrada
            {
                VS = referencia,
                SumaCM = suma,
                SumaTM = suma,
                Mena = "EUR",
                DatumUhrady = datum,
                SposobPlatbyKasa = sposob,
            });
        }

        private static string DajText(CellValue hodnota)
        {
            if (hodnota.IsText)
            {
                return hodnota.TextValue?.Trim();
            }
            if (hodnota.IsNumeric)
            {
                // export niekedy uloží referenciu ako číslo - dlhé čísla faktúr nesmú stratiť presnosť
                return ((long)hodnota.NumericValue).ToString();
            }
            return null;
        }

        private static decimal DajSumu(CellValue hodnota)
        {
            return hodnota.IsNumeric ? (decimal)hodnota.NumericValue : 0;
        }
    }
}
```

- [ ] **Step 3.4: Spustiť testy — musia prejsť**

Run: `dotnet test "C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.sln" --filter KasaParserTests`
Expected: PASS 5/5.

- [ ] **Step 3.5: Commit (repo JurhanLib)**

```bash
cd "C:\Projekty\Private\JurhanProgramy\JurhanLib"
git add JurhanLib/Import/Dopravcovia/KasaParser.cs JurhanLib.Tests/KasaParserTests.cs
git commit -m "Kasa: parser exportu dokladov z pokladnice"
```

---

### Task 4: `UhradyToTxt` + `RozuctovanieCore` — účtovanie kasy

**Files:**
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\UhradyToTxt.cs`
- Modify: `C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib\Import\Dopravcovia\RozuctovanieCore.cs`

Bez unit testov — metódy siahajú na databázu Omegy (rovnako ako existujúce typy); overí sa build, existujúce testy a manuálny test v Task 7.

- [ ] **Step 4.1: `UhradyToTxt.cs` — polia pre aktuálnu úhradu**

Za deklaráciu `private readonly string _nazovPriecinka;` (riadok ~49) pridať:

```csharp
        // kasa: kontext aktuálnej úhrady v PridajPolozkySUhradouFaktur - spôsob platby určuje
        // okruh/rad a protiúčet, dátum úhrady datuje doklad (namiesto dátumu faktúry či výpisu)
        private eSposobPlatbyKasa _sposobPlatbyKasa;
        private DateTime? _datumUhrady;
```

- [ ] **Step 4.2: `UhradyToTxt.cs` — naplniť polia v slučke a hľadať faktúru podľa interného čísla**

V `PridajPolozkySUhradouFaktur` (riadok ~267) hneď na začiatok `foreach (var uhrada in uhrady)` (pred kontrolu prázdneho VS) pridať:

```csharp
            foreach (var uhrada in uhrady)
            {
                _sposobPlatbyKasa = uhrada.SposobPlatbyKasa;
                _datumUhrady = uhrada.DatumUhrady;

                if (string.IsNullOrEmpty(uhrada.VS))
```

O kúsok nižšie zmeniť podmienku výberu vyhľadávania (riadok ~288):

```csharp
                // Referencia kasy je interné číslo faktúry (C030), rovnako ako pri Allegro/Kaufland
                if (Lib.JeTypSuboruAllegroAleboKauflandAleboDobropis(_typSuboru) || Lib.JeTypSuboruKasa(_typSuboru))
                {
                    fakturaHlavicka = _eudHlavickaRepository.DajDokladPodlaInternehoCisla(typDokladovFa, uhrada.VS);
                }
```

- [ ] **Step 4.3: `UhradyToTxt.cs` — okruh/rad podľa spôsobu platby**

V `NastavCisloDokladuADatumy` (riadok ~138) doplniť parameter do volania:

```csharp
            Tuple<eTYP_OKRUH, string, string> okruhKodEVKodCR = Lib.DajKodEvidencieACiselnyRad(_typSuboru, _country, _sposobPlatbyKasa);
```

- [ ] **Step 4.4: `UhradyToTxt.cs` — dátum dokladu z úhrady**

Na začiatok `NastavDatumRozuctovania` (riadok ~868) pridať:

```csharp
        private DateTime NastavDatumRozuctovania(EUDHlavicka fakturaHlavicka)
        {
            if (_datumUhrady.HasValue)
            {
                // kasa: doklad inkasa sa datuje dňom úhrady z exportu, nie dátumom faktúry
                return _datumUhrady.Value.Date;
            }
```

- [ ] **Step 4.5: `UhradyToTxt.cs` — protiúčet V-položky**

V `PridajPolozkySUhradou` nahradiť druhú položku (riadok ~528, blok s `MaDatS = Lib.DajUcetOstatnePohladavky...`):

```csharp
                string ucetMD = Lib.JeTypSuboruKasa(_typSuboru)
                    ? Lib.DajUcetInkasaKasy(_sposobPlatbyKasa)
                    : Lib.DajUcetOstatnePohladavky(_typSuboru, _country);

                polozkaTxt = new EUDPolozkaUctovnyZapisTxt
                {
                    SumaCM = sumaCM,
                    SumaTM = sumaTM,
                    TextPolozky = $"{Constants.TextInkaso} {faHlavicka.CisloInterne}",
                    KodTypSumy = DphConstants.KodTypSumyVolnyZaklad,
                    OddielKVDPH = DphConstants.OddielKVDPHBezDph,
                    MaDatS = ucetMD.Substring(0, 3),
                    MaDatA = ucetMD.Substring(3, 3),
                    DalS = string.Empty,
                    DalA = string.Empty,
                };
                polozkyTxt.Add(polozkaTxt);
```

- [ ] **Step 4.6: `UhradyToTxt.cs` — polia pokladničného dokladu**

V `PridajDoklad` (riadok ~211) za volanie `NastavPartnera(faHlavicka);` pridať `NastavPoliaPokladnicnehoDokladu(faHlavicka);` a doplniť metódu (napr. za `NastavPartnera`):

```csharp
        /// <summary>
        /// Hotovostný doklad kasy (okruh PD) má navyše polia pokladničného dokladu: externé
        /// číslo = VS faktúry a "Prijaté od" = partner. Ostatné doklady tieto polia nemajú.
        /// </summary>
        private void NastavPoliaPokladnicnehoDokladu(EUDHlavicka faHlavicka)
        {
            if (_sposobPlatbyKasa != eSposobPlatbyKasa.Hotovost || faHlavicka == null)
            {
                return;
            }
            _eudHlavickaTxt.ExterneCislo = faHlavicka.CisloExterne;
            _eudHlavickaTxt.PrijateOdVyplateneKomu = faHlavicka.MenoFirmy;
        }
```

- [ ] **Step 4.7: `RozuctovanieCore.cs` — vetva parsera a hláška pre prázdny výsledok**

V `NacitajUhradyZoSuboru` (riadok ~177) na začiatok (pred vetvu Kaufland) pridať:

```csharp
            if (Lib.JeTypSuboruKasa(_ctx.typSuboru))
            {
                return new KasaParser(_ctx).NacitajUhrady();
            }
```

V `Execute` do kaskády hlášok pre prázdne úhrady (za vetvu `JeTypSuboruEmag`, riadok ~126) pridať:

```csharp
            else if (Lib.JeTypSuboruKasa(_ctx.typSuboru))
            {
                ServicesError.ErrorEnd($"Zo súboru kasy {_ctx.fileName} nevznikla ani jedna úhrada - " +
                    $"buď neobsahuje uhradené úhrady faktúr (typ 10), alebo majú nepodporovaný spôsob platby.",
                    _ctx.zobrazenieChyby);
            }
```

- [ ] **Step 4.8: Build + celé testy**

Run: `dotnet test "C:\Projekty\Private\JurhanProgramy\JurhanLib\JurhanLib.sln"`
Expected: Build succeeded, všetky testy PASS.

- [ ] **Step 4.9: Commit (repo JurhanLib)**

```bash
cd "C:\Projekty\Private\JurhanProgramy\JurhanLib"
git add JurhanLib/Import/UhradyToTxt.cs JurhanLib/Import/Dopravcovia/RozuctovanieCore.cs
git commit -m "Kasa: uctovanie inkas - okruh/rad a protiucet podla sposobu platby, datum dokladu z uhrady"
```

---

### Task 5: Služba — priečinok Kasa a mesiac

**Files:**
- Modify: `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\FolderMapping.cs`
- Modify: `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\RozuctovanieEmailov.cs`

- [ ] **Step 5.1: `FolderMapping.cs` — mapovanie priečinka**

Za `case "Emag HU": ...` (riadok ~57) pridať:

```csharp
                // export dokladov z registracnej pokladnice (hotovost + karty)
                case "Kasa": return eTypSuboru.Kasa;
```

(Presný názov IMAP priečinka podľa reálnej schránky — ak sa bude volať inak, upraví sa case.)

- [ ] **Step 5.2: `RozuctovanieEmailov.cs` — kasa nepotrebuje mesiac z názvu súboru**

V `SpracujSubor` (riadok ~642) upraviť podmienku:

```csharp
            short mesiac = NazovSuboru.DajMesiac(Path.GetFileNameWithoutExtension(filePath));
            // kasa ma datum uhrady v kazdom riadku exportu - mesiac z nazvu suboru nepotrebuje
            if (Lib.NastavTypRozuctovania(typSuboru) == eTypRozuctovania.BezZapoctuBanky && mesiac == 0
                && !Lib.JeTypSuboruKasa(typSuboru))
```

POZOR: priečinok "Kasa" sa NEPRIDÁVA do `_pilotnePriecinky` — nový typ beží najprv v simulácii (ostrý import sa zapne po overení, samostatný commit neskôr).

- [ ] **Step 5.3: Build služby**

Run: `dotnet build "C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov.csproj"`
Expected: Build succeeded, 0 errors.

- [ ] **Step 5.4: Commit (repo služby)**

```bash
cd "C:\Projekty\Private\JurhanService\RozuctovanieDopravcov"
git add JurhanService_RozuctovanieDopravcov/FolderMapping.cs JurhanService_RozuctovanieDopravcov/RozuctovanieEmailov.cs
git commit -m "Kasa: priecinok schranky a vynimka z mesiaca v nazve suboru (datumy su v riadkoch)"
```

---

### Task 6: Program — frmVyber

**Files:**
- Modify: `C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\frmVyber.cs`

- [ ] **Step 6.1: Viditeľnosť polí pre kasu**

V `cmbTypSuboru_SelectedIndexChanged` (riadok ~197) upraviť blok logiky viditeľnosti:

```csharp
                bool jeEmag = Lib.JeTypSuboruEmag(Program.typSuboru);
                bool jeAllegroKauflandDobropis = Lib.JeTypSuboruAllegroAleboKauflandAleboDobropis(Program.typSuboru);
                // Packeta: výpis sa sťahuje z jej API, súbor sa z disku nevyberá
                bool jeZApi = Lib.JeTypSuboruZApi(Program.typSuboru);
                // kasa: súbor áno, ale bez mesiaca (dátumy sú v riadkoch) a bez bankového dokladu
                bool jeKasa = Lib.JeTypSuboruKasa(Program.typSuboru);

                bool visibleMesiac = jeAllegroKauflandDobropis || jeEmag;
                bool jeSoZapoctomBanky = !jeAllegroKauflandDobropis && !jeEmag && !jeKasa;
                // Packeta: jeden beh rozúčtuje viac faktúr, každá s vlastným prevodom - doklad sa
                // nezadáva, jadro si ho k každému výpisu nájde samo
                bool visibleBankovyDoklad = jeSoZapoctomBanky && !jeZApi;
                bool visibleSubor = (jeSoZapoctomBanky || jeEmag || jeKasa) && !jeZApi;
```

(Zvyšok metódy — priradenia `.Visible` — ostáva bez zmeny. Filter súboru netreba meniť: `btnFileImport_Click` používa `Lib.JeTypSuboruXlsx`, kam Kasa pribudla v Task 2. Combo text „Kasa" vznikne automaticky z enumu.)

- [ ] **Step 6.2: Build programu**

Run: `dotnet build "C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov\RozuctovanieDopravcov\RozuctovanieDopravcov.csproj"`
Expected: Build succeeded, 0 errors.

- [ ] **Step 6.3: Commit (repo programu)**

```bash
cd "C:\Projekty\Private\JurhanProgramy\RozuctovanieDopravcov"
git add RozuctovanieDopravcov/frmVyber.cs
git commit -m "Kasa: vyber typu vo formulari - subor xlsx, bez mesiaca a bankoveho dokladu"
```

---

### Task 7: Manuálne overenie na testovacej Omege (bez kódu)

Okruh PD (TypOkruh 160) sa cez TXT autoimport nikdy nepoužil — pred nasadením treba overiť na kópii databázy zákazníka:

- [ ] **Step 7.1:** Spustiť program (frmVyber) nad kópiou Omegy zákazníka s reálnym exportom `Doklady 2026-08-28 11-34-01.xlsx`.
- [ ] **Step 7.2:** Skontrolovať hotovostný doklad v Omege proti vzoru zákazníka: okruh PD, evidencia P1, rad P, dátum = dátum úhrady, MD 211001 / DAL 311001, text „Inkaso faktúry …", partner, pole „Prijaté od", externé číslo = VS faktúry, úhrada spárovaná na faktúre (záložka Úhrady). Ak autoimport PD doklad odmietne alebo niektoré pole nenaplní, doplniť polia `EUDHlavickaTxt` podľa protokolu importu (kandidáti: `KS`, `SS`, `InterneCisloUhradzanehoDokladu`).
- [ ] **Step 7.3:** Skontrolovať kartový doklad: IDPB/IDSKP, MD 315102 / DAL 311001, dátum = dátum úhrady, úhrada spárovaná.
- [ ] **Step 7.4:** Predkontácia: porovnať naimportované doklady so vzormi (hotovosť „I-1 PD/BV – Inkaso OF", karta „Z-Al"). Ak import predkontáciu nenastaví a zákazníkovi to prekáža, doriešiť dodatočným UPDATE ako pri 3M úhradách (`T041_EUD_Polozky.C100_PredkontaciaKod`) — samostatná úloha.
- [ ] **Step 7.5:** Pustiť ten istý súbor druhý raz — musí skončiť ako `Duplicita` (kľúč súboru); premenovaný súbor s rovnakými dokladmi musí skončiť `VsetkoUzUhradene` (faktúry už uhradené), bez nových dokladov.
- [ ] **Step 7.6:** Overiť v Omege zákazníka, že číselný rad P1/P je založený pre rok 2026 (na screenshote existuje — doklad P-0283).
- [ ] **Step 7.7:** Po úspešnom overení: pridať priečinok kasy do `_pilotnePriecinky` v `RozuctovanieEmailov.cs` (presný `FullName` podľa schránky) — samostatný commit v repe služby.

Doplnené zo záverečného code review (29.08.2026):

- [ ] **Step 7.8:** Otestovať čiastočnú úhradu (kasa zaplatí menej než suma faktúry) — overiť, že Omega spáruje čiastočne a faktúra ostane čiastočne uhradená.
- [ ] **Step 7.9:** Druhý beh nad prekrývajúcim sa súborom spustiť až PO ostrom importe prvého (v simulácii sa prekryv neodchytí — C220_Uhradene sa nemení, pilotný log môže tie isté doklady „zaúčtovať" viackrát).
- [ ] **Step 7.10:** OTVORENÉ ROZHODNUTIE pred pilotom: email, ktorého export obsahuje LEN nepodporované platby (poukážky/QR/storná), skončí ako `ZiadneUhrady` a ostane v priečinku navždy — služba ho spracuje pri každom behu znova. Rozhodnúť: ručný presun do „Zaúčtované", alebo počítať vykázané nepodporované riadky ako vybavené (úprava kódu).
- [ ] **Step 7.11:** Overiť, že reálny export má hlavičku v prvom použitom riadku hárku (titulný riadok nad hlavičkou by skončil bezpečnou chybou „nenašiel sa stĺpec").
