# Hľadanie bankového dokladu podľa textu hlavičky – implementačný plán

> **Pre agentných workerov:** POVINNÝ SUB-SKILL: použi superpowers:subagent-driven-development (odporúčané) alebo superpowers:executing-plans na implementáciu po úlohách. Kroky používajú checkbox (`- [ ]`) syntax.

**Cieľ:** V službe RozuctovanieDopravcov hľadať bankový doklad podľa textu hlavičky (`C099_VolneDefinovanyText`) obsahujúceho názov priečinka a podľa dátumu vystavenia v rozpätí ±2 dni od dátumu výpisu zisteného zo súboru.

**Architektúra:** Nová čistá trieda `DatumVypisu` (JurhanLib) zisťuje dátum výpisu z obsahu súboru (konštantný dátumový stĺpec → názov súboru → dnešok). `DopravcaToUhrady` ju zavolá a uloží dátum do kontextu. `RozuctovanieCore` sa preusporiada tak, aby parsovanie prebehlo pred hľadaním dokladu, a hľadá doklad novou logikou cez nový repo metódu v `OmegaLib`. Služba doplní názov priečinka do kontextu.

**Tech stack:** C# / .NET 10 (net10.0-windows), Kros.KORM, xUnit (nový testovací projekt).

**Repozitáre (3 samostatné git repo):**
- `JurhanLib` — `C:/Projekty/Private/JurhanProgramy/JurhanLib` (vetva `net10`)
- `OmegaLib` — `C:/Projekty/Private/OmegaLib` (vetva `master` → vytvoriť feature vetvu)
- `RozuctovanieDopravcov` — worktree `…/worktrees/package-numbers-logs-3f6338` (vetva `claude/bank-document-text-search-006264`)

**Štruktúra súborov:**
- Vytvoriť: `JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/DatumVypisu.cs` — čistá logika zistenia dátumu výpisu.
- Vytvoriť: `JurhanProgramy/JurhanLib/JurhanLib.Tests/` — xUnit projekt + testy `DatumVypisu`.
- Zmeniť: `JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/RozuctovanieContext.cs` — nové polia.
- Zmeniť: `JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/DopravcaToUhrady.cs` — výpočet dátumu výpisu.
- Zmeniť: `JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/RozuctovanieCore.cs` — preusporiadanie + nová logika hľadania.
- Zmeniť: `OmegaLib/OmegaLib/Repository/EudHlavickaRepository.cs` — nová metóda.
- Zmeniť: `JurhanService_RozuctovanieDopravcov/RozuctovanieEmailov.cs` (worktree) — nastavenie názvu priečinka.

---

## Task 1: Nová repo metóda v OmegaLib (hľadanie podľa textu hlavičky)

**Files:**
- Modify: `C:/Projekty/Private/OmegaLib/OmegaLib/Repository/EudHlavickaRepository.cs`

- [ ] **Step 1: Vytvoriť feature vetvu v OmegaLib (je na `master`)**

```bash
git -C "C:/Projekty/Private/OmegaLib" checkout -b feature/doklad-podla-textu-hlavicky
```

- [ ] **Step 2: Pridať metódu `DajDokladyPodlaTextuHlavicky`**

Za existujúcu metódu `DajDokladPodlaPoznamky` (riadok 34) doplniť:

```csharp
public IEnumerable<EUDHlavicka> DajDokladyPodlaTextuHlavicky(string text, bool dbJeSQl) =>
    DajDoklady($"{ServicesDatabase.DajStringZMema("C099_VolneDefinovanyText", dbJeSQl)} LIKE @1", "%" + text + "%");
```

Poznámka: `DajStringZMema` vyprodukuje `CAST (C099_VolneDefinovanyText as VARCHAR(255))` (SQL) resp. `CSTR(IIF(...))` (Access), rovnako ako existujúce použitie pri `C097`. `using OmegaLib.Services;` už v súbore je.

- [ ] **Step 3: Overiť build OmegaLib**

Run: `dotnet build "C:/Projekty/Private/OmegaLib/OmegaLib/OmegaLib.csproj" -c Debug`
Expected: Build succeeded (0 Error).

- [ ] **Step 4: Commit**

```bash
git -C "C:/Projekty/Private/OmegaLib" add OmegaLib/Repository/EudHlavickaRepository.cs
git -C "C:/Projekty/Private/OmegaLib" commit -m "feat: DajDokladyPodlaTextuHlavicky - hladanie EUD podla textu C099 (LIKE)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 2: Čistá logika `DatumVypisu` v JurhanLib (TDD s xUnit)

**Files:**
- Create: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj`
- Create: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/DatumVypisuTests.cs`
- Create: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/DatumVypisu.cs`

- [ ] **Step 1: Vytvoriť xUnit testovací projekt**

Vytvoriť `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0-windows</TargetFramework>
    <ImplicitUsings>disable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <IsPackable>false</IsPackable>
    <NoWarn>NU1701;CS0618;CA1416</NoWarn>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.12.0" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\JurhanLib\JurhanLib.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 2: Napísať padajúce testy**

Vytvoriť `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/DatumVypisuTests.cs`:

```csharp
using System;
using System.Collections.Generic;
using JurhanLib.Import.Dopravcovia;
using Xunit;

namespace JurhanLib.Tests
{
    public class DatumVypisuTests
    {
        [Theory]
        [InlineData("2026-07-13", 2026, 7, 13)]
        [InlineData("2026.07.13.", 2026, 7, 13)]   // GLS - koncová bodka
        [InlineData("14.07.2026", 2026, 7, 14)]
        [InlineData("2026-05-17 07:21:34", 2026, 5, 17)] // GoPay - dátum+čas
        [InlineData("\"2026-07-13\"", 2026, 7, 13)]
        public void SkusParsniDatum_platneFormaty(string vstup, int r, int m, int d)
        {
            DateTime? vysledok = DatumVypisu.SkusParsniDatum(vstup);
            Assert.Equal(new DateTime(r, m, d), vysledok);
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("nieje datum")]
        [InlineData("12345678")]
        public void SkusParsniDatum_neplatne_vratiNull(string vstup)
        {
            Assert.Null(DatumVypisu.SkusParsniDatum(vstup));
        }

        [Fact]
        public void ZKonstantnehoStlpca_viacKonstantnych_vratiNajnovsi()
        {
            // GLS: "Dátum doruč." konšt. 08.07, "Dátum prevodu" konšt. 13.07, "Suma" sa mení
            var records = new List<string[]>
            {
                new[] { "Journal No.", "Dátum doruč.", "Suma", "Dátum prevodu" },
                new[] { "122853008", "2026.07.08.", "3530", "2026.07.13." },
                new[] { "122853959", "2026.07.08.", "1310", "2026.07.13." },
                new[] { "5450968329/0800", "", "182434", "" }, // sumárny riadok - prázdne
            };
            DateTime? d = DatumVypisu.ZKonstantnehoDatumovehoStlpca(records, new[] { "Journal No." });
            Assert.Equal(new DateTime(2026, 7, 13), d);
        }

        [Fact]
        public void ZKonstantnehoStlpca_jedenKonstantny_vratiHo()
        {
            // DPD SK: payment_gen_date konšt., delivery_date sa mení
            var records = new List<string[]>
            {
                new[] { "parcelno", "delivery_date", "payment_gen_date" },
                new[] { "0654...21", "06.07.2026", "13.07.2026" },
                new[] { "0654...66", "07.07.2026", "13.07.2026" },
            };
            DateTime? d = DatumVypisu.ZKonstantnehoDatumovehoStlpca(records, new[] { "parcelno" });
            Assert.Equal(new DateTime(2026, 7, 13), d);
        }

        [Fact]
        public void ZKonstantnehoStlpca_ziadnyKonstantnyDatum_vratiNull()
        {
            // DPD CZ: len delivery_date, ktorý sa mení
            var records = new List<string[]>
            {
                new[] { "pl_number", "amount", "delivery_date" },
                new[] { "0654...28", "3105", "2026-07-07" },
                new[] { "0654...52", "925", "2026-06-26" },
            };
            DateTime? d = DatumVypisu.ZKonstantnehoDatumovehoStlpca(records, new[] { "pl_number" });
            Assert.Null(d);
        }

        [Fact]
        public void ZKonstantnehoStlpca_hlavickaNenajdena_vratiNull()
        {
            var records = new List<string[]>
            {
                new[] { "nieco ine", "x" },
                new[] { "a", "2026-07-13" },
            };
            Assert.Null(DatumVypisu.ZKonstantnehoDatumovehoStlpca(records, new[] { "parcelno" }));
        }

        [Fact]
        public void ZNazvuSuboru_dpdVzor_najdeDatum()
        {
            DateTime? d = DatumVypisu.ZNazvuSuboru(
                "2-51147584360-cod_wire_transfer_email_csv-2026_07_14-142042-6282.csv");
            Assert.Equal(new DateTime(2026, 7, 14), d);
        }

        [Fact]
        public void ZNazvuSuboru_bezDatumu_vratiNull()
        {
            Assert.Null(DatumVypisu.ZNazvuSuboru("1244127042.csv"));
        }

        [Fact]
        public void Zisti_prioritaKonstantnyStlpec()
        {
            var records = new List<string[]>
            {
                new[] { "parcelno", "payment_gen_date" },
                new[] { "x", "13.07.2026" },
            };
            DateTime v = DatumVypisu.Zisti(records, new[] { "parcelno" },
                "ignoruje_sa-2026_01_01.csv", new DateTime(2026, 12, 31));
            Assert.Equal(new DateTime(2026, 7, 13), v);
        }

        [Fact]
        public void Zisti_fallbackNazovSuboru()
        {
            var records = new List<string[]>
            {
                new[] { "pl_number", "delivery_date" },
                new[] { "x", "2026-07-07" },
                new[] { "y", "2026-06-26" },
            };
            DateTime v = DatumVypisu.Zisti(records, new[] { "pl_number" },
                "…-2026_07_14-142042.csv", new DateTime(2026, 12, 31));
            Assert.Equal(new DateTime(2026, 7, 14), v);
        }

        [Fact]
        public void Zisti_fallbackDnesnyDatum()
        {
            var records = new List<string[]>
            {
                new[] { "Váš e-shop", "suma" },
                new[] { "x", "10" },
            };
            DateTime dnes = new DateTime(2026, 7, 19);
            DateTime v = DatumVypisu.Zisti(records, new[] { "Váš e-shop" }, "1244127042.csv", dnes);
            Assert.Equal(dnes, v);
        }
    }
}
```
- [ ] **Step 3: Spustiť testy – musia padnúť (trieda ešte neexistuje)**

Run: `dotnet test "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj"`
Expected: FAIL – kompilácia zlyhá, `DatumVypisu` neexistuje.

- [ ] **Step 4: Implementovať `DatumVypisu`**

Vytvoriť `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/DatumVypisu.cs`:

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace JurhanLib.Import.Dopravcovia
{
    /// <summary>
    /// Zisťuje „dátum výpisu" pre rozúčtovanie so zápočtom banky.
    /// Poradie: konštantný dátumový stĺpec v obsahu → dátum z názvu súboru → dnešný dátum.
    /// </summary>
    public static class DatumVypisu
    {
        // formáty pozorované v stĺpcoch dopravcov (bez 8-ciferných, aby sa ID nepomýlilo s dátumom)
        private static readonly string[] _formatyStlpca =
        {
            "yyyy-MM-dd", "yyyy.MM.dd", "yyyy/MM/dd", "dd.MM.yyyy", "d.M.yyyy"
        };

        public static DateTime Zisti(IList<string[]> records, IEnumerable<string> nazvyPrvychStlpcov,
            string fileName, DateTime dnes)
        {
            DateTime? zStlpca = ZKonstantnehoDatumovehoStlpca(records, nazvyPrvychStlpcov);
            if (zStlpca.HasValue)
            {
                return zStlpca.Value.Date;
            }

            DateTime? zNazvu = ZNazvuSuboru(fileName);
            if (zNazvu.HasValue)
            {
                return zNazvu.Value.Date;
            }

            return dnes.Date;
        }

        public static DateTime? ZKonstantnehoDatumovehoStlpca(IList<string[]> records,
            IEnumerable<string> nazvyPrvychStlpcov)
        {
            if (records == null)
            {
                return null;
            }

            int indexHlavicky = -1;
            for (int i = 0; i < records.Count; i++)
            {
                if (records[i].Length > 0 && nazvyPrvychStlpcov.Contains(records[i][0]))
                {
                    indexHlavicky = i;
                    break;
                }
            }
            if (indexHlavicky < 0)
            {
                return null;
            }

            int pocetStlpcov = records[indexHlavicky].Length;
            var datoveRiadky = records.Skip(indexHlavicky + 1).ToList();

            var konstantneDatumy = new List<DateTime>();
            for (int c = 0; c < pocetStlpcov; c++)
            {
                int stlpec = c;
                var neprazdne = datoveRiadky
                    .Where(r => r.Length > stlpec && !string.IsNullOrWhiteSpace(r[stlpec]))
                    .Select(r => r[stlpec])
                    .ToList();
                if (neprazdne.Count == 0)
                {
                    continue;
                }

                var datumy = neprazdne.Select(SkusParsniDatum).ToList();
                if (datumy.Any(d => !d.HasValue))
                {
                    continue; // stĺpec nie je čisto dátumový
                }

                var rozne = datumy.Select(d => d.Value.Date).Distinct().ToList();
                if (rozne.Count == 1)
                {
                    konstantneDatumy.Add(rozne[0]);
                }
            }

            if (konstantneDatumy.Count == 0)
            {
                return null;
            }
            return konstantneDatumy.Max(); // pri viacerých konštantných → najnovší
        }

        public static DateTime? ZNazvuSuboru(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                return null;
            }
            string nazov = Path.GetFileNameWithoutExtension(fileName);
            // RRRR[-_.]MM[-_.]DD (napr. DPD "2026_07_14")
            Match m = Regex.Match(nazov,
                @"(?<![0-9])(20[0-9]{2})[-_.](0[1-9]|1[0-2])[-_.](0[1-9]|[12][0-9]|3[01])(?![0-9])");
            if (m.Success)
            {
                return new DateTime(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value),
                    int.Parse(m.Groups[3].Value));
            }
            return null;
        }

        public static DateTime? SkusParsniDatum(string hodnota)
        {
            if (string.IsNullOrWhiteSpace(hodnota))
            {
                return null;
            }
            string s = hodnota.Trim().Trim('"').Trim();
            int medzera = s.IndexOf(' ');
            if (medzera > 0)
            {
                s = s.Substring(0, medzera); // odseknúť čas ("2026-05-17 07:21:34")
            }
            s = s.TrimEnd('.'); // GLS "2026.07.13."
            if (DateTime.TryParseExact(s, _formatyStlpca, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out DateTime dt))
            {
                return dt;
            }
            return null;
        }
    }
}
```

- [ ] **Step 5: Spustiť testy – musia prejsť**

Run: `dotnet test "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj"`
Expected: PASS (všetky testy zelené).

- [ ] **Step 6: Pridať testovací projekt do JurhanLib.sln (voliteľné, kvôli VS)**

Run: `dotnet sln "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.sln" add "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj"`
Expected: Project added.

- [ ] **Step 7: Commit (JurhanLib repo)**

```bash
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" add JurhanLib/Import/Dopravcovia/DatumVypisu.cs JurhanLib.Tests JurhanLib.sln
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" commit -m "feat: DatumVypisu - zistenie datumu vypisu (konstantny stlpec / nazov suboru / dnes) + testy

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 3: Nové polia v `RozuctovanieContext`

**Files:**
- Modify: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/RozuctovanieContext.cs`

- [ ] **Step 1: Pridať polia `nazovPriecinka` a `datumVypisu`**

Za pole `public string interneCislo;` (riadok 30) doplniť:

```csharp
        /// <summary>Názov IMAP priečinka dopravcu (napr. „DPD SK"); v službe podľa neho hľadáme text hlavičky dokladu.</summary>
        public string nazovPriecinka;
        /// <summary>Dátum výpisu zistený zo súboru (naplní parser); použije sa pri hľadaní bankového dokladu.</summary>
        public DateTime? datumVypisu;
```

Doplniť `using System;` na začiatok súboru, ak tam nie je.

- [ ] **Step 2: Overiť build**

Run: `dotnet build "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/JurhanLib.csproj" -c Debug`
Expected: Build succeeded.

- [ ] **Step 3: Commit**

```bash
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" add JurhanLib/Import/Dopravcovia/RozuctovanieContext.cs
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" commit -m "feat: RozuctovanieContext - polia nazovPriecinka a datumVypisu

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 4: Výpočet dátumu výpisu v `DopravcaToUhrady`

**Files:**
- Modify: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/DopravcaToUhrady.cs`

- [ ] **Step 1: Naplniť `_ctx.datumVypisu` po úspešnej kontrole typu súboru**

V metóde `NacitajUhrady` (preťaženie s `oddelovac`), v `else` vetve hneď za `KontrolaNaSpravnyTypSuboru(...)` (aktuálne riadky 43–45, pred výpočtom `maxIndex`) vložiť:

```csharp
                _ctx.datumVypisu = DatumVypisu.Zisti(records, nazvyPrvychStlpcov, _ctx.fileName, DateTime.Today);
```

Výsledný začiatok `else` vetvy:

```csharp
            else
            {
                _ctx.datumVypisu = DatumVypisu.Zisti(records, nazvyPrvychStlpcov, _ctx.fileName, DateTime.Today);

                var maxIndex = Math.Max(indexBalik, Math.Max(indexVS, Math.Max(indexSumaCM1, Math.Max(indexSumaCM2, indexMena))));
                foreach (var record in records)
                {
```

`using System;` v súbore už je (riadok 6). `DatumVypisu` je v rovnakom namespace `JurhanLib.Import.Dopravcovia`.

- [ ] **Step 2: Overiť build**

Run: `dotnet build "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/JurhanLib.csproj" -c Debug`
Expected: Build succeeded.

- [ ] **Step 3: Commit**

```bash
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" add JurhanLib/Import/Dopravcovia/DopravcaToUhrady.cs
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" commit -m "feat: DopravcaToUhrady - vypocet datumu vypisu zo suboru do kontextu

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 5: Preusporiadanie a nová logika hľadania v `RozuctovanieCore`

**Files:**
- Modify: `C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/Import/Dopravcovia/RozuctovanieCore.cs`

- [ ] **Step 1: Preusporiadať `Execute` (parsovanie pred hľadaním dokladu)**

Nahradiť blok od `_ctx.typRozuctovania = Lib.NastavTypRozuctovania(_ctx.typSuboru);` (riadok 44) po `return eVysledokRozuctovania.ZiadneUhrady;` (riadok 106) týmto:

```csharp
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
            }

            var uhrady = NacitajUhradyZoSuboru();
            if (uhrady == null)
            {
                // chyba uz bola oznamena pri nacitavani suboru
                return eVysledokRozuctovania.Chyba;
            }

            if (uhrady.Any())
            {
                if (_ctx.typRozuctovania == eTypRozuctovania.SoZapoctomBanky)
                {
                    // datum vypisu uz naplnil parser (DopravcaToUhrady) do _ctx.datumVypisu
                    _ctx.dokladPreRozuctovanie = DajDokladPreRozuctovanie();
                    if (_ctx.dokladPreRozuctovanie == null)
                    {
                        // chybu uz oznamil DajDokladPreRozuctovanie
                        return eVysledokRozuctovania.NenajdenyBankovyDoklad;
                    }
                }
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
```

Poznámka k správaniu: pri type SoZapoctomBanky s nulou úhrad sa teraz vráti `ZiadneUhrady` (predtým mohlo hľadanie dokladu vrátiť `NenajdenyBankovyDoklad` ešte pred parsovaním). To je zámer — hlási sa reálna príčina.

- [ ] **Step 2: Prepísať `DajDokladPreRozuctovanie` (vetva služby = podľa textu hlavičky + dátumu)**

Nahradiť celú metódu `DajDokladPreRozuctovanie` (riadky 254–278) týmto:

```csharp
        private EUDHlavicka DajDokladPreRozuctovanie()
        {
            if (_ctx.interneCislo != null)
            {
                // UI: doklad podľa interného čísla (nezmenené)
                EUDHlavicka doklad = _eudHlavickaRepository.DajDokladPodlaInternehoCisla(_ctx.interneCislo);
                if (doklad == null)
                {
                    ServicesError.ErrorEnd($"Doklad s interným číslom {_ctx.interneCislo} neexistuje v evidencii !",
                        _ctx.zobrazenieChyby);
                }
                return doklad;
            }

            // sluzba: doklad podla textu uctovneho zapisu (C099) obsahujuceho nazov priecinka
            // a datumu vystavenia v rozpati +-2 dni od datumu vypisu
            if (string.IsNullOrEmpty(_ctx.nazovPriecinka))
            {
                ServicesError.ErrorEnd("Pre hľadanie bankového dokladu chýba názov priečinka dopravcu !",
                    _ctx.zobrazenieChyby);
                return null;
            }

            DateTime datumVypisu = (_ctx.datumVypisu ?? DateTime.Today).Date;
            var kandidati = _eudHlavickaRepository.DajDokladyPodlaTextuHlavicky(
                _ctx.nazovPriecinka, _ctx.pripojeneFirmy.DbJeSQL());

            EUDHlavicka najblizsi = kandidati
                .Where(d => d.RokVystavenia != 0
                            && Math.Abs((d.DatumVystavenia.Date - datumVypisu).TotalDays) <= 2)
                .OrderBy(d => Math.Abs((d.DatumVystavenia.Date - datumVypisu).TotalDays))
                .ThenByDescending(d => d.DatumVystavenia)
                .FirstOrDefault();

            if (najblizsi == null)
            {
                ServicesError.ErrorEnd($"Nenašiel sa bankový doklad pre priečinok '{_ctx.nazovPriecinka}' " +
                    $"s dátumom výpisu {datumVypisu:dd.MM.yyyy} (±2 dni) pre rozúčtovanie súboru {_ctx.fileName} !",
                    _ctx.zobrazenieChyby);
            }
            return najblizsi;
        }
```

`using System;` a `using System.Linq;` sú v súbore už (riadky 9, 12).

- [ ] **Step 3: Overiť build**

Run: `dotnet build "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib/JurhanLib.csproj" -c Debug`
Expected: Build succeeded.

- [ ] **Step 4: Spustiť testy (regresia sa nedotkla čistej logiky)**

Run: `dotnet test "C:/Projekty/Private/JurhanProgramy/JurhanLib/JurhanLib.Tests/JurhanLib.Tests.csproj"`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" add JurhanLib/Import/Dopravcovia/RozuctovanieCore.cs
git -C "C:/Projekty/Private/JurhanProgramy/JurhanLib" commit -m "feat: RozuctovanieCore - hladanie bankoveho dokladu podla textu hlavicky a datumu vypisu

Parsovanie suboru sa presunulo pred hladanie dokladu. Vetva sluzby hlada doklad
podla C099 (obsahuje nazov priecinka) a DatumVystavenia v rozpati +-2 dni od datumu
vypisu; pri viacerych zhodach vybera najblizsi. Vetva UI (interne cislo) nezmenena.

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 6: Prenos názvu priečinka v službe (`RozuctovanieEmailov`)

**Files:**
- Modify: `…/worktrees/package-numbers-logs-3f6338/JurhanService_RozuctovanieDopravcov/RozuctovanieEmailov.cs`

- [ ] **Step 1: Rozšíriť reťazec volaní o názov priečinka**

`SpracujPriecinok` (riadok 82) — odovzdať `folder.Name` do `SpracujEmail`:

```csharp
                    if (SpracujEmail(message, typSuboru, folder.Name))
```

`SpracujEmail` (riadok 133) — zmeniť signatúru a odovzdať ďalej:

```csharp
        private bool SpracujEmail(MimeMessage message, eTypSuboru typSuboru, string nazovPriecinka)
```

a v tele (riadok 150) volať:

```csharp
                eVysledokRozuctovania vysledok = SpracujSubor(filePath, typSuboru, nazovPriecinka);
```

`SpracujSubor` (riadok 163) — zmeniť signatúru:

```csharp
        private eVysledokRozuctovania SpracujSubor(string filePath, eTypSuboru typSuboru, string nazovPriecinka)
```

- [ ] **Step 2: Nastaviť `nazovPriecinka` do kontextu**

V `SpracujSubor`, v inicializácii `RozuctovanieContext` (riadky 177–186) doplniť pole a spresniť komentár pri `interneCislo`:

```csharp
            RozuctovanieContext ctx = new RozuctovanieContext
            {
                pripojeneFirmy = _pripojeneFirmy,
                typSuboru = typSuboru,
                fileName = filePath,
                mesiac = mesiac,
                interneCislo = null, // sluzba: doklad sa hlada podla textu hlavicky (C099) a datumu vypisu
                nazovPriecinka = nazovPriecinka,
                zobrazenieChyby = eZobrazenieChyby.ZapisDoSuboru,
                typSpustenia = Program.typSpustenia,
            };
```

- [ ] **Step 3: Overiť build služby**

Run: `dotnet build "C:/Projekty/Private/JurhanService/RozuctovanieDopravcov/JurhanService_RozuctovanieDopravcov/JurhanService_RozuctovanieDopravcov.csproj" -c Debug`
Expected: Build succeeded.

Poznámka: build služby vyžaduje ProjectReference na OmegaLib/JurhanLib z hlavného checkoutu (nie z worktree). Ak build z worktree zlyhá na nerozbalených referenciách, spusti build z hlavného checkoutu služby.

- [ ] **Step 4: Commit (worktree, service repo)**

```bash
cd "C:/Projekty/Private/JurhanService/RozuctovanieDopravcov/JurhanService_RozuctovanieDopravcov/.claude/worktrees/package-numbers-logs-3f6338"
git add JurhanService_RozuctovanieDopravcov/RozuctovanieEmailov.cs
git commit -m "feat: RozuctovanieEmailov - prenos nazvu priecinka do kontextu rozuctovania

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

## Task 7: Manuálne overenie na reálnych súboroch

**Files:** žiadne (overenie behu).

- [ ] **Step 1: Overiť výber dátumu výpisu na vzorkách z Log**

Skontrolovať (napr. dočasným logom alebo debug behom), že `DatumVypisu.Zisti` vráti pre reálne súbory:
- GLS `50020890_CZK_…` → 2026-07-13 (stĺpec „Dátum prevodu")
- DPD SK `sales_company…`/`parcelno` → 2026-07-13 („payment_gen_date")
- DPD HU `cod_summary_report…`/`Parcel number` → 2026-07-14 („Reconciliation Date")
- SPS/Express One `TR260714…`/`VS` → 2026-07-14 („DATUM")
- DPD CZ `pl_number` → 2026-07-14 (z názvu súboru, fallback)
- Packeta `Váš e-shop` bez dátumu → dnešný dátum (fallback)

- [ ] **Step 2: Overiť hľadanie dokladu**

Na testovacej DB s bankovým dokladom, ktorý má v `C099_VolneDefinovanyText` názov priečinka (napr. „GLS") a `DatumVystavenia` v rozpätí ±2 dni, spustiť službu a overiť, že sa doklad nájde a súbor sa rozúčtuje; a že mimo rozpätia sa vráti `NenajdenyBankovyDoklad` s hláškou.

---

## Poznámky k dokončeniu

- Zmeny sú v 3 repozitároch — po odsúhlasení použi superpowers:finishing-a-development-branch pre každý repo (OmegaLib, JurhanLib, RozuctovanieDopravcov) samostatne.
- OmegaLib bol na `master` → commituje sa na feature vetve `feature/doklad-podla-textu-hlavicky`.
