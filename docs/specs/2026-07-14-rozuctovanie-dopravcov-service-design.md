# Dizajn: Služba RozuctovanieDopravcov (net10 + net48)

Dátum: 2026-07-14
Stav: schválený návrh (ústne), čaká na review spec dokumentu

## 1. Cieľ

Windows služba, ktorá automaticky rozúčtuje platby dopravcov / platobných brán / marketplace-ov.
Preberá kompletnú funkčnosť WinForms programu RozuctovanieDopravcov. Zdrojom vstupov je
emailová schránka **platby@jurhan.com** (IMAP): priečinky schránky zodpovedajú typom dopravcov,
emaily obsahujú súbory (prílohy) na rozúčtovanie. Po úspešnom rozúčtovaní sa email presunie do
podpriečinka **„Zaúčtované“** v rámci priečinka dopravcu.

Vzniknú dve verzie služby podľa existujúcich vzorov:

| Verzia | Umiestnenie | Vzor |
|---|---|---|
| .NET 10 | `C:\Projekty\Private\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\` | napr. ImportObjednavok (SDK csproj, ProjectReference, sc.exe deploy) |
| .NET 4.8 | `C:\Users\tuleja\source\repos\JurhanService\RozuctovanieDopravcov\JurhanService_RozuctovanieDopravcov\` | napr. ImportFaktur (classic csproj, ProjectInstaller, InstallUtil bat) |

Názvy služieb: `JurhanServiceNew_RozuctovanieDopravcov` (net10), `JurhanService_RozuctovanieDopravcov` (net48).

## 2. Presun zdieľanej logiky do JurhanLib

Parsery a orchestrácia dnes žijú v UI projekte programu a závisia na statickej triede `Program`.
Presunú sa do **JurhanLib** (obe verzie: `C:\Projekty\Private\JurhanProgramy\JurhanLib\` aj
`C:\Users\tuleja\source\repos\JurhanProgramy\JurhanLib\`), nový adresár `Import\Dopravcovia\`:

Presúvané triedy (z `RozuctovanieDopravcov` UI projektu):
- `CsvBase`, `DopravcaToUhrady` (CSV/XLSX dopravcovia a brány)
- `BaseParser`, `AllegroParser`, `KauflandParser` (DB parsery)
- `EmagParser` (Excel + DB)
- orchestrácia z `Rozuctovanie.cs` → nová trieda **`RozuctovanieCore`**

Zásady refaktoringu:
- Žiadna závislosť na `Program.*` statikách — všetko cez parametre / konštruktor.
- Vstupný parameter objekt **`RozuctovanieParametre`**:
  `IDataProvider dataProvider`, `PripojeneFirmy pripojeneFirmy`, `eTypSuboru typSuboru`,
  `string filePath`, `short mesiac`, `string interneCisloDokladu` (nullable),
  `string omegaPath`, `eZobrazenieChyby zobrazenieChyby`, `eTypSpustenia typSpustenia`.
- Hľadanie bankového dokladu pre `eTypRozuctovania.SoZapoctomBanky`:
  - ak je zadané `interneCisloDokladu` (UI program) → `EudHlavickaRepository.DajDokladPodlaInternehoCisla` (dnešné správanie),
  - ak nie je (služba) → **`EudHlavickaRepository.DajDokladPodlaPoznamky(nazovSuboruBezPripony, dbJeSQL)`**
    (stĺpec `C097_Poznamka` v T040_EUD). Nenájdený doklad → výnimka/chybový návrat, volajúci loguje.
- Jadro `UhradyToTxt`, `Lib` (NastavTypRozuctovania, DajKodEvidencieACiselnyRad, …) sa nemení.
- Kontrola duplicity ostáva ako v programe: doklad s `C149_ImportText = nazov suboru bez pripony`
  → súbor už bol rozúčtovaný (výsledok „duplicita“).

Program RozuctovanieDopravcov (obe verzie) sa upraví tak, aby volal `RozuctovanieCore`
— správanie UI sa nemení (interné číslo dokladu naďalej zadáva používateľ).

## 3. Kostra služby (zhodná s ostatnými službami)

```
Program.Main → ServiceRunner.RunService<Service, Execute>(args, ref typSpustenia)
Service : ServiceJurhan          // Interval() = 1 hodina, Name()
Execute : IExecutable            // vytvorí SpustenieServicy a spustí Execute(...)
SpustenieServicy : SpustenieServisyBase   // časové okno, OtvorFirmu (Constants.IdFirma), VykonajFunkciu()
RozuctovanieLogger : Logger      // voliteľné poslanie logu emailom (vzor OrdersLogger)
```

net48 navyše: `ProjectInstaller : ProjectInstallerBase` + `zaregistruj_/odregistruj_RozuctovanieDopravcov.bat`.

## 4. Tok spracovania (`VykonajFunkciu`)

1. **IMAP pripojenie** (MailKit `ImapClient`) na platby@jurhan.com.
   Nové konštanty v `Constants.cs` (JurhanModels, obe verzie):
   `ClientHostImapJurhan` (doplní používateľ, napr. imap host pre jurhan.com),
   `ClientPasswordPlatby` (placeholder, doplní používateľ); adresa = existujúca `MessageToPlatbyJurhan`.
2. **Prechod priečinkov** schránky (`client.GetFolders(PersonalNamespaces[0])`).
   Mapovacia funkcia **`DajTypSuboru(string nazovPriecinka) : eTypSuboru`** — switch s ľudskými
   názvami priečinkov (naplnený odhadmi podľa enumu `eTypSuboru`, používateľ si názvy doplní).
   Nezmapovaný priečinok → log (info) a preskočiť. Podpriečinok „Zaúčtované“ sa nikdy nespracúva.
3. **Pre každý email** v priečinku (chronologicky):
   a. Stiahnuť prílohy do pracovného adresára `C:\Omega\Import\RozuctovanieDopravcov\`
      (vytvorí sa; súbor sa po spracovaní zmaže, zostáva záloha cez `ServicesFile.ZalohujSubor` do Log adresára).
   b. Pre každú prílohu (súbor dopravcu):
      - `mesiac` = **z názvu súboru** funkciou `DajMesiacZNazvuSuboru` — hľadá vzor `MM.RRRR`,
        `MM-RRRR`, `MM_RRRR` alebo `RRRR-MM` v názve; ak nenájde → chyba súboru (log, email ostáva).
      - `typRozuctovania` = `Lib.NastavTypRozuctovania(typSuboru)`.
      - **Duplicita** (`C149_ImportText`): zalogovať upozornenie, email **presunúť do Zaúčtované** (bod 4).
      - **SoZapoctomBanky**: nájsť bankový doklad podľa `C097_Poznamka = nazov suboru bez pripony`.
        **Ak sa nenájde → zapísať do logu a email nechať na mieste** (skúsi sa v ďalšom behu), pokračovať ďalším emailom.
      - Spustiť `RozuctovanieCore.Execute(parametre)` → načítanie úhrad (parser podľa typu),
        `UhradyToTxt` → TXT do `<omega>\Import\` → `MyAutoImport` (Omega autoimport) → kontrola logu importu.
4. **Po úspešnom rozúčtovaní všetkých príloh emailu**: presunúť email do IMAP podpriečinka
   **„Zaúčtované“** daného priečinka (`folder.GetSubfolders` / `folder.Create("Zaúčtované", true)`, `folder.MoveTo(uid, ...)`).
5. **Chyby**: chyba jedného súboru/emailu → log + pokračuje sa ďalším (email ostáva na mieste).
   Neošetrená výnimka behu → štandard `ServiceJurhan` (log + email o chybe).

DB typy (Allegro, Kaufland, Dobropisy): email s prílohou slúži ako spúšťač; úhrady sa čítajú z DB
(ako v programe), z názvu prílohy sa berie mesiac. Emag: príloha (Excel s Reference ID) sa parsuje ako v programe.

## 5. Projekty a závislosti

### net10 (`JurhanService_RozuctovanieDopravcov.csproj`, SDK style)
- `TargetFramework=net10.0-windows`, `OutputType=WinExe`, `UseWindowsForms=true`, `ImplicitUsings/Nullable=disable`.
- ProjectReference: OmegaLib, DataProvider, JurhanModels, JurhanLib.
- PackageReference: `Microsoft.Windows.Compatibility 10.0.9`, `MailKit 4.13.0`.
- HintPath: `C:\Omega\Kros.KORM.dll`, `Kros.Utils.dll` (podľa vzoru ostatných služieb).
- Deploy: doplniť do `Deploy\Install-AllJurhanServices.ps1` (a Publish skriptu) položku
  `JurhanServiceNew_RozuctovanieDopravcov`.

### net48 (classic csproj)
- `TargetFrameworkVersion v4.8`, `OutputType WinExe`, `App.config` (supportedRuntime + bindingRedirects pre MailKit).
- Reference HintPath: JurhanLib, JurhanModels, IndiLib, Kros.KORM, Omega.* (podľa vzoru ImportFaktur),
  MailKit/MimeKit 4.13 + podporné System.* balíky cez `packages.config` (vzor ArchivaciaEmailov).
- `ProjectInstaller` + bat skripty (InstallUtil, nasadenie do `c:\JurhanService\`).
- Pozn.: v starej vetve sa presun tried robí v starom JurhanLib; parsery používajú
  `Encoding.Default` správanie net48 (Windows-1250) — zachovať tamojší stav, v novej vetve
  `ServicesEncoding.Windows1250` (už opravené pri migrácii).

## 6. Logovanie

- `ServicesLog` do `<exe>\Log\JurhanService_RozuctovanieDopravcov_<yyyyMMdd>.log`.
- Logujú sa: štart/koniec behu, každý priečinok, každý spracovaný email/súbor, výsledok
  (rozúčtované / duplicita / nenájdený bankový doklad / chyba parsovania), presuny emailov.
- `RozuctovanieLogger : Logger` — voliteľné odoslanie logu emailom (vzor `OrdersLogger`).

## 7. Testovanie a overenie

- Build oboch služieb, oboch JurhanLib a oboch verzií programu RozuctovanieDopravcov.
- Manuálny beh v režime „Program“ (argument `p`), overenie proti testovacej schránke/firme — vykoná používateľ (prístupové údaje nie sú v repo k dispozícii pre agenta).
- Regresné overenie programu: UI beh s rovnakým vstupným súborom pred/po refaktoringu má produkovať rovnaký TXT import.

## 8. Otvorené body (doplní používateľ)

- Presné názvy IMAP priečinkov → doplniť do `DajTypSuboru`.
- IMAP host a heslo pre platby@jurhan.com → doplniť do `Constants.cs`.
- Presný formát mesiaca v názvoch súborov (spec predpokladá `MM.RRRR`/`MM-RRRR`/`MM_RRRR`/`RRRR-MM`).
