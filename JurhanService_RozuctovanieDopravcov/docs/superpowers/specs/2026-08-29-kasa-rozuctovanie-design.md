# Rozúčtovanie kasy — hotovosť a platobné karty

**Dátum:** 2026-08-29
**Stav:** návrh schválený, čaká na implementačný plán

## Cieľ

Zákazník (Jurhan s. r. o.) platí faktúry aj cez registračnú pokladnicu. Export dokladov
z kasy (`Doklady *.xlsx`) treba automaticky rozúčtovať do Omegy:

1. **Platby v hotovosti** → pokladničné doklady, evidencia **P1**, rad **P**
   (MD 211.001 Pokladnica č. 1 / DAL 311.001 Odberatelia — tuzemsko).
2. **Platby kartou** → interné doklady, evidencia **IDPB**, rad **IDSKP**
   (MD 315.102 pohľadávka voči SK Pay / DAL 311.001) — rovnaké účtovanie, aké dnes
   robí rozúčtovanie SkPay.

Oba typy sú **inkaso faktúry**: doklad sa páruje na faktúru (`CisloUhradzanehoDokladu`)
a datuje sa **dátumom úhrady** z riadku exportu (nie dátumom faktúry ani výpisu).

Súbor prichádza **emailom** (spracuje služba, priečinok schránky „Kasa") aj **ručne**
cez program (frmVyber).

## Vstupný súbor

Xlsx export z kasy, hárok „Doklady", jedna hlavička + riadok na doklad. Kľúčové stĺpce:

| Stĺpec | Význam |
|---|---|
| `Dátum` | dátum a čas úhrady → dátum dokladu |
| `Typ` | spracúvajú sa len riadky `10 - úhrada dokladu/faktúry` |
| `Referencia` | **interné číslo faktúry** v Omege (C030), napr. 1202616885, 2026080047 |
| `Hotovosť` / `Platobné karty` | suma podľa spôsobu platby |
| `Stav úhrady` | spracúvajú sa len riadky `uhradená` |
| `Poukážky`, `QR platba`, `Ostatné` | nepodporované — nenulová hodnota → riadok preskočiť a vykázať |

Vzorové doklady od zákazníka (screenshoty Omegy):

- Hotovosť: P1/P, „1 PD/BV – Inkaso OF, DD", predkontácia I-1, VS/ŠS = interné číslo
  faktúry, externé číslo = VS faktúry, partner z faktúry, text „Inkaso faktúry {číslo}",
  typ sumy V, MD 211.001 / DAL 311.001, úhrada spárovaná na OF.
- Karta: IDPB/IDSKP, „AL ID – Dopravcovia", predkontácia Z-Al, dátum vyhotovenia =
  splatnosť = DVDP = dátum úhrady, riadok S → DAL 311.001, riadok V → MD 315.102,
  úhrada spárovaná na OF.

## Architektúra (schválený prístup A)

Jeden typ súboru `eTypSuboru.Kasa`; jadro rozdelí úhrady podľa spôsobu platby na dve
dávky a zaúčtuje ich dvoma behmi `UhradyToTxt` (hotovosť → PD/P1/P, karta →
ID/IDPB/IDSKP). Zápis do Omegy ostáva cez existujúci TXT autoimport.

### 1. Parser (`KasaParser`, JurhanLib)

- Vzor `EmagParser` — DevExpress Spreadsheet (`Workbook.LoadDocument`).
- Stĺpce sa hľadajú **podľa názvov v hlavičke**, nie podľa pevných indexov.
- Riadok → `Uhrada` s číslom faktúry (Referencia), sumou, dátumom úhrady a spôsobom
  platby. Riadok s hotovosťou aj kartou (delená platba) → dve úhrady.
- Preskočené riadky (iný typ dokladu, neuhradený stav, nepodporovaný spôsob platby,
  prázdna referencia) sa vykážu v logu.

### 2. Rozšírenie modelu

- `Uhrada` (JurhanModels): nové polia `DatumUhrady` (nullable DateTime) a spôsob platby
  (enum Hotovosť/Karta; vyplnené len pre kasu).
- `UhradyToTxt.NastavDatumRozuctovania`: nová vetva — pre kasu dátum dokladu
  (vystavenia, DUUP, DUZP, splatnosti) = dátum úhrady z riadku.

### 3. Registrácia typu a smerovanie

- `eTypSuboru.Kasa` na koniec enumu (`JurhanModels\Import\Enums.cs`).
- `FolderMapping`: priečinok `"Kasa"` → `eTypSuboru.Kasa`.
- `Lib.NastavTypRozuctovania`: kasa → `BezZapoctuBanky` (bez bankového výpisu);
  na rozdiel od ostatných `BezZapoctuBanky` typov **nevyžaduje mesiac z názvu súboru**
  (dátumy sú v riadkoch).
- `Lib.DajKodEvidencieACiselnyRad`: rozšírenie o spôsob platby (parameter/overload):
  Kasa + hotovosť → (PD, "P1", "P"); Kasa + karta → (ID, "IDPB", "IDSKP").
- `Lib.JeTypSuboruXlsx`: kasa medzi xlsx typy; nová pomocná metóda `Lib.JeTypSuboruKasa`
  pre vetvenie v parseri a v `UhradyToTxt` (kasa sa nekonvertuje na CSV).
- Hľadanie faktúry podľa **interného čísla** (`DajDokladPodlaInternehoCisla`, C030) —
  ako Allegro/Kaufland, nie cez externé číslo.
- Účty: hotovosť V-položka MD **211.001** (nové konštanty v `Constants`), karta
  V-položka MD z existujúceho `DajUcetOstatnePohladavky` pre SkPay (315.102).
  S-položka DAL = účet z faktúry (311.001) — existujúca logika.

### 4. Ochrana proti duplicitám

Exporty sa môžu prekrývať (rovnaké doklady vo viacerých súboroch). Dve vrstvy:

1. Existujúca ochrana celého súboru (SHA-256 odtlačok v behu + kľúč
   `dopravca: <názov súboru>` v C112_Zauctoval).
2. **Nová kontrola pred každou úhradou**: ak je faktúra už plne uhradená
   (`FakturaJeUzUhradena`, C220_Uhradene ≥ suma), úhrada sa preskočí a vykáže v logu.

Obmedzenie: pri delenej platbe chráni až plná úhrada faktúry — v dátach zákazníka sa
delená platba zatiaľ nevyskytla, akceptované.

### 5. Služba a program

- Služba: priečinok „Kasa" beží najprv v **simulačnom režime** (nepridáva sa do
  `_pilotnePriecinky`); ostrý import sa zapne po overení.
- Program (frmVyber): typ sa v combe objaví automaticky z enumu; pre kasu sa skryjú
  polia mesiac/bankový doklad; filter dialógu súboru xlsx.

### 6. Logovanie a report

Rovnaký `RozuctovanieLogger`/kontextový logger ako ostatné typy. Na konci súhrn:
počet zaúčtovaných hotovostných a kartových dokladov, počty preskočených s dôvodom
(nespárovaná faktúra, už uhradená, iný typ dokladu, nepodporovaný spôsob platby).

## Overenia pred nasadením (riziká)

1. **Okruh PD (TypOkruh 160) sa cez TXT autoimport nikdy nepoužil** — na teste overiť,
   že import doklad prijme, a zistiť povinné polia navyše (pole „Prijaté od"
   (`PrijateOdVyplateneKomu`), VS/ŠS = interné číslo faktúry, externé číslo = VS
   faktúry — podľa vzoru zákazníka).
2. **Predkontácia PD**: vzor ukazuje „I-1 PD/BV – Inkaso OF"; dnešná konštanta 1-1/3233
   platí pre okruh ID — pre PD zistiť správny kód v Omege zákazníka.
3. Číselný rad P1/P musí v Omege existovať (na screenshote existuje, P-0283) — inak
   import padá, rovnako ako pri GLS HR/SI 27.08.

## Testy

Unit testy v `JurhanLib.Tests` (vzor `RozuctovanieParseryTests`):

- parser nad fixture xlsx (reálny export s anonymizovanými dátami): správne stĺpce,
  dátumy, sumy, rozdelenie hotovosť/karta,
- preskočenie riadkov (iný typ, neuhradené, poukážky/QR, prázdna referencia),
- delená platba → dve úhrady,
- smerovanie okruh/evidencia/rad podľa spôsobu platby,
- dedup: už uhradená faktúra sa preskočí.
