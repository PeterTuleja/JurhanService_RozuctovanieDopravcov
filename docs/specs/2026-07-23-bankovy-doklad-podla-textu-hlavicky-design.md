# Hľadanie bankového dokladu podľa textu hlavičky a dátumu výpisu

Dátum: 2026-07-23
Stav: návrh (schválený používateľom)

## Kontext a motivácia

Služba `RozuctovanieDopravcov` pri type rozúčtovania `eTypRozuctovania.SoZapoctomBanky`
potrebuje nájsť existujúci **bankový doklad** (EUD hlavička, `T040_EUD`), voči ktorému
rozúčtuje platby dopravcu zo súboru (prílohy e-mailu).

**Súčasný stav** (`RozuctovanieCore.DajDokladPreRozuctovanie`, vetva služby `interneCislo == null`):
doklad sa hľadá cez `EudHlavickaRepository.DajDokladPodlaPoznamky` — hľadá sa presná zhoda
`C097_Poznamka == názov súboru bez prípony`. To vyžaduje, aby niekto ručne vyplnil poznámku
dokladu presne rovnako ako názov súboru.

**Cieľ:** prestať sa spoliehať na poznámku a doklad hľadať podľa:
1. **textu účtovného zápisu** (`C099_VolneDefinovanyText`), ktorý má obsahovať **názov priečinka** dopravcu, a
2. **dátumu výpisu** — dátum vystavenia dokladu (`C060–062`) musí byť v rozpätí **±2 dni**
   od dátumu výpisu zisteného zo spracúvaného súboru.

Vetva UI (`interneCislo != null`, hľadanie podľa interného čísla) **ostáva bez zmeny**.

## Zistenia z reálnych súborov (Log priečinok)

„Dátum výpisu" je u každého dopravcu inde. Koncový `_RRRR-MM-DD-HH-MM-SS` v názvoch v `Log`
je **timestamp zálohovania** (`ServicesFile.ZalohujSubor`), nie súčasť pôvodného názvu —
živý (spracúvaný) súbor ho nemá.

| Súbor / 1. stĺpec | Dopravca | Zdroj dátumu výpisu | Hodnota |
|---|---|---|---|
| `Journal No.` | GLS | stĺpec „Dátum prevodu" (konšt.); pozn. aj „Dátum doruč." je konšt. | 2026-07-13 |
| `parcelno` | DPD SK | stĺpec „payment_gen_date" (konšt.) | 2026-07-13 |
| `Parcel number` | DPD HU/HR | stĺpec „Reconciliation Date" (konšt.) | 2026-07-14 |
| `VS` | SPS / Express One | stĺpec „DATUM" (konšt.) | 2026-07-14 |
| `pl_number` | DPD CZ | žiadny konštantný dátumový stĺpec → len názov súboru | 2026-07-14 |
| `ID pohybu` | GoPay | žiadny konštantný (týždenný rozsah) | — |
| `Nr paczki` | GLS PLN | nejasné (viac meniacich sa dátumov) | — |
| `Váš e-shop` | Packeta | bez dátumu v obsahu aj názve | — |

Pozorovanie, na ktorom stojí riešenie: dátum výpisu (prevod / reconciliation / payout) má
**rovnakú hodnotu vo všetkých dátových riadkoch**, zatiaľ čo delivery/pickup/transakčné dátumy
sa líšia. GLS má dva konštantné dátumy — vyhráva **najnovší** (prevod 07-13 > doručenie 07-08).

## Určenie dátumu výpisu (nová logika)

Vrstvená stratégia (v poradí):

1. **Konštantný dátumový stĺpec z obsahu súboru.** Po naparsovaní dátových riadkov sa pre každý
   stĺpec zistí, či sú všetky **neprázdne** hodnoty dátumy a či sú navzájom **rovnaké**.
   - práve 1 konštantný dátumový stĺpec → jeho dátum,
   - viac konštantných (rôzne hodnoty) → **najnovší** (max),
   - žiadny → krok 2.
2. **Dátum z názvu súboru** (napr. DPD CZ `…-2026_07_14-…`). Ak sa nenájde → krok 3.
3. **Dnešný dátum** (deň behu služby) — posledný fallback (Packeta, GoPay-rozsah).

Podporované formáty dátumu pri parsovaní hodnôt: `RRRR-MM-DD`, `RRRR.MM.DD.` (aj s koncovou
bodkou), `DD.MM.RRRR`, `RRRR-MM-DD HH:MM:SS` (berie sa dátumová časť).

Táto logika sa uplatní len pre typy `SoZapoctomBanky` (všetky prechádzajú cez `DopravcaToUhrady`;
`.xls`/`.xlsx` sú v tom čase už skonvertované na CSV, takže dáta sa čítajú jednotne).

## Zmeny po komponentoch

### `RozuctovanieContext`
Pribudnú polia:
- `string nazovPriecinka` — názov IMAP priečinka dopravcu (napr. „DPD SK", „GLS").
- `DateTime? datumVypisu` — dátum výpisu zistený zo súboru (naplní parser).

### `RozuctovanieEmailov`
Pri spracovaní súboru sa do kontextu nastaví `nazovPriecinka = folder.Name`.

### `DopravcaToUhrady` (JurhanLib)
Po naparsovaní dátových riadkov vypočíta dátum výpisu podľa vrstvenej stratégie
(krok 1 z obsahu, inak krok 2 z názvu, inak krok 3 dnes) a uloží ho do `_ctx.datumVypisu`.
Nová privátna pomocná logika: detekcia konštantného dátumového stĺpca + parser dátumu.

### `RozuctovanieCore.Execute` — poradie
Keďže dátum výpisu treba pred hľadaním dokladu, **parsovanie súboru sa presunie pred
hľadanie bankového dokladu** (dnes je hľadanie dokladu pred parsovaním). Duplicitná kontrola
(`C149_ImportText`) ostáva pred hľadaním dokladu.

### `RozuctovanieCore.DajDokladPreRozuctovanie` — vetva služby
Namiesto `DajDokladPodlaPoznamky`:
1. Načítaj kandidátov: doklady, kde `C099_VolneDefinovanyText` obsahuje `nazovPriecinka` (LIKE).
2. Ponechaj tie, kde `DatumVystavenia` je v rozpätí ±2 dni od `datumVypisu`
   (`|DatumVystavenia.Date − datumVypisu.Date| ≤ 2` dni).
3. Viac zhôd → vyber doklad s **najbližším** `DatumVystavenia` k `datumVypisu`.
4. Žiadna zhoda → `NenajdenyBankovyDoklad` + log s názvom priečinka a dátumom výpisu.

Vetva UI (`interneCislo != null`) ostáva nezmenená.

### `EudHlavickaRepository`
Nová metóda:
```csharp
public IEnumerable<EUDHlavicka> DajDokladyPodlaTextuHlavicky(string text, bool dbJeSQL) =>
    DajDoklady($"{ServicesDatabase.DajStringZMema("C099_VolneDefinovanyText", dbJeSQL)} LIKE @1",
               "%" + text + "%");
```
(memo stĺpec `C099` sa obaľuje cez `ServicesDatabase.DajStringZMema` rovnako ako pri `C097`.)

## Chybové stavy a logovanie
- Nenájdený doklad (žiadny kandidát v rozpätí) → `eVysledokRozuctovania.NenajdenyBankovyDoklad`,
  log obsahuje názov priečinka a použitý dátum výpisu.
- Log by mal zaznamenať aj **zdroj dátumu výpisu** (stĺpec / názov súboru / dnešný dátum),
  aby bolo dohľadateľné, ako sa doklad pároval.

## Testovanie
- Detekcia konštantného dátumového stĺpca: 1 stĺpec, viac stĺpcov (→ najnovší), žiadny (→ fallback),
  ignorovanie prázdnych hodnôt (sumárny/total riadok), rôzne formáty dátumu.
- Fallback na názov súboru (DPD CZ) a na dnešný dátum (Packeta).
- Výber dokladu: LIKE nad textom hlavičky, filter ±2 dni, výber najbližšieho pri viacerých.
- Vetva UI (`interneCislo`) sa nezmenila.

## Otvorené / na potvrdenie pri revízii
- Konkrétny parser dátumu z názvu súboru (formáty per dopravca) je potrebný len pre fallback
  (DPD CZ `YYYY_MM_DD`). Rozsah formátov sa doladí podľa reálnych názvov.
- Prípadné obmedzenie kandidátov len na bankový okruh dokladov (zatiaľ neriešené — spolieha sa
  na selektívnosť textu hlavičky + dátumu).
