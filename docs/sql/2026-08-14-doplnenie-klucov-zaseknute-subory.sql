/* =====================================================================
   Doplnenie kľúča rozúčtovania pre zaseknuté súbory
   Databáza: x_455315912        Tabuľka: T040_EUD

   SPUSTIŤ AŽ PO skripte 2026-08-14-migracia-c149-importtext-do-c112-zauctoval.sql
   (ten najprv presunie kľúče z C149_ImportText do C112_Zauctoval).

   Prečo: kľúče presunuté z C149_ImportText sú už orezané na 50 znakov.
   Kontrola v servise porovnáva celý názov súboru, takže orezaný kľúč
   nespozná a súbor by sa pokúsila zaúčtovať znova.
   ===================================================================== */

USE x_455315912;
GO

SET XACT_ABORT ON;
SET NOCOUNT ON;
GO

/* ---------------------------------------------------------------------
   1. DPD SI - SI_cod_wire_transfer_email_xls.104822.20260728.xls
      87 dokladov, výpis 29.07.2026, v Omege sú kompletné.
      Kľúč je orezaný na 50 znakov -> doplníme celý názov.
   --------------------------------------------------------------------- */
PRINT '=== DPD SI: pred ===';
SELECT COUNT(*) AS najdenych
FROM dbo.T040_EUD
WHERE CAST(C112_Zauctoval AS varchar(max)) = 'dopravca: SI_cod_wire_transfer_email_xls.104822.20';

BEGIN TRANSACTION;

UPDATE dbo.T040_EUD
SET C112_Zauctoval = 'dopravca: SI_cod_wire_transfer_email_xls.104822.20260728.xls'
WHERE CAST(C112_Zauctoval AS varchar(max)) = 'dopravca: SI_cod_wire_transfer_email_xls.104822.20';

PRINT '=== DPD SI: zmenených (očakávaných 87) ===';
SELECT @@ROWCOUNT AS zmenenych;

COMMIT TRANSACTION;
GO

/* ---------------------------------------------------------------------
   2. GLS PLN - 1244169682
      18 dokladov, výpis 28.07.2026, v Omege sú kompletné.
      Kľúč sa do 50 znakov zmestil celý -> netreba meniť nič,
      tento SELECT je len na overenie, že po migrácii sedí.
   --------------------------------------------------------------------- */
PRINT '=== GLS PLN: kontrola (očakávaných 18) ===';
SELECT COUNT(*) AS najdenych
FROM dbo.T040_EUD
WHERE CAST(C112_Zauctoval AS varchar(max)) = 'dopravca: 1244169682';
GO

/* ---------------------------------------------------------------------
   3. DPD CZ - 2-51147584360-cod_wire_transfer_email_csv-2026_07_28-134547-5121
      TENTO SÚBOR SA KĽÚČOM OZNAČIŤ NESMIE.
      V Omege z neho nie je 18 dokladov, ale jediný - a ten je pokazený:
      import z 13.08. 13:32 mu nepridelil číslo (C030_CisloInterne = 'DPDCZ07????').
      Zvyšných 17 dokladov chýba.

      Postup: zmazať pokazený doklad v Omege (nie SQL-kom, nech sa upratú
      aj položky), potom nechať servisu súbor zaúčtovať nanovo.
   --------------------------------------------------------------------- */
PRINT '=== DPD CZ: pokazený doklad na zmazanie ===';
SELECT C000_ID, C030_CisloInterne, C024_CisloInterneKodEvidencia AS ev,
       C025_CisloInterneKodCiselnaRada AS rad,
       C060_DenVystavenia AS den, C061_MesVystavenia AS mes, C062_RokVystavenia AS rok,
       C083_CasZaevidovania AS cas_zaevidovania
FROM dbo.T040_EUD
WHERE C030_CisloInterne LIKE '%?%';
GO
