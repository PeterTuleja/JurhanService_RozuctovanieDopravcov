/* =====================================================================
   Migrácia kľúča rozúčtovania:  C149_ImportText  ->  C112_Zauctoval
   Databáza: x_455315912        Tabuľka: T040_EUD

   Čo robí:
     - obsah C149_ImportText, ktorý nie je prázdny a nie je 'api',
       presunie do C112_Zauctoval
     - do C149_ImportText zapíše 'api' (doklad vytvorila servisa, nie človek)

   Bezpečnostné poistky:
     - prepisuje LEN riadky, kde je C112_Zauctoval prázdny (mená ľudí,
       ktorí doklad zaúčtovali ručne, ostanú nedotknuté)
     - pred zmenou vytvorí zálohovú tabuľku s pôvodnými hodnotami
     - beží v transakcii, na konci vypíše kontrolu

   PRED SPUSTENÍM:
     1. zavri Omegu (aby nikto nezapisoval do T040_EUD)
     2. sprav zálohu databázy
   ===================================================================== */

USE x_455315912;
GO

SET XACT_ABORT ON;
SET NOCOUNT ON;
GO

/* ---------------------------------------------------------------------
   1. STAV PRED MIGRÁCIOU
   --------------------------------------------------------------------- */
PRINT '=== STAV PRED ===';

SELECT
    COUNT(*)                                                                       AS riadkov_spolu,
    SUM(CASE WHEN C149_ImportText = 'api' THEN 1 ELSE 0 END)                       AS c149_api,
    SUM(CASE WHEN C149_ImportText IS NOT NULL AND C149_ImportText NOT IN ('', 'api')
             THEN 1 ELSE 0 END)                                                    AS c149_na_presun,
    SUM(CASE WHEN C112_Zauctoval IS NOT NULL AND DATALENGTH(C112_Zauctoval) > 0
             THEN 1 ELSE 0 END)                                                    AS c112_vyplnenych
FROM dbo.T040_EUD;

/* koľko riadkov by kolidovalo (C112 už vyplnený) - očakávaná hodnota 0 */
SELECT COUNT(*) AS kolizie_nemigrujem
FROM dbo.T040_EUD
WHERE C149_ImportText IS NOT NULL
  AND C149_ImportText NOT IN ('', 'api')
  AND C112_Zauctoval IS NOT NULL
  AND DATALENGTH(C112_Zauctoval) > 0;
GO

/* ---------------------------------------------------------------------
   2. ZÁLOHA PÔVODNÝCH HODNÔT
   --------------------------------------------------------------------- */
IF OBJECT_ID('dbo.T040_EUD_zaloha_C149') IS NOT NULL
BEGIN
    RAISERROR('Zálohová tabuľka dbo.T040_EUD_zaloha_C149 už existuje - premenuj ju alebo zmaž.', 16, 1);
    RETURN;
END

SELECT
    C000_ID,
    C149_ImportText                          AS C149_povodne,
    CAST(C112_Zauctoval AS varchar(max))     AS C112_povodne
INTO dbo.T040_EUD_zaloha_C149
FROM dbo.T040_EUD
WHERE C149_ImportText IS NOT NULL
  AND C149_ImportText NOT IN ('', 'api');

PRINT '=== ZÁLOHA VYTVORENÁ: dbo.T040_EUD_zaloha_C149 ===';
SELECT COUNT(*) AS zalohovanych FROM dbo.T040_EUD_zaloha_C149;
GO

/* ---------------------------------------------------------------------
   3. MIGRÁCIA
   --------------------------------------------------------------------- */
BEGIN TRANSACTION;

UPDATE dbo.T040_EUD
SET C112_Zauctoval = C149_ImportText,
    C149_ImportText = 'api'
WHERE C149_ImportText IS NOT NULL
  AND C149_ImportText NOT IN ('', 'api')
  AND (C112_Zauctoval IS NULL OR DATALENGTH(C112_Zauctoval) = 0);

PRINT '=== ZMENENÝCH RIADKOV ===';
SELECT @@ROWCOUNT AS zmenenych;

COMMIT TRANSACTION;
GO

/* ---------------------------------------------------------------------
   4. STAV PO MIGRÁCII
   --------------------------------------------------------------------- */
PRINT '=== STAV PO ===';

SELECT
    SUM(CASE WHEN C149_ImportText = 'api' THEN 1 ELSE 0 END)                       AS c149_api,
    SUM(CASE WHEN C149_ImportText IS NOT NULL AND C149_ImportText NOT IN ('', 'api')
             THEN 1 ELSE 0 END)                                                    AS c149_zvysok_nemigrovany,
    SUM(CASE WHEN C112_Zauctoval IS NOT NULL AND DATALENGTH(C112_Zauctoval) > 0
             THEN 1 ELSE 0 END)                                                    AS c112_vyplnenych
FROM dbo.T040_EUD;

/* mená ľudí musia ostať nedotknuté */
PRINT '=== C112: mená (ručné zaúčtovanie) musia sedieť s pôvodným stavom ===';
SELECT TOP 10 CAST(C112_Zauctoval AS varchar(200)) AS hodnota, COUNT(*) AS pocet
FROM dbo.T040_EUD
WHERE C112_Zauctoval IS NOT NULL AND DATALENGTH(C112_Zauctoval) > 0
  AND CAST(C112_Zauctoval AS varchar(max)) NOT LIKE 'dopravca:%'
GROUP BY CAST(C112_Zauctoval AS varchar(200))
ORDER BY COUNT(*) DESC;
GO


/* =====================================================================
   NÁVRAT SPÄŤ (spustiť len ak treba migráciu vrátiť)
   =====================================================================

UPDATE e
SET e.C149_ImportText = z.C149_povodne,
    e.C112_Zauctoval  = z.C112_povodne
FROM dbo.T040_EUD e
JOIN dbo.T040_EUD_zaloha_C149 z ON z.C000_ID = e.C000_ID;

-- po overení:
-- DROP TABLE dbo.T040_EUD_zaloha_C149;

   ===================================================================== */
