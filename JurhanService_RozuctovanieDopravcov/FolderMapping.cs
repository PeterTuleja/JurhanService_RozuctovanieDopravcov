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
                case "DPD CZ": return eTypSuboru.Dopravca_DPD_Cesko;
                case "DPD HR": return eTypSuboru.Dopravca_DPD_Chorvatsko;
                case "DPD HU": return eTypSuboru.Dopravca_DPD_Madarsko;
                case "DPD PL": return eTypSuboru.Dopravca_DPD_Polsko;
                case "DPD RO": return eTypSuboru.Dopravca_DPD_Rumunsko;
                case "DPD SI": return eTypSuboru.Dopravca_DPD_Slovinsko;
                case "DPD SK": return eTypSuboru.Dopravca_DPD_Slovensko;
                case "Express One": return eTypSuboru.Dopravca_Express_One;
                case "GEIS": return eTypSuboru.Undefined;
                case "GLS": return eTypSuboru.Dopravca_GLS;
                case "GLS CZ": return eTypSuboru.Dopravca_GLS;
                case "GLS HU": return eTypSuboru.Dopravca_GLS;
                case "GLS PLN": return eTypSuboru.Dopravca_GLS_PLN;
                case "GLS RO": return eTypSuboru.Dopravca_GLS;
                case "GLS SK": return eTypSuboru.Dopravca_GLS;
                case "PACKETA": return eTypSuboru.Dopravca_Packeta;
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
