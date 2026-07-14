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
