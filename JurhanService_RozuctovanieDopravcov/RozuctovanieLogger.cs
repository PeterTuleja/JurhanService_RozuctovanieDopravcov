using JurhanLib.Services;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class RozuctovanieLogger : ServicesLogger
    {
        public RozuctovanieLogger() : base("RozuctovanieDopravcov")
        {
        }

        protected override bool Poslat()
        {
            return true;
        }

        protected override string Subject()
        {
            return "Rozúčtovanie dopravcov - log";
        }

        protected override string Body()
        {
            return "Log súbor z rozúčtovania dopravcov.";
        }
    }
}
