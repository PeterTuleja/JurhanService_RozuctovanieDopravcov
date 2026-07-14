using OmegaLib.Services;
using JurhanLib.Services;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class SpustenieServicy : SpustenieServisyBase
    {
        private readonly RozuctovanieLogger _logger;

        internal SpustenieServicy(RozuctovanieLogger logger) : base(logger)
        {
            _logger = logger;
        }

        public override void VykonajFunkciu()
        {
            RozuctovanieEmailov rozuctovanieEmailov = new RozuctovanieEmailov(pripojeneFirmy, _logger);
            rozuctovanieEmailov.Execute();
        }
    }
}
