using JurhanLib.Services;
using JurhanModels.Enum;
using OmegaLib.Services;
using System.Collections.Generic;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class Execute : IExecutable
    {
        private readonly RozuctovanieLogger _logger = new RozuctovanieLogger();

        public void Run()
        {
            ServicesLog.VytvorLogovaciAdresar();

            SpustenieServicy spustenieServicy = new SpustenieServicy(_logger);
            spustenieServicy.Execute(Program.typSpustenia, eTypServisy.IbaProgram, 0, new List<short> { 0, 1, 2, 3, 4, 5, 6 }, 0, 23,
                "Rozúčtovanie dopravcov",
                new List<string> { Constants.MessageToTulejaX });            
        }
    }
}
