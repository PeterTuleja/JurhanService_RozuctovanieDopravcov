using OmegaLib.Services;
using JurhanLib.Services;
using JurhanModels.Enum;
using System;
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
            _logger.NacitajDataZoSuboru();
            try
            {
                spustenieServicy.Execute(Program.typSpustenia, eTypServisy.IbaProgram, 0, new List<short> { 0, 1, 2, 3, 4, 5, 6 }, 0, 23,
                    "Rozúčtovanie dopravcov",
                    new List<string> { Constants.MessageToTulejaX });
            }
            finally
            {
                _logger.ZapisDataDoSuboru();
                _logger.ZapisDataDoLogSuboru();
                _logger.PosliLogSuborEmailom(new List<string> { Constants.MessageToTulejaX }, 22,
                    new List<DayOfWeek> { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday,
                        DayOfWeek.Friday, DayOfWeek.Saturday, DayOfWeek.Sunday });
            }
        }
    }
}
