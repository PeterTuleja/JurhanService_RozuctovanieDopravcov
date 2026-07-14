using OmegaLib.Services;
using JurhanLib.Email;
using JurhanLib.Logger;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class RozuctovanieLogger : Logger
    {
        internal RozuctovanieLogger() : base("RozuctovanieDopravcov")
        {
        }

        public override void PosliLogSuborEmailom(IEnumerable<string> adresy, int hour, IEnumerable<DayOfWeek> days)
        {
            if (DateTime.Now.Hour == hour && days.Contains(DateTime.Now.DayOfWeek))
            {
                ZapisDataDoLogSuboru();
                ServicesFile.ZalohujSubor(fileNameXmlUser);
                ServicesFile.ZalohujSubor(fileNameXmlDeveloper);
                string _fileNameLogDeveloperWithDate = ServicesFile.ZalohujSubor(fileNameLogDeveloper);
                string _fileNameLogUserWithDate = ServicesFile.ZalohujSubor(fileNameLogUser);
                bool odoslane = EmailService.PosliEmail(
                    adresy,
                    null,
                    null,
                    "Rozúčtovanie dopravcov - log",
                    "Log súbor z rozúčtovania dopravcov.",
                    new List<string> { _fileNameLogUserWithDate, _fileNameLogDeveloperWithDate });

                // súbory sa mažú až po úspešnom odoslaní; Email.Send výnimku nehádže,
                // úspech treba kontrolovať cez návratovú hodnotu
                if (odoslane)
                {
                    File.Delete(fileNameXmlUser);
                    File.Delete(fileNameXmlDeveloper);
                    File.Delete(fileNameLogUser);
                    File.Delete(fileNameLogDeveloper);
                }
                else
                {
                    ServicesLog.LogujSDatumom("Log súbory nemažem - e-mail s logmi sa nepodarilo odoslať, skúsim v ďalšom okne.");
                }
            }
        }

    }
}
