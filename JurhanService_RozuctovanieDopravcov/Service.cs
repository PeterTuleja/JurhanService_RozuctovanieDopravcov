using JurhanService;
using System;

namespace JurhanService_RozuctovanieDopravcov
{
    internal class Service : ServiceJurhan
    {
        public Service() : base(new RozuctovanieLogger())
        {
        }

        public override void ExecuteRun()
        {
            Execute execute = new Execute();
            execute.Run();
        }

        public override TimeSpan Interval()
        {
            return new TimeSpan(1, 0, 0); // kazdu hodinu
        }

        public override string Name()
        {
            return Program.serviceName;
        }
    }
}
