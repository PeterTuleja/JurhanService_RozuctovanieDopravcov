using JurhanLib.Services;
using JurhanModels.Enum;

namespace JurhanService_RozuctovanieDopravcov
{
    internal static class Program
    {
        internal static eTypSpustenia typSpustenia = eTypSpustenia.Servica;
        internal static string serviceName = "JurhanServiceNew_RozuctovanieDopravcov";
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            ServiceRunner.RunService<Service, Execute>(args, ref typSpustenia);
        }
    }
}
