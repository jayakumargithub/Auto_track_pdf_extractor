using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace PdfDataExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;

#if DEBUG
            AutoTrackService serv = new AutoTrackService();

            serv.OnDebug();
#else
            ServicesToRun = new ServiceBase[]
            {
                new AutoTrackService()
            };
            ServiceBase.Run(ServicesToRun);
#endif
        }
    }
}
