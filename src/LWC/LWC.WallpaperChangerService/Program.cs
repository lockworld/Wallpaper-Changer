using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace LWC.WallpaperChangerService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            if (!Debugger.IsAttached)
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                new WallpaperChanger()
                };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                log4net.Config.XmlConfigurator.Configure();
                var debugWallpaperChanger = new WallpaperChanger();
                debugWallpaperChanger.RunServiceLoop();
            }
        }
    }
}
