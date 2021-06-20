using System;
using System.Collections.Generic;
using System.Text;


using System.ServiceProcess;
using System.Linq;

namespace WallpaperChangerService
{
    public static class WCServiceTools
    {
        public static void Log(string msg)
        {
            Console.WriteLine(msg);
        }


        public static bool InstallService(string[] args)
        {
            bool ret = false;
            if (args.Length > 0)
            {
                string location = args[0];
                string name = (args.Length > 1 ? args[1] : "WallpaperChangerService");
                ServiceController svc = ServiceController.GetServices().FirstOrDefault(s => s.ServiceName.Equals(name, StringComparison.InvariantCultureIgnoreCase));
                if(svc!=null)
                {
                    WCServiceTools.Log("Service " + name + " is installed.");
                }
                else
                {
                    WCServiceTools.Log("Service " + name + " is not installed.");
                   
                }



                // install service

                try
                {
                    WCServiceTools.Log("");
                }
                catch (Exception ex)
                {
                    WCServiceTools.Log(ex.Message.ToString());
                }

            }
            return ret;
        }
    }
}
