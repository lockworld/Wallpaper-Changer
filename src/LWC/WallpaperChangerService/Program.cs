using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace WallpaperChangerService
{
    public class Program
    {
        public static void Main(string[] args)
        {
            
            CreateHostBuilder(args).Build().Run();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) =>
            Host.CreateDefaultBuilder(args)
                .ConfigureServices((hostContext, services) =>
                {
                    WCServiceTools.InstallService(new string[] { "C:\\GIT", "WallpaperChanger" });
                    Console.WriteLine("Ready to launch application? ");
                    Console.ReadLine();

                    services.AddHostedService<Worker>();
                    services.AddHostedService<WallpaperSwitcherBackgroundService>();

                });
    }
}
