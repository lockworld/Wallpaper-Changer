using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Microsoft.Extensions.Logging;

namespace WallpaperChangerService
{
    public class WallpaperSwitcherBackgroundService : BackgroundService
    {
        //private readonly ILogger<Worker> _logWorker;

        //public WallpaperSwitcherBackgroundService(ILogger<Worker> logWorker)
        //{
        //    _logWorker = logWorker;
        //}

        public override async Task StartAsync(CancellationToken cancelToken)
        {
            await base.StartAsync(cancelToken);
        }

        public override async Task StopAsync(CancellationToken cancelToken)
        {
            await base.StopAsync(cancelToken);
        }

        protected override async Task ExecuteAsync(CancellationToken stopToken)
        {
            while (!stopToken.IsCancellationRequested)
            {
                //_logWorker.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                try
                {
                    Console.WriteLine("Iteration " + DateTime.Now.ToString());
                    await Task.Delay(1000, stopToken);
                }
                catch (Exception ex)
                {
                    await StopAsync(new CancellationToken(true));
                }
            }
        }

        public override void Dispose()
        {
            
        }
    }
}
