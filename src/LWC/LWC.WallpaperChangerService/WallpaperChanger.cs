using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LWC.WallpaperChangerService
{
    public partial class WallpaperChanger : ServiceBase
    {
        private CancellationTokenSource cts = new CancellationTokenSource();
        private Task serviceLoop = null;
        public WallpaperChanger()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Thread.Sleep(20000);
            WallpaperChangerTools.Log("Service started",MessageTypes.Information,"OnStart");
            try
            {
                log4net.Config.XmlConfigurator.Configure();
                serviceLoop = new Task(RunServiceLoop, cts.Token, TaskCreationOptions.LongRunning);
                serviceLoop.Start();
            }
            catch (Exception ex)
            {
                WallpaperChangerTools.Log("Unable to start service. " + ex.Message, MessageTypes.Critical, "OnStart");
            }
        }

        protected override void OnStop()
        {
            WallpaperChangerTools.Log("Service stopped", MessageTypes.Information, "OnStop");
            cts.Cancel();
            serviceLoop.Wait();
        }

        public void RunServiceLoop()
        {
            CancellationToken cancellation = cts.Token;
            TimeSpan interval = TimeSpan.Zero;
            TimeSpan waitAfterSuccessfulInterval = 
                TimeSpan.FromSeconds(WallpaperChangerTools.ReadConfigurationInt("ServiceIntervalInSeconds",5));
            TimeSpan waitAfterErrorInterval = 
                TimeSpan.FromSeconds(WallpaperChangerTools.ReadConfigurationInt("AfterErrorServiceIntervalInSeconds", 20));

            int counter = 0;
            while (!cancellation.WaitHandle.WaitOne(interval))
            {
                try
                {
                    WallpaperChangerTools.Log("Iteration " + counter.ToString(), MessageTypes.Information, "RunServiceLoop");
                    if (cancellation.IsCancellationRequested)
                    {
                        break;
                    }
                    if (counter == 3)
                    {
                        throw new Exception("There was a problem");
                    }
                    interval = waitAfterSuccessfulInterval;
                }
                catch (Exception ex)
                {
                    WallpaperChangerTools.Log(ex.Message, MessageTypes.Error, "RunServiceLoop");
                    interval = waitAfterErrorInterval;
                }
                if (++counter>5)
                {
                    cts.Cancel();
                }
            }
        }
    }
}
