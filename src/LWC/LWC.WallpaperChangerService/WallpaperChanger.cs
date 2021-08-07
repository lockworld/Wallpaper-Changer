using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
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
                TimeSpan.FromSeconds(WallpaperChangerTools.ReadConfigurationInt("ServiceIntervalInSeconds",1));
            TimeSpan waitAfterErrorInterval = 
                TimeSpan.FromSeconds(WallpaperChangerTools.ReadConfigurationInt("AfterErrorServiceIntervalInSeconds", 5));

            int counter = 0;
            while (!cancellation.WaitHandle.WaitOne(interval))
            {
                try
                {
                    string wpdir = ConfigurationManager.AppSettings["WallpaperDirectory"].ToString();
                    string wpdest = ConfigurationManager.AppSettings["WallpaperDestination"].ToString();
                    int pctRoot = 0;
                    int.TryParse(ConfigurationManager.AppSettings["PercentSelectFromRoot"].ToString(), out pctRoot);
                    bool selectRoot = false;
                    Random rnd = new Random();
                    int chance = rnd.Next(1, 100);
                    if (chance <= pctRoot)
                    {
                        // We give users the chance to pull occasional images from the root directory.
                        // This allows them the option to have some images displayed year-round, and other images that are date-specific.
                        selectRoot = true;
                    }
                    else
                    {
                        Dictionary<string, double> wpdirList = new Dictionary<string, double>();
                        foreach (string dir in Directory.EnumerateDirectories(wpdir))
                        {
                            DirectoryInfo di = new DirectoryInfo(dir);
                            string[] dates = di.Name.Split('-');
                            int curMonth = 5;// DateTime.Now.Month;
                            int curDay = 25;// DateTime.Now.Day;
                            int minMonth = 0;
                            int maxMonth = 0;
                            int minDay = 0;
                            int maxDay = 0;
                            int piece = 1;
                            foreach (string date in dates)
                            {
                                string[] pieces = date.Split('_');
                                if (piece == 1)
                                {
                                    int.TryParse(pieces[0], out minMonth);
                                    if (pieces.Length == 1)
                                    {
                                        minDay = 0;
                                    }
                                    else if (pieces.Length == 2)
                                    {
                                        int.TryParse(pieces[1], out minDay);
                                    }
                                }
                                if (dates.Length == 1 || piece == 2)
                                {
                                    int.TryParse(pieces[0], out maxMonth);
                                    if (pieces.Length == 1)
                                    {
                                        maxDay = 32;
                                    }
                                    else if (pieces.Length == 2)
                                    {
                                        int.TryParse(pieces[1], out maxDay);
                                    }
                                }
                                piece++;
                            }
                            if (minMonth > 0 && maxMonth > 0 && minDay > 0 && maxDay > 0)
                            {
                                if (minMonth <= curMonth && maxMonth >= curMonth
                                    && minDay <= curDay && maxDay >= curDay)
                                {
                                    DateTime min = new DateTime(DateTime.Now.Year, minMonth, minDay);
                                    DateTime max = new DateTime(DateTime.Now.Year, maxMonth, maxDay);
                                    TimeSpan diff = max - min;
                                    wpdirList.Add(di.FullName, diff.TotalDays);
                                    WallpaperChangerTools.Log("Setting wallpaper directory to \"" + di.FullName + "\" with a precision of " + diff.TotalDays + ".", MessageTypes.Information);
                                }

                            }
                            else
                            {
                                WallpaperChangerTools.Log("Unable to parse directory \"" + di.FullName + "\" as wallpaper changer directory.", MessageTypes.Debug);
                            }


                        }
                        if (wpdirList.Count > 0)
                        {
                            /* 
                             Sort our list of possible directories based on the precision (The shortest date range equals the highest precision). 
                            
                            This allows users to have multiple overlapping folders and still select the source directory based on the narrowest date range.
                            
                            For example, I have a folder 07_01-07_31 and also a folder 07_17-07_23. 
                            If the current date falls in both ranges, I want the tightest range (07_17-07_23) to be selected.
                            */
                            List<KeyValuePair<string, double>> dList = wpdirList.ToList();
                            dList.Sort(delegate (KeyValuePair<string, double> pair1, KeyValuePair<string, double> pair2) { return pair1.Value.CompareTo(pair2.Value); });
                            WallpaperChangerTools.Log("Wallpaper directory selected is \"" + dList.First().Key + "\"", MessageTypes.Debug);
                        }
                        else
                        {
                            selectRoot = true;
                        }
                    }

                    if (selectRoot)
                    {
                        // We are selecting the root directory based on our percentage or we hit a problem.
                        // Either way, we will pull images from the root directory instead of a sub-directory.
                        WallpaperChangerTools.Log("Wallpaper directory is root directory of " + wpdir, MessageTypes.Debug);
                    }
                    WallpaperChangerTools.Log("Iteration " + counter.ToString(), MessageTypes.Information, "RunServiceLoop");
                    if (cancellation.IsCancellationRequested)
                    {
                        break;
                    }
                    if (counter % 3 == 0)
                    {
                        throw new Exception("There was a problem with iteration " + counter);
                    }
                    interval = waitAfterSuccessfulInterval;
                }
                catch (Exception ex)
                {
                    WallpaperChangerTools.Log(ex.Message, MessageTypes.Error, "RunServiceLoop");
                    interval = waitAfterErrorInterval;
                }
                if (++counter>9)
                {
                    cts.Cancel();
                }
            }
        }
    }
}
