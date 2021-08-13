using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LWC.WallpaperChangerService
{
    public static class WallpaperChangerTools
    {
        private static ILog log = log4net.LogManager.GetLogger(typeof(WallpaperChanger));
        private static readonly MessageTypes minLogLevel = (MessageTypes)Enum.Parse(typeof(MessageTypes), ConfigurationManager.AppSettings["MinimumLogLevel"].ToString());
        public static void Log(string message, MessageTypes type, string source="")
        {
            try
            {
                if (type >= minLogLevel)
                {
                    switch (type)
                    {
                        case MessageTypes.Debug:
                            log.Debug(message);
                            break;
                        case MessageTypes.Warning:
                            log.Warn(message);
                            break;
                        case MessageTypes.Error:
                            log.Error(message);
                            break;
                        case MessageTypes.Critical:
                            log.Fatal(message);
                            break;
                        default:
                            log.Info(message);
                            break;
                    }
                }
                //string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
                //if (!Directory.Exists(path))
                //{
                //    Directory.CreateDirectory(path);
                //}
                //string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\" + DateTime.Now.Date.ToString("yyyy-MM-dd") + ".txt";
                //if (!File.Exists(filepath))
                //{
                //    // Create a file to write to.   
                //    using (StreamWriter sw = File.CreateText(filepath))
                //    {
                //        sw.WriteLine("Timestamp\tSeverity\tSource\tMessage");
                //    }
                //}

                //string line = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss.fff") + "\t" + type.ToString() + "\t" + source + "\t" + message.Replace("\t", "    ");
                //using (StreamWriter sw = File.AppendText(filepath))
                //{
                //        sw.WriteLine(line);                
                //}
            }
            catch
            {
                // Do nothing (File might be in use by another process...don't kill the service for the sake of logging)
            }
        }

        public static void AddConfiguration(string key, string value)
        {
            
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings.Add(key,value);
            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }

        public static string ReadConfigurationString(string key, string defaultValue="")
        {
            if (ConfigurationManager.AppSettings[key] == null)
            {
                WallpaperChangerTools.AddConfiguration(key, defaultValue);
            }
            return ConfigurationManager.AppSettings[key].ToString();
        }

        public static int ReadConfigurationInt(string key, int defaultValue = 0)
        {
            if (ConfigurationManager.AppSettings[key] == null)
            {
                WallpaperChangerTools.AddConfiguration(key, defaultValue.ToString());
            }
            int.TryParse(ConfigurationManager.AppSettings[key].ToString(), out defaultValue);
            return defaultValue;
        }
    }
}
