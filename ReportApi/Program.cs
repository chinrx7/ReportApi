using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using System.Web.Http.SelfHost;
using System.IO;
using Microsoft.Owin.Hosting;
using Microsoft.Extensions.Configuration;
using System.Runtime.ConstrainedExecution;
using System.Diagnostics;
using System.ServiceProcess;
using System.Runtime.CompilerServices;
using Owin;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging.Configuration;
using Microsoft.Extensions.Logging.EventLog;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ReportApi
{
    class Program
    {
        public static IConfigurationRoot Configuration { get; set; }
        public static Config _set;

        public static string path;
        public class Config
        {
            public string Host { get; set; }
            public string CHMIHOST { get; set; }
            public string OUTPUT { get; set; }
            public string User { get; set; }
            public string Pass { get; set; }
            public string Report { get; set; }
            public bool Debug { get; set; }
        }

        

        static void Main(string[] args)
        {
            IHost host = Host.CreateDefaultBuilder(args)
                   .ConfigureServices(services =>
                   {
                       services.AddHostedService<Worker>();

                   })
                   .UseWindowsService()
                   .Build();

            host.Run();
        }

        public static void StartWork()
        {
            Util.Logging("Report API", "Starting");
            LoadConfig();

            //genReport.GetAllReport();

            string domainAddress = "http://" + _set.Host;

            using (WebApp.Start(url: domainAddress))
            {
                //Console.WriteLine("Service Hosted " + domainAddress);
                Util.Logging("Report API", "Started " + domainAddress);
                System.Threading.Thread.Sleep(-1);
            }


            Util.Logging("Info", "Started");
        }



        static void LoadConfig()
        {
            string exe = Process.GetCurrentProcess().MainModule.FileName;
            path = Path.GetDirectoryName(exe);
            Util.Logging("Info", path);

            var builder = new ConfigurationBuilder()
               .SetBasePath(Directory.GetCurrentDirectory())
               .AddJsonFile(path + "\\config.json", optional: false);

            Configuration = builder.Build();

            var _setting = Configuration.GetSection("Setting").Get<Config>();
            _set = _setting;
            Util.Logging("Report API", "Configuration Loaded");
        }

    }
}
