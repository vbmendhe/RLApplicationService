using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net;
using System.Management;
using SimpleService.ServiceReference1;

namespace SimpleService
{
    public partial class ApplicationService : ServiceBase
    {
        public ApplicationService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Simple Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                this.WriteToFile("Simple Service Mode: " + mode + " {0}");

                //Load Data
                LoadSoftwareList();

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("Simple Service scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void LoadSoftwareList()
        {
            ManagementObjectCollection moReturn;
            ManagementObjectSearcher moSearch;

            moSearch = new ManagementObjectSearcher("Select * from Win32_Product");
            moReturn = moSearch.Get();

            var javaScriptSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

            List<data> _data = new List<data>();

            foreach (ManagementObject mo in moReturn)
            {
                _data.Add(new data()
                {
                    Name = mo["Name"].ToString(),
                    Vendor = mo["Vendor"].ToString(),
                    Version = mo["Version"].ToString(),
                    MachineName = Environment.MachineName,
                    InstallDate = mo["InstallDate"].ToString(),
                    CreatedDate = DateTime.Now.ToString()
                });
            }

            string path = @"C:\JasonUpload";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string json = new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(_data);
            
            System.IO.File.WriteAllText(path + "\\Jason_" + DateTime.Now.ToString("yyyyMMdd_hhss") + ".txt", json);

            bool wcfservice = IsReachableUri("http://localhost/RLApplicationService/");

            if (wcfservice)
            {
                string[] array1 = Directory.GetFiles(@"C:\JasonUpload");

                foreach (string name in array1)
                {
                    string jsoncontents = File.ReadAllText(name);

                    Service1Client client = new Service1Client();
                    client.LoadData(jsoncontents);
                    client.Close();

                    File.Delete(name);
                }
            }
        }

        private bool IsReachableUri(string uriInput)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uriInput);
            request.Timeout = 15000;
            request.Method = "HEAD"; // As per Lasse's comment
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    return response.StatusCode == HttpStatusCode.OK;
                }
            }
            catch (WebException)
            {
                return false;
            }
        }

        private void SchedularCallback(object e)
        {
            this.WriteToFile("Simple Service Log: {0}");
            this.ScheduleService();
        }

        private void WriteToFile(string text)
        {
            string path = "C:\\ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }
    }
}
