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
using OfficeOpenXml;
using System.Data.OleDb;
using System.Globalization;

namespace AttendanceService
{
    public partial class AttendanceService : ServiceBase
    {
        public AttendanceService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("AttendanceService started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("AttendanceService stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                this.WriteToFile("AttendanceService Mode: " + mode + " {0}");

                //Load Data
                LoadSoftwareList();

                // Process the list of files found in the directory. 
                string targetDirectory = ConfigurationManager.AppSettings["DestinationPath"];
                string[] fileEntries = Directory.GetFiles(targetDirectory);

                foreach (string fileName in fileEntries)
                    UploadAttendanceFile(fileName);

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

                this.WriteToFile("AttendanceService scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("AttendanceService Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("AttendanceService"))
                {
                    serviceController.Stop();
                }
            }
        }

        public void LoadSoftwareList()
        {
            try
            {
                string conString = @"Provider=Microsoft.JET.OLEDB.4.0; data source=\\252.468.605.226\share\CardV3.mdb";

                // Create an open the connection     
                OleDbConnection conn = new OleDbConnection(conString);

                String TemplatePath = ConfigurationManager.AppSettings["TemplatePath"];
                var existingFile = new FileInfo(TemplatePath);
                String FileName = ConfigurationManager.AppSettings["DestinationPath"] + "PeopleWorks" + DateTime.Now.ToString("ddmyyyyhhmm") + ".xlsx";

                // Open and read the XlSX file.
                using (var package = new ExcelPackage(existingFile))
                {
                    // Get the work book in the file
                    ExcelWorkbook workBook = package.Workbook;
                    if (workBook != null)
                    {
                        if (workBook.Worksheets.Count > 0)
                        {
                            // Get the first worksheet
                            ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

                            for (int rowNumber = 2; rowNumber <= currentWorksheet.Dimension.End.Row; rowNumber++)
                            {
                                String colEmpId = currentWorksheet.Cells[rowNumber, 3].Value.ToString();
                                var EntryExitData=GetEntryExitInfo(colEmpId, conn);
                                var EntryInfo= EntryExitData.Item1;
                                var ExitInfo = EntryExitData.Item2;
                                var WorkHrs = EntryExitData.Item3;

                                currentWorksheet.Cells[rowNumber, 2].Value = DateTime.Now.ToString("dd-MMM-yyyy");
                                currentWorksheet.Cells[rowNumber, 7].Value = EntryInfo;
                                currentWorksheet.Cells[rowNumber, 9].Value = ExitInfo;
                                currentWorksheet.Cells[rowNumber, 12].Value = WorkHrs;
                            }
                        }

                        Byte[] bin = package.GetAsByteArray();
                        File.WriteAllBytes(FileName, bin);
                    }
                }     
            }
            catch (IOException ioEx)
            {
                if (!String.IsNullOrEmpty(ioEx.Message))
                {
                    if (ioEx.Message.Contains("because it is being used by another process."))
                    {
                        this.WriteToFile("AttendanceService could not read attendance data. Please make sure it not open in Excel");
                    }
                }
            }
        }
                
        public Tuple<string, string, string> GetEntryExitInfo(String CardHoldNo, OleDbConnection conn)
        {
            DateTime dt = DateTime.Now.AddDays(-1);

            DataSet dsEntry = new DataSet();
            string QueryEntry = "SELECT * FROM IOData WHERE IOStatus='Entry' AND HOLDERNAME='" + CardHoldNo + "' AND IODate =#" + dt.ToShortDateString() + "# ORDER BY IOTIME ASC";
            OleDbDataAdapter adpEntry = new OleDbDataAdapter(QueryEntry, conn);
            adpEntry.Fill(dsEntry);

            DataSet dsExit = new DataSet();
            string QueryExit = "SELECT * FROM IOData WHERE IOStatus='Exit' AND HOLDERNAME='" + CardHoldNo + "' AND IODate =#" + dt.ToShortDateString() + "# ORDER BY IOTIME DESC";
            OleDbDataAdapter adpExit = new OleDbDataAdapter(QueryExit, conn);
            adpExit.Fill(dsExit);

            DataTable dtEntry = dsEntry.Tables[0];
            DataTable dtExit = dsExit.Tables[0];
            
            TimeSpan ts = TimeSpan.Zero;
            if ((dtEntry.Rows.Count == 0) && (dtExit.Rows.Count == 0))
            {
                return new Tuple<string, string, string>("0", "0", "0");
            }
            else if ((dtEntry.Rows.Count == 0) || (dtExit.Rows.Count == 0))
            {
                return new Tuple<string, string, string>("9.30", "0", "0");
            }
            else if ((dtEntry.Rows.Count > 0) && (dtExit.Rows.Count > 0))
            {
                DateTime dt1 = DateTime.Parse((dtEntry.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
                DateTime dt2 = DateTime.Parse((dtExit.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
                ts = dt2.Subtract(dt1);
            }

            Tuple<string, string, string> tuple = new Tuple<string, string, string>((dtEntry.Rows[0]["IOTime"]).ToString(), (dtExit.Rows[0]["IOTime"]).ToString(), ts.Hours.ToString() + ":" + ts.Minutes.ToString());

            return tuple;
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
            this.WriteToFile("AttendanceService Service Log: {0}");
            this.ScheduleService();
        }

        private void WriteToFile(string text)
        {
            string path = "C:\\AttendanceServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }

        public void UploadAttendanceFile(string fileName)
        {
            Stream requestStream = null;
            FileStream fileStream = null;
            FtpWebResponse response = null;

            try
            {
                // Get the object used to communicate with the server.
                FtpWebRequest request = (FtpWebRequest)FtpWebRequest.Create("ftp://169.225.244.282.564.345/" + Path.GetFileName(fileName));
                request.Method = WebRequestMethods.Ftp.UploadFile;

                // This example assumes the FTP site uses anonymous logon.
                request.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["ftpUserID"], ConfigurationManager.AppSettings["ftpPassword"]);

                // Copy the contents of the file to the request stream.
                StreamReader sourceStream = new StreamReader(fileName);
                byte[] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
                sourceStream.Close();
                request.ContentLength = fileContents.Length;

                requestStream = request.GetRequestStream();
                requestStream.Write(fileContents, 0, fileContents.Length);
                requestStream.Close();

                response = (FtpWebResponse)request.GetResponse();

                File.Copy(fileName, ConfigurationManager.AppSettings["ArchivePath"] + Path.GetFileName(fileName), true);
                File.Delete(fileName);

                response.Close();
            }
            catch (UriFormatException ex)
            {
            }
            catch (IOException ex)
            {
            }
            catch (WebException ex)
            {
            }
            finally
            {
                if (response != null)
                    response.Close();
                if (fileStream != null)
                    fileStream.Close();
                if (requestStream != null)
                    requestStream.Close();
            }
        }
    }
}
