using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceApp
{
    public partial class AttendanceForm : Form
    {
        public AttendanceForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Get the file we are going to process
            //var existingFile = new FileInfo("D:\\AttendanceApp\\DB\\MasterAttendanceRegister.xlsx");

            //Display the date in Short Format
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            
            this.StartPosition = FormStartPosition.CenterScreen;
                        
            // Open and read the XlSX file.
            //using (var package = new ExcelPackage(existingFile))
            //{
            //    // Get the work book in the file
            //    ExcelWorkbook workBook = package.Workbook;
            //    if (workBook != null)
            //    {
            //        foreach (ExcelWorksheet xlworksheet in workBook.Worksheets)
            //        {
            //            comboBox1.Items.Add(xlworksheet.Name);
            //        }
            //    }
            //}
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {                                
            string theDate = dateTimePicker1.Value.ToString("dd");

            DateTime dt = dateTimePicker1.Value;

            try
            {
                if (isWeekend(dt))
                {
                    MessageBox.Show("Selected date is Weekend Date");
                }
                else
                {
                    string conString = @"Provider=Microsoft.JET.OLEDB.4.0; data source=\\192.168.105.26\share\CardV3.mdb";

                    // create an open the connection     
                    OleDbConnection conn = new OleDbConnection(conString);

                    conn.Open();

                    int startRow = 3;
                    int columnNo=0;
                    int columnLOP = 38;
                    int columnWFH = 39;

                    //// Get the file we are going to process
                    var existingFile = new FileInfo("D:\\AttendanceApp\\DB\\MasterAttendanceRegister.xlsx");
                    //var existingFile = new FileInfo(@"\\192.168.105.26\\share\\MasterAttendanceRegister.xlsx");

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
                                    
                                //Skip first four columns
                                columnNo = 4 + Convert.ToInt32(theDate);

                                // read each row from the start of the data (start row + 1 header row) to the end of the spreadsheet.
                                for (int rowNumber = 3; rowNumber <= currentWorksheet.Dimension.End.Row; rowNumber++)
                                {
                                    if (currentWorksheet.Cells[startRow, 1].Value != null)
                                    {
                                        // read some data
                                        String colEmpId = currentWorksheet.Cells[startRow, 1].Value.ToString(); 
                                        String colActive = currentWorksheet.Cells[startRow, 4].Value.ToString();

                                        if ((colEmpId != null) && (colActive.ToString() == "A"))
                                        {
                                            if (GetAttendInfo(colEmpId, dateTimePicker1.Value.ToString("MM/dd/yy"), conn))
                                            {
                                                var tuple = GetTimeInfo(colEmpId, dateTimePicker1.Value.ToString("MM/dd/yy"), conn);
                                                int a = tuple.Item1;
                                                int b = tuple.Item2;

                                                if (tuple.Item1 == 0)
                                                {
                                                    currentWorksheet.Cells[startRow, columnNo].Value = "P";
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                                }
                                                else if (tuple.Item1 >= 8)
                                                {
                                                    currentWorksheet.Cells[startRow, columnNo].Value = "P";
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                                }
                                                else
                                                {
                                                    currentWorksheet.Cells[startRow, columnNo].Value = String.Format("{0:0.0}", tuple.Item1 + "." + tuple.Item2);     
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                    currentWorksheet.Cells[startRow, columnNo].Style.Fill.BackgroundColor.SetColor(Color.LightBlue);                                                    
                                                }
                                            }
                                            else
                                            {
                                                currentWorksheet.Cells[startRow, columnNo].Value = "A";
                                                currentWorksheet.Cells[startRow, columnNo].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                currentWorksheet.Cells[startRow, columnNo].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                            }
                                        }

                                        //Leave Info
                                        object m2 = currentWorksheet.Cells[rowNumber, 38].Value;
                                        object formula1 = currentWorksheet.Cells[rowNumber, 38].Formula;
                                        int LeaveCount = 0;
                                        LeaveCount = Convert.ToInt32(m2);

                                        if (LeaveCount > 5)
                                        {
                                            currentWorksheet.Cells[startRow, columnLOP].Value = LeaveCount;
                                            currentWorksheet.Cells[startRow, columnLOP].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            currentWorksheet.Cells[startRow, columnLOP].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                            currentWorksheet.Cells[startRow, columnLOP].Formula = formula1.ToString();
                                        }
                                        else
                                        {
                                            currentWorksheet.Cells[startRow, columnLOP].Value = LeaveCount;
                                            currentWorksheet.Cells[startRow, columnLOP].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            currentWorksheet.Cells[startRow, columnLOP].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                            currentWorksheet.Cells[startRow, columnLOP].Formula = formula1.ToString();
                                        }

                                        //Work from home Update
                                        DateTime now = DateTime.Now;
                                        int NumberDays = now.Day;
                                        object m = currentWorksheet.Cells[rowNumber, 39].Value;
                                        object formula2 = currentWorksheet.Cells[rowNumber, 39].Formula;
                                        int WFHCount = 0;
                                        WFHCount = Convert.ToInt32(m); 

                                        if (WFHCount > 2)
                                        {
                                            currentWorksheet.Cells[startRow, columnWFH].Value = WFHCount;
                                            currentWorksheet.Cells[startRow, columnWFH].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            currentWorksheet.Cells[startRow, columnWFH].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                            currentWorksheet.Cells[startRow, columnWFH].Formula = formula2.ToString();
                                        }
                                        else
                                        {
                                            currentWorksheet.Cells[startRow, columnWFH].Value = WFHCount;
                                            currentWorksheet.Cells[startRow, columnWFH].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            currentWorksheet.Cells[startRow, columnWFH].Style.Fill.BackgroundColor.SetColor(Color.LightCyan);
                                            currentWorksheet.Cells[startRow, columnWFH].Formula = formula2.ToString();
                                        }                                      
                                    }

                                    startRow = startRow + 1;
                                }
                            }
                            
                            package.Save();

                            MessageBox.Show("Attendance Excel Updated.");
                        }
                    }                  
                }
            }

            catch (IOException ioEx)
            {
                if (!String.IsNullOrEmpty(ioEx.Message))
                {
                    if (ioEx.Message.Contains("because it is being used by another process."))
                    {
                        MessageBox.Show("Could not read attendance data. Please make sure it not open in Excel.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured while reading attendance data.");
            }
            finally
            {
               
            }
        }

        public static bool isWeekend(DateTime dtToValidate)
        {
            if (dtToValidate.DayOfWeek == DayOfWeek.Sunday || dtToValidate.DayOfWeek == DayOfWeek.Saturday)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool GetEntryExitInfo(String CardHoldNo, String AttendDate, OleDbConnection conn)
        {
            Boolean bFlag = false;
            DateTime dt = Convert.ToDateTime(AttendDate);

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

            if ((dtEntry == null) || (dtExit == null))
            {
                bFlag = true;
            }           

            return bFlag;

        }

        //public Double GetTimeInfo(String CardHoldNo, String AttendDate, OleDbConnection conn)
        //{
        //    Double Indicator = -1;

        //    DateTime dt = Convert.ToDateTime(AttendDate);

        //    DataSet dsEntry = new DataSet();
        //    string QueryEntry = "SELECT * FROM IOData WHERE IOStatus='Entry' AND HOLDERNAME='" + CardHoldNo + "' AND IODate =#" + dt.ToShortDateString() + "# ORDER BY IOTIME ASC";
        //    OleDbDataAdapter adpEntry = new OleDbDataAdapter(QueryEntry, conn);
        //    adpEntry.Fill(dsEntry);

        //    DataSet dsExit = new DataSet();
        //    string QueryExit = "SELECT * FROM IOData WHERE IOStatus='Exit' AND HOLDERNAME='" + CardHoldNo + "' AND IODate =#" + dt.ToShortDateString() + "# ORDER BY IOTIME DESC";
        //    OleDbDataAdapter adpExit = new OleDbDataAdapter(QueryExit, conn);
        //    adpExit.Fill(dsExit);

        //    DataTable dtEntry = dsEntry.Tables[0];
        //    DataTable dtExit = dsExit.Tables[0];

        //    Double TotalHours = 0;

        //    if ((dtEntry == null) || (dtExit == null))
        //    {
        //        Indicator = 0;
        //    }
        //    else if ((dtEntry != null) && (dtExit != null))
        //    {
        //        TimeSpan ts = TimeSpan.Zero;
        //        if ((dtEntry.Rows.Count > 0) && (dtExit.Rows.Count > 0))
        //        {
        //            DateTime dt1 = DateTime.Parse((dtEntry.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
        //            DateTime dt2 = DateTime.Parse((dtExit.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
        //            ts = dt2.Subtract(dt1);
                    
        //            TotalHours =  ts.TotalHours;
        //        }

        //        if (TotalHours >= 8)
        //            Indicator = 1;
        //        else
        //            Indicator = TotalHours;
        //    }
            
        //    return Indicator;
        //}

        static Tuple<int, int> GetTimeInfo(String CardHoldNo, String AttendDate, OleDbConnection conn)
        {
            DateTime dt = Convert.ToDateTime(AttendDate);

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

            if ((dtEntry == null) || (dtExit == null))
            {
                return new Tuple<int, int>(0, 0);
            }
            else if ((dtEntry != null) && (dtExit != null))
            {
                if ((dtEntry.Rows.Count > 0) && (dtExit.Rows.Count > 0))
                {
                    DateTime dt1 = DateTime.Parse((dtEntry.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
                    DateTime dt2 = DateTime.Parse((dtExit.Rows[0]["IOTime"]).ToString(), new DateTimeFormatInfo());
                    ts = dt2.Subtract(dt1);
                }
            }

            return new Tuple<int, int>(ts.Hours, ts.Minutes);
        }

        public bool GetAttendInfo(String CardHoldNo, String AttendDate, OleDbConnection conn)
        {
            //if (Convert.ToDateTime(AttendDate) > DateTime.Now)
            //    return true;

            DateTime dt = DateTime.ParseExact(AttendDate, "MM/dd/yy", CultureInfo.InvariantCulture);

            if (isWeekend(dt))
                return true;

            Boolean bFlag = false;

            // create the DataSet
            DataSet ds = new DataSet();
                       
            // create the adapter and fill the DataSet
            string Query = "SELECT * FROM IOData WHERE HOLDERNAME='" + CardHoldNo + "' AND Format(IODate,'MM/DD/YY')='" + AttendDate + "'";
            OleDbDataAdapter adapter = new OleDbDataAdapter(Query, conn);

            adapter.Fill(ds);

            if (ds.Tables[0].Rows.Count > 0)
                bFlag = true;
            else
                bFlag = false;

            return bFlag;
        }

        public int GetWFHInfo(int rowNumber, ExcelWorksheet currentWorksheet)
        {
            DateTime now = DateTime.Now;
            int NumberDays = now.Day;
                       
            object m = currentWorksheet.Cells[rowNumber, 36].Value;
            int wfhCount = Convert.ToInt32(m); 
                      
            return wfhCount;
        }

        public int GetLeaveInfo(int rowNumber, ExcelWorksheet currentWorksheet)
        {
            DateTime now = DateTime.Now;
            int NumberDays = now.Day;

            int LopCount = 0;
            for (int columnNo = 5; columnNo <= NumberDays+1; columnNo++)
            {
                if ((currentWorksheet.Cells[rowNumber, columnNo].Value != null) && (currentWorksheet.Cells[rowNumber, columnNo].Value.ToString() == "A"))
                {
                    LopCount++;
                }
            }
            return LopCount;
        }

        
        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            int NoDays = now.Day;
            String str1 = "here";
            string Str2 = str1;
        }                                  
    }
}
