using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AttendanceMailMerge
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            {
                //string conString = @"Provider=Microsoft.JET.OLEDB.4.0; data source=\\192.168.105.26\share\CardV3.mdb";

                // create an open the connection     
                //OleDbConnection conn = new OleDbConnection(conString);

                //conn.Open();

                int startRow = 3;
                int columnNo = 0;
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
                            //columnNo = 4 + Convert.ToInt32(theDate);

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
    }
}
