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
        Dictionary<string, string> dictionary = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
            LoadEmployeeData();
        }

        public void LoadEmployeeData()
        {
             //Load EmailIDs
             var EmailListFile = new FileInfo("D:\\AttendanceApp\\DB\\ReportingManager.xlsx");
             using (var package = new ExcelPackage(EmailListFile))
             {
                 ExcelWorkbook workBook = package.Workbook;
                 if (workBook != null)
                 {
                     if (workBook.Worksheets.Count > 0)
                     {
                         ExcelWorksheet currentWorksheet = workBook.Worksheets.First();

                         for (int rowNumber = 2; rowNumber <= currentWorksheet.Dimension.End.Row; rowNumber++)
                         {
                             if (currentWorksheet.Cells[rowNumber, 1].Value != null)
                             {
                                 String MgrName = currentWorksheet.Cells[rowNumber, 3].Value.ToString();
                                 String MgrEmailId = currentWorksheet.Cells[rowNumber, 4].Value.ToString();

                                 if (!dictionary.ContainsKey(MgrName))
                                 {
                                     dictionary.Add(MgrName, MgrEmailId);
                                     comboBox1.Items.Add(MgrName);
                                 }
                             }
                         }
                     }
                 }
             }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string conString = @"Provider=Microsoft.JET.OLEDB.4.0; data source=D:\AttendanceApp\DB\CardV3.mdb";
            string month = dateTimeStart.Value.ToString("MMMM");
            string CardHoldNo = "12345";

            // create an open the connection     
            OleDbConnection conn = new OleDbConnection(conString);
            conn.Open();

            // create the DataSet
            DataSet ds = new DataSet();

            // create the adapter and fill the DataSet
            string Query = "SELECT DISTINCT IODate FROM IOData WHERE HOLDERNAME='" + CardHoldNo + "' AND  MONTH(IODate)=" + month + " AND YEAR(IODate)=2015";

            OleDbDataAdapter adapter = new OleDbDataAdapter(Query, conn);

            adapter.Fill(ds);     
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Form1.ActiveForm.Close();
        }
    }
}
