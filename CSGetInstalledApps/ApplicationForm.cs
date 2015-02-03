using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using CSGetInstalledApps.ServiceReference1;

namespace CSGetInstalledApps
{
    public partial class ApplicationForm : Form
    {
        public ApplicationForm()
        {
            InitializeComponent();
        }
       
        private void LoadSoftwareList()
        {
            bool wcfservice = IsReachableUri("http://192.168.150.5/RLApplicationService/");
            
            ManagementObjectCollection moReturn;  
            ManagementObjectSearcher moSearch;

            try
            {
                moSearch = new ManagementObjectSearcher("Select * from Win32_Product");
                moReturn = moSearch.Get();

                var javaScriptSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

                List<data> _data = new List<data>();                   

                foreach (ManagementObject mo in moReturn)
                {                        
                    _data.Add(new data()
                    {
                        Name=mo["Name"].ToString(),
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
            catch (Exception ex)
            {
                throw new Exception(ex.ToString(), ex);
            }
            finally
            {
            }

              MessageBox.Show("Softwares info loaded to Server.....");
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
        
        private void Form1_Load(object sender, EventArgs e)
        {
            //LoadSoftwareList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ApplicationForm.ActiveForm.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {          
            LoadSoftwareList();
        }
    }
}
