using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;

namespace WcfWebService
{
    public class Company
    {
        public List<Employee> Employees { get; set; }
    }
 
    public class Employee
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string Vendor { get; set; }
        public string MachineName { get; set; }
        public string InstallDate { get; set; }
        public string CreatedDate { get; set; }
    }

    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select Service1.svc or Service1.svc.cs at the Solution Explorer and start debugging.
    public class Service1 : IService1
    {
        static private string GetConnectionString()
        {
            return "Data Source=.\\SQLEXPRESS;Initial Catalog=ApplicationDB;Integrated Security=true;";
        }
       
        public void LoadData(string jsonstring)
        {
            var ManagementObj = new System.Web.Script.Serialization.JavaScriptSerializer().Deserialize<List<Employee>>(jsonstring);

            string connectionString = GetConnectionString();
            string sqlIns = "INSERT INTO Application (Name, InstallDate, Version,Vendor,CreatedDate,MachineName,ApplicationID) VALUES (@Name, @InstallDate, @Version, @Vendor,@CreatedDate,@MachineName,@ApplicationID)";

            using (SqlConnection connection = new SqlConnection())
            {
                connection.ConnectionString = connectionString;
                connection.Open();

                try
                {
                    SqlCommand cmdIns = new SqlCommand(sqlIns, connection);

                    foreach (Employee mo in ManagementObj)
                    {
                        Guid ApplicationID = Guid.NewGuid();

                        cmdIns.Parameters.Add("@ApplicationID", ApplicationID);
                        cmdIns.Parameters.Add("@Name", mo.Name);
                        cmdIns.Parameters.Add("@InstallDate", mo.InstallDate);
                        cmdIns.Parameters.Add("@Version", mo.Version);
                        cmdIns.Parameters.Add("@Vendor", mo.Vendor);
                        cmdIns.Parameters.Add("@CreatedDate", mo.CreatedDate);
                        cmdIns.Parameters.Add("@MachineName", Environment.MachineName);

                        cmdIns.ExecuteNonQuery();
                        cmdIns.Parameters.Clear();
                    }
                }

                catch (Exception ex)
                {
                    throw new Exception(ex.ToString(), ex);
                }
                finally
                {
                    connection.Close();
                }
            }        
        }
    }
}
