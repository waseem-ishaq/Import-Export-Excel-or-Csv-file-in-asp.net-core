using ExcelProject.Models.Excel;
using MySqlConnector;
using System.Data;

namespace ExcelProject.DAL
{
    public class DAL
    {
        public List<Customer> GetAllCustomers(string con)
        {
            MySqlDataAdapter da = new MySqlDataAdapter("usp_GetAllCustomers", con);
            da.SelectCommand.CommandType = CommandType.StoredProcedure;
            DataTable dt = new DataTable();
            da.Fill(dt);
            List<Customer> lstCustomer = new List<Customer>();
            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Customer com = new Customer();
                    com.Id = Convert.ToInt32(dt.Rows[i]["id"]);
                    com.CustomerCode = Convert.ToInt32(dt.Rows[i]["CustomerCode"]);
                    com.FirstName = dt.Rows[i]["FirstName"].ToString();
                    com.LastName = dt.Rows[i]["LastName"].ToString();
                    com.Gender = dt.Rows[i]["Gender"].ToString();
                    com.Country = dt.Rows[i]["Country"].ToString();
                    com.Age = Convert.ToInt32(dt.Rows[i]["Age"]);

                    lstCustomer.Add(com);
                }
            }
            if (lstCustomer.Count > 0)
            {
                return lstCustomer;
            }
            else
            {
                return null;
            }
        }
    }
}
