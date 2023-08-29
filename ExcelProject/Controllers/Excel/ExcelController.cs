using ClosedXML.Excel;
using ExcelProject.Models.Excel;
using ExcelProject.Utility;
using Microsoft.AspNetCore.Mvc;
using MySql.Data.MySqlClient;
using System.Data;
using System.Data.OleDb;
using System.Text;
using MySqlCommand = MySql.Data.MySqlClient.MySqlCommand;

namespace ExcelProject.Controllers.Excel
{

    public class ExcelController : Controller
    {
        //IEnumerable<ExcelCustomer> customers = customerDAL.GetAllExcelCustomers();
         
        //private List<Customer> customers = new List<Customer>();
        private readonly string connectionString;
        DAL.DAL x = null;
        public ExcelController()
        {
            connectionString = ConnectionString.CName;
            x = new DAL.DAL();
        }
        public ActionResult Index()
        { 
            return View(x.GetAllCustomers(connectionString));
        }
        public IActionResult ImportExcelFile() 
        {
            return View();
        }
        public IActionResult CSV() 
        {
            List<Customer> customers = x.GetAllCustomers(connectionString);
            var builder = new StringBuilder();
            builder.AppendLine("CustomerCode,First Name,Last Name,Gender,Country,Age");
            foreach (var cus in customers)
            {
                builder.AppendLine($"{cus.CustomerCode},{cus.FirstName},{cus.LastName},{cus.Gender},{cus.Country},{cus.Age}");
            }
            return File(Encoding.UTF8.GetBytes(builder.ToString()),"text/csv","EmployeeInfoCSV.csv");
        }
        public IActionResult EXCEL()
        {
            List<Customer> customers = x.GetAllCustomers(connectionString);
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Customers");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "CustomerCode";
                worksheet.Cell(currentRow, 2).Value = "FirstName";
                worksheet.Cell(currentRow, 3).Value = "LastName";
                worksheet.Cell(currentRow, 4).Value = "Gender";
                worksheet.Cell(currentRow, 5).Value = "Country";
                worksheet.Cell(currentRow, 6).Value = "Age";

                foreach (var cus in customers)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = cus.CustomerCode;
                    worksheet.Cell(currentRow, 2).Value = cus.FirstName;
                    worksheet.Cell(currentRow, 3).Value = cus.LastName;
                    worksheet.Cell(currentRow, 4).Value = cus.Gender;
                    worksheet.Cell(currentRow, 5).Value = cus.Country;
                    worksheet.Cell(currentRow, 6).Value = cus.Age;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmployeeInfoEXCEL.xlsx");
                }
            }



        }

        [HttpPost]
        public IActionResult ImportExcelFile(IFormFile formFile)
        {
            try
            {
                var mainpath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcelFile");
                if (!Directory.Exists(mainpath))
                {
                    Directory.CreateDirectory(mainpath);
                }
                    var filePath = Path.Combine(mainpath, formFile.FileName);
                    using (FileStream stream = new FileStream(filePath, FileMode.Create))
                    {
                        formFile.CopyTo(stream);
                    }
                    var fileName = formFile.FileName;
                    string extension = Path.GetExtension(fileName);
                    string conString = string.Empty;
                    switch (extension)
                    {
                        case ".xls":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0;; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                        case ".xlsx":
                            conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filePath + ";Extended Properties='Excel 8.0; HDR=Yes'";
                            break;
                        case ".csv":
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + Path.GetDirectoryName(filePath) + ";Extended Properties='Text; HDR=Yes; FMT=CSVDelimited'";
                        break;
                }
                    DataTable dt = new DataTable();
                    conString = string.Format(conString, filePath);
                    using (OleDbConnection conExcel = new OleDbConnection(conString))
                    {
                        using (OleDbCommand cmdExcel = conExcel.CreateCommand())
                        {
                            using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                            {
                                cmdExcel.Connection = conExcel;
                                conExcel.Open();
                                DataTable dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                                odaExcel.SelectCommand = cmdExcel;
                                odaExcel.Fill(dt);
                                conExcel.Close();
                            }
                        }
                    }
                    conString = connectionString;
                    MySqlConnection con = new MySqlConnection(conString);
                    Customer customer = new Customer();
                    con.Open();
                    foreach(DataRow row in dt.Rows)
                    {
                        MySqlCommand cmd = new MySqlCommand("AddCustomer", con);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("p_customercode", row["CustomerCode"]);
                        cmd.Parameters.AddWithValue("p_firstname",row["FirstName"]);
                        cmd.Parameters.AddWithValue("p_lastname",row["LastName"]);
                        cmd.Parameters.AddWithValue("p_gender",row["Gender"]);
                        cmd.Parameters.AddWithValue("p_country",row["Country"]);
                        cmd.Parameters.AddWithValue("p_age",row["Age"]);
                        cmd.ExecuteNonQuery();
                    }
                    
                    con.Close();
                    ViewBag.Message = "File Imported Successfully";
                    return RedirectToAction("Index");
                
            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
            return View();
        }
    }
}