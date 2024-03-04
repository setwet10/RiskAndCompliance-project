using OfficeOpenXml;
using RiskComplianceApp.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.EnterpriseServices;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;
using static System.Net.Mime.MediaTypeNames;

namespace RiskComplianceApp.Controllers
{
    public class HomeController : Controller
    {
        RiskAppEntities dbObj = new RiskAppEntities();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult UploadExcel()
        {
            return View();
        }
        //Upload Function for IBCP- working
        [HttpPost]
        public ActionResult UploadFileIBCPAberden(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                // Get the file name without extension
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);

                // Save the file to a temporary location
                var filePath = Path.Combine(Server.MapPath("~/App_Data"), file.FileName);
                file.SaveAs(filePath);

                // Create a table with the same name as the file (if not exists)
                CreateTableIfNotExists(fileNameWithoutExtension, filePath);

                // Process the Excel file data
                //ProcessExcelData(filePath, fileNameWithoutExtension);

                // Optionally: Delete the temporary file
                System.IO.File.Delete(filePath);

                // Redirect or return a success message
                return RedirectToAction("Index", "Home");
            }

            // Handle invalid file or other error scenarios
            return View("Error");
        }

        private void CreateTableIfNotExists(string tableName, string filePath)
        {


            string connectionString = ConfigurationManager.ConnectionStrings["YourConnectionStringName"]?.ConnectionString;

            // Check if the connection string is not null or empty
            if (!string.IsNullOrEmpty(connectionString))
            {
                // Create a new SqlConnection object with the initialized connection string
                using (var sqlConnection = new SqlConnection(connectionString))
                {
                    // Open the SQL connection
                    sqlConnection.Open();

                    // Rest of your code to interact with the database goes here...
                }
            }
            else
            {
                // Handle the case where the connection string is not found or empty
                throw new Exception("Connection string is not found or empty.");
            }



            // Check if the table exists
            if (TableExists(tableName, connectionString))
            {
                // Table exists, proceed to import data
                ImportDataIntoExistingTable(tableName, connectionString, filePath);
            }
            else
            {
                // Table doesn't exist, create the table
                //CreateTable(tableName, connectionString);

                // Import data into the created table
                ImportDataIntoExistingTable(tableName, connectionString, filePath);
            }
        }


        private bool TableExists(string tableName, string connectionString)
        {
            // Check if the table exists in the database
            string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    int count = (int)command.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void CreateTable(string tableName, string connectionString)
        {
            // Create the table if it does not exist
            string createTableQuery = $@"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
        BEGIN
            CREATE TABLE {tableName} (
                ID INT PRIMARY KEY IDENTITY(1,1),
                Column1 NVARCHAR(MAX),
                Column2 NVARCHAR(MAX)
                -- Add other columns as needed
            );
        END";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private List<IBCP> GetExcelData(string filePath)
        {
            List<IBCP> excelData = new List<IBCP>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    bool isActive;
                    bool.TryParse(worksheet.Cells[row, 43].Value?.ToString(), out isActive);
                    IBCP data = new IBCP
                    {
                        Sr_No = Convert.ToInt32(worksheet.Cells[row, 1].Value),
                        Vertical = worksheet.Cells[row, 2].Value?.ToString(),
                        Account = worksheet.Cells[row, 3].Value?.ToString(),
                        Process = worksheet.Cells[row, 4].Value?.ToString(),
                        Sub_Process = worksheet.Cells[row, 5].Value?.ToString(),
                        Activity = worksheet.Cells[row, 6].Value?.ToString(),
                        Head_Count = worksheet.Cells[row, 7].Value?.ToString(),
                        Applications_software = worksheet.Cells[row, 8].Value?.ToString(),
                        Volume = worksheet.Cells[row, 9].Value?.ToString(),
                        Frequency_daily_weekly_monthly = worksheet.Cells[row, 10].Value?.ToString(),
                        Any_Volume_trends = worksheet.Cells[row, 11].Value?.ToString(),
                        SOP_Available = worksheet.Cells[row, 12].Value?.ToString(),
                        No_of_SOP = worksheet.Cells[row, 13].Value?.ToString(),
                        Activity_Description = worksheet.Cells[row, 14].Value?.ToString(),
                        Supplier_who_is_sending = worksheet.Cells[row, 15].Value?.ToString(),
                        Input_info_needs_to_be_processed = worksheet.Cells[row, 16].Value?.ToString(),
                        Process_how_it_is_actually_done = worksheet.Cells[row,17].Value?.ToString(),
                      
                        Output_what_is_the_output_storage = worksheet.Cells[row, 18].Value?.ToString(),
                        Customer_end_client_Onshore_Or_biz = worksheet.Cells[row, 19].Value?.ToString(),
                        SLA_Accuracy_Timelines = worksheet.Cells[row, 20].Value?.ToString(),
                        SLA_Target = worksheet.Cells[row, 21].Value?.ToString(),
                        No_Of_errors = worksheet.Cells[row, 22].Value?.ToString(),
                        Type_of_Errors = worksheet.Cells[row, 23].Value?.ToString(),
                        Risk_Description = worksheet.Cells[row, 24].Value?.ToString(),
                        Type_of_Risk_Financial_Non_Financial_Regulatory = worksheet.Cells[row, 25].Value?.ToString(),
                        Risk_Statement = worksheet.Cells[row, 26].Value?.ToString(),
                        Likelihood = worksheet.Cells[row, 27].Value?.ToString(),
                        Impact = worksheet.Cells[row, 28].Value?.ToString(),
                        Risk_Score = worksheet.Cells[row, 29].Value?.ToString(),
                        Control_Name = worksheet.Cells[row, 30].Value?.ToString(),
                        Control_Description = worksheet.Cells[row, 31].Value?.ToString(),
                        Control_Effectiveness = worksheet.Cells[row, 32].Value?.ToString(),
                        Design_Effectiveness = worksheet.Cells[row, 33].Value?.ToString(),
                        Control_Owner = worksheet.Cells[row, 34].Value?.ToString(),
                        Type_of_control_Preventive_Manual_Detective_Automated_Process_People = worksheet.Cells[row, 35].Value?.ToString(),

                        Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 36].Value?.ToString(),
                        Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 37].Value?.ToString(),


                        Test_steps_Methodology = worksheet.Cells[row, 38].Value?.ToString(),
                        Testing_Objective = worksheet.Cells[row, 39].Value?.ToString(),
                        What_to_Look_at = worksheet.Cells[row, 40].Value?.ToString(),
                        What_to_Look_for = worksheet.Cells[row, 41].Value?.ToString(),
                        What_to_report = worksheet.Cells[row, 42].Value?.ToString(),
                        // IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value),
                        IsActive = isActive,
                        // Add other properties if there are more columns in your Excel file
                    };

                    excelData.Add(data);
                }
            }

            return excelData;
        }
        private void ImportDataIntoExistingTable(string tableName, string connectionString, string filePath)
        {
            // Implement the logic to import data into the existing table
            // Modify this code based on your specific requirements

            string importQuery = $"INSERT INTO {tableName} ( Vertical,Account,Process,Sub_Process,Activity,Head_Count,Applications_software,Volume,Frequency_daily_weekly_monthly,Any_Volume_trends,SOP_Available,No_of_SOP,Activity_Description,Supplier_who_is_sending,Input_info_needs_to_be_processed,Process_how_it_is_actually_done,Output_what_is_the_output_storage,Customer_end_client_Onshore_Or_biz,SLA_Accuracy_Timelines,SLA_Target,No_Of_errors,Type_of_Errors,Risk_Description,Type_of_Risk_Financial_Non_Financial_Regulatory,Risk_Statement,Likelihood,Impact,Risk_Score,Control_Name,Control_Description,Control_Effectiveness,Design_Effectiveness,Control_Owner,Type_of_control_Preventive_Manual_Detective_Automated_Process_People,Residual_Risk_Considering_Control_Effectiveness,Residual_Risk_Considering_Design_Effectiveness,Test_steps_Methodology,Testing_Objective,What_to_Look_at,What_to_Look_for,What_to_report,IsActive) VALUES (@Vertical, @Account,@Process,@Sub_Process,@Activity,@Head_Count,@Applications_software,@Volume,@Frequency_daily_weekly_monthly,@Any_Volume_trends,@SOP_Available,@No_of_SOP,@Activity_Description,@Supplier_who_is_sending,@Input_info_needs_to_be_processed,@Process_how_it_is_actually_done,@Output_what_is_the_output_storage,@Customer_end_client_Onshore_Or_biz,@SLA_Accuracy_Timelines,@SLA_Target,@No_Of_errors,@Type_of_Errors,@Risk_Description,@Type_of_Risk_Financial_Non_Financial_Regulatory,@Risk_Statement,@Likelihood,@Impact,@Risk_Score,@Control_Name,@Control_Description,@Control_Effectiveness,@Design_Effectiveness,@Control_Owner,@Type_of_control_Preventive_Manual_Detective_Automated_Process_People,@Residual_Risk_Considering_Control_Effectiveness,@Residual_Risk_Considering_Design_Effectiveness,@Test_steps_Methodology,@Testing_Objective,@What_to_Look_at,@What_to_Look_for,@What_to_report,@IsActive)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Assuming you have a data source for the Excel data (replace this with your actual data source)
                List<IBCP> excelData = GetExcelData(filePath);

                foreach (var data in excelData)
                {
                    using (SqlCommand command = new SqlCommand(importQuery, connection))
                    {
                     //   command.Parameters.AddWithValue("@Sr_No", data.Sr_No);
                        command.Parameters.AddWithValue("@Vertical", data.Vertical);
                        command.Parameters.AddWithValue("@Account", data.Account);
                        command.Parameters.AddWithValue("@Process", data.Process);
                        command.Parameters.AddWithValue("@Sub_Process", data.Sub_Process);
                        command.Parameters.AddWithValue("@Activity", data.Activity);
                        command.Parameters.AddWithValue("@Head_Count", data.Head_Count);
                        command.Parameters.AddWithValue("@Applications_software", data.Applications_software);
                        command.Parameters.AddWithValue("@Volume", data.Volume);
                        command.Parameters.AddWithValue("@Frequency_daily_weekly_monthly", data.Frequency_daily_weekly_monthly);
                        command.Parameters.AddWithValue("@Any_Volume_trends", data.Any_Volume_trends);
                        command.Parameters.AddWithValue("@SOP_Available", data.SOP_Available);
                        command.Parameters.AddWithValue("@No_of_SOP", data.No_Of_errors);
                        command.Parameters.AddWithValue("@Activity_Description", data.Activity_Description);
                        command.Parameters.AddWithValue("@Supplier_who_is_sending", data.Supplier_who_is_sending);
                        command.Parameters.AddWithValue("@Input_info_needs_to_be_processed", data.Input_info_needs_to_be_processed);
                        command.Parameters.AddWithValue("@Process_how_it_is_actually_done", data.Process_how_it_is_actually_done);
                        command.Parameters.AddWithValue("@Output_what_is_the_output_storage", data.Output_what_is_the_output_storage);
                        command.Parameters.AddWithValue("@Customer_end_client_Onshore_Or_biz", data.Customer_end_client_Onshore_Or_biz);
                        command.Parameters.AddWithValue("@SLA_Accuracy_Timelines", data.SLA_Accuracy_Timelines);
                        command.Parameters.AddWithValue("@SLA_Target", data.SLA_Target);
                        command.Parameters.AddWithValue("@No_Of_errors", data.No_Of_errors);
                        command.Parameters.AddWithValue("@Type_of_Errors", data.Type_of_Errors);
                        command.Parameters.AddWithValue("@Risk_Description", data.Risk_Description);
                        command.Parameters.AddWithValue("@Type_of_Risk_Financial_Non_Financial_Regulatory", data.Type_of_Risk_Financial_Non_Financial_Regulatory);
                        command.Parameters.AddWithValue("@Risk_Statement", data.Risk_Statement);
                        command.Parameters.AddWithValue("@Likelihood", data.Likelihood);
                        command.Parameters.AddWithValue("@Impact", data.Impact);
                        command.Parameters.AddWithValue("@Risk_Score", data.Risk_Score);
                        command.Parameters.AddWithValue("@Control_Name", data.Control_Name);
                        command.Parameters.AddWithValue("@Control_Description", data.Control_Description);
                        command.Parameters.AddWithValue("@Control_Effectiveness", data.Control_Effectiveness);
                        command.Parameters.AddWithValue("@Design_Effectiveness", data.Design_Effectiveness);
                        command.Parameters.AddWithValue("@Control_Owner", data.Control_Owner);
                        command.Parameters.AddWithValue("@Type_of_control_Preventive_Manual_Detective_Automated_Process_People", data.Type_of_control_Preventive_Manual_Detective_Automated_Process_People);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Control_Effectiveness", data.Residual_Risk_Considering_Control_Effectiveness);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Design_Effectiveness", data.Residual_Risk_Considering_Design_Effectiveness);
                        command.Parameters.AddWithValue("@Test_steps_Methodology", data.Test_steps_Methodology);
                        command.Parameters.AddWithValue("@Testing_Objective", data.Testing_Objective);
                        command.Parameters.AddWithValue("@What_to_Look_at", data.What_to_Look_at);
                        command.Parameters.AddWithValue("@What_to_Look_for", data.What_to_Look_for);
                        command.Parameters.AddWithValue("@What_to_report", data.What_to_report);

                        command.Parameters.AddWithValue("@IsActive", data.IsActive);


                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void ProcessExcelData(string filePath, string tableName)
        {
            string connectionString = ConfigurationManager.ConnectionStrings?["YourConnectionStringName"]?.ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Read Excel data using a library like EPPlus or NPOI
                // Example using EPPlus:
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Read data from Excel and insert into the corresponding table
                        int? Sr_No = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        string Vertical = worksheet.Cells[row, 2].Value?.ToString();
                        string Account = worksheet.Cells[row, 3].Value?.ToString();


                        string Process = worksheet.Cells[row, 4].Value?.ToString();
                        string Sub_Process = worksheet.Cells[row, 5].Value?.ToString();
                        string Activity = worksheet.Cells[row, 6].Value?.ToString();
                        string Head_Count = worksheet.Cells[row, 7].Value?.ToString();
                        string Applications_software = worksheet.Cells[row, 8].Value?.ToString();
                        string Volume = worksheet.Cells[row, 9].Value?.ToString();
                        string Frequency_daily_weekly_monthly = worksheet.Cells[row, 10].Value?.ToString();
                        string Any_Volume_trends = worksheet.Cells[row, 11].Value?.ToString();
                        string SOP_Available = worksheet.Cells[row, 12].Value?.ToString();
                        string No_of_SOP = worksheet.Cells[row, 13].Value?.ToString();
                        string Activity_Description = worksheet.Cells[row, 14].Value?.ToString();
                        string Supplier_who_is_sending = worksheet.Cells[row, 15].Value?.ToString();
                        string Input_info_needs_to_be_processed = worksheet.Cells[row, 16].Value?.ToString();
                        string Output_what_is_the_output_storage = worksheet.Cells[row, 17].Value?.ToString();
                        string Customer_end_client_Onshore_Or_biz = worksheet.Cells[row, 18].Value?.ToString();
                        string SLA_Accuracy_Timelines = worksheet.Cells[row, 19].Value?.ToString();
                        string SLA_Target = worksheet.Cells[row, 20].Value?.ToString();
                        string No_Of_errors = worksheet.Cells[row, 21].Value?.ToString();
                        string Type_of_Errors = worksheet.Cells[row, 22].Value?.ToString();
                        string Risk_Description = worksheet.Cells[row, 23].Value?.ToString();
                        string Type_of_Risk_Financial_Non_Financial_Regulatory = worksheet.Cells[row, 24].Value?.ToString();
                        string Risk_Statement_I_have_added_this_column_as_required_for_RCSA = worksheet.Cells[row, 25].Value?.ToString();
                        string Likelihood = worksheet.Cells[row, 26].Value?.ToString();
                        string Impact = worksheet.Cells[row, 27].Value?.ToString();
                        string Risk_Score = worksheet.Cells[row, 28].Value?.ToString();
                        string Control_Name = worksheet.Cells[row, 29].Value?.ToString();
                        string Control_Description = worksheet.Cells[row, 30].Value?.ToString();
                        string Control_Effectiveness = worksheet.Cells[row, 31].Value?.ToString();
                        string Design_Effectiveness = worksheet.Cells[row, 32].Value?.ToString();
                        string Control_Owner = worksheet.Cells[row, 33].Value?.ToString();
                        string Type_of_control_Preventive_Manual_Detective_Automated_Process_People = worksheet.Cells[row, 34].Value?.ToString();

                        string Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 35].Value?.ToString();
                        string Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 36].Value?.ToString();


                        string Test_steps_Methodology = worksheet.Cells[row, 37].Value?.ToString();
                        string Testing_Objective = worksheet.Cells[row, 38].Value?.ToString();
                        string What_to_Look_at = worksheet.Cells[row, 39].Value?.ToString();
                        string What_to_Look_for = worksheet.Cells[row, 40].Value?.ToString();
                        string What_to_report = worksheet.Cells[row, 41].Value?.ToString();
                        bool IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value);
                        // Insert data into the corresponding table
                        //InsertDataIntoTable(connection, tableName, Sr_No, Vertical, Account);
                    }
                }
            }
        }

        private void InsertDataIntoTable(SqlConnection connection, string tableName, int? Sr_No, string Vertical, string Account, string Process, bool? IsActive)
        {
            // Implement the logic to insert data into the specified table
            // You may use parameterized queries or an ORM like Dapper or Entity Framework

            string query = $"INSERT INTO {tableName} (Sr_No, Vertical,Account) VALUES (@Sr_No, @Vertical,@Account)";

            using (SqlCommand command = new SqlCommand(query, connection))
            {


                command.Parameters.AddWithValue("@Sr_No", Sr_No);
                command.Parameters.AddWithValue("@Vertical", Vertical);
                command.Parameters.AddWithValue("@Account", Account);



                command.ExecuteNonQuery();
            }
        }



       // Templete_download_Code_Shubham
        public ActionResult DownloadTemplateKRI()
        {
            // Path to your existing template file
            string templatePath = Server.MapPath("~/Templets/KRI.xlsx");

            // Return the file
            return File(templatePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "KRI.xlsx");
        }

        public ActionResult DownloadTemplateIBCP()
        {
            // Path to your existing template file
            string templatePath = Server.MapPath("~/Templets/IBCP.xlsx");

            // Return the file
            return File(templatePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "IBCP.xlsx");
        }
        public ActionResult DownloadTemplateTRR()
        {
            // Path to your existing template file
            string templatePath = Server.MapPath("~/Templets/TRR.xlsx");

            // Return the file
            return File(templatePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TRR.xlsx");
        }
        public ActionResult DownloadTemplateNTRR()
        {
            // Path to your existing template file
            string templatePath = Server.MapPath("~/Templets/NTRR.xlsx");

            // Return the file
            return File(templatePath, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "NTRR.xlsx");
        }
        //Upload function for KRI_Shubham

        [HttpGet]
        public ActionResult UploadExcelKRI()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFileKRI(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                // Get the file name without extension
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);

                // Save the file to a temporary location
                var filePath = Path.Combine(Server.MapPath("~/App_Data"), file.FileName);
                file.SaveAs(filePath);

                // Create a table with the same name as the file (if not exists)
                CreateTableIfNotExist(fileNameWithoutExtension, filePath);

                // Process the Excel file data
                //ProcessExcelData(filePath, fileNameWithoutExtension);

                // Optionally: Delete the temporary file
                System.IO.File.Delete(filePath);

                // Redirect or return a success message
                return RedirectToAction("Index", "Home");
            }

            // Handle invalid file or other error scenarios
            return View("Error");
        }

        private void CreateTableIfNotExist(string tableName, string filePath)
        {


            string connectionString = ConfigurationManager.ConnectionStrings["YourConnectionStringName"]?.ConnectionString;

            // Check if the connection string is not null or empty
            if (!string.IsNullOrEmpty(connectionString))
            {
                // Create a new SqlConnection object with the initialized connection string
                using (var sqlConnection = new SqlConnection(connectionString))
                {
                    // Open the SQL connection
                    sqlConnection.Open();

                    // Rest of your code to interact with the database goes here...
                }
            }
            else
            {
                // Handle the case where the connection string is not found or empty
                throw new Exception("Connection string is not found or empty.");
            }



            // Check if the table exists
            if (TableExist(tableName, connectionString))
            {
                // Table exists, proceed to import data
                ImportDataIntoExistingTabl(tableName, connectionString, filePath);
            }
            else
            {
                // Table doesn't exist, create the table
                //CreateTable(tableName, connectionString);

                // Import data into the created table
                ImportDataIntoExistingTabl(tableName, connectionString, filePath);
            }
        }


        private bool TableExist(string tableName, string connectionString)
        {
            // Check if the table exists in the database
            string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    int count = (int)command.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void CreateTabl(string tableName, string connectionString)
        {
            // Create the table if it does not exist
            string createTableQuery = $@"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
        BEGIN
            CREATE TABLE {tableName} (
                ID INT PRIMARY KEY IDENTITY(1,1),
                Column1 NVARCHAR(MAX),
                Column2 NVARCHAR(MAX)
                -- Add other columns as needed
            );
        END";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private List<KRI> GetExcelDat(string filePath)
        {
            List<KRI> excelData = new List<KRI>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    bool isActive;
                    bool.TryParse(worksheet.Cells[row, 22].Value?.ToString(), out isActive);
                    KRI data = new KRI
                    {
                        Sr_NO = Convert.ToInt32(worksheet.Cells[row, 1].Value),
                        Vertical = worksheet.Cells[row, 2].Value?.ToString(),
                        Account = worksheet.Cells[row, 3].Value?.ToString(),
                        Process = worksheet.Cells[row, 4].Value?.ToString(),
                        Sub_Process = worksheet.Cells[row, 5].Value?.ToString(),
                        Activity = worksheet.Cells[row, 6].Value?.ToString(),
                        Activity_Discription = worksheet.Cells[row, 7].Value?.ToString(),
                       
                        Risk_Statement = worksheet.Cells[row, 8].Value?.ToString(),
                        Likelyhood = worksheet.Cells[row, 9].Value?.ToString(),
                        Impact = worksheet.Cells[row, 10].Value?.ToString(),
                        Risk_Score = worksheet.Cells[row, 11].Value?.ToString(),
                        Control_Name = worksheet.Cells[row, 12].Value?.ToString(),
                        Control_Description = worksheet.Cells[row, 13].Value?.ToString(),
                        Control_Effectiveness = worksheet.Cells[row, 14].Value?.ToString(),
                        Design_Effectiveness = worksheet.Cells[row, 15].Value?.ToString(),

                        Residual_Risk_Considering_Conrol_Effectiveness = worksheet.Cells[row, 16].Value?.ToString(),
                        Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 17].Value?.ToString(),


                        KRI1 = worksheet.Cells[row, 18].Value?.ToString(),
                        Green = worksheet.Cells[row, 19].Value?.ToString(),
                        Amber = worksheet.Cells[row, 20].Value?.ToString(),
                        Red = worksheet.Cells[row, 21].Value?.ToString(),
                        // IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value),
                        IsActive = isActive,
                        // Add other properties if there are more columns in your Excel file
                    };

                    excelData.Add(data);
                }
            }

            return excelData;
        }
        private void ImportDataIntoExistingTabl(string tableName, string connectionString, string filePath)
        {
            // Implement the logic to import data into the existing table
            // Modify this code based on your specific requirements

            string importQuery = $"INSERT INTO {tableName} ( Vertical,Account,Process,Sub_Process,Activity,Activity_Discription,Risk_Statement,Likelyhood,Impact,Risk_Score,Control_Name,Control_Description,Control_Effectiveness,Design_Effectiveness,Residual_Risk_Considering_Conrol_Effectiveness,Residual_Risk_Considering_Design_Effectiveness,KRI,Green,Amber,Red,IsActive) VALUES (@Vertical, @Account,@Process,@Sub_Process,@Activity,@Activity_Discription,@Risk_Statement,@Likelyhood,@Impact,@Risk_Score,@Control_Name,@Control_Description,@Control_Effectiveness,@Design_Effectiveness,@Residual_Risk_Considering_Conrol_Effectiveness,@Residual_Risk_Considering_Design_Effectiveness,@KRI,@Green,@Amber,@Red,@IsActive)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Assuming you have a data source for the Excel data (replace this with your actual data source)
                List<KRI> excelData = GetExcelDat(filePath);

                foreach (var data in excelData)
                {
                    using (SqlCommand command = new SqlCommand(importQuery, connection))
                    {
                        //   command.Parameters.AddWithValue("@Sr_No", data.Sr_No);
                        command.Parameters.AddWithValue("@Vertical", data.Vertical);
                        command.Parameters.AddWithValue("@Account", data.Account);
                        command.Parameters.AddWithValue("@Process", data.Process);
                        command.Parameters.AddWithValue("@Sub_Process", data.Sub_Process);
                        command.Parameters.AddWithValue("@Activity", data.Activity);
                        command.Parameters.AddWithValue("@Activity_Discription", data.Activity_Discription);
                       
                        command.Parameters.AddWithValue("@Risk_Statement", data.Risk_Statement);
                        command.Parameters.AddWithValue("@Likelyhood", data.Likelyhood);
                        command.Parameters.AddWithValue("@Impact", data.Impact);
                        command.Parameters.AddWithValue("@Risk_Score", data.Risk_Score);
                        command.Parameters.AddWithValue("@Control_Name", data.Control_Name);
                        command.Parameters.AddWithValue("@Control_Description", data.Control_Description);
                        command.Parameters.AddWithValue("@Control_Effectiveness", data.Control_Effectiveness);
                        command.Parameters.AddWithValue("@Design_Effectiveness", data.Design_Effectiveness);
                 
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Conrol_Effectiveness", data.Residual_Risk_Considering_Conrol_Effectiveness);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Design_Effectiveness", data.Residual_Risk_Considering_Design_Effectiveness);
                        command.Parameters.AddWithValue("@KRI", data.KRI1);
                        command.Parameters.AddWithValue("@Green", data.Green);
                        command.Parameters.AddWithValue("@Amber", data.Amber);
                        command.Parameters.AddWithValue("@Red", data.Red);
                        

                        command.Parameters.AddWithValue("@IsActive", data.IsActive);


                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void ProcessExcelDat(string filePath, string tableName)
        {
            string connectionString = ConfigurationManager.ConnectionStrings?["YourConnectionStringName"]?.ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Read Excel data using a library like EPPlus or NPOI
                // Example using EPPlus:
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Read data from Excel and insert into the corresponding table
                        int? Sr_NO = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        //Sr_NO = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        string Vertical = worksheet.Cells[row, 2].Value?.ToString();
                        string Account = worksheet.Cells[row, 3].Value?.ToString();
                        string Process = worksheet.Cells[row, 4].Value?.ToString();
                        string Sub_Process = worksheet.Cells[row, 5].Value?.ToString();
                        string Activity = worksheet.Cells[row, 6].Value?.ToString();
                        string Activity_Discription = worksheet.Cells[row, 7].Value?.ToString();

                        string Risk_Statement = worksheet.Cells[row, 8].Value?.ToString();
                        string Likelyhood = worksheet.Cells[row, 9].Value?.ToString();
                        string Impact = worksheet.Cells[row, 10].Value?.ToString();
                        string Risk_Score = worksheet.Cells[row, 11].Value?.ToString();
                        string Control_Name = worksheet.Cells[row, 12].Value?.ToString();
                        string Control_Description = worksheet.Cells[row, 13].Value?.ToString();
                        string Control_Effectiveness = worksheet.Cells[row, 14].Value?.ToString();
                        string Design_Effectiveness = worksheet.Cells[row, 15].Value?.ToString();

                        string Residual_Risk_Considering_Conrol_Effectiveness = worksheet.Cells[row, 16].Value?.ToString();
                        string Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 17].Value?.ToString();


                        string KRI1 = worksheet.Cells[row, 18].Value?.ToString();
                        string Green = worksheet.Cells[row, 19].Value?.ToString();
                        string Amber = worksheet.Cells[row, 20].Value?.ToString();
                        string Red = worksheet.Cells[row, 21].Value?.ToString();
                        bool IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value);
                        // Insert data into the corresponding table
                        //InsertDataIntoTable(connection, tableName, Sr_No, Vertical, Account);
                    }
                }
            }
        }

        private void InsertDataIntoTabl(SqlConnection connection, string tableName, int? Sr_No, string Vertical, string Account, string Process, bool? IsActive)
        {
            // Implement the logic to insert data into the specified table
            // You may use parameterized queries or an ORM like Dapper or Entity Framework

            string query = $"INSERT INTO {tableName} (Sr_No, Vertical,Account) VALUES (@Sr_No, @Vertical,@Account)";

            using (SqlCommand command = new SqlCommand(query, connection))
            {


                command.Parameters.AddWithValue("@Sr_No", Sr_No);
                command.Parameters.AddWithValue("@Vertical", Vertical);
                command.Parameters.AddWithValue("@Account", Account);



                command.ExecuteNonQuery();
            }
        }

        //Upload function for NTRR

        [HttpGet]
        public ActionResult UploadExcelNTRR()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFileNTRR(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                // Get the file name without extension
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);

                // Save the file to a temporary location
                var filePath = Path.Combine(Server.MapPath("~/App_Data"), file.FileName);
                file.SaveAs(filePath);

                // Create a table with the same name as the file (if not exists)
                CreateTableIfNotExistNTRR(fileNameWithoutExtension, filePath);

                // Process the Excel file data
                //ProcessExcelData(filePath, fileNameWithoutExtension);

                // Optionally: Delete the temporary file
                System.IO.File.Delete(filePath);

                // Redirect or return a success message
                return RedirectToAction("Index", "Home");
            }

            // Handle invalid file or other error scenarios
            return View("Error");
        }

        private void CreateTableIfNotExistNTRR(string tableName, string filePath)
        {


            string connectionString = ConfigurationManager.ConnectionStrings["YourConnectionStringName"]?.ConnectionString;

            // Check if the connection string is not null or empty
            if (!string.IsNullOrEmpty(connectionString))
            {
                // Create a new SqlConnection object with the initialized connection string
                using (var sqlConnection = new SqlConnection(connectionString))
                {
                    // Open the SQL connection
                    sqlConnection.Open();

                    // Rest of your code to interact with the database goes here...
                }
            }
            else
            {
                // Handle the case where the connection string is not found or empty
                throw new Exception("Connection string is not found or empty.");
            }



            // Check if the table exists
            if (TableExistNTRR(tableName, connectionString))
            {
                // Table exists, proceed to import data
                ImportDataIntoExistingTableNTRR(tableName, connectionString, filePath);
            }
            else
            {
                // Table doesn't exist, create the table
                //CreateTable(tableName, connectionString);

                // Import data into the created table
                ImportDataIntoExistingTableNTRR(tableName, connectionString, filePath);
            }
        }


        private bool TableExistNTRR(string tableName, string connectionString)
        {
            // Check if the table exists in the database
            string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    int count = (int)command.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void CreateTableNTRR(string tableName, string connectionString)
        {
            // Create the table if it does not exist
            string createTableQuery = $@"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
        BEGIN
            CREATE TABLE {tableName} (
                ID INT PRIMARY KEY IDENTITY(1,1),
                Column1 NVARCHAR(MAX),
                Column2 NVARCHAR(MAX)
                -- Add other columns as needed
            );
        END";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private List<NTRR> GetExcelDataNTRR(string filePath)
        {
            List<NTRR> excelData = new List<NTRR>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    bool isActive;
                    bool.TryParse(worksheet.Cells[row, 36].Value?.ToString(), out isActive);
                    NTRR data = new NTRR
                    {
                        Sr_No = Convert.ToInt32(worksheet.Cells[row, 1].Value),
                        Vertical = worksheet.Cells[row, 2].Value?.ToString(),
                        Account = worksheet.Cells[row, 3].Value?.ToString(),
                        Process = worksheet.Cells[row, 4].Value?.ToString(),
                        Sub_Process = worksheet.Cells[row, 5].Value?.ToString(),
                        Activity = worksheet.Cells[row, 6].Value?.ToString(),
                       
                        Applications_software = worksheet.Cells[row, 7].Value?.ToString(),
                       
                        Frequency_daily_weekly_monthly = worksheet.Cells[row, 8].Value?.ToString(),
                        Activity_Description = worksheet.Cells[row, 9].Value?.ToString(),
                        Supplier_who_is_sending = worksheet.Cells[row, 10].Value?.ToString(),
                       
                        Input_info_needs_to_be_processed = worksheet.Cells[row, 11].Value?.ToString(),
                        Process_how_it_is_actually_done = worksheet.Cells[row, 12].Value?.ToString(),

                        Output_what_is_the_output_storage = worksheet.Cells[row, 13].Value?.ToString(),
                        Customer_End_Client_Client_Or_biz = worksheet.Cells[row, 14].Value?.ToString(),
                        SLA_Accuracy_Timelines = worksheet.Cells[row, 15].Value?.ToString(),
                        SLA_Target = worksheet.Cells[row, 16].Value?.ToString(),
                       
                        Risk_Description = worksheet.Cells[row, 17].Value?.ToString(),
                        Type_of_Risk_Financial_Non_Financial_Regulatory = worksheet.Cells[row, 18].Value?.ToString(),
                        Risk_Statement = worksheet.Cells[row, 19].Value?.ToString(),
                        Likelyhood = worksheet.Cells[row, 20].Value?.ToString(),
                        Impact = worksheet.Cells[row, 21].Value?.ToString(),
                        Risk_Score = worksheet.Cells[row, 22].Value?.ToString(),
                        Control_Name = worksheet.Cells[row, 23].Value?.ToString(),
                        Control_Description = worksheet.Cells[row, 24].Value?.ToString(),
                        Control_Effectiveness = worksheet.Cells[row, 25].Value?.ToString(),
                        Design_Effectiveness = worksheet.Cells[row, 26].Value?.ToString(),
                        Control_Owner = worksheet.Cells[row, 27].Value?.ToString(),
                        Type_of_control_Preventive_Manual_Detective_Automated_Process_People = worksheet.Cells[row, 28].Value?.ToString(),

                        Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 29].Value?.ToString(),
                        Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 30].Value?.ToString(),


                        Test_steps_Methodology = worksheet.Cells[row, 31].Value?.ToString(),
                        Testing_Objective = worksheet.Cells[row, 32].Value?.ToString(),
                        What_to_Look_at = worksheet.Cells[row, 33].Value?.ToString(),
                        What_to_Look_for = worksheet.Cells[row, 34].Value?.ToString(),
                        What_to_report = worksheet.Cells[row, 35].Value?.ToString(),
                        // IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value),
                        IsActive = isActive,
                        // Add other properties if there are more columns in your Excel file
                    };

                    excelData.Add(data);
                }
            }

            return excelData;
        }
        private void ImportDataIntoExistingTableNTRR(string tableName, string connectionString, string filePath)
        {
            // Implement the logic to import data into the existing table
            // Modify this code based on your specific requirements

            string importQuery = $"INSERT INTO {tableName} (  Vertical,Account,Process,Sub_Process,Activity,Applications_software,Frequency_daily_weekly_monthly,Activity_Description,Supplier_who_is_sending,Input_info_needs_to_be_processed,Process_how_it_is_actually_done,Output_what_is_the_output_storage,Customer_End_Client_Client_Or_biz,SLA_Accuracy_Timelines,SLA_Target,Risk_Description,Type_of_Risk_Financial_Non_Financial_Regulatory,Risk_Statement,Likelyhood,Impact,Risk_Score,Control_Name,Control_Description,Control_Effectiveness,Design_Effectiveness,Control_Owner,Type_of_control_Preventive_Manual_Detective_Automated_Process_People,Residual_Risk_Considering_Control_Effectiveness,Residual_Risk_Considering_Design_Effectiveness,Test_steps_Methodology,Testing_Objective,What_to_Look_at,What_to_Look_for,What_to_report,IsActive) VALUES (@Vertical, @Account,@Process,@Sub_Process,@Activity,@Applications_software,@Frequency_daily_weekly_monthly,@Activity_Description,@Supplier_who_is_sending,@Input_info_needs_to_be_processed,@Process_how_it_is_actually_done,@Output_what_is_the_output_storage,@Customer_End_Client_Client_Or_biz,@SLA_Accuracy_Timelines,@SLA_Target,@Risk_Description,@Type_of_Risk_Financial_Non_Financial_Regulatory,@Risk_Statement,@Likelyhood,@Impact,@Risk_Score,@Control_Name,@Control_Description,@Control_Effectiveness,@Design_Effectiveness,@Control_Owner,@Type_of_control_Preventive_Manual_Detective_Automated_Process_People,@Residual_Risk_Considering_Control_Effectiveness,@Residual_Risk_Considering_Design_Effectiveness,@Test_steps_Methodology,@Testing_Objective,@What_to_Look_at,@What_to_Look_for,@What_to_report,@IsActive)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Assuming you have a data source for the Excel data (replace this with your actual data source)
                List<NTRR> excelData = GetExcelDataNTRR(filePath);

                foreach (var data in excelData)
                {
                    using (SqlCommand command = new SqlCommand(importQuery, connection))
                    {
                        //   command.Parameters.AddWithValue("@Sr_No", data.Sr_No);
                        command.Parameters.AddWithValue("@Vertical", data.Vertical);
                        command.Parameters.AddWithValue("@Account", data.Account);
                        command.Parameters.AddWithValue("@Process", data.Process);
                        command.Parameters.AddWithValue("@Sub_Process", data.Sub_Process);
                        command.Parameters.AddWithValue("@Activity", data.Activity);
                        command.Parameters.AddWithValue("@Applications_software", data.Applications_software);
                      
                        command.Parameters.AddWithValue("@Frequency_daily_weekly_monthly", data.Frequency_daily_weekly_monthly);
                      
                        command.Parameters.AddWithValue("@Activity_Description", data.Activity_Description);
                        command.Parameters.AddWithValue("@Supplier_who_is_sending", data.Supplier_who_is_sending);
                        command.Parameters.AddWithValue("@Input_info_needs_to_be_processed", data.Input_info_needs_to_be_processed);
                        command.Parameters.AddWithValue("@Process_how_it_is_actually_done", data.Process_how_it_is_actually_done);
                        command.Parameters.AddWithValue("@Output_what_is_the_output_storage", data.Output_what_is_the_output_storage);
                        command.Parameters.AddWithValue("@Customer_End_Client_Client_Or_biz", data.Customer_End_Client_Client_Or_biz);
                        command.Parameters.AddWithValue("@SLA_Accuracy_Timelines", data.SLA_Accuracy_Timelines);
                        command.Parameters.AddWithValue("@SLA_Target", data.SLA_Target);
     
                        command.Parameters.AddWithValue("@Risk_Description", data.Risk_Description);
                        command.Parameters.AddWithValue("@Type_of_Risk_Financial_Non_Financial_Regulatory", data.Type_of_Risk_Financial_Non_Financial_Regulatory);
                        command.Parameters.AddWithValue("@Risk_Statement", data.Risk_Statement);
                        command.Parameters.AddWithValue("@Likelyhood", data.Likelyhood);
                        command.Parameters.AddWithValue("@Impact", data.Impact);
                        command.Parameters.AddWithValue("@Risk_Score", data.Risk_Score);
                        command.Parameters.AddWithValue("@Control_Name", data.Control_Name);
                        command.Parameters.AddWithValue("@Control_Description", data.Control_Description);
                        command.Parameters.AddWithValue("@Control_Effectiveness", data.Control_Effectiveness);
                        command.Parameters.AddWithValue("@Design_Effectiveness", data.Design_Effectiveness);
                        command.Parameters.AddWithValue("@Control_Owner", data.Control_Owner);
                        command.Parameters.AddWithValue("@Type_of_control_Preventive_Manual_Detective_Automated_Process_People", data.Type_of_control_Preventive_Manual_Detective_Automated_Process_People);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Control_Effectiveness", data.Residual_Risk_Considering_Control_Effectiveness);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Design_Effectiveness", data.Residual_Risk_Considering_Design_Effectiveness);
                        command.Parameters.AddWithValue("@Test_steps_Methodology", data.Test_steps_Methodology);
                        command.Parameters.AddWithValue("@Testing_Objective", data.Testing_Objective);
                        command.Parameters.AddWithValue("@What_to_Look_at", data.What_to_Look_at);
                        command.Parameters.AddWithValue("@What_to_Look_for", data.What_to_Look_for);
                        command.Parameters.AddWithValue("@What_to_report", data.What_to_report);

                        command.Parameters.AddWithValue("@IsActive", data.IsActive);


                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void ProcessExcelDataNTRR(string filePath, string tableName)
        {
            string connectionString = ConfigurationManager.ConnectionStrings?["YourConnectionStringName"]?.ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Read Excel data using a library like EPPlus or NPOI
                // Example using EPPlus:
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Read data from Excel and insert into the corresponding table
                        int? Sr_No = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        string Vertical = worksheet.Cells[row, 2].Value?.ToString();
                        string Account = worksheet.Cells[row, 3].Value?.ToString();


                        string Process = worksheet.Cells[row, 4].Value?.ToString();
                        string Sub_Process = worksheet.Cells[row, 5].Value?.ToString();
                        string Activity = worksheet.Cells[row, 6].Value?.ToString();
                        
                        string Applications_software = worksheet.Cells[row, 7].Value?.ToString();
                        string Frequency_daily_weekly_monthly = worksheet.Cells[row, 8].Value?.ToString();
                      
                        string Activity_Description = worksheet.Cells[row,9].Value?.ToString();
                        string Supplier_who_is_sending = worksheet.Cells[row, 10].Value?.ToString();
                        string Input_info_needs_to_be_processed = worksheet.Cells[row, 11].Value?.ToString();
                        string Output_what_is_the_output_storage = worksheet.Cells[row, 12].Value?.ToString();
                        string Customer_End_Client_Client_Or_biz = worksheet.Cells[row, 13].Value?.ToString();
                        string SLA_Accuracy_Timelines = worksheet.Cells[row, 14].Value?.ToString();
                        string SLA_Target = worksheet.Cells[row, 15].Value?.ToString();
                       
                        string Risk_Description = worksheet.Cells[row, 16].Value?.ToString();
                        string Type_of_Risk_Financial_Non_Financial_Regulatory = worksheet.Cells[row, 17].Value?.ToString();
                        string Risk_Statement_I_have_added_this_column_as_required_for_RCSA = worksheet.Cells[row, 18].Value?.ToString();
                        string Likelyhood = worksheet.Cells[row, 19].Value?.ToString();
                        string Impact = worksheet.Cells[row, 20].Value?.ToString();
                        string Risk_Score = worksheet.Cells[row, 21].Value?.ToString();
                        string Control_Name = worksheet.Cells[row, 22].Value?.ToString();
                        string Control_Description = worksheet.Cells[row, 23].Value?.ToString();
                        string Control_Effectiveness = worksheet.Cells[row, 24].Value?.ToString();
                        string Design_Effectiveness = worksheet.Cells[row, 25].Value?.ToString();
                        string Control_Owner = worksheet.Cells[row, 26].Value?.ToString();
                        string Type_of_control_Preventive_Manual_Detective_Automated_Process_People = worksheet.Cells[row, 27].Value?.ToString();

                        string Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 28].Value?.ToString();
                        string Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 29].Value?.ToString();


                        string Test_steps_Methodology = worksheet.Cells[row, 30].Value?.ToString();
                        string Testing_Objective = worksheet.Cells[row, 31].Value?.ToString();
                        string What_to_Look_at = worksheet.Cells[row, 32].Value?.ToString();
                        string What_to_Look_for = worksheet.Cells[row, 33].Value?.ToString();
                        string What_to_report = worksheet.Cells[row, 34].Value?.ToString();
                        bool IsActive = Convert.ToBoolean(worksheet.Cells[row, 35].Value);
                        // Insert data into the corresponding table
                        //InsertDataIntoTable(connection, tableName, Sr_No, Vertical, Account);
                    }
                }
            }
        }

        private void InsertDataIntoTableNTRR(SqlConnection connection, string tableName, int? Sr_No, string Vertical, string Account, string Process, bool? IsActive)
        {
            // Implement the logic to insert data into the specified table
            // You may use parameterized queries or an ORM like Dapper or Entity Framework

            string query = $"INSERT INTO {tableName} (Sr_No, Vertical,Account) VALUES (@Sr_No, @Vertical,@Account)";

            using (SqlCommand command = new SqlCommand(query, connection))
            {


                command.Parameters.AddWithValue("@Sr_No", Sr_No);
                command.Parameters.AddWithValue("@Vertical", Vertical);
                command.Parameters.AddWithValue("@Account", Account);



                command.ExecuteNonQuery();
            }
        }

        //upload for TRR

        [HttpGet]
        public ActionResult UploadExcelTRR()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadFileTRR(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                // Get the file name without extension
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.FileName);

                // Save the file to a temporary location
                var filePath = Path.Combine(Server.MapPath("~/App_Data"), file.FileName);
                file.SaveAs(filePath);

                // Create a table with the same name as the file (if not exists)
                CreateTableIfNotExitTRR(fileNameWithoutExtension, filePath);

                // Process the Excel file data
                //ProcessExcelData(filePath, fileNameWithoutExtension);

                // Optionally: Delete the temporary file
                System.IO.File.Delete(filePath);

                // Redirect or return a success message
                return RedirectToAction("Index", "Home");
            }

            // Handle invalid file or other error scenarios
            return View("Error");
        }

        private void CreateTableIfNotExitTRR(string tableName, string filePath)
        {


            string connectionString = ConfigurationManager.ConnectionStrings["YourConnectionStringName"]?.ConnectionString;

            // Check if the connection string is not null or empty
            if (!string.IsNullOrEmpty(connectionString))
            {
                // Create a new SqlConnection object with the initialized connection string
                using (var sqlConnection = new SqlConnection(connectionString))
                {
                    // Open the SQL connection
                    sqlConnection.Open();

                    // Rest of your code to interact with the database goes here...
                }
            }
            else
            {
                // Handle the case where the connection string is not found or empty
                throw new Exception("Connection string is not found or empty.");
            }



            // Check if the table exists
            if (TableExitTRR(tableName, connectionString))
            {
                // Table exists, proceed to import data
                ImportDataIntoExistingTableTRR(tableName, connectionString, filePath);
            }
            else
            {
                // Table doesn't exist, create the table
                //CreateTable(tableName, connectionString);

                // Import data into the created table
                ImportDataIntoExistingTableTRR(tableName, connectionString, filePath);
            }
        }


        private bool TableExitTRR(string tableName, string connectionString)
        {
            // Check if the table exists in the database
            string query = $"SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}'";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    int count = (int)command.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void CreateTableTRR(string tableName, string connectionString)
        {
            // Create the table if it does not exist
            string createTableQuery = $@"
        IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{tableName}')
        BEGIN
            CREATE TABLE {tableName} (
                ID INT PRIMARY KEY IDENTITY(1,1),
                Column1 NVARCHAR(MAX),
                Column2 NVARCHAR(MAX)
                -- Add other columns as needed
            );
        END";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                using (SqlCommand command = new SqlCommand(createTableQuery, connection))
                {
                    command.ExecuteNonQuery();
                }
            }
        }

        private List<TRR> GetExcelDTRR(string filePath)
        {
            List<TRR> excelData = new List<TRR>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    bool isActive;
                    bool.TryParse(worksheet.Cells[row, 43].Value?.ToString(), out isActive);
                    TRR data = new TRR
                    {
                        Sr_No = Convert.ToInt32(worksheet.Cells[row, 1].Value),
                        Vertical = worksheet.Cells[row, 2].Value?.ToString(),
                        Account = worksheet.Cells[row, 3].Value?.ToString(),
                        Process = worksheet.Cells[row, 4].Value?.ToString(),
                        Sub_Process = worksheet.Cells[row, 5].Value?.ToString(),
                        Activity = worksheet.Cells[row, 6].Value?.ToString(),
                        Head_Count = worksheet.Cells[row, 7].Value?.ToString(),
                        Applications = worksheet.Cells[row, 8].Value?.ToString(),
                        Volume = worksheet.Cells[row, 9].Value?.ToString(),
                        Frequency = worksheet.Cells[row, 10].Value?.ToString(),

                        Any_Volume_trends = worksheet.Cells[row, 11].Value?.ToString(),
                        SOP_Available = worksheet.Cells[row, 12].Value?.ToString(),
                        No_of_SOP = worksheet.Cells[row, 13].Value?.ToString(),
                        Activity_Description = worksheet.Cells[row, 14].Value?.ToString(),
                        Supplier = worksheet.Cells[row, 15].Value?.ToString(),
                        Input = worksheet.Cells[row, 16].Value?.ToString(),
                        Process2 = worksheet.Cells[row, 17].Value?.ToString(),
                        Output = worksheet.Cells[row, 18].Value?.ToString(),

                        Customer = worksheet.Cells[row, 19].Value?.ToString(),
                        SLA = worksheet.Cells[row, 20].Value?.ToString(),
                        SLA_Target = worksheet.Cells[row, 21].Value?.ToString(),
                        No_Of_errors = worksheet.Cells[row, 22].Value?.ToString(),
                        Type_of_Errors = worksheet.Cells[row, 23].Value?.ToString(),
                        Risk_Description = worksheet.Cells[row, 24].Value?.ToString(),
                        Type_of_Risk = worksheet.Cells[row, 25].Value?.ToString(),



                        Risk_Statement = worksheet.Cells[row, 26].Value?.ToString(),
                        Likelyhood = worksheet.Cells[row, 27].Value?.ToString(),
                        Impact = worksheet.Cells[row, 28].Value?.ToString(),
                        Risk_Score = worksheet.Cells[row, 29].Value?.ToString(),
                        Control_Name = worksheet.Cells[row, 30].Value?.ToString(),
                        Control_Description = worksheet.Cells[row, 31].Value?.ToString(),
                        Control_Effectiveness = worksheet.Cells[row, 32].Value?.ToString(),
                        Design_Effectiveness = worksheet.Cells[row, 33].Value?.ToString(),

                       


                        Control_Owner = worksheet.Cells[row, 34].Value?.ToString(),
                        Type_of_control = worksheet.Cells[row, 35].Value?.ToString(),
                        Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 36].Value?.ToString(),
                        Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 37].Value?.ToString(),
                        Test_steps_Methodology = worksheet.Cells[row, 38].Value?.ToString(),
                        Testing_Objective = worksheet.Cells[row, 39].Value?.ToString(),
                        What_to_Look_at = worksheet.Cells[row, 40].Value?.ToString(),
                        What_to_Look_for = worksheet.Cells[row, 41].Value?.ToString(),
                        What_to_report = worksheet.Cells[row, 42].Value?.ToString(),
                        // IsActive = Convert.ToBoolean(worksheet.Cells[row, 42].Value),
                        IsActive = isActive,
                        // Add other properties if there are more columns in your Excel file
                    };

                    excelData.Add(data);
                }
            }

            return excelData;
        }
        private void ImportDataIntoExistingTableTRR(string tableName, string connectionString, string filePath)
        {
            // Implement the logic to import data into the existing table
            // Modify this code based on your specific requirements

            string importQuery = $"INSERT INTO {tableName} ( Vertical,Account,Process, Sub_Process,Activity,Head_Count,Applications,Volume,Frequency,Any_Volume_trends,SOP_Available,No_of_SOP,Activity_Description,Supplier,Input,Process2,Output,Customer,SLA,SLA_Target,No_Of_errors,Type_of_Errors,Risk_Description,Type_of_Risk,Risk_Statement,Likelyhood,Impact,Risk_Score,Control_Name,Control_Description,Control_Effectiveness,Design_Effectiveness,Control_Owner,Type_of_control,Residual_Risk_Considering_Control_Effectiveness,Residual_Risk_Considering_Design_Effectiveness,Test_steps_Methodology,Testing_Objective,What_to_Look_at,What_to_Look_for,What_to_report,IsActive) VALUES (@Vertical, @Account,@Process,@Sub_Process,@Activity,@Head_Count,@Applications,@Volume,@Frequency,@Any_Volume_trends,@SOP_Available,@No_of_SOP,@Activity_Description,@Supplier,@Input,@Process2,@Output,@Customer,@SLA,@SLA_Target,@No_Of_errors,@Type_of_Errors,@Risk_Description,@Type_of_Risk,@Risk_Statement,@Likelyhood,@Impact,@Risk_Score,@Control_Name,@Control_Description,@Control_Effectiveness,@Design_Effectiveness,@Control_Owner,@Type_of_control,@Residual_Risk_Considering_Control_Effectiveness,@Residual_Risk_Considering_Design_Effectiveness,@Test_steps_Methodology,@Testing_Objective,@What_to_Look_at,@What_to_Look_for,@What_to_report,@IsActive)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Assuming you have a data source for the Excel data (replace this with your actual data source)
                List<TRR> excelData = GetExcelDTRR(filePath);

                foreach (var data in excelData)
                {
                    using (SqlCommand command = new SqlCommand(importQuery, connection))
                    {
                        //   command.Parameters.AddWithValue("@Sr_No", data.Sr_No);
                        command.Parameters.AddWithValue("@Vertical", data.Vertical);
                        command.Parameters.AddWithValue("@Account", data.Account);
                        command.Parameters.AddWithValue("@Process", data.Process);
                        command.Parameters.AddWithValue("@Sub_Process", data.Sub_Process);
                        command.Parameters.AddWithValue("@Activity", data.Activity);
                        command.Parameters.AddWithValue("@Head_Count", data.Head_Count);
                        command.Parameters.AddWithValue("@Applications", data.Applications);
                        command.Parameters.AddWithValue("@Volume", data.Volume);
                        command.Parameters.AddWithValue("@Frequency", data.Frequency);
                        command.Parameters.AddWithValue("@Any_Volume_trends", data.Any_Volume_trends);
                        command.Parameters.AddWithValue("@SOP_Available", data.SOP_Available);
                        command.Parameters.AddWithValue("@No_of_SOP", data.No_of_SOP);
                        command.Parameters.AddWithValue("@Activity_Description", data.Activity_Description);
                        command.Parameters.AddWithValue("@Supplier", data.Supplier);
                        command.Parameters.AddWithValue("@Input", data.Input);
                        command.Parameters.AddWithValue("@Process2", data.Process2);
                        command.Parameters.AddWithValue("@Output", data.Output);
                        command.Parameters.AddWithValue("@Customer", data.Customer);
                        command.Parameters.AddWithValue("@SLA", data.SLA);
                        command.Parameters.AddWithValue("@SLA_Target", data.SLA_Target);
                        command.Parameters.AddWithValue("@No_Of_errors", data.No_Of_errors);
                        command.Parameters.AddWithValue("@Type_of_Errors", data.Type_of_Errors);
                        command.Parameters.AddWithValue("@Risk_Description", data.Risk_Description);
                        command.Parameters.AddWithValue("@Type_of_Risk", data.Type_of_Risk);

                        command.Parameters.AddWithValue("@Risk_Statement", data.Risk_Statement);
                        command.Parameters.AddWithValue("@Likelyhood", data.Likelyhood);
                        command.Parameters.AddWithValue("@Impact", data.Impact);
                        command.Parameters.AddWithValue("@Risk_Score", data.Risk_Score);
                        command.Parameters.AddWithValue("@Control_Name", data.Control_Name);
                        command.Parameters.AddWithValue("@Control_Description", data.Control_Description);
                        command.Parameters.AddWithValue("@Control_Effectiveness", data.Control_Effectiveness);
                        command.Parameters.AddWithValue("@Design_Effectiveness", data.Design_Effectiveness);
                        command.Parameters.AddWithValue("@Control_Owner", data.Control_Owner);
                        command.Parameters.AddWithValue("@Type_of_control", data.Type_of_control);
                     
                        

                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Control_Effectiveness", data.Residual_Risk_Considering_Control_Effectiveness);
                        command.Parameters.AddWithValue("@Residual_Risk_Considering_Design_Effectiveness", data.Residual_Risk_Considering_Design_Effectiveness);
                        command.Parameters.AddWithValue("@Test_steps_Methodology", data.Test_steps_Methodology);
                        command.Parameters.AddWithValue("@Testing_Objective", data.Testing_Objective);
                        command.Parameters.AddWithValue("@What_to_Look_at", data.What_to_Look_at);
                        command.Parameters.AddWithValue("@What_to_Look_for", data.What_to_Look_for);
                        command.Parameters.AddWithValue("@What_to_report", data.What_to_report);

                        command.Parameters.AddWithValue("@IsActive", data.IsActive);


                        command.ExecuteNonQuery();
                    }
                }
            }
        }

        private void ProcessExcelDTRR(string filePath, string tableName)
        {
            string connectionString = ConfigurationManager.ConnectionStrings?["YourConnectionStringName"]?.ConnectionString;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                // Read Excel data using a library like EPPlus or NPOI
                // Example using EPPlus:
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Read data from Excel and insert into the corresponding table
                        int? Sr_NO = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        //Sr_NO = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                        string Vertical = worksheet.Cells[row, 2].Value?.ToString(); ;
                        string Account = worksheet.Cells[row, 3].Value?.ToString();
                        string Process = worksheet.Cells[row, 4].Value?.ToString();
                        string Sub_Process = worksheet.Cells[row, 5].Value?.ToString();
                        string Activity = worksheet.Cells[row, 6].Value?.ToString();
                        string Head_Count = worksheet.Cells[row, 7].Value?.ToString();
                        string Applications = worksheet.Cells[row, 8].Value?.ToString();
                        string Volume = worksheet.Cells[row, 9].Value?.ToString();
                        string Frequency = worksheet.Cells[row, 10].Value?.ToString();

                        string Any_Volume_trends = worksheet.Cells[row, 11].Value?.ToString();
                        string SOP_Available = worksheet.Cells[row, 12].Value?.ToString();
                        string No_of_SOP = worksheet.Cells[row, 13].Value?.ToString();
                        string Activity_Description = worksheet.Cells[row, 14].Value?.ToString();
                        string Supplier = worksheet.Cells[row, 15].Value?.ToString();
                        string Input = worksheet.Cells[row, 16].Value?.ToString();
                        string Process2 = worksheet.Cells[row, 17].Value?.ToString();
                        string Output = worksheet.Cells[row, 18].Value?.ToString();

                        string Customer = worksheet.Cells[row, 19].Value?.ToString();
                        string SLA = worksheet.Cells[row, 20].Value?.ToString();
                        string SLA_Target = worksheet.Cells[row, 21].Value?.ToString();
                        string No_Of_errors = worksheet.Cells[row, 22].Value?.ToString();
                        string Type_of_Errors = worksheet.Cells[row, 23].Value?.ToString();
                        string Risk_Description = worksheet.Cells[row, 24].Value?.ToString();
                        string Type_of_Risk = worksheet.Cells[row, 25].Value?.ToString();
                        string Risk_Statement = worksheet.Cells[row, 26].Value?.ToString();
                        string Likelyhood = worksheet.Cells[row, 27].Value?.ToString();
                        string Impact = worksheet.Cells[row, 28].Value?.ToString();
                        string Risk_Score = worksheet.Cells[row, 29].Value?.ToString();
                        string Control_Name = worksheet.Cells[row, 30].Value?.ToString();
                        string Control_Description = worksheet.Cells[row, 31].Value?.ToString();
                        string Control_Effectiveness = worksheet.Cells[row, 32].Value?.ToString();
                        string Design_Effectiveness = worksheet.Cells[row, 33].Value?.ToString();
                        string Control_Owner = worksheet.Cells[row, 34].Value?.ToString();
                        string Type_of_control = worksheet.Cells[row, 35].Value?.ToString();
                        string Residual_Risk_Considering_Control_Effectiveness = worksheet.Cells[row, 36].Value?.ToString();
                        string Residual_Risk_Considering_Design_Effectiveness = worksheet.Cells[row, 37].Value?.ToString();
                        string Test_steps_Methodology = worksheet.Cells[row, 38].Value?.ToString();
                        string Testing_Objective = worksheet.Cells[row, 39].Value?.ToString();
                        string What_to_Look_at = worksheet.Cells[row, 40].Value?.ToString();
                        string What_to_Look_for = worksheet.Cells[row, 41].Value?.ToString();
                        string What_to_report = worksheet.Cells[row, 42].Value?.ToString();
                        bool IsActive = Convert.ToBoolean(worksheet.Cells[row, 43].Value);
                        // Insert data into the corresponding table
                        //InsertDataIntoTable(connection, tableName, Sr_No, Vertical, Account);
                    }
                }
            }
        }

        private void InsertDataIntoTaTRR(SqlConnection connection, string tableName, int? Sr_No, string Vertical, string Account, string Process, bool? IsActive)
        {
            // Implement the logic to insert data into the specified table
            // You may use parameterized queries or an ORM like Dapper or Entity Framework

            string query = $"INSERT INTO {tableName} (Sr_No, Vertical,Account) VALUES (@Sr_No, @Vertical,@Account)";

            using (SqlCommand command = new SqlCommand(query, connection))
            {


                command.Parameters.AddWithValue("@Sr_No", Sr_No);
                command.Parameters.AddWithValue("@Vertical", Vertical);
                command.Parameters.AddWithValue("@Account", Account);



                command.ExecuteNonQuery();
            }
        }
        /// <summary>
        /// below code is related to generate the templete EXCeL
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>

        

        [HttpGet]
        public ActionResult TRRAmex(TRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectAmex(TRR model)
        {
            if (ModelState.IsValid)
            {
                TRR obj = new TRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;

                obj.Head_Count = model.Head_Count;
                obj.Applications = model.Applications;
                obj.Volume = model.Volume;
                obj.Frequency = model.Frequency;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier = model.Supplier;
                obj.Input = model.Input;
                obj.Process2 = model.Process2;
                obj.Output = model.Output;

                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;

                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.TRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("TRRAmexList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult TRRAmexList(string filterValue)
        {
            // Get the data from the database
            var TRRAmexList = dbObj.TRRs.Where(p => p.Account.Contains("Amex")).ToList(); ;

            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(TRRAmexList);
        }
        public ActionResult Assign1(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("TRRAmexList");
        }

        [HttpGet]
        public ActionResult TRRAegon(TRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectTRRAegon(TRR model)
        {
            if (ModelState.IsValid)
            {
                TRR obj = new TRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;

                obj.Head_Count = model.Head_Count;
                obj.Applications = model.Applications;
                obj.Volume = model.Volume;
                obj.Frequency = model.Frequency;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier = model.Supplier;
                obj.Input = model.Input;
                obj.Process2 = model.Process2;
                obj.Output = model.Output;

                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;

                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.TRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("TRRAegonList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult TRRAegonList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var TRRAegonList = dbObj.TRRs.Where(p => p.Account.Contains("Aegon")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(TRRAegonList);
        }
        public ActionResult Assign2(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("TRRAegonList");
        }

        [HttpGet]
        public ActionResult TRRNovartis(TRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectTRRNovartis(TRR model)
        {
            if (ModelState.IsValid)
            {
                TRR obj = new TRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;

                obj.Head_Count = model.Head_Count;
                obj.Applications = model.Applications;
                obj.Volume = model.Volume;
                obj.Frequency = model.Frequency;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier = model.Supplier;
                obj.Input = model.Input;
                obj.Process2 = model.Process2;
                obj.Output = model.Output;

                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;

                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.TRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("TRRNovartisList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult TRRNovartisList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var TRRNovartisList = dbObj.TRRs.Where(p => p.Account.Contains("Novartis")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(TRRNovartisList);
        }
        public ActionResult Assign3(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("TRRNovartisList");
        }
        [HttpGet]
        public ActionResult TRRAberden(TRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectTRRAberden(TRR model)
        {
            if (ModelState.IsValid)
            {
                TRR obj = new TRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;

                obj.Head_Count = model.Head_Count;
                obj.Applications = model.Applications;
                obj.Volume = model.Volume;
                obj.Frequency = model.Frequency;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier = model.Supplier;
                obj.Input = model.Input;
                obj.Process2 = model.Process2;
                obj.Output = model.Output;

                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;

                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.TRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("TRRAberdenList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult TRRAberdenList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var TRRAberdenList = dbObj.TRRs.Where(p => p.Account.Contains("Abrdn")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(TRRAberdenList);
        }
        public ActionResult Assign4(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("TRRAberdenList");
        }

        [HttpGet]
        public ActionResult TRRMultiplan(TRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectTRRMultiplan(TRR model)
        {
            if (ModelState.IsValid)
            {
                TRR obj = new TRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;

                obj.Head_Count = model.Head_Count;
                obj.Applications = model.Applications;
                obj.Volume = model.Volume;
                obj.Frequency = model.Frequency;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier = model.Supplier;
                obj.Input = model.Input;
                obj.Process2 = model.Process2;
                obj.Output = model.Output;

                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;

                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.TRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("TRRMultiplanList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult TRRMultiplanList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var TRRMultiplanList = dbObj.TRRs.Where(p => p.Account.Contains("Multiplan")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(TRRMultiplanList);
        }
        public ActionResult Assign5(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("TRRMultiplanList");
        }
        [HttpGet]
        public ActionResult NTRRMultiplan(NTRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectNTRRMultiplan(NTRR model)
        {
            if (ModelState.IsValid)
            {
                NTRR obj = new NTRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Applications_software = model.Applications_software;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;

                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_End_Client_Client_Or_biz = model.Customer_End_Client_Client_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.NTRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("NTRRMultiplanList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult NTRRMultiplanList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var NTRRMultiplanList = dbObj.NTRRs.Where(p => p.Account.Contains("Multiplan")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(NTRRMultiplanList);
        }
        public ActionResult Assign6(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("NTRRMultiplanList");
        }

        [HttpGet]
        public ActionResult NTRRNovartis(NTRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectNTRRNovartis(NTRR model)
        {
            if (ModelState.IsValid)
            {
                NTRR obj = new NTRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Applications_software = model.Applications_software;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;

                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_End_Client_Client_Or_biz = model.Customer_End_Client_Client_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.NTRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("NTRRNovartisList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult NTRRNovartisList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var NTRRNovartisList = dbObj.NTRRs.Where(p => p.Account.Contains("Novartis")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(NTRRNovartisList);
        }
        public ActionResult Assign7(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("NTRRNovartisList");
        }
        [HttpGet]
        public ActionResult NTRRAmex(NTRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectNTRRAmex(NTRR model)
        {
            if (ModelState.IsValid)
            {
                NTRR obj = new NTRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Applications_software = model.Applications_software;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;

                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_End_Client_Client_Or_biz = model.Customer_End_Client_Client_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.NTRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("NTRRAmexList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult NTRRAmexList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var NTRRAmexList = dbObj.NTRRs.Where(p => p.Account.Contains("Amex")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(NTRRAmexList);
        }
        public ActionResult Assign8(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("NTRRAmexList");
        }

        [HttpGet]
        public ActionResult NTRRAberden(NTRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectNTRRAberden(NTRR model)
        {
            if (ModelState.IsValid)
            {
                NTRR obj = new NTRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Applications_software = model.Applications_software;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;

                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_End_Client_Client_Or_biz = model.Customer_End_Client_Client_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.NTRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("NTRRAberdenList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult NTRRAberdenList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var NTRRAberdenList = dbObj.NTRRs.Where(p => p.Account.Contains("abrdn")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(NTRRAberdenList);
        }
        public ActionResult Assign9(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("NTRRAberdenList");
        }
        [HttpGet]
        public ActionResult NTRRAegon(NTRR obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectNTRRAegon(NTRR model)
        {
            if (ModelState.IsValid)
            {
                NTRR obj = new NTRR();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Applications_software = model.Applications_software;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;

                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_End_Client_Client_Or_biz = model.Customer_End_Client_Client_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.NTRRs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("NTRRAegonList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult NTRRAegonList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var NTRRAegonList = dbObj.NTRRs.Where(p => p.Account.Contains("Aegon")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(NTRRAegonList);
        }
        public ActionResult Assign10(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("NTRRAegonList");
        }
        [HttpGet]
        public ActionResult IBCPAegon(IBCP obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectIBCPAegon(IBCP model)
        {
            if (ModelState.IsValid)
            {
                IBCP obj = new IBCP();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Head_Count   = model.Head_Count;
                obj.Applications_software = model.Applications_software;
                obj.Volume = model.Volume;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_end_client_Onshore_Or_biz = model.Customer_end_client_Onshore_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.No_Of_errors = model.No_Of_errors;
                obj.Type_of_Errors = model.Type_of_Errors;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelihood = model.Likelihood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.IBCPs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("IBCPAegonList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult IBCPAegonList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var IBCPAegonList = dbObj.IBCPs.Where(p => p.Account.Contains("Aegon")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(IBCPAegonList);
        }
        public ActionResult Assign11(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("IBCPAegonList");
        }
        [HttpGet]
        public ActionResult IBCPAmex(IBCP obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectIBCPAmex(IBCP model)
        {
            if (ModelState.IsValid)
            {
                IBCP obj = new IBCP();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Head_Count = model.Head_Count;
                obj.Applications_software = model.Applications_software;
                obj.Volume = model.Volume;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_end_client_Onshore_Or_biz = model.Customer_end_client_Onshore_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.No_Of_errors = model.No_Of_errors;
                obj.Type_of_Errors = model.Type_of_Errors;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelihood = model.Likelihood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.IBCPs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("IBCPAmexList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult IBCPAmexList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var IBCPAmexList = dbObj.IBCPs.Where(p => p.Account.Contains("Amex")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(IBCPAmexList);
        }
        public ActionResult Assign12(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("IBCPAmexList");
        }
        [HttpGet]
        public ActionResult IBCPMultiplan(IBCP obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectIBCPMultiplan(IBCP model)
        {
            if (ModelState.IsValid)
            {
                IBCP obj = new IBCP();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Head_Count = model.Head_Count;
                obj.Applications_software = model.Applications_software;
                obj.Volume = model.Volume;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_end_client_Onshore_Or_biz = model.Customer_end_client_Onshore_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.No_Of_errors = model.No_Of_errors;
                obj.Type_of_Errors = model.Type_of_Errors;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelihood = model.Likelihood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.IBCPs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("IBCPMultiplanList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult IBCPMultiplanList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var IBCPMultiplanList = dbObj.IBCPs.Where(p => p.Account.Contains("Multiplan")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(IBCPMultiplanList);
        }
        public ActionResult Assign13(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("IBCPMultiplanList");
        }
        [HttpGet]
        public ActionResult IBCPNovartis(IBCP obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectIBCPNovartis(IBCP model)
        {
            if (ModelState.IsValid)
            {
                IBCP obj = new IBCP();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Head_Count = model.Head_Count;
                obj.Applications_software = model.Applications_software;
                obj.Volume = model.Volume;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_end_client_Onshore_Or_biz = model.Customer_end_client_Onshore_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.No_Of_errors = model.No_Of_errors;
                obj.Type_of_Errors = model.Type_of_Errors;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelihood = model.Likelihood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.IBCPs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("IBCPNovartisList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult IBCPNovartisList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var IBCPNovartisList = dbObj.IBCPs.Where(p => p.Account.Contains("Novartis")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(IBCPNovartisList);
        }
        public ActionResult Assign14(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("IBCPNovartisList");
        }
        [HttpGet]
        public ActionResult IBCPAberden(IBCP obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectIBCPAberden(IBCP model)
        {
            if (ModelState.IsValid)
            {
                IBCP obj = new IBCP();

                obj.Sr_No = model.Sr_No;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Head_Count = model.Head_Count;
                obj.Applications_software = model.Applications_software;
                obj.Volume = model.Volume;
                obj.Frequency_daily_weekly_monthly = model.Frequency_daily_weekly_monthly;
                obj.Any_Volume_trends = model.Any_Volume_trends;
                obj.SOP_Available = model.SOP_Available;
                obj.No_of_SOP = model.No_of_SOP;
                obj.Activity_Description = model.Activity_Description;
                obj.Supplier_who_is_sending = model.Supplier_who_is_sending;
                obj.Input_info_needs_to_be_processed = model.Input_info_needs_to_be_processed;
                obj.Process_how_it_is_actually_done = model.Process_how_it_is_actually_done;
                obj.Output_what_is_the_output_storage = model.Output_what_is_the_output_storage;
                obj.Customer_end_client_Onshore_Or_biz = model.Customer_end_client_Onshore_Or_biz;
                obj.SLA_Accuracy_Timelines = model.SLA_Accuracy_Timelines;
                obj.SLA_Target = model.SLA_Target;
                obj.No_Of_errors = model.No_Of_errors;
                obj.Type_of_Errors = model.Type_of_Errors;
                obj.Risk_Description = model.Risk_Description;
                obj.Type_of_Risk_Financial_Non_Financial_Regulatory = model.Type_of_Risk_Financial_Non_Financial_Regulatory;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelihood = model.Likelihood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Control_Owner = model.Control_Owner;
                obj.Type_of_control_Preventive_Manual_Detective_Automated_Process_People = model.Type_of_control_Preventive_Manual_Detective_Automated_Process_People;
                obj.Residual_Risk_Considering_Control_Effectiveness = model.Residual_Risk_Considering_Control_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.Test_steps_Methodology = model.Test_steps_Methodology;
                obj.Testing_Objective = model.Testing_Objective;
                obj.What_to_Look_at = model.What_to_Look_at;
                obj.What_to_Look_for = model.What_to_Look_for;
                obj.What_to_report = model.What_to_report;


                dbObj.IBCPs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("IBCPAberdenList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult IBCPAberdenList(string searchText)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var IBCPAberdenListt = dbObj.IBCPs.Where(p => p.Account.Contains("abrdn")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(searchText))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();
                IBCPAberdenListt = IBCPAberdenListt.Where(item => item.Status.Contains(searchText)).ToList();

            }

            return View(IBCPAberdenListt);
        }
        public ActionResult Assign15(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("IBCPAberdenList");
        }
        [HttpGet]
        public ActionResult KRIAberden(KRI obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectKRIAberden(KRI model)
        {
            if (ModelState.IsValid)
            {
                KRI obj = new KRI();
                obj.Sr_NO = model.Sr_NO;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Activity_Discription = model.Activity_Discription;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;               
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;             
                obj.Residual_Risk_Considering_Conrol_Effectiveness = model.Residual_Risk_Considering_Conrol_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.KRI1 = model.KRI1;
                obj.Green = model.Green;
                obj.Amber = model.Amber;
                obj.Red = model.Red;
                dbObj.KRIs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("KRIAberdenList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult KRIAberdenList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var KRIAberdenList = dbObj.KRIs.Where(p => p.Account.Contains("abrdn")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(KRIAberdenList);
        }
        public ActionResult Assign16(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("KRIAberdenList");
        }
        [HttpGet]
        public ActionResult KRINovartis(KRI obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectKRINovartis(KRI model)
        {
            if (ModelState.IsValid)
            {
                KRI obj = new KRI();
                obj.Sr_NO = model.Sr_NO;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Activity_Discription = model.Activity_Discription;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Residual_Risk_Considering_Conrol_Effectiveness = model.Residual_Risk_Considering_Conrol_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.KRI1 = model.KRI1;
                obj.Green = model.Green;
                obj.Amber = model.Amber;
                obj.Red = model.Red;
                dbObj.KRIs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("KRINovartisList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult KRINovartisList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var KRINovartisList = dbObj.KRIs.Where(p => p.Account.Contains("Novartis")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(KRINovartisList);
        }
        public ActionResult Assign17(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("KRINovartisList");
        }
        [HttpGet]
        public ActionResult KRIAmex(KRI obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectKRIAmex(KRI model)
        {
            if (ModelState.IsValid)
            {
                KRI obj = new KRI();
                obj.Sr_NO = model.Sr_NO;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Activity_Discription = model.Activity_Discription;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Residual_Risk_Considering_Conrol_Effectiveness = model.Residual_Risk_Considering_Conrol_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.KRI1 = model.KRI1;
                obj.Green = model.Green;
                obj.Amber = model.Amber;
                obj.Red = model.Red;
                dbObj.KRIs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("KRIAmexList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult KRIAmexList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var KRIAmexList = dbObj.KRIs.Where(p => p.Account.Contains("Amex")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(KRIAmexList);
        }
        public ActionResult Assign18(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("KRIAmexList");
        }
        [HttpGet]
        public ActionResult KRIMultiplan(KRI obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectKRIMultiplan(KRI model)
        {
            if (ModelState.IsValid)
            {
                KRI obj = new KRI();
                obj.Sr_NO = model.Sr_NO;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Activity_Discription = model.Activity_Discription;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Residual_Risk_Considering_Conrol_Effectiveness = model.Residual_Risk_Considering_Conrol_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.KRI1 = model.KRI1;
                obj.Green = model.Green;
                obj.Amber = model.Amber;
                obj.Red = model.Red;
                dbObj.KRIs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("KRIMultiplanList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult KRIMultiplanList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var KRIMultiplanList = dbObj.KRIs.Where(p => p.Account.Contains("Multiplan")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(KRIMultiplanList);
        }
        public ActionResult Assign19(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("KRIMultiplanList");
        }
        [HttpGet]
        public ActionResult KRIAegon(KRI obj)
        {
            return View(obj);

        }
        [HttpPost]
        public ActionResult AddProjectKRIAegon(KRI model)
        {
            if (ModelState.IsValid)
            {
                KRI obj = new KRI();
                obj.Sr_NO = model.Sr_NO;
                obj.Vertical = model.Vertical;
                obj.Account = model.Account;
                obj.Process = model.Process;
                obj.Sub_Process = model.Sub_Process;
                obj.Activity = model.Activity;
                obj.Activity_Discription = model.Activity_Discription;
                obj.Risk_Statement = model.Risk_Statement;
                obj.Likelyhood = model.Likelyhood;
                obj.Impact = model.Impact;
                obj.Risk_Score = model.Risk_Score;
                obj.Control_Name = model.Control_Name;
                obj.Control_Description = model.Control_Description;
                obj.Control_Effectiveness = model.Control_Effectiveness;
                obj.Design_Effectiveness = model.Design_Effectiveness;
                obj.Residual_Risk_Considering_Conrol_Effectiveness = model.Residual_Risk_Considering_Conrol_Effectiveness;
                obj.Residual_Risk_Considering_Design_Effectiveness = model.Residual_Risk_Considering_Design_Effectiveness;
                obj.KRI1 = model.KRI1;
                obj.Green = model.Green;
                obj.Amber = model.Amber;
                obj.Red = model.Red;
                dbObj.KRIs.Add(obj);
                dbObj.SaveChanges();
            }
            ModelState.Clear();
            return View("KRIAegonList");
        }
        //public ActionResult ProjectList() 
        //{
        //    var res = dbObj.tbl_Project.ToList();

        //    return View(res);
        //}
        public ActionResult KRIAegonList(string filterValue)
        {
            // Get the data from the database
            //var TRRAegonList = dbObj.TRRs.ToList();
            var KRIAegonList = dbObj.KRIs.Where(p => p.Account.Contains("Multiplan")).ToList();
            // Apply filtering if a filter value is provided
            if (!string.IsNullOrEmpty(filterValue))
            {
                //projectList = projectList.Where(p =>
                //p.New_Amendment.Contains(filterValue) ||
                //p.Technology.Contains(filterValue) ||
                //p.Sub_Function_Name.Contains(filterValue)).ToList();


            }

            return View(KRIAegonList);
        }
        public ActionResult Assign20(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("KRIAegonList");
        }

        [HttpPost]
        public ActionResult SearchStatus(string searchText)
        {
            var filteredData = dbObj.IBCPs.Where(item => item.Status.Contains(searchText)).ToList();
            return PartialView("_TableRowsPartial", filteredData);
        }

    }
}
