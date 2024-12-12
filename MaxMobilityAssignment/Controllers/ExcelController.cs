using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using MaxMobilityAssignment.Models;

namespace MaxMobilityAssignment.Controllers
{
    public class ExcelController : Controller
    {
        [HttpGet]
        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null || Path.GetExtension(file.FileName).ToLower() != ".xlsx")
            {
                ViewBag.Message = "Please upload a valid Excel file.";
                return View();
            }

            string uploadPath = Server.MapPath("~/Uploads/");
            if (!Directory.Exists(uploadPath))
                Directory.CreateDirectory(uploadPath);

            string filePath = Path.Combine(uploadPath, file.FileName);
            file.SaveAs(filePath);

            var result = ProcessExcelFile(filePath);

            string updatedFilePath = Path.Combine(uploadPath, "Processed_" + file.FileName);
            GenerateUpdatedExcel(filePath, updatedFilePath, result);

            string downloadLink = Url.Content("~/Uploads/Processed_" + file.FileName);
            ViewBag.Message = "File uploaded and processed successfully.";
            ViewBag.DownloadLink = downloadLink;
            ViewBag.Result = result;

            return View();
        }
        private List<UploadResult> ProcessExcelFile(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            List<UploadResult> results = new List<UploadResult>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null) return results;

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string name = worksheet.Cells[row, 1].Text;
                    string email = worksheet.Cells[row, 2].Text;
                    string phone = worksheet.Cells[row, 3].Text;
                    string address = worksheet.Cells[row, 4].Text;

                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(email) ||
                        string.IsNullOrWhiteSpace(phone) || string.IsNullOrWhiteSpace(address) ||
                        !IsValidEmail(email))
                    {
                        results.Add(new UploadResult { SerialNo = row - 1, Name = name, Email = email, Address = address, PhoneNo = phone, Row = row, Status = "Failed" });
                        continue;
                    }

                    SaveToDatabase(name, email, phone, address);
                    results.Add(new UploadResult { SerialNo = row - 1, Row = row, Name = name, Email = email, Address = address, PhoneNo = phone, Status = "Success" });
                }
            }
            return results;
        }
        private bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }
        private void SaveToDatabase(string name, string email, string phone, string address)
        {
            using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                using (var command = new SqlCommand("SaveExcelData", connection))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Name", name);
                    command.Parameters.AddWithValue("@Email", email);
                    command.Parameters.AddWithValue("@Phone", phone);
                    command.Parameters.AddWithValue("@Address", address);
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
        private void GenerateUpdatedExcel(string originalFilePath, string updatedFilePath, List<UploadResult> results)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var originalPackage = new ExcelPackage(new FileInfo(originalFilePath)))
            using (var updatedPackage = new ExcelPackage())
            {
                ExcelWorksheet originalSheet = originalPackage.Workbook.Worksheets.FirstOrDefault();
                if (originalSheet == null) return;

                ExcelWorksheet updatedSheet = updatedPackage.Workbook.Worksheets.Add("Processed Data");

                int rowCount = originalSheet.Dimension.Rows;
                int colCount = originalSheet.Dimension.Columns;

                updatedSheet.Cells[1, 1].Value = "SL No";
                for (int col = 1; col <= colCount; col++)
                {
                    updatedSheet.Cells[1, col + 1].Value = originalSheet.Cells[1, col].Text;
                }
                updatedSheet.Cells[1, colCount + 2].Value = "Status";

                for (int row = 2; row <= rowCount; row++)
                {
                    UploadResult result = results.FirstOrDefault(r => r.Row == row);

                    updatedSheet.Cells[row, 1].Value = row - 1; 
                    for (int col = 1; col <= colCount; col++)
                    {
                        updatedSheet.Cells[row, col + 1].Value = originalSheet.Cells[row, col].Text;
                    }
                    updatedSheet.Cells[row, colCount + 2].Value = result != null ? result.Status : "Unknown";
                }

                updatedPackage.SaveAs(new FileInfo(updatedFilePath));
            }
        }

    }
}