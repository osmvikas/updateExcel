using GemBox.Spreadsheet;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class UpdateExcelSheetController : ApiController
    {
        /// <summary>
        /// Update excel sheet.
        /// </summary>
        /// <param name="formNo"></param>
        /// <returns></returns>
        [HttpPost]
        public object Writexcel(FormNo formNo)
        {
            try
            {
                //Getting Base URL.
                string baseUrl = HttpContext.Current.Request.Url.Scheme + "://" + HttpContext.Current.Request.Url.Authority + HttpContext.Current.Request.ApplicationPath.TrimEnd('/') + "/";

                //Raw file path.
                string fileRawpath = ConfigurationManager.AppSettings.Get("FilePath");

                //Complete file path.
                string filePath = HttpContext.Current.Server.MapPath(fileRawpath);

                //SaveAs file path.
                string saveAsPath = ConfigurationManager.AppSettings.Get("SaveAs");

                //Complete file path.
                string saveAsFilePth = HttpContext.Current.Server.MapPath(saveAsPath);

                //convert path to url to download file.
                string downloadFile = baseUrl + saveAsPath;

                using (ExcelEngine excel = new ExcelEngine())
                {
                    //Instantiate the Excel application object
                    IApplication application = excel.Excel;

                    //Set the default application version
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Load the existing Excel workbook into IWorkbook
                    IWorkbook workbook = application.Workbooks.Open(filePath);

                    //Get the first worksheet in the workbook into IWorksheet
                    IWorksheet worksheet = workbook.Worksheets[0];

                    worksheet.Range["D5"].Text = formNo.NameAddress;
                    worksheet.Range["D6"].Text = formNo.PAN;
                    worksheet.Range["D7"].Text = formNo.FinancialYear;
                    worksheet.Range["A47"].Text = "Place : " + formNo.Place;
                    worksheet.Range["A48"].Text = "Date : " + formNo.Date;
                    worksheet.Range["A49"].Text = "Designation : " + formNo.Designation;

                    workbook.SaveAs(saveAsFilePth);
                }

                return new { status = 1, message = downloadFile };
            }
            catch(Exception ex)
            {
                return new { status = 0, message = ex.Message };
            }
            

        }
    }
}
