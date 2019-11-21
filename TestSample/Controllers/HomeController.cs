using System;
using System.Collections.Generic;
using System.Linq;
using TestSample.Models;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using Syncfusion.EJ2.Base;
using System.Data;
using System.IO;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;

namespace TestSample.Controllers
{

    public class HomeController : Controller
    {

        public IActionResult Index()
        {
            ViewBag.datasource = OrdersDetails.GetAllRecords().ToList();
           return View();
        }
        public async Task<IActionResult> Save()
        {
            string filePath = "App_Data/TempData/";
            string directoryPath = Path.Combine(new FileInfo(filePath).Directory.FullName);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            try
            {
                if (HttpContext.Request.Form.Files.Count > 0)
                {
                    for (int i = 0; i < HttpContext.Request.Form.Files.Count; ++i)
                    {
                        IFormFile httpPostedFile = HttpContext.Request.Form.Files[i];

                        if (httpPostedFile != null)
                        {
                            filePath = Path.Combine(directoryPath, httpPostedFile.FileName);

                            if (!System.IO.File.Exists(filePath))
                            {
                                using (var fileStream = new FileStream(filePath, FileMode.Create))
                                {
                                    await httpPostedFile.CopyToAsync(fileStream);
                                    ExcelEngine excelEngine = new ExcelEngine();

                                    //Loads or open an existing workbook through Open method of IWorkbooks
                                    fileStream.Position = 0;
                                    IWorkbook workbook = excelEngine.Excel.Workbooks.Open(httpPostedFile.OpenReadStream());
                                    IWorksheet worksheet = workbook.Worksheets[0];

                                    // Read data from the worksheet and Export to the DataTable.

                                    DataTable table = worksheet.ExportDataTable(worksheet.UsedRange.Row, worksheet.UsedRange.Column, worksheet.UsedRange.LastRow, worksheet.UsedRange.LastColumn, ExcelExportDataTableOptions.ColumnNames | ExcelExportDataTableOptions.ComputedFormulaValues);
                                    string JSONString = string.Empty;
                                    JSONString = JsonConvert.SerializeObject(table);
                                    ViewBag.data = JsonConvert.SerializeObject(table, Formatting.Indented, new JsonSerializerSettings { Converters = new[] { new Newtonsoft.Json.Converters.DataTableConverter() } });

                                    //return View();
                                }

                                return Ok(ViewBag.data);
                            }
                            else
                            {
                                return BadRequest("File already exists");
                            }
                        }
                    }
                }

                return BadRequest("No file in request"); ;
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }
        }

        public IActionResult gridState([FromBody]DataManagerRequest dm)
        {
            var query = dm;
            return Json(query);
        }

        public IActionResult GridDatasource([FromBody]DataManagerRequest dm)
        {
            var DataSource = OrdersDetails.GetAllRecords().ToList();

            DataOperations operation = new DataOperations();
            int count = DataSource.Count();
            return dm.RequiresCounts ? Json(new { result = DataSource.Skip(dm.Skip).Take(dm.Take), count = count }) : Json(DataSource);
        }
        public ActionResult Update([FromBody]CRUDModel<OrdersDetails> value)
        {
            var ord = value.Value;
            OrdersDetails val = OrdersDetails.GetAllRecords().Where(or => or.OrderID == ord.OrderID).FirstOrDefault();
            val.OrderID = ord.OrderID;
            val.ShipName = ord.ShipName;
            val.CustomerID = ord.CustomerID;
            val.ShipCountry = ord.ShipCountry;

            return Json(value.Value);
        }
        //insert the record
        public ActionResult Insert([FromBody]CRUDModel<OrdersDetails> value)
        {

            OrdersDetails.GetAllRecords().Insert(0, value.Value);
            return Json(value.Value);
        }
        //Delete the record
        public ActionResult Delete([FromBody]CRUDModel<OrdersDetails> value)
        {
            OrdersDetails.GetAllRecords().Remove(OrdersDetails.GetAllRecords().Where(or => or.OrderID == int.Parse(value.Key.ToString())).FirstOrDefault());
            return Json(value);
        }

        public class Data
        {

            public bool requiresCounts { get; set; }
            public int skip { get; set; }
            public int take { get; set; }
        }
        public class CRUDModel<T> where T : class
        {
            public string Action { get; set; }

            public string Table { get; set; }

            public string KeyColumn { get; set; }

            public object Key { get; set; }

            public T Value { get; set; }

            public List<T> Added { get; set; }

            public List<T> Changed { get; set; }

            public List<T> Deleted { get; set; }

            public IDictionary<string, object> @params { get; set; }
        }
    }
}
