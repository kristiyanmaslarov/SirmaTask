using Infrastructure;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Sirma.Web.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sirma.Web.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IHostingEnvironment _hostingEnvironment;
        public List<Employee> employees = new List<Employee>();

        public HomeController(ILogger<HomeController> logger, IHostingEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public ActionResult Download()

        {

            string Files = "wwwroot/UploadExcel/employees";

            byte[] fileBytes = System.IO.File.ReadAllBytes(Files);

            System.IO.File.WriteAllBytes(Files, fileBytes);

            MemoryStream ms = new MemoryStream(fileBytes);

            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, "employee.xlsx");

        }

        public ActionResult Import()

        {


            IFormFile file = Request.Form.Files[0];

            string folderName = "UploadExcel";

            string webRootPath = _hostingEnvironment.WebRootPath;

            string newPath = Path.Combine(webRootPath, folderName);



            StringBuilder sb = new StringBuilder();

            if (!Directory.Exists(newPath))

            {

                Directory.CreateDirectory(newPath);

            }

            if (file.Length > 0)

            {

                string sFileExtension = Path.GetExtension(file.FileName).ToLower();

                ISheet sheet;

                string fullPath = Path.Combine(newPath, file.FileName);

                using (var stream = new FileStream(fullPath, FileMode.Create))

                {

                    file.CopyTo(stream);


                    stream.Position = 0;

                    if (sFileExtension == ".xls")

                    {
                        HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  

                        sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook

                        for (int r = 0; r < sheet.LastRowNum; r++)
                        {
                            var row = sheet.GetRow(r);
                            var employee = new Employee();

                            employee.ID = int.Parse(row.GetCell(0).ToString());
                            employee.ProjectID = int.Parse(row.GetCell(1).ToString());
                            employee.DateFrom = DateTime.Parse(row.GetCell(2).ToString());
                            employee.DateTo = DateTime.Parse(row.GetCell(3).ToString());

                            employees.Add(employee);
                        }
                    }

                    else

                    {
                        XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  

                        sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  



                        for (int r = 0; r < sheet.LastRowNum; r++)
                        {
                            var row = sheet.GetRow(r);
                            var employee = new Employee();

                            employee.ID = int.Parse(row.GetCell(0).ToString());
                            employee.ProjectID = int.Parse(row.GetCell(1).ToString());
                            employee.DateFrom = DateTime.Parse(row.GetCell(2).ToString());
                            employee.DateTo = DateTime.Parse(row.GetCell(3).ToString());

                            employees.Add(employee);
                        }

                    }

                    IRow headerRow = sheet.GetRow(0); //Get Header Row

                    int cellCount = headerRow.LastCellNum;

                    sb.Append("<table class='table table-bordered'><tr>");

                    for (int j = 0; j < cellCount; j++)

                    {

                        NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);

                        if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;

                        sb.Append("<th>" + cell.ToString() + "</th>");

                    }

                    sb.Append("</tr>");

                    sb.AppendLine("<tr>");

                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File

                    {

                        IRow row = sheet.GetRow(i);

                        if (row == null) continue;

                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;

                        for (int j = row.FirstCellNum; j < cellCount; j++)

                        {

                            if (row.GetCell(j) != null)

                                sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");

                        }

                        sb.AppendLine("</tr>");

                    }

                    sb.Append("</table>");

                }

            }


            return this.Content(sb.ToString());
        }

        public async Task<IActionResult> Export()

        {

            string sWebRootFolder = _hostingEnvironment.WebRootPath;

            string sFileName = @"Employees.xlsx";

            string URL = string.Format("{0}://{1}/{2}", Request.Scheme, Request.Host, sFileName);

            FileInfo file = new FileInfo(Path.Combine(sWebRootFolder, sFileName));

            var memory = new MemoryStream();

            using (var fs = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Create, FileAccess.Write))

            {

                IWorkbook workbook;

                workbook = new XSSFWorkbook();

                ISheet excelSheet = workbook.CreateSheet("employee");

                IRow row = excelSheet.CreateRow(0);


                row.CreateCell(0).SetCellValue("ID#1");

                row.CreateCell(1).SetCellValue("Id#2");

                row.CreateCell(2).SetCellValue("ProjectId");

                row.CreateCell(3).SetCellValue("DaysWorked");

                var employee = employees.FirstOrDefault();

                row = excelSheet.CreateRow(1);

                row.CreateCell(0).SetCellValue(employee.ID);

                row.CreateCell(1).SetCellValue(employee.ID);

                row.CreateCell(2).SetCellValue(employee.ProjectID);

                row.CreateCell(3).SetCellValue("13");






                workbook.Write(fs);

            }

            using (var stream = new FileStream(Path.Combine(sWebRootFolder, sFileName), FileMode.Open))

            {

                await stream.CopyToAsync(memory);

            }

            memory.Position = 0;

            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", sFileName);

        }

    }
}