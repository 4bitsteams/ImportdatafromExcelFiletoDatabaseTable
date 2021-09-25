using ImportdatafromExcelFiletoDatabaseTable.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ImportdatafromExcelFiletoDatabaseTable.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public async Task<List<Subject>> Import(IFormFile formFIle)
        {
            var SubjectList = new List<Subject>();
            using (var stream = new MemoryStream())
            {
                await formFIle.CopyToAsync(stream);
                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet excelWorksheet = package.Workbook.Worksheets[0];
                    var rowCount = excelWorksheet.Dimension.Rows;
                    for (int row = 0; row < rowCount; row++)
                    {
                        SubjectList.Add(new Subject
                        {

                            RefId = (int)excelWorksheet.Cells[row, 1].Value,
                            Code = excelWorksheet.Cells[row, 2].Value.ToString().Trim(),
                            Name = excelWorksheet.Cells[row, 3].Value.ToString().Trim(),
                            Description = excelWorksheet.Cells[row, 3].Value.ToString().Trim(),
                        });
                    }
                }
            }

            return SubjectList;
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
    }
}
