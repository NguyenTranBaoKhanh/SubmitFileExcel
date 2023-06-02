using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using MVC.Models;
using OfficeOpenXml;

namespace MVC.Controllers;

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

    public async Task<List<countries>> Import(IFormFile file)
    {
        var list = new List<countries>();
        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);
            using (var package = new ExcelPackage(stream))
            {

                // must install EPPLus version 4.5.2.1 will not be licences
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var rowcount = worksheet.Dimension.Rows;
                for (int row = 2; row < rowcount; row++)
                {
                    list.Add(new countries
                    {
                        CountryId = worksheet.Cells[row, 1].Value.ToString().Trim(),
                        CountryName = worksheet.Cells[row, 2].Value.ToString().Trim(),
                        TowCharCountryCode = worksheet.Cells[row, 3].Value.ToString().Trim(),
                        ThreeCharCountryCode = worksheet.Cells[row, 4].Value.ToString().Trim(),
                    });
                }
            }
        }
        return list;
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
