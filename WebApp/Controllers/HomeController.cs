using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WebApp.Data;
using WebApp.Models;


public class HomeController : Controller
{
    private readonly ApplicationDbContext _dbContext;

    public HomeController(ApplicationDbContext dbContext)
    {
        _dbContext = dbContext;
    }
    public IActionResult Index()
    {
        return View();
    }
    

    [HttpPost]
    public IActionResult UploadExcel(IFormFile file)
    {
        if (file == null || file.Length <= 0)
            return BadRequest("Invalid file");

        using (var stream = new MemoryStream())
        {
            file.CopyTo(stream);
            stream.Position = 0;

            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                var excelDataList = new List<ExcelData>();

                for (int row = 2; row <= rowCount; row++)
                {
                    var excelData = new ExcelData
                    {
                        FirstName = worksheet.Cells[row, 1].Value?.ToString(),
                        LastName = worksheet.Cells[row, 2].Value?.ToString(),
                        CheckoutDate = DateTime.Parse(worksheet.Cells[row, 3].Value?.ToString()),
                        Email = worksheet.Cells[row, 4].Value?.ToString(),
                        TodaysDate = DateTime.Parse(worksheet.Cells[row, 5].Value?.ToString())
                    };

                    excelDataList.Add(excelData);
                }

                _dbContext.ExcelData.AddRange(excelDataList);
                _dbContext.SaveChanges();
            }
        }

        return Ok("Excel data uploaded successfully");
    }
}