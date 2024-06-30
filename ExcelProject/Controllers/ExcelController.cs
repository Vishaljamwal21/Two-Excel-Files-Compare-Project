using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExcelProject.Controllers
{
    public class ExcelController : Controller
    {
        private readonly ExcelService _excelService;

        public ExcelController(ExcelService excelService)
        {
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult UploadFiles(IFormFile firstFile, IFormFile secondFile)
        {
            if (firstFile == null || secondFile == null)
            {
                ModelState.AddModelError("", "Please select both files.");
                return View("Index");
            }
            if (Path.GetExtension(firstFile.FileName) != ".xlsx" || Path.GetExtension(secondFile.FileName) != ".xlsx")
            {
                ModelState.AddModelError("", "Please upload Excel files with .xlsx extension.");
                return View("Index");
            }

            // Validate file size (max 10 MB)
            if (firstFile.Length > 10 * 1024 * 1024 || secondFile.Length > 10 * 1024 * 1024)
            {
                ModelState.AddModelError("", "File size exceeds the maximum limit of 10 MB.");
                return View("Index");
            }

            var uploadsFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            Directory.CreateDirectory(uploadsFolderPath);

            var firstFilePath = Path.Combine(uploadsFolderPath, firstFile.FileName);
            var secondFilePath = Path.Combine(uploadsFolderPath, secondFile.FileName);

            using (var firstFileStream = new FileStream(firstFilePath, FileMode.Create))
            {
                firstFile.CopyTo(firstFileStream);
            }

            using (var secondFileStream = new FileStream(secondFilePath, FileMode.Create))
            {
                secondFile.CopyTo(secondFileStream);
            }

            var resultFilePath = Path.Combine(uploadsFolderPath, "merged.xlsx");

            try
            {
                var mergedData = _excelService.CompareAndMerge(firstFilePath, secondFilePath, resultFilePath);
                ViewBag.MergedFileName = Path.GetFileName(resultFilePath);
                ViewBag.ThankYouMessage = "Files uploaded successfully!";
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"An error occurred while comparing and merging Excel files: {ex.Message}");
                ModelState.AddModelError("", "An error occurred while processing the files.");
            }

            return View("Index");
        }


        public IActionResult GetMergedFileContent()
        {
            try
            {
                var resultFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads", "merged.xlsx");

                if (!System.IO.File.Exists(resultFilePath))
                {
                    Console.WriteLine($"Merged file does not exist at path: {resultFilePath}");
                    return NotFound();
                }
                var mergedData = _excelService.ReadExcelData(resultFilePath);
                if (mergedData.Count == 0)
                {
                    Console.WriteLine("No data found in the merged file.");
                    return NoContent();
                }
                _excelService.RemoveDuplicates(resultFilePath); // Remove duplicates from merged file
                _excelService.HighlightBlankCells(resultFilePath); // Highlight blank cells
                var cleanedMergedFileContent = System.IO.File.ReadAllBytes(resultFilePath);
                return File(cleanedMergedFileContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "cleaned_merged.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating cleaned merged file content: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }
    }
}