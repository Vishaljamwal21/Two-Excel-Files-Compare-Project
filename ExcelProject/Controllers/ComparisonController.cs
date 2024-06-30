using System;
using System.IO;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using ExcelProject;

namespace ExcelProject.Controllers
{
    public class ComparisonController : Controller
    {
        private readonly ComparisonService _comparisonService;

        public ComparisonController(ComparisonService comparisonService)
        {
            _comparisonService = comparisonService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CompareAndGenerateOutputs(IFormFile firstFile, IFormFile secondFile)
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

            var firstOutputFilePath = Path.Combine(uploadsFolderPath, "first_output.xlsx");
            var secondOutputFilePath = Path.Combine(uploadsFolderPath, "second_output.xlsx");

            try
            {
                _comparisonService.GenerateFirstOutput(firstFilePath, secondFilePath, firstOutputFilePath);
                _comparisonService.GenerateSecondOutput(firstFilePath, secondFilePath, secondOutputFilePath);

                ViewBag.FirstOutputFileName = Path.GetFileName(firstOutputFilePath);
                ViewBag.SecondOutputFileName = Path.GetFileName(secondOutputFilePath);
                ViewBag.ThankYouMessage = "Comparison completed successfully!";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while comparing and generating outputs: {ex.Message}");
                ModelState.AddModelError("", "An error occurred while processing the files.");
            }

            return View("Index");
        }

        public IActionResult GetFirstOutputFileContent()
        {
            var uploadsFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            var firstOutputFilePath = Path.Combine(uploadsFolderPath, "first_output.xlsx");

            if (!System.IO.File.Exists(firstOutputFilePath))
            {
                return NotFound();
            }

            var memory = new MemoryStream();
            using (var stream = new FileStream(firstOutputFilePath, FileMode.Open))
            {
                stream.CopyTo(memory);
            }
            memory.Position = 0;

            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Path.GetFileName(firstOutputFilePath));
        }

        public IActionResult GetSecondOutputFileContent()
        {
            var uploadsFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
            var secondOutputFilePath = Path.Combine(uploadsFolderPath, "second_output.xlsx");

            if (!System.IO.File.Exists(secondOutputFilePath))
            {
                return NotFound();
            }

            var memory = new MemoryStream();
            using (var stream = new FileStream(secondOutputFilePath, FileMode.Open))
            {
                stream.CopyTo(memory);
            }
            memory.Position = 0;

            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Path.GetFileName(secondOutputFilePath));
        }
    }
}
