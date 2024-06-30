using Microsoft.EntityFrameworkCore.ChangeTracking;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelProject
{
    public class ExcelService
    {
        public void RemoveDuplicates(string filePath)
        {
            var data = ReadExcelData(filePath);
            var distinctData = data.Distinct(new ListComparer()).ToList();
            WriteExcelData(filePath, distinctData);
        }
        public List<List<string>> CompareAndMerge(string firstFilePath, string secondFilePath, string resultFilePath)
        {
            try
            {
                var firstData = ReadExcelData(firstFilePath);
                var secondData = ReadExcelData(secondFilePath);
                var mergedData = firstData.Concat(secondData).ToList();
                var distinctData = mergedData.Distinct(new ListComparer()).ToList();
                WriteExcelData(resultFilePath, distinctData);
                return distinctData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while comparing and merging Excel files: {ex.Message}");
                return null;
            }
        }
        public List<List<string>> ReadExcelData(string filePath)
        {
            var data = new List<List<string>>();

            try
            {
                if (!System.IO.File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File not found at path: {filePath}");
                    return data;
                }

                var fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        Console.WriteLine($"Error: No worksheet found in the Excel file: {filePath}");
                        return data;
                    }
                    var startRow = worksheet.Dimension.Start.Row;
                    var endRow = worksheet.Dimension.End.Row;
                    var startColumn = worksheet.Dimension.Start.Column;
                    var endColumn = worksheet.Dimension.End.Column;
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = startColumn; col <= endColumn; col++)
                        {
                            rowData.Add(worksheet.Cells[row, col].Value?.ToString());
                        }
                        data.Add(rowData);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while reading Excel data from file '{filePath}': {ex.Message}");
            }
            return data;
        }

        public void WriteExcelData(string filePath, List<List<string>> data)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Merged Data");
                    for (int row = 0; row < data.Count; row++)
                    {
                        for (int col = 0; col < data[row].Count; col++)
                        {
                            worksheet.Cells[row + 1, col + 1].Value = data[row][col];
                        }
                    }
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while writing Excel data to file '{filePath}': {ex.Message}");
            }
        }
        public void HighlightBlankCells(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        Console.WriteLine($"Error: No worksheet found in the Excel file: {filePath}");
                        return;
                    }
                    var startRow = worksheet.Dimension.Start.Row;
                    var endRow = worksheet.Dimension.End.Row;
                    var startColumn = worksheet.Dimension.Start.Column;
                    var endColumn = worksheet.Dimension.End.Column;

                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startColumn; col <= endColumn; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            if (cell.Value == null || string.IsNullOrEmpty(cell.Text))
                            {
                                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                            }
                        }
                    }
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while highlighting blank cells in Excel file '{filePath}': {ex.Message}");
            }
        }

    }
}
