using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelProject
{
    public class ComparisonService
    {
        public void GenerateFirstOutput(string firstFilePath, string secondFilePath, string outputFilePath)
        {
            var mergedData = CompareAndMerge(firstFilePath, secondFilePath);
            using (var firstPackage = new ExcelPackage(new FileInfo(firstFilePath)))
            {
                var worksheet = firstPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    Console.WriteLine($"Error: No worksheet found in the Excel file: {firstFilePath}");
                    return;
                }
                RemoveDuplicatesAndHighlightBlanks(mergedData, worksheet);
            }

            WriteExcelData(outputFilePath, mergedData);
        }


        public void GenerateSecondOutput(string firstFilePath, string secondFilePath, string outputFilePath)
        {
            var differences = GetDifferences(firstFilePath, secondFilePath);
            RemoveDuplicates(differences);
            WriteExcelData(outputFilePath, differences);
        }

        private List<List<string>> CompareAndMerge(string firstFilePath, string secondFilePath)
        {
            var firstData = ReadExcelData(firstFilePath);
            var secondData = ReadExcelData(secondFilePath);
            return firstData.Concat(secondData).ToList();
        }

        private List<List<string>> GetDifferences(string firstFilePath, string secondFilePath)
        {
            var firstData = ReadExcelData(firstFilePath);
            var secondData = ReadExcelData(secondFilePath);

            var differences = firstData.Except(secondData, new ListComparer()).ToList();
            differences.AddRange(secondData.Except(firstData, new ListComparer()));

            return differences;
        }
        public void RemoveDuplicatesAndHighlightBlanks(List<List<string>> data, ExcelWorksheet worksheet)
        {
            var distinctData = data.Distinct(new ListComparer()).ToList();
            for (int i = 0; i < data.Count; i++)
            {
                bool hasBlankCell = false;
                for (int j = 0; j < data[i].Count; j++)
                {
                    if (string.IsNullOrWhiteSpace(data[i][j]))
                    {
                        hasBlankCell = true;
                        break;
                    }
                }
                if (!hasBlankCell)
                {
                    distinctData.Remove(data[i]);
                }
                else
                {
                    for (int j = 0; j < data[i].Count; j++)
                    {
                        if (string.IsNullOrWhiteSpace(data[i][j]))
                        {
                            var cell = worksheet.Cells[i + 1, j + 1];
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                        }
                    }
                }
            }
            data.Clear();
            data.AddRange(distinctData);
        }


        private void RemoveDuplicates(List<List<string>> data)
        {
            data.RemoveAll(row => data.Count(r => r.SequenceEqual(row)) > 1);
        }

        private List<List<string>> ReadExcelData(string filePath)
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

        private void WriteExcelData(string filePath, List<List<string>> data)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Output Data");
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

    }
}
