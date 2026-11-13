using ClosedXML.Excel;
using Kalite.Models;

namespace Kalite.Services
{
    public class ExcelService
    {
        private readonly string _excelFolder;

        public ExcelService(IWebHostEnvironment env)
        {
            _excelFolder = Path.Combine(env.ContentRootPath, "App_Data", "ExcelFiles");
            Directory.CreateDirectory(_excelFolder);
        }

        public List<ExcelFileInfo> GetExcelFiles()
        {
            return Directory.GetFiles(_excelFolder, "*.xlsx")
                .Select(f => new FileInfo(f))
                .Select(fi => new ExcelFileInfo
                {
                    FileName = fi.Name,
                    LastModified = fi.LastWriteTime,
                    Size = fi.Length
                })
                .OrderByDescending(f => f.LastModified)
                .ToList();
        }

        public ExcelViewModel ReadExcel(string fileName)
        {
            var result = new ExcelViewModel { FileName = fileName };
            var path = Path.Combine(_excelFolder, fileName);

            if (!File.Exists(path))
            {
                result.ErrorMessage = "Hata: Seçilen dosya bulunamadı.";
                return result;
            }

            var fileInfo = new FileInfo(path);
            result.IsReadOnly = fileInfo.IsReadOnly;

            try
            {
                using var workbook = new XLWorkbook(path);

                // string sheetName = Path.GetFileNameWithoutExtension(fileName);
                // var worksheet = workbook.Worksheet(1);

                string sheetName = Path.GetFileNameWithoutExtension(fileName);
                IXLWorksheet worksheet;

                if (workbook.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    worksheet = workbook.Worksheet(sheetName);
                }
                else if (workbook.Worksheets.Count >= 2)
                {
                    // Fallback: second sheet (common in YÖKSİS files)
                    worksheet = workbook.Worksheet(2);
                }
                else
                {
                    // Last fallback: first sheet
                    worksheet = workbook.Worksheet(1);
                }


                var range = worksheet.RangeUsed();
                if (range == null)
                {
                    result.ErrorMessage = "Hata: Seçilen çalışma sayfasın boştur.";
                    return result;
                }

                int maxScanRow = Math.Min(range.LastRow().RowNumber(), 50);

                // Find the real header row (row with the most non-empty cells)
                int bestHeaderRow = 1;
                int maxFilledCells = 0;

                for (int r = 1; r <= maxScanRow; r++)
                {
                    int filled = 0;
                    for (int c = 1; c <= 25; c++)
                    {
                        if (!worksheet.Cell(r, c).IsEmpty())
                            filled++;
                    }

                    if (filled > maxFilledCells)
                    {
                        maxFilledCells = filled;
                        bestHeaderRow = r;
                    }
                }

                if (maxFilledCells < 2)
                {
                    result.ErrorMessage = "Hata: Uygun bir başlık satırı bulunamadı.";
                    return result;
                }

                // Title = first non-empty cell in column A (before the header)
                for (int r = 1; r < bestHeaderRow; r++)
                {
                    var cellValue = worksheet.Cell(r, 1).GetString().Trim();
                    if (!string.IsNullOrEmpty(cellValue) && string.IsNullOrEmpty(result.Title))
                    {
                        result.Title = cellValue;
                        break;
                    }
                }

                // Description lines = all non-empty text in column A between title and header
                for (int r = 1; r < bestHeaderRow; r++)
                {
                    var cellValue = worksheet.Cell(r, 1).GetString().Trim();
                    if (!string.IsNullOrEmpty(cellValue) && cellValue != result.Title)
                        result.Descriptions.Add(cellValue);
                }

                // Read real headers
                int lastCol = 0;
                for (int c = 1; c <= 25; c++)
                {
                    var value = worksheet.Cell(bestHeaderRow, c).GetString().Trim();
                    if (string.IsNullOrEmpty(value) && lastCol > 0)
                        break;

                    if (!string.IsNullOrEmpty(value))
                        lastCol = c;

                    result.Headers.Add(string.IsNullOrEmpty(value) ? $"Column {c}" : value);
                }

                // Remove trailing empty column headers
                while (result.Headers.Count > 0 && result.Headers.Last().StartsWith("Column "))
                    result.Headers.RemoveAt(result.Headers.Count - 1);

                lastCol = result.Headers.Count;

                // Read data rows (skip ZORUNLU row)
                // int consecutiveEmpty = 0;

                // Read all non-empty rows (no early stop)
                int maxRows = 5000; // safety limit

                for (int r = bestHeaderRow + 1; r <= Math.Min(range.LastRow().RowNumber(), bestHeaderRow + maxRows); r++)
                {
                    var rowValues = new List<string>();
                    bool isEmpty = true;
                    bool isRequirementRow = false;

                    for (int c = 1; c <= lastCol; c++)
                    {
                        var cell = worksheet.Cell(r, c);
                        string value = cell.IsEmpty() ? "" : cell.GetFormattedString().Trim();
                        rowValues.Add(value);

                        if (!string.IsNullOrWhiteSpace(value))
                            isEmpty = false;

                        if (value.Contains("ZORUNLU", StringComparison.OrdinalIgnoreCase) ||
                            value.Contains("Zorunlu", StringComparison.OrdinalIgnoreCase) ||
                            value.Contains("Zorun Alan", StringComparison.OrdinalIgnoreCase))
                            isRequirementRow = true;
                    }

                    // Skip completely empty rows, but do not stop
                    if (isEmpty)
                        continue;

                    if (!isRequirementRow)
                    {
                        result.Rows.Add(rowValues);
                        result.RowNumbers.Add(r);
                    }
                }

                if (!result.Rows.Any() && string.IsNullOrEmpty(result.ErrorMessage))
                    result.ErrorMessage = "No data rows found.";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Error reading Excel: {ex.Message}";
            }

            return result;
        }

        public string GetExcelFolder() => _excelFolder;

        public string AddRowAndSaveAsNewFile(string originalFileName, List<string> newRowValues)
        {
            var originalPath = Path.Combine(_excelFolder, originalFileName);
            if (!File.Exists(originalPath))
                throw new FileNotFoundException("Hata: Orijinal dosya bulunamadı.");

            var newFileName = Path.GetFileNameWithoutExtension(originalFileName)
                              + "_yeni_"
                              + DateTime.Now.ToString("yyyyMMdd_HHmm")
                              + ".xlsx";
            var newPath = Path.Combine(_excelFolder, newFileName);

            using var workbook = new XLWorkbook(originalPath);
            var worksheet = workbook.Worksheet(1);

            // Find the last used row
            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            int newRowNumber = lastRow + 1;

            // Write the new row
            for (int i = 0; i < newRowValues.Count; i++)
            {
                worksheet.Cell(newRowNumber, i + 1).Value = newRowValues[i] ?? "";
            }

            workbook.SaveAs(newPath);
            return newFileName;
        }
    }
}
