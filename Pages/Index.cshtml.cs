using Kalite.Models;
using Kalite.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Kalite.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ExcelService _excelService;

        public IndexModel(ExcelService excelService)
        {
            _excelService = excelService;
        }

        public List<ExcelFileInfo> Files { get; set; } = new();
        public ExcelViewModel? ExcelData { get; set; }

        public void OnGet(string? fileName)
        {
            Files = _excelService.GetExcelFiles();

            if (!string.IsNullOrWhiteSpace(fileName))
            {
                ExcelData = _excelService.ReadExcel(fileName);
            }
        }

        public async Task<IActionResult> OnPostUploadAsync(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                TempData["Message"] = "Lütfen bir dosya seçin.";
                return RedirectToPage();
            }

            if (!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                TempData["Message"] = "Sadece .xlsx dosyaları yüklenebilir.";
                return RedirectToPage();
            }

            var path = Path.Combine(_excelService.GetExcelFolder(), file.FileName);
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            TempData["Message"] = $"Dosya '{file.FileName}' başarıyla yüklendi";
            return RedirectToPage();
        }

        public IActionResult OnGetDownloadAll()
        {
            var files = _excelService.GetExcelFiles();
            if (!files.Any())
            {
                TempData["Message"] = "İndirilecek dosya bulunamadı.";
                return RedirectToPage();
            }

            using var memoryStream = new MemoryStream();
            using (var archive = new System.IO.Compression.ZipArchive(memoryStream, System.IO.Compression.ZipArchiveMode.Create, true))
            {
                var folder = _excelService.GetExcelFolder();

                foreach (var file in files)
                {
                    var filePath = Path.Combine(folder, file.FileName);
                    if (System.IO.File.Exists(filePath))
                    {
                        var entry = archive.CreateEntry(file.FileName);
                        using var entryStream = entry.Open();
                        using var fileStream = System.IO.File.OpenRead(filePath);
                        fileStream.CopyTo(entryStream);
                    }
                }
            }

            memoryStream.Position = 0;
            var zipName = $"ExcelDosyalari_{DateTime.Now:yyyyMMdd_HHmm}.zip";
            return File(memoryStream.ToArray(), "application/zip", zipName);
        }

        public IActionResult OnGetDownloadFile(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return RedirectToPage();

            var path = Path.Combine(_excelService.GetExcelFolder(), fileName);
            if (!System.IO.File.Exists(path))
            {
                TempData["Message"] = "Dosya bulunamadı.";
                return RedirectToPage();
            }

            var bytes = System.IO.File.ReadAllBytes(path);
            return File(bytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName);
        }

        public IActionResult OnPostAddRow(string fileName, List<string> values)
        {
            if (string.IsNullOrWhiteSpace(fileName) || values == null || !values.Any(v => !string.IsNullOrWhiteSpace(v)))
            {
                TempData["Message"] = "Lütfen en az bir alan doldurun.";
                return RedirectToPage(new { fileName });
            }

            var path = Path.Combine(_excelService.GetExcelFolder(), fileName);
            if (!System.IO.File.Exists(path))
            {
                TempData["Message"] = "Dosya bulunamadı.";
                return RedirectToPage(new { fileName });
            }

            var fileInfo = new FileInfo(path);
            if (fileInfo.IsReadOnly)
            {
                TempData["Message"] = "Bu dosya salt okunurdur. Yeni satır eklenemez.";
                return RedirectToPage(new { fileName });
            }

            try
            {
                using var workbook = new ClosedXML.Excel.XLWorkbook(path);

                string sheetName = Path.GetFileNameWithoutExtension(fileName);
                ClosedXML.Excel.IXLWorksheet worksheet;

                if (workbook.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                    worksheet = workbook.Worksheet(sheetName);
                else if (workbook.Worksheets.Count >= 2)
                    worksheet = workbook.Worksheet(2);
                else
                    worksheet = workbook.Worksheet(1);

                // 1. Find the header row (row with the most filled cells in the top part)
                int headerRow = 1;
                int maxFilled = 0;
                int scanLimit = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 50, 50);

                for (int r = 1; r <= scanLimit; r++)
                {
                    int filled = 0;
                    for (int c = 1; c <= 20; c++)
                    {
                        if (!worksheet.Cell(r, c).IsEmpty()) filled++;
                    }
                    if (filled > maxFilled)
                    {
                        maxFilled = filled;
                        headerRow = r;
                    }
                }

                // 2. Find the last real data row (stop after 10 consecutive empty rows)
                int lastDataRow = headerRow;
                int consecutiveEmpty = 0;
                int maxCol = Math.Max(values.Count, 15);
                int absoluteMax = worksheet.LastRowUsed()?.RowNumber() ?? headerRow + 100;

                for (int r = headerRow + 1; r <= absoluteMax; r++)
                {
                    bool isEmpty = true;
                    for (int c = 1; c <= maxCol; c++)
                    {
                        if (!worksheet.Cell(r, c).IsEmpty())
                        {
                            isEmpty = false;
                            break;
                        }
                    }

                    if (isEmpty)
                    {
                        consecutiveEmpty++;
                        if (consecutiveEmpty >= 10)
                            break; // stop – we are past the real data
                    }
                    else
                    {
                        consecutiveEmpty = 0;
                        lastDataRow = r;
                    }
                }

                // 3. Add the new row right after the last real data row
                int newRow = lastDataRow + 1;

                for (int i = 0; i < values.Count; i++)
                {
                    worksheet.Cell(newRow, i + 1).Value = values[i] ?? "";
                }

                workbook.Save();

                TempData["Message"] = "Yeni satır eklendi.";
                return RedirectToPage(new { fileName });
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Hata oluştu: " + ex.Message;
                return RedirectToPage(new { fileName });
            }
        }

        public IActionResult OnPostDeleteRows(string fileName, List<int> rowNumbers)
        {
            if (string.IsNullOrWhiteSpace(fileName) || rowNumbers == null || !rowNumbers.Any())
            {
                TempData["Message"] = "Lütfen silmek için en az bir satır seçin.";
                return RedirectToPage(new { fileName });
            }

            var path = Path.Combine(_excelService.GetExcelFolder(), fileName);
            if (!System.IO.File.Exists(path))
            {
                TempData["Message"] = "Dosya bulunamadı.";
                return RedirectToPage(new { fileName });
            }

            var fileInfo = new FileInfo(path);
            if (fileInfo.IsReadOnly)
            {
                TempData["Message"] = "Bu dosya salt okunurdur. Satır silinemez.";
                return RedirectToPage(new { fileName });
            }

            try
            {
                using var workbook = new ClosedXML.Excel.XLWorkbook(path);

                string sheetName = Path.GetFileNameWithoutExtension(fileName);
                ClosedXML.Excel.IXLWorksheet worksheet;

                if (workbook.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                    worksheet = workbook.Worksheet(sheetName);
                else if (workbook.Worksheets.Count >= 2)
                    worksheet = workbook.Worksheet(2);
                else
                    worksheet = workbook.Worksheet(1);

                // Delete from highest row to lowest so indexes stay correct
                foreach (var rowNumber in rowNumbers.OrderByDescending(r => r))
                {
                    if (rowNumber > 0 && rowNumber <= (worksheet.LastRowUsed()?.RowNumber() ?? 0))
                    {
                        worksheet.Row(rowNumber).Delete();
                    }
                }

                workbook.Save();   // Save changes to the same file

                TempData["Message"] = "Seçilen satırlar silindi.";
                return RedirectToPage(new { fileName });
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Hata oluştu: " + ex.Message;
                return RedirectToPage(new { fileName });
            }
        }

        public IActionResult OnPostPrepareEdit(string fileName, List<int> rowNumbers)
        {
            if (rowNumbers == null || rowNumbers.Count != 1)
            {
                TempData["Message"] = "Lütfen düzenlemek için tam olarak bir satır seçin.";
                return RedirectToPage(new { fileName });
            }

            // Store the selected row number so the page can show the edit form
            TempData["EditRowNumber"] = rowNumbers[0];
            return RedirectToPage(new { fileName });
        }

        public IActionResult OnPostUpdateRow(string fileName, int rowNumber, List<string> values)
        {
            if (string.IsNullOrWhiteSpace(fileName) || rowNumber <= 0 || values == null)
            {
                TempData["Message"] = "Geçersiz istek.";
                return RedirectToPage(new { fileName });
            }

            var path = Path.Combine(_excelService.GetExcelFolder(), fileName);
            if (!System.IO.File.Exists(path))
            {
                TempData["Message"] = "Dosya bulunamadı.";
                return RedirectToPage(new { fileName });
            }

            var fileInfo = new FileInfo(path);
            if (fileInfo.IsReadOnly)
            {
                TempData["Message"] = "Bu dosya salt okunurdur. Güncelleme yapılamaz.";
                return RedirectToPage(new { fileName });
            }

            try
            {
                using var workbook = new ClosedXML.Excel.XLWorkbook(path);

                string sheetName = Path.GetFileNameWithoutExtension(fileName);
                ClosedXML.Excel.IXLWorksheet worksheet;

                if (workbook.Worksheets.Any(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
                    worksheet = workbook.Worksheet(sheetName);
                else if (workbook.Worksheets.Count >= 2)
                    worksheet = workbook.Worksheet(2);
                else
                    worksheet = workbook.Worksheet(1);

                for (int i = 0; i < values.Count; i++)
                {
                    worksheet.Cell(rowNumber, i + 1).Value = values[i] ?? "";
                }

                workbook.Save();

                TempData["Message"] = "Satır güncellendi.";
                return RedirectToPage(new { fileName });
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Hata oluştu: " + ex.Message;
                return RedirectToPage(new { fileName });
            }
        }

    }
}
