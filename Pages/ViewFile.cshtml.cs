using Kalite.Models;
using Kalite.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace Kalite.Pages
{
    public class ViewFileModel : PageModel
    {
        private readonly ExcelService _excelService;

        public ViewFileModel(ExcelService excelService)
        {
            _excelService = excelService;
        }

        public ExcelViewModel ExcelData { get; set; } = new();

        public void OnGet(string fileName)
        {
            if (!string.IsNullOrEmpty(fileName))
                ExcelData = _excelService.ReadExcel(fileName);
        }
    }
}
