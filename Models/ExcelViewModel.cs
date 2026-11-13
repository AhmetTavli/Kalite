namespace Kalite.Models
{
    public class ExcelViewModel
    {
        public string FileName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public List<string> Descriptions { get; set; } = new();
        public List<string> Headers { get; set; } = new();
        public List<List<string>> Rows { get; set; } = new();
        public string? ErrorMessage { get; set; }
        public Dictionary<string, string> NewRow { get; set; } = new();
        public bool IsReadOnly { get; set; }
        public List<int> RowNumbers { get; set; } = new();
    }

    public class ExcelFileInfo
    {
        public string FileName { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
        public long Size { get; set; }
    }
}
