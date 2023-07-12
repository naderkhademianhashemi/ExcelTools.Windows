using System.Collections.Generic;

namespace PayamHannan.ExcelTools
{
    public class SheetTable
    {
        public List<string> Headers { get; set; }
        public List<Dictionary<string, string>> Rows { get; set; }
    }
}