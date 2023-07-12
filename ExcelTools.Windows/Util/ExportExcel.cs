using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTools.Windows.Util
{
    public static class ExportExcel
    {
        private static DataTable RowsToDataTable(List<Dictionary<string, object>> Rows, Dictionary<string, string> columnTitles, bool filterByTitles)
        {
            var columns = Rows.SelectMany(k => k.Keys).Distinct().OrderBy(k => k).ToArray().Distinct().ToList();
            
            var dataTable = new DataTable();
            dataTable.Columns.AddRange(columns.Select(c => new DataColumn(c) { }).ToArray());
            var i = 1;
            Rows.ForEach(item =>
            {
                var row = dataTable.NewRow();
                columns.ForEach(c =>
                {
                    var hasColumn = item.TryGetValue(c, out object data);

                    row[c] = data;
                });
                dataTable.Rows.Add(row);

                i++;
            });
            return dataTable;
        }

        private static Workbook Export(DataTable table)
        {
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns)
                {
                    column.ReadOnly = false;
                    if (row[column].GetType() == typeof(string))
                    {
                        var value = row[column].ToString();
                        if (value.Length > 32572)
                            value = value.Substring(0, 32572);
                        if (value.StartsWith("=")) value = " " + value;
                        row[column] = System.Text.RegularExpressions.Regex.Replace(value, @"<(.|\n)*?>", "");

                    }
                }
            }

            Workbook workbook = new Workbook
            {
                Version = ExcelVersion.Version2010
            };
            var sheet = workbook.Worksheets[0];

            sheet.InsertDataTable(table, true, 1, 1);
            return workbook;
        }

        public static Workbook Export(this List<Dictionary<string, object>> Rows, Dictionary<string, string> columnTitles = null, bool filterByTitles = false)
        {
            DataTable dataTable = RowsToDataTable(Rows, columnTitles, filterByTitles);

            Workbook workbook = Export(dataTable);
            return workbook;
        }

        public static void SaveToFile(this List<Dictionary<string, object>> rows, string filename)
        {
            var workbook = Export(rows);
            workbook.SaveToFile(filename, ExcelVersion.Version2010);

        }

    }
}
