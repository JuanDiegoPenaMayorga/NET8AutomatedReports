using OfficeOpenXml;
using System.Data;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace NET8AutomatedReports
{
    public static class ExcelExporter
    {
        public static void Export(string filePath, Dictionary<string, DataTable> tables)
        {
            ExcelPackage.License.SetNonCommercialPersonal("JP");

            using var package = new ExcelPackage();
            var sheet = package.Workbook.Worksheets.Add("Sheet1");

            int startRow = 1;
            int startCol = 1;

            foreach (var kvp in tables)
            {
                string tableName = kvp.Key;
                DataTable dt = kvp.Value;

                var titleCell = sheet.Cells[startRow, startCol];
                titleCell.Value = tableName;
                titleCell.Style.Font.Bold = true;
                titleCell.Style.Font.Size = 14;

                startRow++;

                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    sheet.Cells[startRow, startCol + col].Value = dt.Columns[col].ColumnName;
                    sheet.Cells[startRow, startCol + col].Style.Font.Bold = true;
                }

                startRow++;

                if (dt.Rows.Count > 0)
                {
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        for (int col = 0; col < dt.Columns.Count; col++)
                        {
                            sheet.Cells[startRow + row, startCol + col].Value = dt.Rows[row][col];
                        }
                    }
                }
                else
                {
                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        sheet.Cells[startRow, startCol + col].Value = "N/A";
                        sheet.Cells[startRow, startCol + col].Style.Font.Italic = true;
                    }
                }
                int endRow = startRow + Math.Max(dt.Rows.Count, 1) - 1;
                int endCol = startCol + dt.Columns.Count - 1;

                if (dt.Columns.Count > 0)
                {
                    string safeTableName = SanitizeTableName(tableName + "Table");
                    var tableRange = sheet.Cells[startRow - 1, startCol, endRow, endCol];
                    var excelTable = sheet.Tables.Add(tableRange, safeTableName);
                    excelTable.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                }
                startRow = endRow + 3;
            }
            if (sheet.Dimension != null)
            {
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            }

            package.SaveAs(new FileInfo(filePath));
        }

        private static string SanitizeTableName(string name)
        {
            string clean = Regex.Replace(name, @"[^\w]", "_");
            if (char.IsDigit(clean[0]))
                clean = "T_" + clean;
            return clean.Length > 255 ? clean.Substring(0, 255) : clean;
        }
    }
}
