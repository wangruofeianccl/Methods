using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace M.OpenXmlHelper
{
    public static class ExcelHelper
    {
        /// <summary>
        /// The title.
        /// </summary>
        static string title = string.Empty;

        /// <summary>
        /// The date.
        /// </summary>
        static string date = string.Empty;

        /// <summary>
        /// The user.
        /// </summary>
        static string user = string.Empty;
        /// <summary>
        /// The a.
        /// </summary>
        private const string A = "A", B = "B", C = "C", D = "D", E = "E", F = "F", G = "G";

        static DateTime excelStartTime = DateTime.Parse("1900-01-01").AddDays(-2);
        /// <summary>
        /// 读取Excel数据
        /// </summary>
        /// <param name="rows">
        /// The rows.
        /// </param>
        /// <param name="sharedStringTable">
        /// The shared string table.
        /// </param>
        /// <returns>
        /// DataTable
        /// </returns>
        private static System.Data.DataTable ReadExcelData(List<Row> rows, SharedStringTablePart sharedStringTable, System.Data.DataTable dt)
        {
            //var dt = CreateDataTable();

            ReadExcelTitle(rows, sharedStringTable);

            ReadExcelRows(rows, sharedStringTable, dt);

            return dt;
        }
        /// <summary>
        /// The read excel title.
        /// </summary>
        /// <param name="rows">
        /// The rows.
        /// </param>
        /// <param name="sharedStringTable">
        /// The shared string table.
        /// </param>
        private static void ReadExcelTitle(List<Row> rows, SharedStringTablePart sharedStringTable)
        {
            title = rows.GetCells(1).GetCellValue("A1", sharedStringTable);
            var row2Cells = rows.GetCells(2);
            date = row2Cells.GetCellValue("A2", sharedStringTable);
            user = row2Cells.GetCellValue("G2", sharedStringTable);
        }

        private static void ReadExcelRows(List<Row> rows, SharedStringTablePart sharedStringTable, System.Data.DataTable dt)
        {
            for (var i = 0; i < rows.Where(x => x.RowIndex.Value > 3).GetRowsCount(); i++)
            {
                var row = dt.NewRow();
                int rowIndex = 4 + i;
                var cells = rows.GetCells(rowIndex);
                row[A] = cells.GetCellValue(A + rowIndex, sharedStringTable);
                row[B] = cells.GetCellValue(B + rowIndex, sharedStringTable);
                row[C] = cells.GetCellValue(C + rowIndex, sharedStringTable);
                row[D] = cells.GetCellValue(D + rowIndex, sharedStringTable);
                row[F] = cells.GetCellValue(F + rowIndex, sharedStringTable);
                row[G] = cells.GetCellValue(G + rowIndex, sharedStringTable);
                var eVal = cells.GetCellValue(E + rowIndex, sharedStringTable);
                DateTime timeVal;
                double doubleVal;

                DateTime.TryParse(eVal, out timeVal);
                double.TryParse(eVal, out doubleVal);
                if (timeVal > DateTime.MinValue)
                {
                    row[E] = timeVal;
                }
                else if (doubleVal > 0)
                {
                    row[E] = excelStartTime.AddDays(doubleVal);
                }
                else
                {
                    row[E] = "时间格式不正确";
                }

                dt.Rows.Add(row);
            }
        }

    }
}
