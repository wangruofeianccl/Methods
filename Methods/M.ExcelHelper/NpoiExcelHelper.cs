using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace M.ExcelHelper
{
    public static class NpoiExcelHelper
    {

        #region
        public static void WriteSteamToFile(MemoryStream ms, string FileName)
        {
            FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write);
            byte[] data = ms.ToArray();

            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();

            data = null;
            ms = null;
            fs = null;
        }
        public static void WriteSteamToFile(byte[] data, string FileName)
        {
            FileStream fs = new FileStream(FileName, FileMode.Create, FileAccess.Write);
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
            data = null;
            fs = null;
        }
        public static Stream WorkBookToStream(HSSFWorkbook InputWorkBook)
        {
            MemoryStream ms = new MemoryStream();
            InputWorkBook.Write(ms);
            ms.Flush();
            ms.Position = 0;
            return ms;
        }
        public static HSSFWorkbook StreamToWorkBook(Stream InputStream)
        {
            HSSFWorkbook WorkBook = new HSSFWorkbook(InputStream);
            return WorkBook;
        }
        public static HSSFWorkbook MemoryStreamToWorkBook(MemoryStream InputStream)
        {
            HSSFWorkbook WorkBook = new HSSFWorkbook(InputStream as Stream);
            return WorkBook;
        }
        public static MemoryStream WorkBookToMemoryStream(HSSFWorkbook InputStream)
        {
            //Write the stream data of workbook to the root directory
            MemoryStream file = new MemoryStream();
            InputStream.Write(file);
            return file;
        }
        public static Stream FileToStream(string FileName)
        {
            FileInfo fi = new FileInfo(FileName);
            if (fi.Exists == true)
            {
                FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                return fs;
            }
            else return null;
        }
        public static Stream MemoryStreamToStream(MemoryStream ms)
        {
            return ms as Stream;
        }
        #endregion

        #region
        /// <summary>
        /// 將DataTable轉成Stream輸出.
        /// </summary>
        /// <param name="SourceTable">The source table.</param>
        /// <returns></returns>
        public static Stream RenderDataTableToExcel(DataTable SourceTable)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            //InitializeWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet();
            IRow headerRow = sheet.CreateRow(0);

            // handling header.
            foreach (DataColumn column in SourceTable.Columns)
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            // handling value.
            int rowIndex = 1;

            foreach (DataRow row in SourceTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in SourceTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }

                rowIndex++;
            }

            workbook.Write(ms);
            ms.Flush();
            ms.Position = 0;

            sheet = null;
            headerRow = null;
            workbook = null;

            return ms;
        }
        /// <summary>
        /// 將DataTable轉成Workbook(自定資料型態)輸出.
        /// </summary>
        /// <param name="SourceTable">The source table.</param>
        /// <returns></returns>
        public static HSSFWorkbook RenderDataTableToWorkBook(DataTable SourceTable)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();
            //InitializeWorkbook();
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.CreateSheet();
            IRow headerRow = sheet.CreateRow(0);

            // handling header.
            foreach (DataColumn column in SourceTable.Columns)
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

            // handling value.
            int rowIndex = 1;

            foreach (DataRow row in SourceTable.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in SourceTable.Columns)
                {
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                }

                rowIndex++;
            }
            return workbook;
        }

        /// <summary>
        /// 將DataTable資料輸出成檔案.
        /// </summary>
        /// <param name="SourceTable">The source table.</param>
        /// <param name="FileName">Name of the file.</param>
        public static void RenderDataTableToExcel(DataTable SourceTable, string FileName)
        {
            MemoryStream ms = RenderDataTableToExcel(SourceTable) as MemoryStream;
            WriteSteamToFile(ms, FileName);
        }

        /// <summary>
        /// 從位元流讀取資料到DataTable.
        /// </summary>
        /// <param name="ExcelFileStream">The excel file stream.</param>
        /// <param name="SheetName">Name of the sheet.</param>
        /// <param name="HeaderRowIndex">Index of the header row.</param>
        /// <param name="HaveHeader">if set to <c>true</c> [have header].</param>
        /// <returns></returns>
        public static DataTable RenderDataTableFromExcel(Stream ExcelFileStream, string SheetName, int HeaderRowIndex, bool HaveHeader)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);
            //InitializeWorkbook();
            ISheet sheet = workbook.GetSheet(SheetName);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                string ColumnName = (HaveHeader == true) ? headerRow.GetCell(i).StringCellValue : "f" + i.ToString();
                DataColumn column = new DataColumn(ColumnName);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;
            int RowStart = (HaveHeader == true) ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;
            for (int i = RowStart; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                    dataRow[j] = row.GetCell(j).ToString();
            }

            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return table;
        }

        /// <summary>
        /// 從位元流讀取資料到DataTable.
        /// </summary>
        /// <param name="ExcelFileStream">The excel file stream.</param>
        /// <param name="SheetIndex">Index of the sheet.</param>
        /// <param name="HeaderRowIndex">Index of the header row.</param>
        /// <param name="HaveHeader">if set to <c>true</c> [have header].</param>
        /// <returns></returns>
        public static DataTable RenderDataTableFromExcel(Stream ExcelFileStream, int SheetIndex, int HeaderRowIndex, bool HaveHeader)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);

            ISheet sheet = workbook.GetSheetAt(SheetIndex);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                string ColumnName = (HaveHeader == true) ? headerRow.GetCell(i).StringCellValue : "f" + i.ToString();
                DataColumn column = new DataColumn(ColumnName);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;
            int RowStart = (HaveHeader == true) ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;
            for (int i = RowStart; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();
                }

                table.Rows.Add(dataRow);
            }

            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return table;
        }

        #endregion

        #region
        /// <summary>
        /// 建立datatable
        /// </summary>
        /// <param name="ColumnName">欄位名用逗號分隔</param>
        /// <param name="value">data陣列 , rowmajor</param>
        /// <returns>DataTable</returns>
        public static DataTable CreateDataTable(string ColumnName, string[,] value)
        {
            /*  輸入範例
            string cname = " name , sex ";
            string[,] aaz = new string[4, 2];
            for (int q = 0; q < 4; q++)
                for (int r = 0; r < 2; r++)
                    aaz[q, r] = "1";
            dataGridView1.DataSource = NewMediaTest1.Model.Utility.DataSetUtil.CreateDataTable(cname, aaz);
            */
            int i, j;
            DataTable ResultTable = new DataTable();
            string[] sep = new string[] { "," };

            string[] TempColName = ColumnName.Split(sep, StringSplitOptions.RemoveEmptyEntries);
            DataColumn[] CName = new DataColumn[TempColName.Length];
            for (i = 0; i < TempColName.Length; i++)
            {
                DataColumn c1 = new DataColumn(TempColName[i].ToString().Trim(), typeof(object));
                ResultTable.Columns.Add(c1);
            }
            if (value != null)
            {
                for (i = 0; i < value.GetLength(0); i++)
                {
                    DataRow newrow = ResultTable.NewRow();
                    for (j = 0; j < TempColName.Length; j++)
                    {
                        newrow[j] = string.Copy(value[i, j].ToString());

                    }
                    ResultTable.Rows.Add(newrow);
                }
            }
            return ResultTable;
        }
        /// <summary>
        /// Creates the string array.
        /// </summary>
        /// <param name="dt">The dt.</param>
        /// <returns></returns>
        public static string[,] CreateStringArray(DataTable dt)
        {
            int ColumnNum = dt.Columns.Count;
            int RowNum = dt.Rows.Count;
            string[,] result = new string[RowNum, ColumnNum];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    result[i, j] = string.Copy(dt.Rows[i][j].ToString());
                }
            }
            return result;
        }
        /// <summary>
        /// 將陣列輸出成位元流.
        /// </summary>
        /// <param name="ColumnName">Name of the column.</param>
        /// <param name="SourceTable">The source table.</param>
        /// <returns></returns>
        public static Stream RenderArrayToExcel(string ColumnName, string[,] SourceTable)
        {
            DataTable dt = CreateDataTable(ColumnName, SourceTable);
            return RenderDataTableToExcel(dt);
        }
        /// <summary>
        /// 將陣列輸出成檔案.
        /// </summary>
        /// <param name="FileName">Name of the file.</param>
        /// <param name="ColumnName">Name of the column.</param>
        /// <param name="SourceTable">The source table.</param>
        public static void RenderArrayToExcel(string FileName, string ColumnName, string[,] SourceTable)
        {
            DataTable dt = CreateDataTable(ColumnName, SourceTable);
            RenderDataTableToExcel(dt, FileName);
        }
        /// <summary>
        /// 將陣列輸出成WorkBook(自訂資料型態).
        /// </summary>
        /// <param name="ColumnName">Name of the column.</param>
        /// <param name="SourceTable">The source table.</param>
        /// <returns></returns>
        public static HSSFWorkbook RenderArrayToWorkBook(string ColumnName, string[,] SourceTable)
        {
            DataTable dt = CreateDataTable(ColumnName, SourceTable);
            return RenderDataTableToWorkBook(dt);
        }

        /// <summary>
        /// 將位元流資料輸出成陣列.
        /// </summary>
        /// <param name="ExcelFileStream">The excel file stream.</param>
        /// <param name="SheetName">Name of the sheet.</param>
        /// <param name="HeaderRowIndex">Index of the header row.</param>
        /// <param name="HaveHeader">if set to <c>true</c> [have header].</param>
        /// <returns></returns>
        public static string[,] RenderArrayFromExcel(Stream ExcelFileStream, string SheetName, int HeaderRowIndex, bool HaveHeader)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);
            ISheet sheet = workbook.GetSheet(SheetName);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;
            int RowStart = (HaveHeader == true) ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;
            for (int i = RowStart; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                    dataRow[j] = row.GetCell(j).ToString();
            }

            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return CreateStringArray(table);
        }

        /// <summary>
        /// 將位元流資料輸出成陣列.
        /// </summary>
        /// <param name="ExcelFileStream">The excel file stream.</param>
        /// <param name="SheetIndex">Index of the sheet.</param>
        /// <param name="HeaderRowIndex">Index of the header row.</param>
        /// <param name="HaveHeader">if set to <c>true</c> [have header].</param>
        /// <returns></returns>
        public static string[,] RenderArrayFromExcel(Stream ExcelFileStream, int SheetIndex, int HeaderRowIndex, bool HaveHeader)
        {
            HSSFWorkbook workbook = new HSSFWorkbook(ExcelFileStream);

            ISheet sheet = workbook.GetSheetAt(SheetIndex);

            DataTable table = new DataTable();

            IRow headerRow = sheet.GetRow(HeaderRowIndex);
            int cellCount = headerRow.LastCellNum;

            for (int i = headerRow.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                table.Columns.Add(column);
            }

            int rowCount = sheet.LastRowNum;
            int RowStart = (HaveHeader == true) ? sheet.FirstRowNum + 1 : sheet.FirstRowNum;
            for (int i = RowStart; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                DataRow dataRow = table.NewRow();

                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null)
                        dataRow[j] = row.GetCell(j).ToString();
                }

                table.Rows.Add(dataRow);
            }

            ExcelFileStream.Close();
            workbook = null;
            sheet = null;
            return CreateStringArray(table);
        }

        #endregion

    }

}
