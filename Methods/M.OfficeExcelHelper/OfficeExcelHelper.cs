using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace M.OfficeExcelHelper
{
    public static class OfficeExcelHelper
    {
        public static Dictionary<string, int> GetSheet(string excelFilePath)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Sheets sheets;
            object oMissiong = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            try
            {
                if (app == null) return null;
                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong,
                    oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                sheets = workbook.Worksheets;
                dic.Add("请选择...", 0);
                foreach (Worksheet item in sheets)
                {
                    dic.Add(item.Name, item.Index);
                }
                return dic;

            }
            catch { return null; }
            finally
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }
        }



        public static System.Data.DataTable GetDataFromExcelByCom(string excelFilePath, bool hasTitle = false, int sheetIndex = 1)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Sheets sheets;
            object oMissiong = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            System.Data.DataTable dt = new System.Data.DataTable();

            try
            {
                if (app == null) return null;
                workbook = app.Workbooks.Open(excelFilePath, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong,
                    oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                sheets = workbook.Worksheets;

                //将数据读入到DataTable中
                Microsoft.Office.Interop.Excel.Worksheet worksheet =
                    (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(sheetIndex); //读取第一张表  
                if (worksheet == null) return null;

                int iRowCount = worksheet.UsedRange.Rows.Count;
                int iColCount = worksheet.UsedRange.Columns.Count;
                //生成列头
                for (int i = 0; i < iColCount; i++)
                {
                    var name = "column" + i;
                    if (hasTitle)
                    {
                        var txt = ((Microsoft.Office.Interop.Excel.Range)worksheet.Cells[1, i + 1]).Text.ToString();
                        if (!string.IsNullOrWhiteSpace(txt)) name = txt;
                    }
                    while (dt.Columns.Contains(name)) name = name + "_1"; //重复行名称会报错。
                    dt.Columns.Add(new DataColumn(name, typeof(string)));
                }
                //生成行数据
                Microsoft.Office.Interop.Excel.Range range;
                int rowIdx = hasTitle ? 2 : 1;
                for (int iRow = rowIdx; iRow <= iRowCount; iRow++)
                {
                    DataRow dr = dt.NewRow();
                    for (int iCol = 1; iCol <= iColCount; iCol++)
                    {
                        range = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[iRow, iCol];
                        dr[iCol - 1] = (range.Value2 == null) ? "" : range.Text.ToString();
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            catch (Exception EX)
            {
                return null;
            }
            finally
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }
        }

        //把数据表的内容导出到Excel文件中  
        public static void OutDataToExcel(System.Data.DataTable srcDataTable, string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            object missing = System.Reflection.Missing.Value;

            //导出到execl   
            try
            {
                if (xlApp == null)
                {
                    //MessageBox.Show("无法创建Excel对象，可能您的电脑未安装Excel!");
                    return;
                }

                Microsoft.Office.Interop.Excel.Workbooks xlBooks = xlApp.Workbooks;
                Microsoft.Office.Interop.Excel.Workbook xlBook = xlBooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
                Microsoft.Office.Interop.Excel.Worksheet xlSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlBook.Worksheets[1];

                //让后台执行设置为不可见，为true的话会看到打开一个Excel，然后数据在往里写  
                xlApp.Visible = false;

                object[,] objData = new object[srcDataTable.Rows.Count + 1, srcDataTable.Columns.Count];
                //首先将数据写入到一个二维数组中  
                for (int i = 0; i < srcDataTable.Columns.Count; i++)
                {
                    objData[0, i] = srcDataTable.Columns[i].ColumnName;
                }
                if (srcDataTable.Rows.Count > 0)
                {
                    for (int i = 0; i < srcDataTable.Rows.Count; i++)
                    {
                        for (int j = 0; j < srcDataTable.Columns.Count; j++)
                        {
                            objData[i + 1, j] = srcDataTable.Rows[i][j];
                        }
                    }
                }

                string startCol = "A";
                int iCnt = (srcDataTable.Columns.Count / 26);
                string endColSignal = (iCnt == 0 ? "" : ((char)('A' + (iCnt - 1))).ToString());
                string endCol = endColSignal + ((char)('A' + srcDataTable.Columns.Count - iCnt * 26 - 1)).ToString();
                Microsoft.Office.Interop.Excel.Range range = xlSheet.get_Range(startCol + "1", endCol + (srcDataTable.Rows.Count - iCnt * 26 + 1).ToString());

                range.Value = objData; //给Exccel中的Range整体赋值  
                range.EntireColumn.AutoFit(); //设定Excel列宽度自适应  
                xlSheet.get_Range(startCol + "1", endCol + "1").Font.Bold = 1;//Excel文件列名 字体设定为Bold  

                //设置禁止弹出保存和覆盖的询问提示框  
                xlApp.DisplayAlerts = false;
                xlApp.AlertBeforeOverwriting = false;

                if (xlSheet != null)
                {
                    xlSheet.SaveAs(excelFilePath, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    xlApp.Quit();  
                    //KillProcess(xlApp);
                    
                }
            }
            catch (Exception ex)
            {
                xlApp.Quit(); 
                //KillProcess(xlApp);
                throw ex;
            }
        }  


    }
}
