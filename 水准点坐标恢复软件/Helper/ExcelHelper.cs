using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Runtime.InteropServices;


    public enum ExcelColorIndex
    {
        无色 = -4142, 自动 = -4105, 黑色 = 1, 褐色 = 53, 橄榄 = 52, 深绿 = 51, 深青 = 49,
        深蓝 = 11, 靛蓝 = 55, 灰色80 = 56, 深红 = 9, 橙色 = 46, 深黄 = 12, 绿色 = 10,
        青色 = 14, 蓝色 = 5, 蓝灰 = 47, 灰色50 = 16, 红色 = 3, 浅橙色 = 45, 酸橙色 = 43,
        海绿 = 50, 水绿色 = 42, 浅蓝 = 41, 紫罗兰 = 13, 灰色40 = 48, 粉红 = 7,
        金色 = 44, 黄色 = 6, 鲜绿 = 4, 青绿 = 8, 天蓝 = 33, 梅红 = 54, 灰色25 = 15,
        玫瑰红 = 38, 茶色 = 40, 浅黄 = 36, 浅绿 = 35, 浅青绿 = 34, 淡蓝 = 37, 淡紫 = 39,
        白色 = 2
    }

    class ExcelHelper
    {
        private Application m_ExcelApp = null;

        /// <summary>
        /// 是否与输入版本相同或高于该版本，相同或高于返回true,否则返回false;
        /// 2003对应11,2007对应12,2010对应14 2012对应15
        /// </summary>
        /// <param name="iVersion"></param>
        /// <returns></returns>
        public bool IsRightVesrion(int version)
        {
            RegistryKey rk = Registry.LocalMachine;

            RegistryKey f = null;

            f = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\" + version + @".0\Excel\InstallRoot\");

            if (f != null)
            {
                string file03 = f.GetValue("Path").ToString();
                if (File.Exists(file03 + "EXCEL.exe"))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 是否与输入版本相同或高于该版本，相同或高于返回true,否则返回false;
        /// 2003对应11,2007对应12,2010对应14 2012对应15
        /// </summary>
        /// <param name="iVersion"></param>
        /// <returns></returns>
        static public bool IsRightVesrion(int beginVersion, int endVersion)
        {
            RegistryKey rk = Registry.LocalMachine;

            RegistryKey f = null;
            //目前
            for (int i = beginVersion; i <= endVersion; i++)
            {
                f = rk.OpenSubKey(@"SOFTWARE\Microsoft\Office\" + i + @".0\Excel\InstallRoot\");

                if (f != null)
                {
                    string file03 = f.GetValue("Path").ToString();
                    if (File.Exists(f + "Excel.exe"))
                    {
                        return true;
                    }
                }
            }

            return false;
        }


        public ExcelHelper(bool isVisible)
        {
            try
            {
                if (m_ExcelApp == null)
                {
                    m_ExcelApp = new ApplicationClass();
                    m_ExcelApp.Visible = isVisible;
                    m_ExcelApp.DisplayAlerts = false;

                }
            }
            catch (Exception excep)
            {
                throw new Exception("创建Excel应用程序出错,可能原因为：" + excep.Message);
            }
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out   int ID); 

        /// <summary>
        /// 关闭Excel应用程序
        /// </summary>
        /// <param name="isSaveChanges">是否保存修改</param>
        public void CloseExcelApplication(bool isSaveChanges)
        {
            try
            {
                object saveChanges = isSaveChanges;
                object missing = System.Type.Missing;
                if (m_ExcelApp != null)
                {
                    if (m_ExcelApp.Workbooks.Count > 0)
                    {
                        if (isSaveChanges)
                        {
                            m_ExcelApp.Workbooks[1].Save();
                        }

                        m_ExcelApp.Workbooks.Close();
                    }
                    m_ExcelApp.Quit();

                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(m_ExcelApp);
                  
                   // m_ExcelApp = null;


                }
            }
            catch (Exception excep)
            {
                throw new Exception("关闭Excel应用程序失败,可能原因为:" + excep.Message);
            }
            finally//释放资源
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                IntPtr t = new IntPtr(m_ExcelApp.Hwnd);
                int k = 0;
                GetWindowThreadProcessId(t, out k);
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);
                p.Kill();


            }
        }
        /// <summary>
        /// 打开Excel文件
        /// </summary>
        /// <param name="sFileName">EXCEL文件名</param>
        /// <returns></returns>
        public Workbook OpenExcel(string sFileName)
        {
            try
            {
                //检查文件是否存在

                if (!File.Exists(sFileName))
                {
                    throw new Exception("Excel文件不存在");
                }
                else
                {
                    FileInfo pFInfo = new FileInfo(sFileName);
                    pFInfo.IsReadOnly = false;
                }

                object missing = System.Type.Missing;

                Workbook excelBook = null;

                object readOnly = false;
                object editable = true;
                object addToRecentFiles = false;


                excelBook = m_ExcelApp.Workbooks.Open(sFileName, missing, readOnly, missing, missing, missing, missing,
                    missing, missing, editable, missing, missing, addToRecentFiles, missing, missing);
                return excelBook;

            }
            catch (Exception excep)
            {
                throw new Exception("打开Excel文档失败,可能原因为:" + excep.Message);
            }
        }
        /// <summary>
        /// 取sheet
        /// </summary>
        /// <param name="excelBook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Worksheet GetSheetByName(Workbook excelBook, string sheetName)
        {
            Worksheet reSheet = null;
            if (sheetName == null)
            {
                return (Worksheet)excelBook.Sheets[1];
            }
            for (int i = 1; i < excelBook.Sheets.Count + 1; i++)
            {
                reSheet = (Worksheet)excelBook.Sheets[i];
                if (reSheet.Name == sheetName)
                {
                    return reSheet;
                }
            }
            return null;
        }
        /// <summary>
        /// 取sheet
        /// </summary>
        /// <param name="excelBook"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public Worksheet GetDefaultSheet(Workbook excelBook)
        {
            if (excelBook.Sheets.Count > 0)
            {
                return (Worksheet)excelBook.Sheets[1];
            }
            else
            {
                return null;
            }

        }
        /// <summary>
        /// 插入行

        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        public void InsertRows(Worksheet sheet, int rowIndex)
        {
            try
            {
                object missing = System.Type.Missing;
                Range range = (Range)sheet.Rows[rowIndex, missing];
                range.Insert(XlInsertShiftDirection.xlShiftDown, missing);
            }
            catch (Exception excep)
            {
                throw new Exception("插入行失败," + excep.Message);
            }
        }
        /// <summary>
        /// 删除行

        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        public void DeleteRows(Worksheet sheet, int rowIndex)
        {
            try
            {
                object missing = System.Type.Missing;
                Range range = (Range)sheet.Rows[rowIndex, missing];
                range.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
            catch (Exception excep)
            {
                throw new Exception("删除行失败," + excep.Message);
            }
        }

        /// <summary>
        /// 关闭Excel文档
        /// </summary>
        /// <param name="wordDoc">文档</param>
        /// <param name="saveChanges">是否保存修改</param>
        public void CloseWorkBook(Workbook wBook, bool isSaveChanges)
        {
            try
            {
                object missing = System.Type.Missing;
                if (wBook != null && wBook.Application != null)
                {
                    object saveChanges = isSaveChanges;
                    wBook.Close(saveChanges, missing, missing);
                    wBook = null;
                }
            }
            catch (Exception excep)
            {
                throw new Exception("关闭Excel文档失败,可能原因为:" + excep.Message);
            }
             finally//释放资源
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        /// <summary>
        /// 打开Excel文件,并获取第一个sheet
        /// </summary>
        /// <param name="sFileName">EXCEL文件</param>
        /// <returns></returns>
        public Worksheet OpenExcelDefaultSheet(string sFileName)
        {
            try
            {
                Workbook excelBook = null;
                //检查文件是否存在

                if (!File.Exists(sFileName))
                {
                    throw new Exception("文件已存在");
                }
                else
                {
                    FileInfo pFInfo = new FileInfo(sFileName);
                    pFInfo.IsReadOnly = false;
                }
                object missing = System.Type.Missing;
                object readOnly = false;
                object editable = true;
                object addToRecentFiles = false;

                excelBook = m_ExcelApp.Workbooks.Open(sFileName, missing, readOnly, missing, missing, missing, missing,
                    missing, missing, editable, missing, missing, addToRecentFiles, missing, missing);
                Worksheet reSheet = null;
                reSheet = (Worksheet)excelBook.Sheets[1];
                return reSheet;
            }
            catch (Exception excep)
            {
                throw new Exception("打开Excel文档失败,可能原因为:" + excep.Message);
            }
        }

        public void CreateExcel(string filepath)
        {
            try
            {
                Workbook excelBook = null;
                excelBook = m_ExcelApp.Workbooks.Add(Missing.Value);//添加新工作簿

                Worksheet reSheet = null;
                reSheet = (Worksheet)m_ExcelApp.Worksheets.Add(Missing.Value);//添加新工作表

                excelBook.SaveAs(filepath,XlFileFormat.xlExcel8, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                //保存文件
                excelBook.Close(false, Missing.Value, Missing.Value);
                //关闭工作簿

                m_ExcelApp.Quit();

            }
            catch (Exception excep)
            {
                throw new Exception("打开Excel文档失败,可能原因为:" + excep.Message);
            }
        }

        public bool IsExcelExist(string FilePath)
        {
            bool isexist = false;

            return isexist;
        }

        /// <summary>
        /// 保存excel文档
        /// </summary>
        /// <param name="wBook"></param>
        public void SaveWorkBook(Workbook wBook)
        {
            try
            {
                if (wBook != null)
                {
                    wBook.Save();
                }
            }
            catch (Exception excep)
            {
                throw new Exception("保存Excel文档失败,可能原因为:" + excep.Message);
            }
            finally//释放资源
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        /// <summary>
        /// 保存excel文档
        /// </summary>
        /// <param name="wBook"></param>
        public void SaveWorkBookAs(Workbook wBook, string sFileName)
        {
            try
            {
                object missing = System.Type.Missing;
                if (wBook != null && wBook.Application != null)
                {
                    wBook.SaveAs(sFileName, missing, missing, missing, missing, missing, XlSaveAsAccessMode.xlNoChange, missing,
                        missing, missing, missing, missing);
                }
            }
            catch (Exception excep)
            {
                throw new Exception("保存Excel文档失败,可能原因为:" + excep.Message);
            }
            finally//释放资源
            {
               

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        /// <summary>
        /// 在指定的行、列格子中填充字符

        /// </summary>
        /// <param name="wSheet">excel worksheet</param>
        /// <param name="rows">所在行，从1开始</param>
        /// <param name="cols"></param>
        /// <param name="fillText"></param>
        public void FillCellText(Worksheet wSheet, int rows, int cols, string fillText)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                wSheet.Cells[rows, cols] = fillText;

            }
            catch (Exception excep)
            {
                throw new Exception("填充单元格失败,可能原因为:" + excep.Message);
            }
        }

        /// <summary>
        /// 获取指定的行、列格子的值。

        /// </summary>
        /// <param name="wSheet">excel worksheet</param>
        /// <param name="rows">所在行，从1开始</param>
        /// <param name="cols"></param>
        /// <param name="fillText"></param>
        public string GetCellText(Worksheet wSheet, int rows, int cols)
        {
            try
            {
                string text = "";

                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                Range eRange = (Range)wSheet.Cells.get_Item(rows, cols);

                if (eRange.Value != null)
                {
                    text = eRange.Value.ToString();
                }

                //text = ((Excel.Range)wSheet.Cells[rows, cols]).Value2.ToString();
                //Excel.Range eRange = wSheet.get_Range(wSheet.Cells[rows, cols], wSheet.Cells[rows, cols]);
                //text = eRange.Value2.ToString();

                return text;
            }
            catch (Exception excep)
            {
                throw new Exception("获取单元格内容失败,可能原因为:" + excep.Message);
            }

            return "";
        }
        /// <summary>
        /// 设置某行背景色

        /// </summary>
        /// <param name="wSheet">工作表</param>
        /// <param name="rows">行数</param>
        /// <param name="rColor">颜色</param>
        public void SetRowBackgroundColor(Worksheet wSheet, int rows, ExcelColorIndex rColor)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                Range eRange = ((Range)wSheet.Rows[rows.ToString() + ":" + rows.ToString(), System.Type.Missing]);
                eRange.Interior.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定行背景色失败,可能原因为:" + excep.Message);
            }
        }

        public void SetColumnBackgroundColor(Worksheet wSheet, int iColumn, ExcelColorIndex rColor)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                int iRow = GetRowsCount(wSheet);

                Range eRange = ((Range)wSheet.get_Range(wSheet.Cells[1, iColumn], wSheet.Cells[iRow, iColumn]));
                eRange.Interior.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定行背景色失败,可能原因为:" + excep.Message);
            }
        }

        public void SetColumnFontColor(Worksheet wSheet, int iColumn, ExcelColorIndex rColor)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                int iRow = GetRowsCount(wSheet);

                Range eRange = ((Range)wSheet.get_Range(wSheet.Cells[1, iColumn], wSheet.Cells[iRow, iColumn]));
                eRange.Font.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定行背景色失败,可能原因为:" + excep.Message);
            }
        }

        public void SetRowFontColor(Worksheet wSheet, int iRow, ExcelColorIndex rColor)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                int iCol = GetColumnsCount(wSheet);

                Range eRange = ((Range)wSheet.get_Range(wSheet.Cells[iRow, 1], wSheet.Cells[iRow, iCol]));
                eRange.Font.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定行背景色失败,可能原因为:" + excep.Message);
            }
        }


        public void MergeCells(Worksheet wSheet, int sRow, int sCol, int dRow, int dCol)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                int iCol = GetColumnsCount(wSheet);

                Range excelRange = wSheet.get_Range(wSheet.Cells[sRow, sCol], wSheet.Cells[dRow, dCol]);
                excelRange.Merge();
            }
            catch (Exception excep)
            {
                throw new Exception("合并单元格失败,可能原因为:" + excep.Message);
            }

        }
        public void SetRowHeight(Worksheet wSheet, int sRow, int height)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }

                int iCol = GetColumnsCount(wSheet);

                Range eRange = ((Range)wSheet.get_Range(wSheet.Cells[sRow, 1], wSheet.Cells[sRow, iCol]));
                eRange.RowHeight = (double)height;
            }
            catch (Exception excep)
            {
                throw new Exception("设置行高失败,可能原因为:" + excep.Message);
            }
        }
        /// <summary>
        /// 指定单元格的背景色

        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <param name="rColor"></param>
        public void SetCellBackgroundColor(Worksheet wSheet, int rows, int cols, ExcelColorIndex rColor)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                //Excel.Range eRange = wSheet.get_Range(wSheet.Cells[rows, cols], wSheet.Cells[rows, cols]);
                Range eRange = (Range)wSheet.Cells.get_Item(rows, cols);
                eRange.Interior.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定单元格背景色失败,可能原因为:" + excep.Message);
            }
        }
        /// <summary>
        /// 设定单元格字体颜色

        /// </summary>
        /// <param name="wSheet"></param>
        /// <param name="rows"></param>
        /// <param name="cols"></param>
        /// <param name="rColor"></param>
        public void SetCellFontColor(Worksheet wSheet, int rows, int cols, ExcelColorIndex rColor)//会修改上一行及指定行？
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                Range eRange = wSheet.get_Range(wSheet.Cells[rows, cols], wSheet.Cells[rows, cols]);
                eRange.Font.ColorIndex = rColor;
            }
            catch (Exception excep)
            {
                throw new Exception("设置指定单元格字体颜色失败,可能原因为:" + excep.Message);
            }
        }


        public void SetCellLineStyle(Worksheet wSheet, int rows, int cols)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                Range eRange = wSheet.get_Range(wSheet.Cells[rows, cols], wSheet.Cells[rows, cols]);

                //单元格边框线类型(线型,虚线型) 
                eRange.Borders.LineStyle = 1;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin; ;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeTop).ColorIndex = 1;
                //指定单元格下边框线粗细,和色彩

                eRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeBottom).ColorIndex = 1;

                eRange.Borders.get_Item(XlBordersIndex.xlEdgeLeft).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeLeft).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeLeft).ColorIndex = 1;

                eRange.Borders.get_Item(XlBordersIndex.xlEdgeRight).LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeRight).Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                eRange.Borders.get_Item(XlBordersIndex.xlEdgeRight).ColorIndex = 1;

            }
            catch (Exception excep)
            {
                throw new Exception("设置指定单元格字体颜色失败,可能原因为:" + excep.Message);
            }
        }

        public void SetCellFont(Worksheet wSheet, int rows, int cols, int size)
        {
            try
            {
                if (wSheet == null)
                {
                    throw new Exception("工作表为空");
                }
                Range eRange = wSheet.get_Range(wSheet.Cells[rows, cols], wSheet.Cells[rows, cols]);
                eRange.Font.Size = size;
                eRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;//居中
                eRange.EntireColumn.AutoFit(); //自适应

            }
            catch (Exception excep)
            {
                throw new Exception("设置指定单元格字体颜色失败,可能原因为:" + excep.Message);
            }
        }
        /// <summary>
        /// 获取记录条数
        /// </summary>
        /// <param name="wSheet"></param>
        /// <returns></returns>
        public int GetRowsCount(Worksheet wSheet)
        {
            return wSheet.UsedRange.CurrentRegion.Rows.Count;
        }

        /// <summary>
        /// 获取列数
        /// </summary>
        /// <param name="wSheet"></param>
        /// <returns></returns>
        public int GetColumnsCount(Worksheet wSheet)
        {
            return wSheet.UsedRange.CurrentRegion.Columns.Count;
        }

        /// <summary>
        /// 获取列数
        /// </summary>
        /// <param name="wSheet"></param>
        /// <returns></returns>
        public int GetColumnIndexByName(Worksheet wSheet, string columnName)
        {
            int iColCount = GetColumnsCount(wSheet);

            for (int i = 1; i <= iColCount; i++)
            {
                if (columnName == GetCellText(wSheet, 1, i))
                {
                    return i;
                }
            }

            return -1;
        }

        public string[] GetColumnAllValue(Worksheet wSheet, string columnName)
        {
            int iColIndex = GetColumnIndexByName(wSheet, columnName);
            int iRowCount = GetRowsCount(wSheet);

            string[] Values = new string[iRowCount - 1];

            for (int i = 2; i <= iRowCount; i++)
            {
                Values[i - 2] = GetCellText(wSheet, i, iColIndex);
            }

            return Values;
        }

        #region - 由数字转换为Excel中的列字母 -

        public static int ToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+")) { throw new Exception("invalid parameter"); }

            int index = 0;
            char[] chars = columnName.ToUpper().ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                index += ((int)chars[i] - (int)'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1);
            }
            return index - 1;
        }


        public static string ToName(int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + (int)'A')).ToString());
                index = (int)((index - index % 26) / 26);
            } while (index > 0);

            return String.Join(string.Empty, chars.ToArray());
        }
        #endregion


        /// <summary>  
        /// 导入文件的具体方法  
        /// </summary>  
        /// <param name="file">要导入的文件</param>  
        /// <returns></returns>  
        public static System.Data.DataTable ImportExcel(string file)
        {
            FileInfo fileInfo = new FileInfo(file);
            if (!fileInfo.Exists) return null;
            string strConn = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + file + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'";
            //if (!fileInfo.Exists) return null; 
            //string strConn = @"Provider=Microsoft.Jet.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'";
            OleDbConnection objConn = new OleDbConnection(strConn);
            System.Data.DataTable dsExcel = new System.Data.DataTable();
            try
            {
                objConn.Open();
                //string strSql = "select * from [Sheet1$]";
                //第一个sheet，不涉及表名称

                string strSql = "select * from [Sheet1$]";
                OleDbDataAdapter odbcExcelDataAdapter = new OleDbDataAdapter(strSql, objConn);

                odbcExcelDataAdapter.Fill(dsExcel);

                //重新整理数据表

                DataRow firstRow = dsExcel.Rows[0];
                for (int i = 0; i < dsExcel.Columns.Count; i++)
                {
                    dsExcel.Columns[i].ColumnName = firstRow[i].ToString();
                }
                dsExcel.Rows.RemoveAt(0);

                return dsExcel;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


    

    }

