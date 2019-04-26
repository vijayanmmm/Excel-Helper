using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Data;
//using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelUtility
{
    /// <summary>
    /// Excel帮助类
    /// <para>用法1：使用using</para> 
    /// <para>用法2：不使用using，使用结束后显式调用Dispose()</para> 
    /// </summary>
    public class ExcelHelper : IDisposable
    {
        private Application appExcel = new Application();
        private Workbooks wbs = null;
        private Workbook wb = null;
        private Worksheet ws = null;

        private string filePath = "";
        private bool visible = false;
        private bool readOnly = false;
        private bool displayAlerts = false;
        private bool isSave = false;

        #region 构造函数
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="pFilePath"></param>
        /// <param name="pVisible"></param>
        /// <param name="pReadOnly"></param>
        /// <param name="pDisplayAlerts"></param>
        public ExcelHelper(string pFilePath, bool pVisible = false, bool pReadOnly = false, bool pIsSave = false, bool pDisplayAlerts = false)
        {
            filePath = pFilePath;
            visible = pVisible;
            displayAlerts = pDisplayAlerts;
            readOnly = pReadOnly;
            isSave = pIsSave;

            InitExcel();
        }

        private void InitExcel()
        {
            appExcel.Visible = visible;
            appExcel.DisplayAlerts = displayAlerts;
            wbs = appExcel.Workbooks;
            if (File.Exists(filePath))
                wb = wbs.Open(filePath, ReadOnly: readOnly);
            else
                wb = wbs.Add();
        }
        #endregion

        #region 方法继承
        /// <summary>
        /// 实现IDisposable中的Dispose方法
        /// </summary>
        public void Dispose()
        {
            if (isSave)
            {
                Save();
            }

            if (wb != null)
            {
                wb.Close();
                Marshal.FinalReleaseComObject(wb);
            }

            if (ws != null)
                Marshal.FinalReleaseComObject(ws);

            if (wbs != null)
                Marshal.FinalReleaseComObject(wbs);

            if (appExcel != null)
            {
                appExcel.Quit();
                Marshal.FinalReleaseComObject(appExcel);
            }
        }
        #endregion

        #region 方法实现

        #region 替换方法
        /// <summary>
        /// 替换标签值
        /// </summary>
        /// <param name="pDicData"></param>
        /// <param name="pSheetName"></param>
        public void ReplaceData(Dictionary<string, string> pDicData, string pSheetName, int pRow = 0)
        {
            if (string.IsNullOrEmpty(pSheetName))
                return;

            List<string> sheetNameList = new List<string>() { pSheetName };
            ReplaceData(pDicData, sheetNameList, pRow);
        }

        /// <summary>
        /// 替换标签值
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="DicData"></param>
        /// <param name="SheetNameList"></param>
        public void ReplaceData(Dictionary<string, string> pDicData, List<string> pSheetNameList, int pRow = 0)
        {
            if (pSheetNameList == null || pSheetNameList.Count == 0)
                return;

            foreach (string sheetName in pSheetNameList)
            {
                ws = wb.Worksheets[sheetName];
                Range ReplaceArea = null;
                if (pRow == 0)
                    ReplaceArea = ws.Cells;
                else
                {
                    Range startRange = ws.Cells[1, 1];
                    Range endRange = ws.Cells[1, ws.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column];
                    ReplaceArea = ws.Range[startRange, endRange];
                    ReplaceArea.Select();
                    Marshal.FinalReleaseComObject(startRange);
                    Marshal.FinalReleaseComObject(endRange);
                }

                foreach (var data in pDicData)
                {
                    object tag = data.Key;
                    object value = data.Value;
                    ReplaceArea.Replace(tag, value, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
                Marshal.FinalReleaseComObject(ReplaceArea);
                Marshal.FinalReleaseComObject(ws);
            }
        }

        #endregion

        #region 删除方法
        /// <summary>
        /// 清空单元格内容
        /// </summary>
        /// <param name="DelDic"></param>
        public void DeleteData(string pSheetName, int pRowNo, int pCellNo)
        {
            ws = wb.Sheets[pSheetName];
            Range rngCell = ws.Cells[pRowNo, pCellNo];
            rngCell.Clear();
            Marshal.FinalReleaseComObject(rngCell);
        }

        /// <summary>
        /// 选择性清空工作表内容
        /// </summary>
        /// <param name="FilePath"></param>
        /// <param name="DelDic"></param>
        public void DeleteData(string FilePath, Dictionary<string, string> DelDic)
        {
            Range range = null;
            foreach (var Dic in DelDic)
            {
                ws = wb.Worksheets[Dic.Key];
                ws.Select();

                //清除筛选条件
                if (ws.AutoFilter != null && ws.AutoFilter.FilterMode)
                    ws.AutoFilterMode = false;

                int BeginRow = int.Parse(Dic.Value.Split(',')[0]);
                int EndColumn = int.Parse(Dic.Value.Split(',')[1]);
                Range startRange = ws.Cells[BeginRow, 1];
                Range endRange = ws.Cells[ws.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row, EndColumn];
                range = ws.Range[startRange, endRange];
                range.Select();
                range.ClearContents();

                Marshal.FinalReleaseComObject(startRange);
                Marshal.FinalReleaseComObject(endRange);
                Marshal.FinalReleaseComObject(range);
            }
        }

        /// <summary>
        /// 清空工作表内容
        /// </summary>
        /// <param name="DelDic"></param>
        public void DeleteAllData(List<string> pSheetNames)
        {
            foreach (var sheetName in pSheetNames)
            {
                ws = wb.Worksheets[sheetName];
                ws.Select();
                ws.Cells.Select();
                ws.Cells.ClearContents();
                Marshal.FinalReleaseComObject(ws);
            }
        }

        /// <summary>
        /// 删除列
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pColumns"></param>
        public void DropColumns(string pSheetName, List<int> pColumns)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            pColumns = pColumns.OrderByDescending(m => m).ToList();
            foreach (int columnNo in pColumns)
            {
                Range rngColumn = ws.Columns[columnNo];
                rngColumn.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                Marshal.FinalReleaseComObject(rngColumn);
            }

            Marshal.FinalReleaseComObject(ws);
        }

        /// <summary>
        /// 删除最后一行
        /// </summary>
        /// <param name="pSheetName"></param>
        public void DropLastRow(string pSheetName)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            int rowNum = ws.UsedRange.CurrentRegion.Rows.Count;
            Range rngRow = ws.Rows[rowNum];
            rngRow.Delete(XlDeleteShiftDirection.xlShiftUp);
            Marshal.FinalReleaseComObject(rngRow);
            Marshal.FinalReleaseComObject(ws);
        }
        /// <summary>
        /// 删除行
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pColumns"></param>
        public void DropRows(string pSheetName, List<int> pRows)
        {
            ws = wb.Sheets[pSheetName];
            ws.Select();

            pRows = pRows.OrderByDescending(m => m).ToList();
            foreach (int RowNo in pRows)
            {
                Range rngRow = ws.Rows[RowNo];
                rngRow.Delete(XlDeleteShiftDirection.xlShiftUp);
                Marshal.FinalReleaseComObject(rngRow);
            }

            Marshal.FinalReleaseComObject(ws);
        }
        #endregion

        #region 设置工作表是否可见
        /// <summary>
        /// 设置工作表是否可见
        /// </summary>
        /// <param name="pSheetName"></param>
        /// <param name="pVisible">1显示 0隐藏</param>
        public void VisibleSheet(string pSheetName, bool pVisible)
        {
            if (string.IsNullOrEmpty(pSheetName))
                return;

            ws = wb.Worksheets[pSheetName];
            ws.Visible = pVisible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
        }

        /// <summary>
        /// 设置工作表是否可见
        /// </summary>
        /// <param name="pShowSheets"></param>
        /// <param name="pVisible">1显示 0隐藏</param>
        public void VisibleSheet(List<string> pShowSheets, bool pVisible)
        {
            if (pShowSheets == null || pShowSheets.Count == 0)
                return;

            XlSheetVisibility visibility = pVisible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
            foreach (string item in pShowSheets)
            {
                ws = wb.Worksheets[item];
                ws.Visible = visibility;
            }
        }
        #endregion

        #region 保存
        private void Save()
        {
            wb.RefreshAll();
            if (File.Exists(filePath))
                wb.Save();
            else
                wb.SaveAs(filePath);
        }
        #endregion

        #endregion
    }

    #region Models
    public class DataSescribeModel
    {
        public string Name { get; set; }
        public dynamic Value { get; set; }
        public int RowNo { get; set; }
        public int ColNo { get; set; }
        public int RowEndNo { get; set; }
        public int ColEndNo { get; set; }
    }

    public class DataDTToExcelModel
    {
        public string SheetName;
        public System.Data.DataTable DataDT;
        public int BeginRow = 0;
        public int BeginCell = 0;
        public int InsertRowNo = 0;
        public int InsertRows = 0;
        public List<int> FormulaCell = null;
        public List<int> NumberFormatCell = new List<int>();
    }

    public class ColumnDataToExcel
    {
        public int ColumnNo = 0;
        public string ColumnName = "";
        public dynamic ColumnValue;
    }
    #endregion
}