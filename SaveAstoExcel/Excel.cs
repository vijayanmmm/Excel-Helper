using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelUtility
{
    public class Excel
    {
        public string saveAsExcel_FromActiveMHTFormat(string strFolderPath, string strWorkBookName)
        {
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;

            //validate whether the active book path is correct or not
            if (oWB.Name != strWorkBookName + ".MHTML")
            {
                //Loop throught the workbooks and find out is it there or not, if not then reply as error
                foreach(var wb in oXL.Workbooks)
                {
                    if (oWB.Name != strWorkBookName + ".MHTML")
                    {
                        blnWBFound = true;
                        oWB = (Microsoft.Office.Interop.Excel.Workbook)wb;
                        break;
                    }
                }

                if (blnWBFound == false)
                {
                    return "Excel workbook" +  strWorkBookName + " is not found";
                }
            }

            //Delete if the file is already exist
            if (System.IO.File.Exists(strFolderPath + "\\" + strWorkBookName + ".xlsx") == true)
            {
                System.IO.File.Delete(strFolderPath + "\\" + strWorkBookName + ".xlsx");
            }                
            oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault);
            oWB.Close();
            return "Success";
        }        
        public string Excel_To_CSV_Conversion(string strFolderPath, string strWorkBookName, string strSheetName) {
            //Gets Excel and gets Activeworkbook and worksheet
            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel.Workbook oWB;
            oXL = (Microsoft.Office.Interop.Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Microsoft.Office.Interop.Excel.Workbook)oXL.ActiveWorkbook;
            bool blnWBFound = false;

            //validate whether the active book path is correct or not
            if (oWB.Name != strWorkBookName + ".xlsx") {
                //Loop throught the workbooks and find out is it there or not, if not then reply as error
                foreach (var wb in oXL.Workbooks) {
                    if (oWB.Name == strWorkBookName + ".xlsx") {
                        blnWBFound = true;
                        oWB = (Microsoft.Office.Interop.Excel.Workbook)wb;
                        break;
                    }
                }

                if (blnWBFound == false) {
                    //Open the new excel
                    //oWB = oXL.WorkbookOpen()
                    return "Excel workbook" + strWorkBookName + " is not found";
                }
            }

            //Delete if the file is already exist
            if (System.IO.File.Exists(strFolderPath + "\\" + strWorkBookName + ".csv") == true) {
                System.IO.File.Delete(strFolderPath + "\\" + strWorkBookName + ".csv");
            }

            //Save with the mentioned sheet
            bool blnSaved = false;

            //If it has only one sheet, then no need to check
            if (oWB.Worksheets.Count == 1) {
                oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange);
                blnSaved = true;
            }
            else {
                foreach (Microsoft.Office.Interop.Excel.Worksheet sht in oWB.Worksheets) {
                    if (sht.Name.Trim().ToLower() == strSheetName.Trim().ToLower()) {
                        //Check whether this sheet is alreayd selected or not, if not then select it
                        if (oWB.ActiveSheet.name != sht.Name)
                                sht.Select();
                        oWB.SaveAs(strFolderPath + "\\" + strWorkBookName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,);
                        blnSaved = true;
                    }
                }
            }      
            
            if (blnSaved == false) {
                return "Excel Sheet with the name:" + strSheetName + " is not found";
            }
            else {
                oWB.Close(false);
                return "Success";
            }
        }


    }
}
