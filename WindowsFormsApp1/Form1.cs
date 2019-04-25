using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelUtility;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel objsave = new Excel();
            objsave.ExportToCSV("D:\\1.xlsx", "D:\\1.csv");
            //objsave.Excel_DropLastRow(@"C:\Users\abipcadmin\Desktop\CNA8_AR01_2019_3.xlsx", "Sheet1");
        }

        private void button2_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.Excel_To_CSV_Conversion("\\\\ap1chndh111\\Data Center\\.Net Projects Files", "CN56税费计算表-201805", "附表（二）");
        }

        private void button3_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.CSV_To_Excel_Converstion("D:\\Users\\28066351\\Documents\\Projects\\Book1.csv");
        }

        private void button4_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            //objsave.Excel_Delete_BlankColumns("D:\\Users\\28066351\\Documents\\tax1111\\CN93_AR01_201810.xls", "资产");
            objsave.Excel_Delete_BlankColumns("D:\\taxRawData\\CN33_AR01_2018_12.xls", "资产","Yes");
        }

        private void btnGetHeaderClmn_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            //objsave.Excel_Delete_BlankColumns("D:\\Users\\28066351\\Documents\\tax1111\\CN93_AR01_201810.xls", "资产");
           MessageBox.Show( objsave.Excel_GetHeader_Index("CN62_AR01_2019_1.xls", "资产","1","1"));
        }

        private void button5_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            //objsave.Excel_RunMacro("COGNOS_FA_Flow.xlsm","FileMerge");//, "Macro1", "", "");
            //objsave.Excel_RunMacro("TstMacro.xlsm", "Macro1");//, "Macro1", "", "");
            objsave.Excel_RunMacro("COGNOS_FA_Flow.xlsm", "FileMerge");//, "Macro1", "", "");
        }

        private void btnCreateNewExcel_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.Excel_CreateNewExcel("D:\\Users\\28066351\\Documents\\Testing\\TestExcel.xlsx");
        }

        private void btnDeleteRow_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.Excel_Delete_Row("D:\\Users\\28066351\\Documents\\Testing\\家乐福数据.xlsx", "2", "2");
        }

        private void button6_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.Excel_Copy_Data("D:\\Users\\28066351\\Documents\\Testing\\FS_CN16.xlsx", "D:\\Users\\28066351\\Documents\\Testing\\FS Report and Recon_AC1901_CN16.xlsx", "1", "4","A1:G500","A1"); ;
        }

        private void button7_Click(object sender, EventArgs e) {

            Excel objsave = new Excel();
            objsave.Htm_To_Excel_Converstion("D:\\TaxRawData\\Rachel\\ABC.htm"); 
        }

        private void button8_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            //objsave.Excel_Delete_BlankColumns("D:\\Users\\28066351\\Documents\\tax1111\\CN93_AR01_201810.xls", "资产");
            objsave.Excel_Delete_DuplicateHeadings(@"D:\Users\28066351\Documents\Testing\TestDeleteHeadings1.xlsx", "公司代码");
        }
    }
}
