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
            objsave.saveAsExcel_FromActiveMHTFormat("D:\\TaxRawData\\201808", "expor1t");
        }

        private void button2_Click(object sender, EventArgs e) {
            Excel objsave = new Excel();
            objsave.Excel_To_CSV_Conversion("\\\\ap1chndh111\\Data Center\\.Net Projects Files", "CN56税费计算表-201805", "附表（二）");
        }
    }
}
