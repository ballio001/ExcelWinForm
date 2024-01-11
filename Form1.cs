using System;
using System.Windows.Forms;
using ExcelWinForm.Excel;
using Microsoft.Office.Interop.Excel;

namespace ExcelWinForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void CmdRead_Click(object sender, EventArgs e)
        {
            ReadExcel.readExcel();
        }

        private void CmdWrite_Click(object sender, EventArgs e)
        {
            WriteExcel.writeExcel();
        }
    }
}
