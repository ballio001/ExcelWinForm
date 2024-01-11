using System;
using System.Windows.Forms;
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
            readExcel();
        }

        private void readExcel()
        {
            string filePath = "C:\\Users\\Onur\\source\\repos\\ExcelWinForm\\ExcelWinForm\\TestExcel.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //var to hold the objects
            Workbook wb;
            Worksheet ws;

            //opens workbook and stores in wb, same with sheet
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            //iterates through the given range values in Excel
            Microsoft.Office.Interop.Excel.Range cell = ws.Range["A1:B1"];
            foreach(string Result in cell.Value)
            {
                MessageBox.Show(Result);
            }
        }
        /*
        private void CmdWrite_Click(object sender, EventArgs e)
        {
            WriteExcel.writeExcel();
        }*/
    }
}
