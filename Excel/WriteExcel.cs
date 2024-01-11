using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelWinForm.Excel
{
    class WriteExcel
    {
        public static void writeExcel()
        {
            //filepath to the location of the Excel
            string filePath = "C:\\Users\\Onur\\source\\repos\\ExcelWinForm\\ExcelWinForm\\TestExcel.csv";
            string filePathEdited = "C:\\Users\\Onur\\source\\repos\\ExcelWinForm\\ExcelWinForm\\TestExcel2.csv";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //var to hold the objects
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            //write in array
            Range cellRange = ws.Range["C1:C4"];
            string [] Landmarks = new [] {"Vondel Park","Eifel Tower", "Mona lisa", "Cambodia"};

            cellRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, Landmarks);

            wb.SaveAs(filePathEdited);
            wb.Close();

            MessageBox.Show("File are been written to:  " + filePathEdited);
            // unocmment if you want to open excel after executing write
            //Process.Start(filePathEdited);
        }
    }
}
