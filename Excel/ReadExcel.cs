using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelWinForm.Excel
{
    class ReadExcel
    {
        public static void readExcel()
        {
            string filePath = Files.OriginalFilePath;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            //var to hold the objects
            Workbook wb;
            Worksheet ws;

            //opens workbook and stores in wb, same with sheet
            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            //iterates through the given range values in Excel
            Microsoft.Office.Interop.Excel.Range cell = ws.Range["A1:B1"];
            foreach (string Result in cell.Value)
            {
                MessageBox.Show(Result);
            }
        }
    }
}
