using System;
using System.Windows.Forms;
using System.Threading;

namespace ExcelWinForm
{
    public class Program
    {
        [STAThread]
        public static void Main()
        {
            Application.Run(new Form1());
        }
    }
}
