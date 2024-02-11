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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void CmdPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fileDialog.FileName;
                //RichTextBox1.Text = Path.GetFileName(fileDialog.FileName);
            }
        }
    }
}
