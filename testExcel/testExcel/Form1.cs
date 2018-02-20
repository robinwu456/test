using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace testExcel
{
    public partial class Form1 : Form
    {
        string fileTest = "C:\\Users\\Robin-PC\\Desktop\\test5.xlsx";

        public Form1()
        {
            InitializeComponent();
        }

        //creat excel
        private void button1_Click(object sender, EventArgs e)
        {
            if (File.Exists(fileTest))
            {
                File.Delete(fileTest);
            }

            Excel.Application oApp = new Excel.Application();            
            Excel.Workbook oBook = oApp.Workbooks.Add();
            Excel.Worksheet oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);

            oSheet.Cells[1, 1] = "aaa";
            
            oBook.SaveAs(fileTest);
            oBook.Close();
            oApp.Quit();
        }

        //write import
        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(fileTest);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;

            int countRecords = userRange.Rows.Count;
            int add = countRecords + 1;
            x.Cells[add, 1] = "Total Rows" + countRecords;

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

        //read text
        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(fileTest);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;

            label1.Text = x.UsedRange.Cells[2, 1].value;

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(fileTest);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;

            Console.WriteLine(userRange.Rows.Count);
            Console.WriteLine(userRange.Columns.Count);

            for (int i = 1; i <= userRange.Rows.Count; i++)
            {
                int index = 0;

                for(int j = 1; j <= userRange.Columns.Count; j++)
                {
                    index += (int)userRange.Cells[i, j].value;

                    if (j == userRange.Columns.Count)
                    {
                        x.Cells[i, userRange.Columns.Count + 1] = index;
                    }
                }
            }

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
