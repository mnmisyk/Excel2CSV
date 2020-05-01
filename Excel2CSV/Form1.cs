using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;

namespace Excel2CSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
           string filePath = @"D:\Projects\Excel2CSV\excel2csv.xls";
            Workbook wb = new Workbook(filePath);            
            Cells cells = wb.Worksheets[0].Cells;
            int row = 0, column = 0;
            while (cells[0, column].StringValue!="")
            {
                column++;
            }
            while (cells[row, 0].StringValue!="")
            {
                row++;
            }
            StreamWriter sw = new StreamWriter(@"D:\Projects\Excel2CSV\output.csv");
            for (int i = 0; i < row; i++)
            {
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j < column; j++)
                {
                    if (j == column - 1)
                    {
                        sb = sb.Append(cells[i, j].StringValue );
                    }
                    else
                    {
                        sb = sb.Append(cells[i, j].StringValue + ";");
                    }
                   
                }
                
                sw.WriteLine(sb);
            }
            sw.Close(); 
 

        }
    }
}
