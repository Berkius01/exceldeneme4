using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using excel = Microsoft.Office.Interop.Excel;
namespace exceldeneme4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        int satır = 1;
        int sutun = 1;
        string[] veriler;
        bool degisken = true;
        private void button1_Click(object sender, EventArgs e)
        {
            
            veriler = new string[3];
            excel.Application uygulama = new excel.Application();
            uygulama.Visible = true;
            object Missing = Type.Missing;
            excel.Workbook kitap = uygulama.Workbooks.Add(System.Reflection.Missing.Value);
            excel.Worksheet sayfa1 = (excel.Worksheet)kitap.Sheets[1];
            veriler[0] = textBox1.Text;
            veriler[1] = textBox2.Text;
            veriler[2] = textBox3.Text;
            for (int i = 0; i < 3; i++)
            {
                excel.Range alan = (excel.Range)sayfa1.Cells[satır,sutun+i];
                alan.Value2 = dataGridView1.Columns[i].HeaderText;
            }
            //satır++;
            for(int j = 0; j < 3; j++)
            {
                excel.Range range2 = (excel.Range)sayfa1.Cells[satır + 1, sutun + j];
                range2.Value2 = veriler[j];
            }
            //tablo(degisken, veriler);
            //degisken = false;
        }
        /*private void tablo(bool degisken1, string[] veriler2) 
        {
            if (degisken1)
            {

            }
        }*/
    }
}
