using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelRead
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelRead read = new ExcelRead();
            var result = read.ToDataTable(@"E:\Aaronguo\test\ExcelReadWrite\ExcelRead\bin\2014淘宝双11购物清单.xls", "2014淘宝双11购物清单.xls");
             
            foreach (var item in result)
            {
                richTextBox1.AppendText(string.Format("{0},{1},{2}\n", item.proName, item.examine, item.result));
            }
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.ScrollToCaret();
        }
    }
}
