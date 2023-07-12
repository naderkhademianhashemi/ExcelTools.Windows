using ExcelTools.Windows.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTools.Windows.FRM
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            var rows = new List<Dictionary<string, object>>();
            rows.Add(new Dictionary<string, object>() { { "1", "1" } });
            ExportExcel.SaveToFile(rows, textBox1.Text);
            MessageBox.Show("success");
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            var svFlDlg = new SaveFileDialog();

            if (svFlDlg.ShowDialog() == DialogResult.OK)
                textBox1.Text = svFlDlg.FileName;
        }
    }
}
