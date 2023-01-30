using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PlanLEFileLoadHelper;
namespace PlanLEFileCheck
{
    public partial class Form1 : Form
    {
        ProcessExcel plnleHelper = new ProcessExcel();
        string sMinRowCount = ConfigurationManager.AppSettings["MinTableRowCount"].ToString();
        string sFilePattern = "Ecom Daily Plan*.xlsx";
        public Form1()
        {
            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
            openFileDialog1.Filter = "Plan or LE File(" + sFilePattern + ")|*.xlsx";
            openFileDialog1.FileName = "Ecom Daily Plan*.xlsx";
            openFileDialog1.ShowDialog();

            if (openFileDialog1.FileName !="")
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           

            label4.Text = "File Should have columns " + plnleHelper.HeaderString;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int minRowCount = 1000;
            if (sMinRowCount != "")
            {
                minRowCount = int.Parse(sMinRowCount);
            }
            if (plnleHelper.Validate(textBox1.Text, minRowCount))
            {
                MessageBox.Show("File looks good");
            }
        }
    }
}
