using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelEditor
{
    public partial class DataViewer : Form
    {
        public FileSelector ParentWindow;
        public string ExcelFileName;

        public DataViewer()
        {
            InitializeComponent();            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void DataViewer_FormClosed(object sender, EventArgs e)
        {
            cleanUpTheDataStructures();
            this.ParentWindow.Show();            
        }

        private void DataViewer_Load(object sender, EventArgs e)
        {
            ParseDataFromExcelFile(this.ExcelFileName);

            PlotLabelsForTitles();

        }        
        
    }
}
