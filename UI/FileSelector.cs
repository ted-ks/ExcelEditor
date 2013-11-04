using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelEditor
{
    public partial class FileSelector : Form
    {
    
        public FileSelector()
        {
            InitializeComponent();
        }

        private void openExcelFile_HelpRequest(object sender, EventArgs e)
        {

        }

        private void openExcelFile_FileOk(object sender, CancelEventArgs e)
        {
            
            
        }

        private void openFileButton_Click(object sender, EventArgs e)
        {
            int size = -1;

            DialogResult result = openExcelFile.ShowDialog();

            if (DialogResult.OK == result)
            {
                string file = openExcelFile.FileName;
                try
                {
                    string text = File.ReadAllText(file);
                    size = text.Length;

                    if (ValidateOpenedFile(file) == false)
                        MessageBox.Show("Currently one XL files are supported");
                    else
                    {
                        this.Hide();
                        var DataViewerObject = new DataViewer();
                        DataViewerObject.ParentWindow = this;
                        DataViewerObject.ExcelFileName = file;
                        DataViewerObject.Show(); 
                    }
                    
                }
                catch(IOException)
                {
                }                
                Console.WriteLine(size.ToString());
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        
    }
}
