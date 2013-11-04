using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelEditor
{
    public partial class DataViewer
    {

        public static Dictionary<string, string> excelFileData = new Dictionary<string, string>();

        public static Dictionary<string, string[]> excelFileActualData = new Dictionary<string, string[]>();

        private static void ReleaseObject(Object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception e)
            {
                obj = null;
                MessageBox.Show("Unable To release the object  " + e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        internal void PlotLabelsForTitles()
        {
            int widthForWindowFactor;
            
            if (excelFileData.ContainsKey("TitlesCount"))
            {
                widthForWindowFactor = int.Parse(excelFileData["TitlesCount"]) ;
                this.Width = widthForWindowFactor * 100 + 20 * (widthForWindowFactor+2) + 10;
                this.Height = 300;

                int titleCount = 0;
                int centerAlignment = 0;

                FontFamily family = new FontFamily("Times New Roman");
                Font font = new Font(family, 13.0f, FontStyle.Bold );

                for (; titleCount < widthForWindowFactor; titleCount++)
                {
                    Label newLabel = new Label();
                    TextBox newTextBox = new TextBox();

                    string temp = "Title" + (titleCount + 1).ToString();
                    if (excelFileData.ContainsKey(temp))
                    {
                        newLabel.Name = "label" + excelFileData[temp];
                        newLabel.Text = excelFileData[temp];

                        newTextBox.Name = "textBox" + excelFileData[temp];
                        newTextBox.Text = "You can search here";                                                
                    }
                    newLabel.Location = new Point(100 * titleCount + 20 * (titleCount+1) + centerAlignment, 30);
                    newLabel.Font = font;
                    newLabel.TextAlign = System.Drawing.ContentAlignment.BottomRight;

                    newTextBox.Location = new Point(100 * titleCount + 20 * (titleCount+1), 30+30);
                    newLabel.AutoSize = true;

                    newTextBox.Width = 110;
                    
                    this.Controls.Add(newLabel);
                    this.Controls.Add(newTextBox);
                    this.MaximizeBox = false;
                    this.MinimizeBox = false;
                    this.FormBorderStyle = FormBorderStyle.FixedSingle;
                }

            }

        }

        internal void cleanUpTheDataStructures()
        {

            excelFileData.Clear();
        }

        internal void ParseDataFromExcelFile(string excelFileName)
        {
            Excel.Application excelApp;
            Excel.Workbook excelWorkbook;

            Excel.Worksheet excelWorksheet;
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            excelApp = new Excel.ApplicationClass();
            excelWorkbook = excelApp.Workbooks.Open(excelFileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1);

            range = excelWorksheet.UsedRange;

            excelFileData.Add("TitlesCount", range.Columns.Count.ToString());            

            
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {
                 str = (range.Cells[1, cCnt] as Excel.Range).Value2.ToString();
                 excelFileData.Add("Title" + cCnt.ToString(), str);                 
            }
            
            for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
            {                
                for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {                 
                    str = (range.Cells[rCnt, cCnt] as Excel.Range).Value2.ToString();
                    MessageBox.Show(str);
                }
            }

            excelWorkbook.Close(true, null, null);

            excelApp.Quit();

            ReleaseObject(excelWorksheet);
            ReleaseObject(excelWorkbook);
            ReleaseObject(excelApp);            

            

        }
    }
}
