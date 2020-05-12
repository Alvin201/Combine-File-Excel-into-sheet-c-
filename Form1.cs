using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button1.Text = "Generate";
            browse_button.Text = "Browse";
            upload_button.Text = "Upload";
            listBox1.MouseDoubleClick += new MouseEventHandler(listBox1_DoubleClick);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pathlocation = @"C:\pandora\data\export\";
            string filename = "Reconcile_Paperless_*.xlsx";
          
            foreach (string file in Directory.EnumerateFiles(pathlocation, filename))
            {

                if (File.Exists(file))
                {
                        MergeExcelNew();
                }
                else
                {
                    MessageBox.Show("file not exist", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }
            }


        }

        private static void MergeExcelNew()       
        {
            #region merged
            //    var app = new Microsoft.Office.Interop.Excel.Application();
            //    Workbook bookDest = null;
            //    Worksheet sheetDest = null;
            //    Workbook bookSource = null;
            //    Worksheet sheetSource = null;

            //    string sTempPath = @"C:\pandora\data\export\Reconcile_Paperless.xlsx";
            //    string sFinalPath = @"C:\pandora\data\export\Reconcile_Paperless_IWID.xlsx";
            //    try
            //    {
            //        //OpenBook
            //        bookDest = app.Workbooks._Open(sFinalPath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //        bookSource = app.Workbooks._Open(sTempPath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //        sheetSource = (Worksheet)bookDest.Worksheets[1];
            //        sheetDest = (Worksheet)bookSource.Worksheets[1];
            //        //CopyData
            //        sheetDest = (Worksheet)bookDest.Worksheets[1];
            //        int ss = bookDest.Sheets.Count;
            //        sheetSource.Copy(After: bookSource.Worksheets[ss]);
            //        bookDest.Close(false, Missing.Value, Missing.Value);
            //        //Save
            //        bookDest.Saved = true;
            //        bookSource.Saved = true;
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            //    }
            //    finally
            //    {
            //        app.ActiveWorkbook.Save();
            //        app.Quit();
            //        Marshal.ReleaseComObject(app);
            //        GC.Collect();
            //    }
            //}
            #endregion merged
            var app = new Excel.Application();
            app.Visible = true;
            string pathlocation = @"C:\pandora\data\export\";
            string excelfilename = "Reconcile_Paperless_";
            string CDM = "CDM.xlsx";
            string IWID = "IWID.xlsx";
            string Count = "Count.xlsx";

            Excel.Workbook w1 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless.xlsx");
            Excel.Workbook w2 = app.Workbooks.Add(pathlocation + excelfilename + CDM);
            Excel.Workbook w3 = app.Workbooks.Add(pathlocation + excelfilename + IWID);
            Excel.Workbook w4 = app.Workbooks.Add(pathlocation + excelfilename + Count);

            for (int i = 2; i <= app.Workbooks.Count; i++)
            {
                for (int j = 1; j <= app.Workbooks[i].Worksheets.Count; j++)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)app.Workbooks[i].Worksheets[j];
                    ws.Copy(app.Workbooks[1].Worksheets[1]);
                }
            }
            app.Worksheets["Sheet1"].Delete();

            string filenameresult = "Reconcile_Paperless" + DateTime.Now.ToString("dd-MMMM-yyyy HHmmss") + ".xlsx";
            app.Workbooks[1].SaveAs(pathlocation + filenameresult, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
            Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlUserResolution, true,
            Missing.Value, Missing.Value, Missing.Value);
            w1.Close(0);
            w2.Close(0);
            w3.Close(0);
            w4.Close(0);
            app.Workbooks.Close();
            app.Quit();

            Marshal.ReleaseComObject(app);
            GC.Collect();

            //delete multiple files
            var dir = new DirectoryInfo(pathlocation);
            foreach (var file in dir.EnumerateFiles("Reconcile_Paperless_*.xlsx"))
            {
                file.Delete();
            }

            MessageBox.Show("Complete", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void browse_button_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.Filter = "XLS files|*.xlsx";
            this.openFileDialog1.Title = "Please Select Excel Source File(s) for Consolidation";
            this.openFileDialog1.FilterIndex = 2;
            this.openFileDialog1.RestoreDirectory = true;
            this.openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK) // Test result.
            {
                foreach (string FileName in openFileDialog1.SafeFileNames)
                {
                    listBox1.Items.Add(System.IO.Path.GetFileNameWithoutExtension(FileName));
                }
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }


        private void upload_button_Click(object sender, EventArgs e)
        {
            int count = 0;
            string[] FilenameName;

            if (listBox1.Items.Count > 0)
            {

                DialogResult result = MessageBox.Show("Please do check again your files?", "Confirmation", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    // do what you want!!
                    foreach (string item in openFileDialog1.FileNames)
                    {
                        FilenameName = item.Split('\\');
                        File.Copy(item, @"C:\pandora\data\export\" + FilenameName[FilenameName.Length - 1]);
                        count++;
                    }
                    MessageBox.Show(Convert.ToString(count) + " File(s) copied");
                    listBox1.Items.Clear();

                }
                else if (result == DialogResult.No)
                {
                    //do what you want!!
                    MessageBox.Show("Canceled");
                }
            }
            else
            {
                // It doesn't
                MessageBox.Show("Please upload file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);


            }

        }

        private void listBox1_DoubleClick(object sender, MouseEventArgs e)
        {
            int index = this.listBox1.IndexFromPoint(e.Location);
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                //MessageBox.Show(listBox1.SelectedItem.ToString());
                listBox1.Items.Remove(listBox1.SelectedItem.ToString());
                listBox1.SelectedIndex = -1;
            }
        }


    }
}
