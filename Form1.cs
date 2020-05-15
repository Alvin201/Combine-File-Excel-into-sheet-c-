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
           
            listBox2.MouseDoubleClick += new MouseEventHandler(listBox2_DoubleClick);
            listBox1.MouseDoubleClick += new MouseEventHandler(listBox1_DoubleClick);
        }

        private static void delete_multiple_file()
        {
            string pathlocation = @"C:\pandora\data\export\source\temp\";
            var dir = new DirectoryInfo(pathlocation);
            foreach (var file in dir.EnumerateFiles("*.xlsx"))
            {
                file.Delete();
            }
        }

        private static void delete_multiple_file_reconcille()
        {
            string pathlocation = @"C:\pandora\data\export\source\";
            var dir = new DirectoryInfo(pathlocation);
            foreach (var file in dir.EnumerateFiles("Reconcile_Paperless_*.xlsx"))
            {
                file.Delete();
            }
        }

        private void generate_button_Click(object sender, EventArgs e)
        {

            string pathlocation = @"C:\pandora\data\export\source\";
            string filename = "Reconcile_Paperless_*.xlsx";

            if (listBox2.Items.Count > 0)
            {
                foreach (string file in Directory.EnumerateFiles(pathlocation, filename))
                {
                    if (File.Exists(file))
                    {
                        MergeExcelNew();
                        listBox2.Items.Clear();
                    }
                    else
                    {
                        MessageBox.Show("file not exist", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            else
            {
                MessageBox.Show("Please upload file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void List_1()
        {
            DirectoryInfo dinfo = new DirectoryInfo(@"C:\\pandora\\data\\export\\source\\temp\\");
            FileInfo[] Files = dinfo.GetFiles("*.xlsx");
            foreach (FileInfo file in Files)
            {
                listBox1.Items.Add(file.FullName);
            }
        }

        private void List_2()
        {
            DirectoryInfo dinfo = new DirectoryInfo(@"C:\\pandora\\data\\export\\source\\");
            FileInfo[] Files = dinfo.GetFiles("*.xlsx");
            foreach (FileInfo file in Files)
            {
                listBox2.Items.Add(file.FullName); //note FullName, not Name
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            generate_button.Text = "Generate";
            browse_button.Text = "Browse";
            upload_button.Text = "Upload";

            List_1();
            List_2();
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

            //var excelnames = new string[] {"*.xlsx"};
            var app = new Excel.Application();
            
            app.Visible = true;
            string pathlocation = @"C:\pandora\data\export\source\";
            string resultpathlocation = @"C:\pandora\data\export\";
            string excelfilename = "Reconcile_Paperless_";
            string CDM = "CDM.xlsx";
            string IWID = "IWID.xlsx";
            string Count = "Count.xlsx";

            Excel.Workbook w1 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless.xlsx");
            Excel.Workbook w2 = app.Workbooks.Add(pathlocation + excelfilename + CDM);
            Excel.Workbook w3 = app.Workbooks.Add(pathlocation + excelfilename + IWID);
            Excel.Workbook w4 = app.Workbooks.Add(pathlocation + excelfilename + Count);


            //string[] ext = new string[2] { "*.xlsx", "*.xls" };

            //foreach (string found in ext)
            //{
            //    string[] extracted = Directory.GetFiles(@"C:\pandora\data\export\source", found, System.IO.SearchOption.AllDirectories);
            //    foreach (string file in extracted)
            //    {
            //        Console.WriteLine(file);
            //    }
            //}



            int wb_count = app.Workbooks.Count;
                    for (int i = 2; i <= wb_count; i++)
                    {
                        for (int j = 1; j <= app.Workbooks[i].Worksheets.Count; j++)
                        {
                            Excel.Worksheet ws = (Excel.Worksheet)app.Workbooks[i].Worksheets[j];
                            ws.Copy(Before: app.Workbooks[1].Worksheets[1]);
                        }
                    }
            app.Worksheets["Sheet1"].Delete();
            string filenameresult = "Reconcile_Paperless" + DateTime.Now.ToString("dd-MMMM-yyyy HHmmss") + ".xlsx";
            app.Workbooks[1].SaveAs(resultpathlocation + filenameresult, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
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

            //delete multiple files contains reconcille
            delete_multiple_file_reconcille();

            MessageBox.Show("Complete", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
        }

        private void browse_button_Click(object sender, EventArgs e)
        {
            
            string pathlocation = @"C:\pandora\data\export\source\temp\";
            this.openFileDialog1.Filter = "XLS files|*.xlsx";
            this.openFileDialog1.Title = "Please Select Excel Source File(s) for Consolidation";
            this.openFileDialog1.FilterIndex = 2;
            this.openFileDialog1.RestoreDirectory = true;
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.ValidateNames = false;
            this.openFileDialog1.CheckFileExists = false;
            this.openFileDialog1.CheckPathExists = true;

            this.openFileDialog1.FileOk += delegate (object s, CancelEventArgs ev) {
                var size = new FileInfo(this.openFileDialog1.FileName).Length;
                if (size > 250000)
                {
                    MessageBox.Show("Sorry, file is too large");
                    ev.Cancel = true;             // <== here
                }
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK) // Test result.
            {
                //string[] FilenameName;
                int count = 0;
                string name = "Reconcile_Paperless_";
                string ext = ".xlsx";
                string[] myFiles = Directory.GetFiles(pathlocation);

                foreach (string item in openFileDialog1.FileNames)
                {
                    string[] splitName = item.Split('\\');
                    string fileName = splitName[splitName.Length - 1];
                    
                    //validate filename
                    if (fileName.StartsWith(name) && fileName.EndsWith(ext))
                    {
                        listBox1.Items.Add(System.IO.Path.Combine(pathlocation, splitName[splitName.Length - 1]));
                        File.Copy(item, pathlocation + splitName[splitName.Length - 1]);
                        count++;
                        #region
                        //foreach (string item in openFileDialog1.FileNames)
                        //{
                        //    FilenameName = item.Split('\\');
                        //    File.Copy(item, extractPath + FilenameName[FilenameName.Length - 1]);
                        //    listBox1.Items.Add(System.IO.Path.Combine(extractPath, FilenameName[FilenameName.Length - 1]));
                        //    count++;
                        //}
                        #endregion
                    }
                    else
                    {
                        MessageBox.Show("Filename must be {" + name + "*.xlsx}", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                }
            }
        }


        private void upload_button_Click(object sender, EventArgs e)
        {
            //int count = 0;
            //string[] FilenameName;
            string pathlocation = @"C:\pandora\data\export\source\";

            if (listBox1.Items.Count > 0)
            {
                DialogResult result = MessageBox.Show("Please do check again your files?", "Confirmation", MessageBoxButtons.YesNoCancel);
                if (result == DialogResult.Yes)
                {
                    // check list value loop listbox
                    foreach (string item in listBox1.Items) //openFileDialog1.FileNames)
                    {
                        string[] splitName = item.Split('\\');
                        string fileName = splitName[splitName.Length - 1];

                        File.Copy(item, pathlocation + splitName[splitName.Length - 1]);
                        listBox2.Items.Add(System.IO.Path.Combine(pathlocation, splitName[splitName.Length - 1]));
                        //count++;
                    }
                    MessageBox.Show(Convert.ToString(listBox1.Items.Count) + " File(s) copied");
                    #region move list2 with selected
                    //for (int intCount = listBox1.SelectedItems.Count - 1; intCount >= 0; intCount--)
                    //{
                    //    listBox2.Items.Add(listBox1.SelectedItems[intCount]);
                    //    listBox1.Items.Remove(listBox1.SelectedItems[intCount]);
                    //}
                    #endregion

                    ////move list2 all
                    #region
                    //DirectoryInfo dinfo = new DirectoryInfo(pathlocation);
                    //FileInfo[] Files = dinfo.GetFiles("*.xlsx");

                    //foreach (FileInfo file in Files)
                    //{
                    //}
                    #endregion
                    listBox1.Items.Clear();
                    delete_multiple_file();

                    #region
                    //for (int i = 0; i < listBox1.Items.Count; i++)
                    //{
                    //    listBox2.Items.Add(listBox1.Items[i].ToString());
                    //    listBox1.Items.Clear();
                    //}
                    #endregion
                }
                else if (result == DialogResult.No)
                {
                    MessageBox.Show("Canceled");
                }
            }
            else
            {
                MessageBox.Show("Please upload file", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

        }

        
        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex != -1)
            {
                string filepath = listBox2.Items[listBox2.SelectedIndex].ToString();

                if (File.Exists(filepath))
                    File.Delete(filepath);
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
            }
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                string filepath = listBox1.Items[listBox1.SelectedIndex].ToString();

                if (File.Exists(filepath))
                    File.Delete(filepath);
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
        }
    }
}
