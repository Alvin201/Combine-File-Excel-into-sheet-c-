using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MergeExcelNew();
        }

        private static void MergeExcelNew()
        {
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
            var app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook w1 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless.xlsx");
            Excel.Workbook w2 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless_CDM.xlsx");
            Excel.Workbook w3 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless_IWID.xlsx");
            Excel.Workbook w4 = app.Workbooks.Add(@"C:\pandora\data\export\Reconcile_Paperless_Count.xlsx");
            for (int i = 2; i <= app.Workbooks.Count; i++)
            {
                for (int j = 1; j <= app.Workbooks[i].Worksheets.Count; j++)
                {
                    Excel.Worksheet ws = (Excel.Worksheet)app.Workbooks[i].Worksheets[j];
                    ws.Copy(app.Workbooks[1].Worksheets[1]);
                }
            }
            app.Worksheets["Sheet1"].Delete();
            string filename = "Reconcile_Paperless_" + DateTime.Now.ToString("ddMMyyyy") + ".xlsx";

            app.Workbooks[1].SaveAs(@"C:\pandora\data\export\" + filename, Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
            Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlUserResolution, true,
            Missing.Value, Missing.Value, Missing.Value);
            w1.Close(0);
            w2.Close(0);
            w3.Close(0);
            w4.Close(0);
            app.Quit();
            Marshal.ReleaseComObject(app);
            GC.Collect();
            MessageBox.Show("Complete", "Message Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

       
    }
}
