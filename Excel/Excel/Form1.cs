using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;


namespace Excel
{
    public partial class Form1 : Form
    {
        #region Excel Process ID
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> testData = new List<string>()
            { "Excel", "Access", "Word", "OnceNote"};

            // Save Excel Variable
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook wb = null;
            Worksheet ws = null;

            String filePath = null;
            String data = null;
            uint processId = 0;

            try
            {
                // Excel get the first worksheet
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wb = excelApp.Workbooks.Add();
                ws = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;

                filePath = @"C:\Users\Hongmin\Desktop/test.xls";

                // Data Input
                int r = 1;
                foreach (var d in testData)
                {
                    ws.Cells[r, 1] = d;
                    r++;
                }

                if (File.Exists(filePath))
                    File.Delete(filePath);

                wb.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal); // ** xlsx 확장자 안되는 것 확인 ㅠ

                wb = excelApp.Workbooks.Open(Filename: @filePath);
                excelApp.Visible = false;
                Range range = ws.UsedRange;


                for (int i = 1; i <= range.Rows.Count; ++i)
                {
                    for (int j = 1; j <= range.Columns.Count; ++j)
                    {
                        data += ((range.Cells[i, j] as Range).Value2.ToString() + " ");
                    }
                    data += "\n";
                }

                textBox1.Text = data;

                // close Excel
                wb.Close(true);
                excelApp.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {

                GetWindowThreadProcessId(new IntPtr(excelApp.Hwnd), out processId); // Get the Excel PID

                if (processId != 0) //Excel Process kill
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }

                // Clean up
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
                
            }
        }

        private static void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
