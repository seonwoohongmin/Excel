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
                //ws = wb.Worksheets[1];

                filePath = @"C:\Users\Hongmin\Desktop/Test.xlsx";
                //filePath = @"C:\Users\Hongmin\Desktop/test.xls"; // xls

                // Data Input
                int c = 1;
                foreach (var d in testData)
                {
                    ws.Cells[1, c] = d;
                    c++;
                }

                if (File.Exists(filePath))
                    File.Delete(filePath);

                //wb.SaveAs(filePath);
                wb.SaveAs(filePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);

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

                object[,] Griddata = range.Value;

                int rowCount = Griddata.GetLength(0);
                int columnCount = Griddata.GetLength(1);

                //엑셀 크기보다 dataGirdView가 작다면 넓혀준다.

                if (dataGridView1.RowCount < rowCount)
                {
                    dataGridView1.RowCount = rowCount;
                }

                if (dataGridView1.ColumnCount < columnCount)
                {
                    dataGridView1.ColumnCount = columnCount;
                }

                for (int row = 0; row < Griddata.GetLength(0); ++row)
                {
                    for (int column = 0; column < Griddata.GetLength(1); ++column)
                    {
                        dataGridView1[column, row].Value = Griddata[row + 1, column + 1];
                    }
                }

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
