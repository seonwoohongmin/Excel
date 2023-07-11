using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> testData = new List<string>()
            { "Excel", "Access", "Word", "OnceNote"};

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook wb = null;
            Worksheet ws = null;

            try
            {
                // Excel get the first worksheet
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                wb = excelApp.Workbooks.Add();
                ws = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;

                // Data Input
                int r = 1;
                foreach (var d in testData)
                {
                    ws.Cells[r, 1] = d;
                    r++;
                }

                // Save Excel
                wb.SaveAs(@"C:\Users\Hongmin\Desktop/test.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal); // ** xlsx 확장자 안되는 것 확인 ㅠ
                wb.Close(true);
                excelApp.Quit();

            }
            catch (Exception ex)
            {

            }
            finally
            {
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
