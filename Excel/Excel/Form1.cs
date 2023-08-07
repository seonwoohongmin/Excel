using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace Excel
{
    public partial class Form1 : Form
    {
        #region Excel Process ID
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        #endregion

        #region DB Connection variables desktop
        //string strServer = @"DESKTOP-8HL44SC\SQLEXPRESS"; // 서버주소 -- sql 서버 접속시 사용한 주소
        //string strDatabase = "TESTDB"; // 사용할 DATABASE 이름
        //string strUid = "testid"; // DB 접속 아이디
        //string strPassword = "1234"; // DB 접속 비밀번호
        #endregion

        #region DB Connection variables Surface
        string strServer = @"DESKTOP-9J5VP01\HONGMIN"; // 서버주소 -- sql 서버 접속시 사용한 주소
        string strDatabase = "TESTDB"; // 사용할 DATABASE 이름
        string strUid = "testid"; // DB 접속 아이디
        string strPassword = "1234"; // DB 접속 비밀번호
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region excel
            List<string> testData = new List<string>()
            { "Excel", "Access", "Word", "OnceNote"};

            // Save Excel Variable
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook wb = null;
            Worksheet ws = null;

            String strFilePath = null;
            uint processId = 0;

            DataGridViewCheckBoxColumn checkcolumn = new DataGridViewCheckBoxColumn();
            OpenFileDialog ofdfilepath = new OpenFileDialog();

            try
            {

                ofdfilepath.InitialDirectory = @"C:\Users\Hongmin\Desktop";
                if (ofdfilepath.ShowDialog() == DialogResult.OK)
                {
                    strFilePath = ofdfilepath.FileName;
                }
                //strFilePath = @"C:\Users\Hongmin\Desktop/test.xls"; // xls
                //filePath = @"C:\Users\Hongmin\Desktop/Test.xlsx";
                #region test Data Input
                //// Excel get the first worksheet
                //excelApp = new Microsoft.Office.Interop.Excel.Application();
                //wb = excelApp.Workbooks.Add();
                //ws = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet;
                ////ws = wb.Worksheets[1];



                //// Data Input
                //int c = 1;
                //foreach (var d in testData)
                //{
                //    ws.Cells[1, c] = d;
                //    c++;
                //}

                //if (File.Exists(strFilePath))
                //    File.Delete(strFilePath);

                ////wb.SaveAs(filePath);
                //wb.SaveAs(strFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);

                //wb = excelApp.Workbooks.Open(Filename: strFilePath);
                //excelApp.Visible = false;

                //Range range = ws.UsedRange;
                //for (int i = 1; i <= range.Rows.Count; ++i)
                //{
                //    for (int j = 1; j <= range.Columns.Count; ++j)
                //    {
                //        data += ((range.Cells[i, j] as Range).Value2.ToString() + " ");
                //    }
                //    data += "\n";
                //}

                //textBox1.Text = data;
                #endregion

                excelApp = new Microsoft.Office.Interop.Excel.Application();                // 엑셀 어플리케이션 생성
                wb = excelApp.Workbooks.Open(Filename: @strFilePath);                          // 워크북 열기
                ws = wb.Worksheets.get_Item(1) as Microsoft.Office.Interop.Excel.Worksheet; // 엑셀 첫번째 워크시트 가져오기

                Range range = ws.UsedRange;    // 사용중인 셀 범위를 가져오기

                object[,] Griddata = range.Value;

                int rowCount = Griddata.GetLength(0);
                int columnCount = Griddata.GetLength(1);

                //엑셀 크기보다 dataGirdView가 작다면 넓혀준다.

                if (dataGridView1.RowCount < rowCount)
                {
                    dataGridView1.RowCount = rowCount - 1;
                }

                if (dataGridView1.ColumnCount < columnCount)
                {
                    dataGridView1.ColumnCount = columnCount - 1;
                }

                for (int column = 0; column < columnCount; ++column) //dgv 컬럼 헤더이름 넣어주기
                {
                    if ("부가세포함" == Griddata[1, column + 1].ToString())
                        dataGridView1.Columns.Insert(column, checkcolumn);

                    dataGridView1.Columns[column].HeaderText = Griddata[1, column + 1].ToString();
                }

                for (int row = 0; row < rowCount; ++row) //내용 채우기 //rowCount -1 because headername
                {
                    for (int column = 0; column < columnCount; ++column)
                    {
                        dataGridView1[column, row].Value = Griddata[row + 2, column + 1];
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                // close Excel
                wb.Close(true);
                excelApp.Quit();

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
            #endregion

            #region DB
            string strConnectString = string.Format("Server={0};Database={1};Uid ={2};Pwd={3};", strServer, strDatabase, strUid, strPassword);
            //string connectString = "Data Source = (local); Initial Catalog = pubs; Connection Timeout=5;Integrated Security = SSPI"; // Windows 인증

            SqlConnection conn = new SqlConnection(strConnectString);
            SqlCommand cmd = new SqlCommand();
            string strColumnName = string.Empty;

            System.Data.DataTable dtSource = new System.Data.DataTable();

            try
            {

                #region DB
                DBConnection(conn);
                dtSource = GetDataTableFromDGV(dataGridView1);

                if (dtSource != null)
                {
                    using (var bulk = new SqlBulkCopy(conn))
                    {
                        bulk.DestinationTableName = "TESTTABLE";
                        bulk.WriteToServer(dtSource);
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                DBClosed(conn);
            }
            #endregion


        }

        public System.Data.DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            try
            {
                var dt = new System.Data.DataTable();
                foreach (DataGridViewColumn column in dgv.Columns)
                {
                    dt.Columns.Add(column.HeaderText);
                }

                object[] cellValues = new object[dgv.Columns.Count];
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        cellValues[i] = row.Cells[i].Value;
                    }
                    dt.Rows.Add(cellValues);
                }

                return dt;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void DBConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State != System.Data.ConnectionState.Open)
                {
                    conn.Open();
                    tbDBStatus.Text = "Open";
                }
            }
            catch (Exception ex)
            {
                tbDBStatus.Text = "Open Failure";
            }
        }

        public void DBClosed(SqlConnection conn)
        {
            if (conn.State == System.Data.ConnectionState.Open)
            {
                conn.Close();
                tbDBStatus.Text = "Close";
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