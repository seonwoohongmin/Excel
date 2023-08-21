using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace Excel
{
    public partial class Form1 : Form
    {
        #region Excel Process ID
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        #endregion

#if desktop
        #region DB Connection variables Surface
        string strServer = @"DESKTOP-9J5VP01\HONGMIN"; // 서버주소 -- sql 서버 접속시 사용한 주소
        string strDatabase = "TESTDB"; // 사용할 DATABASE 이름
        string strUid = "testid"; // DB 접속 아이디
        string strPassword = "1234"; // DB 접속 비밀번호
        #endregion
#else

        #region DB Connection variables desktop
        string strServer = @"DESKTOP-8HL44SC\SQLEXPRESS"; // 서버주소 -- sql 서버 접속시 사용한 주소
        string strDatabase = "TESTDB"; // 사용할 DATABASE 이름
        string strUid = "testid"; // DB 접속 아이디
        string strPassword = "1234"; // DB 접속 비밀번호
        #endregion
#endif

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
            DataGridViewComboBoxColumn comboboxcell = new DataGridViewComboBoxColumn();
            OpenFileDialog ofdfilepath = new OpenFileDialog();

            try
            {

                ofdfilepath.InitialDirectory = @"C:\Users\Hongmin\Desktop";
                ofdfilepath.Filter = "엑셀 파일 (*.xls)|*.xls";
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

                int rowCount = Griddata.GetLength(0) - 1; // -1 because checkcolumn
                int columnCount = Griddata.GetLength(1) - 1; //-1 because headername

                //엑셀 크기보다 dataGirdView가 작다면 넓혀준다.

                if (dataGridView1.RowCount < rowCount)
                {
                    dataGridView1.RowCount = rowCount;
                }

                if (dataGridView1.ColumnCount < columnCount)
                {
                    dataGridView1.ColumnCount = columnCount;
                }

                for (int column = 0; column < columnCount; ++column) //dgv 컬럼 헤더이름 넣어주기
                {
                    switch (Griddata[1, column + 1].ToString())
                    {
                        case "부가세포함":
                            dataGridView1.Columns.Insert(column, checkcolumn);
                            break;

                        case "기타출고구분":
                            comboboxcell.Items.Add("판촉");
                            comboboxcell.Items.Add("샘플출고");
                            dataGridView1.Columns.Insert(column, comboboxcell);
                            break;

                    }

                    dataGridView1.Columns[column].HeaderText = Griddata[1, column + 1].ToString();
                }

                for (int row = 0; row < rowCount; ++row)
                {
                    for (int column = 0; column < columnCount; ++column)
                    {
                        dataGridView1[column, row].Value = Griddata[row + 2, column + 1];
                    }
                }

                if (dataGridView1.Rows.Count > 0)
                {
                    SaveFileDialog save = new SaveFileDialog();
                    save.Filter = "PDF (*.pdf)|*.pdf";
                    save.FileName = "dgv.pdf";
                    bool ErrorMessage = false;
                    if (save.ShowDialog() == DialogResult.OK)
                    {
                        if (File.Exists(save.FileName))
                        {
                            try
                            {
                                File.Delete(save.FileName);
                            }
                            catch (Exception ex)
                            {
                                ErrorMessage = true;
                                MessageBox.Show("Unable to wride data in disk" + ex.Message);
                            }
                        }
                        if (!ErrorMessage)
                        {
                            try
                            {
                                PdfPTable pTable = new PdfPTable(dataGridView1.Columns.Count);
                                pTable.DefaultCell.Padding = 2;
                                pTable.WidthPercentage = 100;
                                pTable.HorizontalAlignment = Element.ALIGN_LEFT;
                                foreach (DataGridViewColumn col in dataGridView1.Columns)
                                {
                                    PdfPCell pCell = new PdfPCell(new Phrase(col.HeaderText));
                                    pTable.AddCell(pCell);
                                }
                                foreach (DataGridViewRow viewRow in dataGridView1.Rows)
                                {
                                    foreach (DataGridViewCell dcell in viewRow.Cells)
                                    {
                                        if (dcell.Value != null)
                                            pTable.AddCell(dcell.Value.ToString());
                                    }
                                }
                                using (FileStream fileStream = new FileStream(save.FileName, FileMode.Create))
                                {
                                    Document document = new Document(PageSize.A4, 8f, 16f, 16f, 8f);
                                    PdfWriter.GetInstance(document, fileStream);
                                    document.Open();
                                    document.Add(pTable);
                                    document.Close();
                                    fileStream.Close();
                                }
                                MessageBox.Show("Data Export Successfully", "info");
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error while exporting Data" + ex.Message);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("No Record Found", "Info");
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
            //string strConnectString = string.Format("Server={0};Database={1};Uid ={2};Pwd={3};", strServer, strDatabase, strUid, strPassword);
            ////string connectString = "Data Source = (local); Initial Catalog = pubs; Connection Timeout=5;Integrated Security = SSPI"; // Windows 인증

            //SqlConnection conn = new SqlConnection(strConnectString);
            //SqlCommand cmd = new SqlCommand();
            //string strColumnName = string.Empty;

            //System.Data.DataTable dtSource = new System.Data.DataTable();

            //try
            //{

            //    #region DB
            //    DBConnection(conn);
            //    dtSource = GetDataTableFromDGV(dataGridView1);

            //    if (dtSource != null)
            //    {
            //        using (var bulk = new SqlBulkCopy(conn))
            //        {
            //            bulk.DestinationTableName = "TESTTABLE";
            //            bulk.WriteToServer(dtSource);
            //        }
            //    }

            //    #endregion
            //}
            //catch (Exception ex)
            //{
            //}
            //finally
            //{
            //    DBClosed(conn);
            //}
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
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void DBClosed(SqlConnection conn)
        {
            if (conn.State == System.Data.ConnectionState.Open)
            {
                conn.Close();
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