using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;

namespace profitReport.AppCode
{
    class FunctionClass
    {
        public static string GetExcelFile()
        {
            string fileName = "";
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx";
            bool? result = dlg.ShowDialog();
            if (result == true)
                fileName = dlg.FileName;
            return fileName;
        }

        public static DataTable ExcelToDataTable(string sql, string file)
        {
            DataTable dt = new DataTable();
            try
            {
                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties='Excel 12.0 Xml;HDR=YES'");
                OleDbDataAdapter da = new OleDbDataAdapter(sql, conn);
                da.Fill(dt);
                dt.TableName = "tmp";
                conn.Close();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.ToString(), "Message");
            }
            finally
            {               
            }
            return dt;
        }

        public static void DataTableToExcelFile(DataTable dt, string path, string fileName)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;
            if (dt.Rows.Count > 0)
            {
                if (dt.TableName != string.Empty)
                    ws = wb.CreateSheet(dt.TableName);
                else
                    ws = wb.CreateSheet("Sheet1");

                ws.CreateRow(0);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
                string date = DateTime.Now.Date.Year.ToString() + DateTime.Now.Date.Month.ToString() + DateTime.Now.Date.Day.ToString();
                try
                {
                    FileStream file;                    
                    file = new FileStream(path + fileName + "_" + date + ".xlsx", FileMode.Create);
                    wb.Write(file);
                    file.Close();
                    dt = null;
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("無法儲存檔案", "Message");
                }
            }
            else
                System.Windows.Forms.MessageBox.Show("無資料", "Message");
        }

        public static bool DataTableToDb(DataTable dt, string sqlDataTable)
        {
            bool result = false;       
            if (dt.Rows.Count > 0)
            {
                try
                {
                    using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
                    {
                        conn.Open();
                        using (SqlBulkCopy copy = new SqlBulkCopy(conn))
                        {
                            int columnCount = dt.Columns.Count;
                            for (int i = 0; i < columnCount; i++)
                            {
                                copy.ColumnMappings.Add(i, i);
                            }
                            copy.DestinationTableName = sqlDataTable;
                            copy.WriteToServer(dt);
                        }
                        conn.Close();
                    }
                    result = true;
                }
                catch (Exception ex)
                {
                    result = false;
                }
                finally
                {
                    dt = null;
                }
            }
            return result;
        }
    }
}
