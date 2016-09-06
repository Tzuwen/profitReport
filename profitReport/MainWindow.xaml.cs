using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;

namespace profitReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static DataTable dtResult;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnGetFolder_Click(object sender, RoutedEventArgs e)
        {
            ChooseFolder();
        }

        private void ChooseFolder()
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (!string.IsNullOrWhiteSpace(fbd.SelectedPath))
                this.tbShowPath.Text = fbd.SelectedPath;   
        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            string path = this.tbShowPath.Text.Trim();
            if (path == "")
                System.Windows.Forms.MessageBox.Show("請先選擇目錄", "Message"); 
            else
            {
                int filesCount = 0;
                dtResult = new DataTable();
                SetDataTable(dtResult);
                // Process the list of files found in the directory.
                string[] fileEntries = Directory.GetFiles(path, "既有產品報價表*.xlsx");
                if (fileEntries.Length != 0)
                {
                    foreach (string filePath in fileEntries)
                    {
                        ProcessFile(filePath);
                        filesCount++;
                    }
                    DataTableToExcelFile(dtResult, path + "\\");
                    dtResult = null;
                    SendMsg("流水帳已完成，共匯入 " + filesCount.ToString() + " 筆檔案");
                    this.svMsg.ScrollToBottom();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("目錄內無檔案", "Message");
                }
            }
        }

        // Do somthing with each excel file.
        private void ProcessFile(string path)
        {
            string fileName = Path.GetFileName(path);
            SendMsg("已匯入檔案：" + fileName);           
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

            using (OleDbConnection conn = new OleDbConnection(ConnectionString))
            {
                conn.Open();
                DataTable dtExcelSchema = conn.GetSchema("Tables");
                conn.Close();
                foreach (DataRow dr in dtExcelSchema.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    if (sheetName.Length >= 4 && sheetName.Substring(sheetName.Length - 4, 2) == "PO")
                        ProcessSheet(path, sheetName);
                }
                dtExcelSchema = null;
            }
        }

        // Do somthing with each sheet.
        private static void ProcessSheet(string path, string sheetName)
        {
            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }

            ISheet sheet = xssfwb.GetSheet(sheetName.Substring(1, sheetName.Length - 3));
            int itemNumber = (sheet.LastRowNum - 7) / 12; // total item number
            int purchaseOrderPosition = 38;
            switch (itemNumber)
            {
                case 1:
                    purchaseOrderPosition = 14;
                    break;
                case 2:
                    purchaseOrderPosition = 26;
                    break;
                default:
                    break;
            }

            string soldNumber = ""; // 銷貨單號, null
            string soldDate = ""; // 銷貨日期, null
            string purchaseOrder = sheet.GetRow(purchaseOrderPosition).GetCell(5).StringCellValue; // 訂單編號, (38,5)
            string customerPo = sheetName.Substring(1, sheetName.Length - 3).Replace("PO", ""); // 客戶訂單編號, (sheet name)
            string customerCode = ""; // 客戶編號, null
            string salesName = sheet.GetRow(0).GetCell(10).StringCellValue; // 業務主辦, (0,10)
            string customerName = sheet.GetRow(0).GetCell(1).StringCellValue; // 客戶, (0,1)
            string soldTo = sheet.GetRow(0).GetCell(4).StringCellValue; // sold to, (0,4)

            decimal rate = Convert.ToDecimal(sheet.GetRow(0).GetCell(35).NumericCellValue); // 匯率, (0,35)
            string term = sheet.GetRow(0).GetCell(30).StringCellValue; // Term, (0,30)
            string currency = sheet.GetRow(0).GetCell(33).StringCellValue; // 幣別, (0,33)

            // Set excel start position
            int rowAdd12 = 12, rowAdd18 = 18;
            int productCodeArrayRow = 3, productCodeArrayColume = 0;
            int productCodeMinorRow = 2, productCodeMinorColumn = 3;
            int packageRowA = 3, packageRowB = 4, packageColumn = 27;
            int boxesRowA = 6, boxesRowB = 7, boxesColumn = 1;
            int quantityRow = 12, quantityColumn = 36;
            int basePriceRow = 8, suggestPriceRow = 9, salesPriceRow = 10, projectPriceRow = 12, priceColumn = 34;

            for (int i = 0; i < itemNumber; i++)
            {
                // 
                string[] productCodeArray = (sheet.GetRow(productCodeArrayRow).GetCell(productCodeArrayColume).StringCellValue).Split(' ');// (3,0)
                if (productCodeArray[0] != "")
                {
                    try
                    {
                        string productCode = productCodeArray[0]; // 品項(產品編號)                
                        string productCodeMinor = sheet.GetRow(productCodeMinorRow).GetCell(productCodeMinorColumn).StringCellValue; // 細項(產品編號), (2,3)
                        for (int j = 4; j < 12; j++)
                        {
                            string nextProductCodeMinor = sheet.GetRow(productCodeMinorRow).GetCell(j).StringCellValue;
                            if (nextProductCodeMinor != "")
                                productCodeMinor += "+" + nextProductCodeMinor;
                        }
                        string customerProductCode = productCodeArray.Length == 2 ? productCodeArray[1] : ""; // 客戶品項編號
                        string package = sheet.GetRow(packageRowA).GetCell(packageColumn).StringCellValue
                            + (sheet.GetRow(packageRowB).GetCell(packageColumn).StringCellValue == "" ? "" : " + " + sheet.GetRow(packageRowB).GetCell(packageColumn).StringCellValue); // 包材/包裝方式, (3,27) + (4,27)
                                                                                                                                                                                        //string boxes = Convert.ToInt32(sheet.GetRow(boxesRowA).GetCell(boxesColumn).NumericCellValue).ToString() + "/" + Convert.ToInt32(sheet.GetRow(boxesRowB).GetCell(boxesColumn).NumericCellValue).ToString(); // 裝箱數, (6,2) + "/" + (7,2)
                        sheet.GetRow(boxesRowA).GetCell(boxesColumn).SetCellType(CellType.String);
                        sheet.GetRow(boxesRowB).GetCell(boxesColumn).SetCellType(CellType.String);
                        string boxes = sheet.GetRow(boxesRowA).GetCell(boxesColumn).StringCellValue + "/" + sheet.GetRow(boxesRowB).GetCell(boxesColumn).StringCellValue; // 裝箱數, (6,2) + "/" + (7,2)
                                                                                                                                                                          //
                        int quantity = Convert.ToInt32(sheet.GetRow(quantityRow).GetCell(quantityColumn).NumericCellValue); // 製單數量, (12,36)
                        decimal basePrice = Math.Round(Convert.ToDecimal(sheet.GetRow(basePriceRow).GetCell(priceColumn).NumericCellValue), 2); // 底價, (8,34)
                        decimal suggestPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(suggestPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 建議報價, (9,34)
                        decimal salesPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(salesPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 業務報價, (10,34)
                        decimal projectPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(projectPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 專案價格, (12,34)
                        string priceOrigin = ""; // 單價(原), null
                        string price = ""; // 單價, null
                        string totalOrigin = ""; // 應收金額(原), null
                        string total = ""; // 應收金額, null
                        string tax = ""; // 應收稅額, null
                        string totalCost = ""; // 總成本, null
                        string profit = ""; // 利潤, null
                        string remark = ""; // 備註, null

                        // Write to datatable
                        WriteToDt(soldNumber, soldDate, purchaseOrder, customerPo, customerCode, salesName, customerName,
                            soldTo, productCode, productCodeMinor, customerProductCode, package, boxes,
                            rate, term, currency, quantity, basePrice, suggestPrice, salesPrice, projectPrice,
                            priceOrigin, price, totalOrigin, total, tax, totalCost, profit, remark

                           );

                        if (i != 2)
                        {
                            // row + 12
                            productCodeArrayRow += rowAdd12;
                            productCodeMinorRow += rowAdd12;
                            packageRowA += rowAdd12; packageRowB += rowAdd12;
                            boxesRowA += rowAdd12; boxesRowB += rowAdd12;
                            quantityRow += rowAdd12;
                            basePriceRow += rowAdd12; suggestPriceRow += rowAdd12; salesPriceRow += rowAdd12; projectPriceRow += rowAdd12;
                        }
                        else
                        {
                            // row + 18
                            // 跳過簽核欄位共六欄，所以比一般多加六行
                            productCodeArrayRow += rowAdd18;
                            productCodeMinorRow += rowAdd18;
                            packageRowA += rowAdd18; packageRowB += rowAdd18;
                            boxesRowA += rowAdd18; boxesRowB += rowAdd18;
                            quantityRow += rowAdd18;
                            basePriceRow += rowAdd18; suggestPriceRow += rowAdd18; salesPriceRow += rowAdd18; projectPriceRow += rowAdd18;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(ex.ToString(), "Message");
                    }
                    finally { }
                }
            }
        }

        private static void WriteToDt(string soldNumber, string soldDate, string purchaseOrder,
            string customerPo, string customerCode, string salesName, string customerName, string soldTo,
            string productCode, string productCodeMinor, string customerProductCode, string package, string boxes,
            decimal rate, string term, string currency, int quantity, decimal basePrice, decimal suggestPrice,
            decimal salesPrice, decimal projectPrice, string priceOrigin, string price, string totalOrigin,
            string total, string tax, string totalCost, string profit, string remark
            )
        {
            int i = 0;
            DataRow dtRow = dtResult.NewRow();
            dtRow[i++] = soldNumber;
            dtRow[i++] = soldDate;
            dtRow[i++] = purchaseOrder;
            dtRow[i++] = customerPo;
            dtRow[i++] = customerCode;
            dtRow[i++] = salesName;
            dtRow[i++] = customerName;
            dtRow[i++] = soldTo;
            dtRow[i++] = productCode;
            dtRow[i++] = productCodeMinor;
            dtRow[i++] = customerProductCode;
            dtRow[i++] = package;
            dtRow[i++] = boxes;
            dtRow[i++] = rate;
            dtRow[i++] = term;
            dtRow[i++] = currency;
            dtRow[i++] = quantity;
            dtRow[i++] = basePrice;
            dtRow[i++] = suggestPrice;
            dtRow[i++] = salesPrice;
            dtRow[i++] = projectPrice;
            dtRow[i++] = priceOrigin;
            dtRow[i++] = price;
            dtRow[i++] = totalOrigin;
            dtRow[i++] = total;
            dtRow[i++] = tax;
            dtRow[i++] = totalCost;
            dtRow[i++] = profit;
            dtRow[i] = remark;
            dtResult.Rows.Add(dtRow);
        }

        private static void DataTableToExcelFile(DataTable dt, string path)
        {
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            if (dt.TableName != string.Empty)
                ws = wb.CreateSheet(dt.TableName);
            else
                ws = wb.CreateSheet("Sheet1");

            ws.CreateRow(0);// Create row 1 for column name
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
                FileStream file = new FileStream(path + "流水帳_系統產出_" + date + ".xlsx", FileMode.Create);
                wb.Write(file);
                file.Close();
                dt = null;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("無法儲存檔案", "Message");
            }
        }

        private void SendMsg(string msg)
        {            
            this.tbShowMsg.Inlines.Add(new LineBreak());
            this.tbShowMsg.Inlines.Add(msg);
        }

        private static void SetDataTable(DataTable dt)
        {
            DataColumn column;
            //string // soldNumber 銷貨單號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "銷貨單號";
            dt.Columns.Add(column);
            //string // soldDate, 銷貨日期
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "銷貨日期";
            dt.Columns.Add(column);
            //string // purchaseOrder, 訂單編號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "訂單編號";
            dt.Columns.Add(column);
            //string // customerPo, 客戶訂單單號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶訂單單號";
            dt.Columns.Add(column);
            //string // customerCode, 客戶代碼
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶代碼";
            dt.Columns.Add(column);
            //string // salesName, 業務主辦
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "業務主辦";
            dt.Columns.Add(column);
            //string // customerName, 客戶
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶";
            dt.Columns.Add(column);
            //string // soldTo, 市場/品牌/SellTo
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "市場/品牌/SellTo";
            dt.Columns.Add(column);
            //string // productCode, 品項   
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "品項";
            dt.Columns.Add(column);
            //string // productCodeMinor, 細項
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "細項";
            dt.Columns.Add(column);
            //string // customerProductCode, 客戶品號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶品號";
            dt.Columns.Add(column);
            //string // package, 包材/包裝方式
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "包材/包裝方式";
            dt.Columns.Add(column);
            //string // boxes, 裝箱數
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "裝箱數";
            dt.Columns.Add(column);
            //decimal // rate, 匯率
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Decimal");
            column.ColumnName = "匯率";
            dt.Columns.Add(column);
            //string // term, Term
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "Term";
            dt.Columns.Add(column);
            //string // currency, 幣別
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "幣別";
            dt.Columns.Add(column);
            //int // quantity, 製單數量
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "製單數量";
            dt.Columns.Add(column);
            //decimal // basePrice, 底價
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Decimal");
            column.ColumnName = "底價";
            dt.Columns.Add(column);
            //decimal // suggestPrice, 建議報價 1.05%
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Decimal");
            column.ColumnName = "建議報價";
            dt.Columns.Add(column);
            //decimal // salesPrice, 業務報價
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Decimal");
            column.ColumnName = "業務報價";
            dt.Columns.Add(column);
            //decimal // projectPrice, 專案價格
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Decimal");
            column.ColumnName = "專案價格";
            dt.Columns.Add(column);
            //string // priceOrigin, 單價(原)
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "單價(原)";
            dt.Columns.Add(column);
            //string // price, 單價
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "單價";
            dt.Columns.Add(column);
            //string // totalOrigin, 應收金額(原)
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "應收金額(原)";
            dt.Columns.Add(column);
            //string // total, 應收金額
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "應收金額";
            dt.Columns.Add(column);
            //string // tax, 應收稅額
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "應收稅額";
            dt.Columns.Add(column);
            //string // totalCost, 總成本
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "總成本";
            dt.Columns.Add(column);
            //string // profit, 利潤
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "利潤";
            dt.Columns.Add(column);
            //string // remark, 備註
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "備註";
            dt.Columns.Add(column);
        }
    }
}
