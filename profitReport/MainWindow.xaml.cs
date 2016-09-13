using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using profitReport.AppCode;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;
using profitReport.AppCode.BLL;
using System.Windows.Media;

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
                tbShowPath.Text = fbd.SelectedPath;
        }

        private void btnGo_Click(object sender, RoutedEventArgs e)
        {
            string path = this.tbShowPath.Text.Trim();
            if (path == "")
                System.Windows.Forms.MessageBox.Show("請先選擇目錄", "Message");
            else
            {
                int filesCount = 0;
                int errorFileCount = 0;
                string msgColor = "black";
                dtResult = new DataTable();
                SetDataTable(dtResult);
                string[] fileEntries = Directory.GetFiles(path, "既有產品報價表*.xlsx");
                if (fileEntries.Length != 0)
                {
                    foreach (string filePath in fileEntries)
                    {
                        bool error = ProcessFile(filePath);
                        filesCount++;
                        if (error == true)
                            errorFileCount++;
                    }
                    // Delete from database first
                    ProfitReportBLL.deleteAllProfitReport();

                    // Insert into database
                    if (FunctionClass.DataTableToDb(dtResult, "ProfitReport"))
                    {
                        // Create excel journal report
                        string fileName = "";
                        if (rbTypePo.IsChecked == true)
                            fileName = "1_PO流水帳_系統產出";
                        else
                            fileName = "1_總表流水帳_系統產出";
                        FunctionClass.DataTableToExcelFile(dtResult, path + "\\", fileName);

                        // Create excel profit report                        
                        // 1.get data from database
                        dtResult = ProfitReportBLL.getProfitReport("2016");
                        // 2.create excel
                        FunctionClass.DataTableToExcelFile(dtResult, path + "\\", "2_利潤表");
                    }
                    else
                        System.Windows.Forms.MessageBox.Show("流水匯入資料庫失敗", "Message");

                    if (errorFileCount != 0)
                        msgColor = "red";
                    SendMsg("共處理 " + filesCount.ToString() + " 筆檔案，共有" + errorFileCount + " 筆資料有錯誤", msgColor);
                    svMsg.ScrollToBottom();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("目錄內無檔案", "Message");
                }
                dtResult = null;
            }
        }

        // Do somthing with each excel file.
        private bool ProcessFile(string path)
        {
            bool error = false;
            string sheetMsg = "";
            string functionType = "PO";
            int sheetCount = 0;
            if (rbTypeTo.IsChecked == true)
                functionType = "總表";

            string fileName = Path.GetFileName(path);
            string ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

            using (OleDbConnection conn = new OleDbConnection(ConnectionString))
            {
                conn.Open();
                DataTable dtExcelSchema = conn.GetSchema("Tables");
                conn.Close();
                foreach (DataRow dr in dtExcelSchema.Select().SkipWhile(e => e.ItemArray[2].ToString().StartsWith("_xlnm")))
                {
                    string sheetName = dr["TABLE_NAME"].ToString().Replace('#', '.');
                    if (sheetName.Length >= 4 && sheetName.Substring(sheetName.Length - 4, 2).ToUpper() == functionType)
                    {
                        sheetMsg = ProcessSheet(path, sheetName);
                        sheetCount++;
                    }
                }
                dtExcelSchema = null;
                SendMsg("檔案：" + fileName + " 有 " + sheetCount + " 張" + functionType + "表單", "black");
                if (sheetMsg.Length != 0)
                {
                    SendMsg(sheetMsg, "red");
                    error = true;
                }
            }
            return error;
        }

        // Do somthing with each sheet.
        private static string ProcessSheet(string path, string sheetName)
        {
            string result = "";
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

            string invoice = ""; // 銷貨單號, null
            string invoiceDate = ""; // 銷貨日期, null
            string purchaseOrder = sheet.GetRow(purchaseOrderPosition).GetCell(5).StringCellValue; // 訂單編號, (38,5)
            string customerPurchaseOrder = sheetName.Substring(1, sheetName.Length - 3).Replace("PO", ""); // 客戶訂單編號, (sheet name)
            string customerId = ""; // 客戶編號, null
            string salesName = sheet.GetRow(0).GetCell(10).StringCellValue; // 業務主辦, (0,10)
            string customerName = sheet.GetRow(0).GetCell(1).StringCellValue; // 客戶, (0,1)
            string soldTo = sheet.GetRow(0).GetCell(4).StringCellValue; // sold to, (0,4)

            decimal rate = Convert.ToDecimal(sheet.GetRow(0).GetCell(35).NumericCellValue); // 匯率, (0,35)
            string term = sheet.GetRow(0).GetCell(30).StringCellValue; // Term, (0,30)
            string currency = sheet.GetRow(0).GetCell(33).StringCellValue; // 幣別, (0,33)

            // Set excel start position
            int rowAdd12 = 12, rowAdd18 = 18;
            int productIdArrayRow = 3, productIdArrayColume = 0;
            int productIdMinorRow = 2, productIdMinorColumn = 3;
            int packageRowA = 3, packageRowB = 4, packageColumn = 27;
            int boxesRowA = 6, boxesRowB = 7, boxesColumn = 1;
            int quantityRow = 12, quantityColumn = 36;
            int basePriceRow = 8, suggestPriceRow = 9, salesPriceRow = 10, projectPriceRow = 12, priceColumn = 34;


            for (int i = 0; i < itemNumber; i++)
            {
                // 
                string[] productIdArray = (sheet.GetRow(productIdArrayRow).GetCell(productIdArrayColume).StringCellValue).Split(' ');// (3,0)
                if (productIdArray[0] != "")
                {
                    try
                    {
                        string productId = productIdArray[0]; // 品項(產品編號)
                        sheet.GetRow(productIdMinorRow).GetCell(productIdMinorColumn).SetCellType(CellType.String);
                        string productIdMinor = sheet.GetRow(productIdMinorRow).GetCell(productIdMinorColumn).StringCellValue; // 細項(產品編號), (2,3)
                        for (int j = 4; j < 12; j++)
                        {
                            sheet.GetRow(productIdMinorRow).GetCell(j).SetCellType(CellType.String);
                            string nextProductIdMinor = sheet.GetRow(productIdMinorRow).GetCell(j).StringCellValue;
                            if (nextProductIdMinor != "")
                                productIdMinor += "+" + nextProductIdMinor;
                        }
                        string customerProductId = productIdArray.Length == 2 ? productIdArray[1] : ""; // 客戶品項編號

                        sheet.GetRow(packageRowA).GetCell(packageColumn).SetCellType(CellType.String);
                        sheet.GetRow(packageRowB).GetCell(packageColumn).SetCellType(CellType.String);
                        string package = sheet.GetRow(packageRowA).GetCell(packageColumn).StringCellValue
                            + (sheet.GetRow(packageRowB).GetCell(packageColumn).StringCellValue == "" ? "" : " + " + sheet.GetRow(packageRowB).GetCell(packageColumn).StringCellValue); // 包材/包裝方式, (3,27) + (4,27)
                                                                                                                                                                                        //string boxes = Convert.ToInt32(sheet.GetRow(boxesRowA).GetCell(boxesColumn).NumericCellValue).ToString() + "/" + Convert.ToInt32(sheet.GetRow(boxesRowB).GetCell(boxesColumn).NumericCellValue).ToString(); // 裝箱數, (6,2) + "/" + (7,2)
                        sheet.GetRow(boxesRowA).GetCell(boxesColumn).SetCellType(CellType.Numeric);
                        sheet.GetRow(boxesRowB).GetCell(boxesColumn).SetCellType(CellType.Numeric);
                        int innerBox = Convert.ToInt32(sheet.GetRow(boxesRowA).GetCell(boxesColumn).NumericCellValue); // 內箱數, (6,2)
                        int outerBox = Convert.ToInt32(sheet.GetRow(boxesRowB).GetCell(boxesColumn).NumericCellValue); // 外箱數, (7,2)

                        int setQuantity = 0;
                        if (sheet.GetRow(quantityRow).GetCell(quantityColumn) != null)
                        {
                            sheet.GetRow(quantityRow).GetCell(quantityColumn).SetCellType(CellType.Numeric);
                            setQuantity = Convert.ToInt32(sheet.GetRow(quantityRow).GetCell(quantityColumn).NumericCellValue);
                        }

                        int quantity = setQuantity; // 製單數量, (12,36)

                        sheet.GetRow(basePriceRow).GetCell(priceColumn).SetCellType(CellType.Numeric);
                        decimal basePrice = Math.Round(Convert.ToDecimal(sheet.GetRow(basePriceRow).GetCell(priceColumn).NumericCellValue), 2); // 底價, (8,34)

                        sheet.GetRow(suggestPriceRow).GetCell(priceColumn).SetCellType(CellType.Numeric);
                        decimal suggestPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(suggestPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 建議報價, (9,34)

                        sheet.GetRow(salesPriceRow).GetCell(priceColumn).SetCellType(CellType.Numeric);
                        decimal salesPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(salesPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 業務報價, (10,34)

                        sheet.GetRow(projectPriceRow).GetCell(priceColumn).SetCellType(CellType.Numeric);
                        decimal projectPrice = Math.Round(Convert.ToDecimal(sheet.GetRow(projectPriceRow).GetCell(priceColumn).NumericCellValue), 2); // 專案價格, (12,34)

                        decimal priceOrigin = 0;// 單價(原), null
                        decimal price = 0; // 單價, null
                        decimal totalOrigin = 0; // 應收金額(原), null
                        decimal total = 0; // 應收金額, null
                        decimal tax = 0; // 應收稅額, null
                        decimal totalCost = 0; // 總成本, null
                        decimal profit = 0; // 利潤, null
                        string remark = ""; // 備註, null

                        // Write to datatable
                        WriteToDt(invoice, invoiceDate, purchaseOrder, customerPurchaseOrder, customerId, salesName, customerName,
                            soldTo, productId, productIdMinor, customerProductId, package, innerBox, outerBox,
                            rate, term, currency, quantity, basePrice, suggestPrice, salesPrice, projectPrice,
                            priceOrigin, price, totalOrigin, total, tax, totalCost, profit, remark
                           );

                        if (i != 2)
                        {
                            // row + 12
                            productIdArrayRow += rowAdd12;
                            productIdMinorRow += rowAdd12;
                            packageRowA += rowAdd12; packageRowB += rowAdd12;
                            boxesRowA += rowAdd12; boxesRowB += rowAdd12;
                            quantityRow += rowAdd12;
                            basePriceRow += rowAdd12; suggestPriceRow += rowAdd12; salesPriceRow += rowAdd12; projectPriceRow += rowAdd12;
                        }
                        else
                        {
                            // row + 18
                            // 跳過簽核欄位共六欄，所以比一般多加六行
                            productIdArrayRow += rowAdd18;
                            productIdMinorRow += rowAdd18;
                            packageRowA += rowAdd18; packageRowB += rowAdd18;
                            boxesRowA += rowAdd18; boxesRowB += rowAdd18;
                            quantityRow += rowAdd18;
                            basePriceRow += rowAdd18; suggestPriceRow += rowAdd18; salesPriceRow += rowAdd18; projectPriceRow += rowAdd18;
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorItem = productIdArray[0].ToString() + (productIdArray.Length == 2 ? productIdArray[1] : "").ToString();
                        //System.Windows.Forms.MessageBox.Show(ex.ToString(), "Message");
                        //System.Windows.Forms.MessageBox.Show("資料有誤，請檢查項目：" + errorItem + "，位置：" + productIdArrayRow + " 行", "Message");
                        result = "資料有誤，請檢查項目：" + errorItem + "，位置：" + productIdArrayRow + " 行";
                        break;
                    }
                    finally { }
                }
            }
            return result;
        }

        private static void WriteToDt(string invoice, string invoiceDate, string purchaseOrder,
            string customerPurchaseOrder, string customerId, string salesName, string customerName, string soldTo,
            string productId, string productIdMinor, string customerProductId, string package, int innerBox, int outerBox,
            decimal rate, string term, string currency, int quantity, decimal basePrice, decimal suggestPrice,
            decimal salesPrice, decimal projectPrice, decimal priceOrigin, decimal price, decimal totalOrigin,
            decimal total, decimal tax, decimal totalCost, decimal profit, string remark
            )
        {
            int i = 0;
            DataRow dtRow = dtResult.NewRow();
            dtRow[i++] = invoice;
            dtRow[i++] = invoiceDate;
            dtRow[i++] = purchaseOrder;
            dtRow[i++] = customerPurchaseOrder;
            dtRow[i++] = customerId;
            dtRow[i++] = salesName;
            dtRow[i++] = customerName;
            dtRow[i++] = soldTo;
            dtRow[i++] = productId;
            dtRow[i++] = productIdMinor;
            dtRow[i++] = customerProductId;
            dtRow[i++] = package;
            dtRow[i++] = innerBox;
            dtRow[i++] = outerBox;
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

        private void SendMsg(string msg, string color)
        {
            Run runtext = new Run();
            switch (color)
            {
                case "red":
                    runtext = new Run(msg) { Foreground = Brushes.Red };
                    break;
                case "black":
                    runtext = new Run(msg) { Foreground = Brushes.Black };
                    break;
                case "green":
                    runtext = new Run(msg) { Foreground = Brushes.Green };
                    break;
                case "blue":
                    runtext = new Run(msg) { Foreground = Brushes.Blue };
                    break;
                default:
                    runtext = new Run(msg) { Foreground = Brushes.Black };
                    break;
            }

            tbShowMsg.Inlines.Add(new LineBreak());
            tbShowMsg.Inlines.Add(runtext);
        }

        private static void SetDataTable(DataTable dt)
        {
            DataColumn column;
            //string // Invoice 銷貨單號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "銷貨單號";
            dt.Columns.Add(column);
            //string // InvoiceDate, 銷貨日期
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "銷貨日期";
            dt.Columns.Add(column);
            //string // purchaseOrder, 訂單編號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "訂單編號";
            dt.Columns.Add(column);
            //string // customerPurchaseOrder, 客戶訂單單號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶訂單單號";
            dt.Columns.Add(column);
            //string // customerId, 客戶代碼
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
            //string // productId, 品項   
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "品項";
            dt.Columns.Add(column);
            //string // productIdMinor, 細項
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "細項";
            dt.Columns.Add(column);
            //string // customerProductId, 客戶品號
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "客戶品號";
            dt.Columns.Add(column);
            //string // package, 包材/包裝方式
            column = new DataColumn();
            column.DataType = Type.GetType("System.String");
            column.ColumnName = "包材/包裝方式";
            dt.Columns.Add(column);
            //string // InnerBox, 內箱數
            column = new DataColumn();
            column.DataType = Type.GetType("System.Int32");
            column.ColumnName = "內箱數";
            dt.Columns.Add(column);
            //string // OuterBox, 外箱數
            column = new DataColumn();
            column.DataType = Type.GetType("System.Int32");
            column.ColumnName = "外箱數";
            dt.Columns.Add(column);
            //decimal rate, 匯率
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
            //decimal // unitPrice1, 單價(原)
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "單價(原)";
            dt.Columns.Add(column);
            //decimal // unitPrice2, 單價
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "單價";
            dt.Columns.Add(column);
            //decimal // total1, 應收金額(原)
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "應收金額(原)";
            dt.Columns.Add(column);
            //decimal // total2, 應收金額
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "應收金額";
            dt.Columns.Add(column);
            //decimal // tax, 應收稅額
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "應收稅額";
            dt.Columns.Add(column);
            //decimal // totalCost, 總成本
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
            column.ColumnName = "總成本";
            dt.Columns.Add(column);
            //decimal // profit, 利潤
            column = new DataColumn();
            column.DataType = Type.GetType("System.Decimal");
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