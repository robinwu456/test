using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace SalesOrdersExcel {
    public partial class Form1 : Form {

        //計算同商品數量
        //執行檔路徑 SalesOrdersExcel\SalesOrdersExcel\bin\Debug\SalesOrdersExcel.exe

        string salesOrderFormPath = string.Empty;       //訂單路徑
        string storeCostFormPath = string.Empty;        //成本表路徑

        int totalProfit = 0;    //總利潤

        List<string[]> productCountList = new List<string[]>();     //商品種數列表

        bool processPermit = true;

        //richbox顯示商品種數列表
        public void PrintProductCountList() {

            string spacetxt = string.Empty;

            richTextBox1.AppendText("-----------------------------------------------------------------------------------\n");

            for (int i = 0; i < productCountList.Count; i++)
            {
                if (i == 0)
                {
                    richTextBox1.AppendText(" 共有" + productCountList.Count + "種商品：\n");
                    richTextBox1.AppendText("\n");
                    richTextBox1.AppendText("      數量\t商品名稱\n");
                    richTextBox1.AppendText("\n");
                }

                spacetxt = (i + 1) < 10 ? ".    " : ".  ";
                
                richTextBox1.AppendText(" " + (i + 1) + spacetxt + productCountList[i][1] + "\t" + productCountList[i][0] + "\n");
            }

            richTextBox1.AppendText("\n");

            spacetxt = (totalProfit > 0) ? (" 總利潤：" + totalProfit + "\n") : (" 總利潤：有東西無成本，無法判斷\n");
 
            richTextBox1.AppendText(spacetxt);
        }

        //計算商品數量
        public void ProductQuantityCount(string productName, int productOption, int productQuantity) {

            int totalQuantity = productOption * productQuantity;            

            if (productCountList.Count == 0)
            {
                string[] list = { productName, totalQuantity.ToString() };
                productCountList.Add(list);
            }
            else
            {
                for (int i = 0; i < productCountList.Count; i++)
                {
                    string[] nowlist = productCountList[i];
                    if (nowlist[0] == productName)
                    {
                        int oriQuan = int.Parse(productCountList[i][1]);
                        oriQuan += totalQuantity;
                        productCountList[i][1] = oriQuan.ToString();
                        break;
                    }
                    if (i == productCountList.Count - 1)
                    {
                        string[] list = { productName, totalQuantity.ToString() };
                        productCountList.Add(list);
                        break;
                    }
                }
            }

        }

        //切割商品資訊至字串陣列
        public string[] ProductInfoTransfer(string str) {

            string[] infos = str.Split(';');

            for (int i = 0; i < infos.Length; i++)
            {
                if (infos[i].Contains("商品名稱"))
                {
                    infos[i] = infos[i].Substring(infos[i].IndexOf(":") + 1);
                }
                else if (infos[i].Contains("商品選項名稱"))
                {
                    infos[i] = infos[i].Substring(infos[i].IndexOf(":") + 1).Replace("入", "");

                    if (!Int32.TryParse(infos[i], out int number))
                    {
                        string[] word = { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" };

                        if (infos[i] == "兩") { infos[i] = "二"; }
                        if (infos[i] == "") { infos[i] = "一"; }

                        //項目顯示其他處理為"一入"
                        for (int j = 0; j < word.Length; j++)
                        {
                            if (infos[i] == word[j])
                            {
                                break;
                            }
                            else if (j == word.Length - 1)
                            {
                                infos[i] = "一";
                            }
                        }

                        //幾入國字轉數字
                        for (int j = 0; j < word.Length; j++)
                        {
                            if (infos[i] == word[j])
                            {
                                infos[i] = (j + 1).ToString();
                            }
                        }
                    }
                }
                else if (infos[i].Contains("價格"))
                {
                    infos[i] = infos[i].Substring(infos[i].IndexOf("$") + 2).Replace(",", "");
                }
                else if (infos[i].Contains("數量"))
                {
                    infos[i] = infos[i].Substring(infos[i].IndexOf(":") + 2);
                }
            }

            return infos;
        }

        //多項商品字串切割
        public List<string[]> ProductQuantity(string[] strs) {

            List<string[]> list = new List<string[]>();

            string s = string.Empty;

            for (int i = 0; i < strs.Length; i++)
            {
                if (i % 4 == 0 && i > 0)
                {
                    s += "|";
                }
                s += strs[i] + ";";
            }

            string[] items = s.Split('|');

            for (int i = 0; i < items.Length - 1; i++)
            {
                string[] itemsItem = items[i].Split(';');
                list.Add(itemsItem);
            }

            return list;
        }

        //下載成本表
        public string[,] DownloadStoreCostForm(string path) {

            richTextBox1.AppendText("從 " + storeCostFormPath + " 下載成本表...\n");
            MessageBox.Show("成本表下載");
            Excel.Application excel = new Excel.Application();
            Excel.Workbook sheet = excel.Workbooks.Open(storeCostFormPath);
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;

            Excel.Range userRange = x.UsedRange;

            string[,] form = new string[userRange.Rows.Count + 1, userRange.Columns.Count + 1];

            string copytext = richTextBox1.Text;

            try
            {
                //存入商品名稱
                for (int i = 1; i <= userRange.Rows.Count; i++)
                {
                    if (userRange.Cells[i, 1].value != null)
                    {
                        string item = userRange.Cells[i, 1].value;
                        form[i, 1] = item;

                        float progressRate = ((float)i / (float)userRange.Rows.Count) * 50;
                        richTextBox1.Text = String.Format("{0}下載成本表...( {1}% )\n", copytext, (int)progressRate);
                        progressBar1.Value = ((int)progressRate * 15) / 50 + 20;
                    }
                }

                //存入商品成本
                for (int i = 1; i <= userRange.Rows.Count; i++)
                {
                    if (userRange.Cells[i, 2].value != null)
                    {
                        double item = userRange.Cells[i, 2].value;
                        form[i, 2] = item.ToString();

                        float progressRate = ((float)i / (float)userRange.Rows.Count) * 50 + 50;
                        richTextBox1.Text = String.Format("{0}下載成本表...( {1}% )\n", copytext, (int)progressRate);
                        progressBar1.Value = ((int)progressRate * 15) / 100 + 20;
                    }
                }
            }
            catch (Exception excp)
            {
                richTextBox1.AppendText("錯誤資訊：" + excp.ToString() + "\n");
            }
            MessageBox.Show("刪除成本");
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            
            richTextBox1.AppendText("成本表下載完成\n");

            return form;
        }

        //成本判斷
        public int ProductStoreCost(string[,] costForm, string productName) {

            int storeCostRows = costForm.GetLength(0);
            int productStoreCost = 0;

            //判斷成本表商品名稱
            for (int i = 1; i < storeCostRows; i++)
            {
                if (productName.Contains(costForm[i, 1]))
                {
                    if (Int32.TryParse(costForm[i, 2], out int result))
                    {
                        productStoreCost = result;
                        break;
                    }
                    else
                    {
                        //richTextBox1.AppendText("\n該商品(" + productName + ")成本價有問題");
                        break;
                    }
                }
            }

            return productStoreCost;
        }

        //計算單筆訂單利潤
        public int SingleOrderProfit(List<string[]> orderList, string[,] costForm, string payment) {

            int profit = 0;     //利潤       

            double fee = 0;     //手續費
            double paymentRef = payment == "信用卡/VISA金融卡" ? 0.015 : 0.005;       //手續費乘數

            //利潤加總
            for (int i = 0; i < orderList.Count; i++)
            {
                //商品加入種類清單以方便計算
                ProductQuantityCount(orderList[i][0], Int32.Parse(orderList[i][1]), Int32.Parse(orderList[i][3]));

                if (ProductStoreCost(costForm, orderList[i][0]) == 0)
                {
                    profit = 0;
                    goto exit;
                }
                else
                {
                    //顯示商品各計算數字
                    //Console.WriteLine("({0}-{1}*{2})*{3}", Int32.Parse(orderList[i][2]), ProductStoreCost(costForm, orderList[i][0]), Int32.Parse(orderList[i][1]), Int32.Parse(orderList[i][3]));
                    //利潤 = (價格 - 成本 * 選項) * 數量
                    profit += ((Int32.Parse(orderList[i][2]) - (ProductStoreCost(costForm, orderList[i][0]) * Int32.Parse(orderList[i][1])))) * Int32.Parse(orderList[i][3]);                    
                }
                
            }

            //計算手續費
            for (int i = 0; i < orderList.Count; i++)
            {
                fee += (Int32.Parse(orderList[i][2]) * Int32.Parse(orderList[i][3]));

                if (i == orderList.Count - 1)
                {
                    fee = Math.Round(fee * paymentRef, 1);
                }
            }

            //Console.WriteLine("fee = " + Convert.ToInt32(fee));

            profit -= Convert.ToInt32(fee);

            exit:

            //利潤加總判斷
            TotalProfitCount(profit);

            return profit;
        }

        //利潤加總
        public void TotalProfitCount(int profit) {

            if (totalProfit >= 0)
            {
                if (profit == 0)
                {
                    totalProfit = -1;
                }
                else
                {
                    totalProfit += profit;
                }
            }
        }

        //瀏覽Excel路徑
        public string OpenFile() {

            string filePath = string.Empty;

            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();

            OpenFileDialog1.Filter = "Excel Worksheets| *.xlsx";

            if (OpenFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = OpenFileDialog1.FileName.ToString();
            }

            return filePath;
        }

        //資料路徑"\"轉換
        public string PathTransfer(string path) {

            path = path.Replace(@"\", @"\\");

            return path;
        }

        //Excel運行
        public void ExcelWorking(Excel.Application excel, Excel.Workbook sheet) {

            richTextBox1.AppendText("程序開始...\n");

            richTextBox1.AppendText("從 " + salesOrderFormPath + " 讀取訂單...\n");
            
            Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            richTextBox1.AppendText("name=" + x.Name + "\n");
            Excel.Range userRange = x.UsedRange;

            richTextBox1.AppendText("讀取成功\n");
            progressBar1.Value = 20;

            string[,] costForm = DownloadStoreCostForm(storeCostFormPath);

            int singleTotalIdx = 0;
            int productInfoIdx = 0;
            int paymentMethodIdx = 0;

            //判斷運算項目欄位
            for (int i = 1; i <= userRange.Columns.Count; i++)
            {
                switch (userRange.Cells[1, i].value)
                {
                    case "訂單小計 (TWD)":
                        singleTotalIdx = i;
                        break;
                    case "商品資訊":
                        productInfoIdx = i;
                        break;
                    case "付款方式":
                        paymentMethodIdx = i;
                        break;
                }
            }

            try
            {
                string copytext = richTextBox1.Text;

                for (int testRows = 2; testRows <= userRange.Rows.Count; testRows++)
                {
                    if (testRows == 2)
                    {
                        x.Cells[1, userRange.Columns.Count + 1] = "利潤";
                    }

                    string productInfoString = userRange.Cells[testRows, productInfoIdx].value;
                    string paymentMethodString = userRange.Cells[testRows, paymentMethodIdx].value;

                    //單項商品List
                    string[] productList = ProductInfoTransfer(productInfoString);

                    List<string[]> list = new List<string[]>();

                    //判斷商品數量
                    if (productList.Length - 1 > 4)
                    {
                        list = ProductQuantity(productList);
                    }
                    else
                    {
                        list.Add(productList);
                    }

                    int profit = SingleOrderProfit(list, costForm, paymentMethodString);

                    //填入Excel                        
                    x.Cells[testRows, userRange.Columns.Count + 1] = profit == 0 ? "找不到成本" : profit.ToString();

                    float progressRate = ((float)testRows / (float)userRange.Rows.Count) * 100;
                    richTextBox1.Text = String.Format("{0}計算中...{1}/{2} ( {3}% )\n", copytext, testRows, userRange.Rows.Count, (int)progressRate);
                    progressBar1.Value = ((int)progressRate * 50) / 100 + 50;
                }

                //Test();
            }
            catch (Exception a)
            {
                richTextBox1.AppendText("錯誤資訊：" + a.ToString() + "\n");
            }
        }

        //測試
        public void Test() {
            /*
            int testIndex = 2;      //測試第幾筆訂單
            string productInfoString = userRange.Cells[testIndex, productInfoIdx].value;
            string paymentMethodString = userRange.Cells[testIndex, paymentMethodIdx].value;

            //單項商品List
            string[] productList = ProductInfoTransfer(productInfoString);

            List<string[]> list = new List<string[]>();

            //判斷商品數量
            if (productList.Length - 1 > 4)
            {
                list = ProductQuantity(productList);
            }
            else
            {
                list.Add(productList);
            }

            int profit = SingleOrderProfit(list, costForm, paymentMethodString);

            //Console.WriteLine("profit = " + profit);
            
            */
        }

        //修改Excel
        public void ExcelUpdata() {

            if (salesOrderFormPath == "" || storeCostFormPath == "")
            {
                richTextBox1.AppendText("請選取路徑\n");
            }
            else
            {
                MessageBox.Show("修改excel");
                Excel.Application excel = new Excel.Application();
                Excel.Workbook sheet = excel.Workbooks.Open(salesOrderFormPath);

                ExcelWorking(excel,sheet);
                MessageBox.Show("刪除修改excel");
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();
                Thread.Sleep(3000);
                richTextBox1.AppendText("程序結束。\n");

                //顯示種類清單
                PrintProductCountList();
            }
        }

        //另存Excel
        public void ExcelSaveAs() {

            if (salesOrderFormPath == "" || storeCostFormPath == "")
            {
                richTextBox1.AppendText("請選取路徑\n");
            }
            else
            {
                SaveFileDialog save = new SaveFileDialog();

                save.Filter = "Excel Worksheets| *.xlsx";

                if (save.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application excel = new Excel.Application();
                    Excel.Workbook sheet = excel.Workbooks.Open(salesOrderFormPath);

                    ExcelWorking(excel, sheet);

                    sheet.SaveAs(save.FileName);

                    sheet.Close(true, Type.Missing, Type.Missing);
                    excel.Quit();

                    richTextBox1.AppendText("程序結束。\n");

                    //顯示種類清單
                    PrintProductCountList();
                }
            }
        }              

        public Form1() {
            InitializeComponent();

            progressBar1.Value = 0;
            processPermit = true;
        }

        public void CheckExcelProcess() {

            foreach(System.Diagnostics.Process proc in System.Diagnostics.Process.GetProcesses())
            {
                if(proc.ProcessName == "EXCEL")
                {
                    processPermit = false;
                    richTextBox1.AppendText("請關閉Excel");
                }
            }

        }

        //修改訂單
        private void button2_Click(object sender, EventArgs e) {

            richTextBox1.Text = string.Empty;
            productCountList.Clear();

            CheckExcelProcess();

            if (processPermit)
            {
                ExcelUpdata();
            }
        }

        //訂單瀏覽
        private void button1_Click(object sender, EventArgs e) {

            salesOrderFormPath = PathTransfer(OpenFile());

            textBox1.Text = salesOrderFormPath;
        }

        //成本表瀏覽
        private void button3_Click(object sender, EventArgs e) {

            storeCostFormPath = PathTransfer(OpenFile());

            textBox2.Text = storeCostFormPath;
        }

        //另存訂單
        private void button4_Click(object sender, EventArgs e) {

            richTextBox1.Text = string.Empty;
            productCountList.Clear();

            CheckExcelProcess();

            if (processPermit)
            {
                ExcelSaveAs();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e) {

        }
    }
}
