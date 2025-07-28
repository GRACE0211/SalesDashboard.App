using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace SalesDashboard
{
    public partial class SalesMonthlyUserControl : UserControl
    {
        private string username;
        private string connectionString = "server=localhost;user=root;password=0905656870;database=sales_dashboard;port=3306;SslMode=None";
        private int highlightedPieIndex = -1; // 用於記錄高亮顯示的行索引
        private int highlightedBarIndex = -1; // 用於記錄高亮顯示的行索引
        private int highlightedRowIndex = -1; // 用於記錄高亮顯示的行索引

        public SalesMonthlyUserControl(string username)
        {
            InitializeComponent();
            this.username = username; // 設定使用者名稱
            dateTimePickerChart.Value = new DateTime(2024, 12, 1);
        }


        // MainForm_Load 設定預設內容, for panel 的 ToolTip, 滑鼠移入會顯示完整內容
        private void MainForm_Load(object sender, EventArgs e)
        {
            toolTip1.SetToolTip(labelSalesOrderCustomer, "客戶：全部");
            toolTip1.SetToolTip(labelSalesOrderProduct, "商品：全部");
            toolTip1.SetToolTip(labelSalesRevenueCustomer, "客戶：全部");
            toolTip1.SetToolTip(labelSalesRevenueProduct, "商品：全部");

            for (int i = 0; i < checkedListBoxProducts_sales.Items.Count; i++)
            {
                checkedListBoxProducts_sales.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedProducts_sales = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();

            for (int i = 0; i < checkedListBoxCustomers_sales.Items.Count; i++)
            {
                checkedListBoxCustomers_sales.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedCustomers_sales = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();

            InitTrendChart();



        }


        private void SalesMonthlyUserControl_Load(object sender, EventArgs e)
        {
            labelSales.Text = $"{username} 的月報表"; // 顯示使用者名稱
            List<string> selectedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers);
            LoadMonthlyOrdersChartByCustomer(username, selectedProducts, selectedCustomers);
            LoadMonthlyRevenueDetailsChart(username, selectedProducts, selectedCustomers);
            for (int i = 0; i < checkedListBoxProducts_sales.Items.Count; i++)
            {
                checkedListBoxProducts_sales.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedProducts_sales = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();

            for (int i = 0; i < checkedListBoxCustomers_sales.Items.Count; i++)
            {
                checkedListBoxCustomers_sales.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedCustomers_sales = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();

            InitTrendChart();
        }


        // ------------------ SalesChartPage_Monthly -- tabPage3 ------------------


        // tabPage3, 右中圓餅圖 -- 計算某月某產品的總銷售量
        private DataTable GetMonthlyProductsData(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份

            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";
            var customerFilter = selectedCustomers.Count > 0 ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")" : "";


            string query = $@"
                select p.name,sum(o.amount) as total_amount
                 from orders o
                    left join products p on o.product_id = p.product_id
                    left join customers c on o.customer_id = c.customer_id
                    left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth 
                      and u.username = @username
                      {productFilter}
                      {customerFilter}
                group by o.product_id
                order by total_amount desc;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);

                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);

                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadMonthlyProductsChart(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DataTable dt = GetMonthlyProductsData(username, selectedProducts, selectedCustomers);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartSalesProducts.Titles.Clear();
                chartSalesProducts.Titles.Add("無資料顯示!!");
                chartSalesProducts.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartSalesProducts.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartSalesProducts.Series.Clear();
                return;
            }
            decimal totalAmount = 0;
            foreach (DataRow row in dt.Rows)
            {
                totalAmount += Convert.ToDecimal(row["total_amount"]);
            }
            chartSalesProducts.Titles.Clear();
            chartSalesProducts.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}產品銷售占比"); // 設定圖表標題
            chartSalesProducts.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartSalesProducts.Series.Clear();
            var series = chartSalesProducts.Series.Add("Products");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "產品名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式

            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }

            foreach (DataRow row in dt.Rows)
            {
                string product = row["name"].ToString();
                int amount = Convert.ToInt32(row["total_amount"]);
                double percent = totalAmount == 0 ? 0 : (double)(amount / totalAmount) * 100;
                int idx = series.Points.AddXY(product, amount);
                series.Points[idx].LegendText = $"{product}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].Label = $"{percent:F1}%";
                series.Points[idx].ToolTip = $"{product} - {amount}(個)";
            }
            // 若有 highlihgtedPieIndex，爆開+變色
            if (highlightedPieIndex >= 0 && highlightedPieIndex < series.Points.Count)
            {
                series.Points[highlightedPieIndex]["Exploded"] = "true"; // 爆開
                series.Points[highlightedPieIndex].Color = Color.MediumTurquoise; // 高亮
            }

        }

        // tabPage3, 右上橫條圖 -- 計算某月某客戶的訂單數量
        private DataTable GetMonthlyOrdersData(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份

            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";
            var customerFilter = selectedCustomers.Count > 0 ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")" : "";


            string query = $@"
                select c.name as customer_name, count(*) as o_count
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                and u.username = @username
                {productFilter}
                {customerFilter}
                group by o.customer_id
                order by o_count desc;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }


        private void LoadMonthlyOrdersChartByCustomer(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {

            DataTable dt = GetMonthlyOrdersData(username, selectedProducts, selectedCustomers);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartSalesCustomersOrders.Titles.Clear();
                chartSalesCustomersOrders.Titles.Add("無資料顯示!!");
                chartSalesCustomersOrders.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartSalesCustomersOrders.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartSalesCustomersOrders.Series.Clear();
                return;
            }
            chartSalesCustomersOrders.Titles.Clear();
            chartSalesCustomersOrders.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}客戶訂單量"); // 設定圖表標題
            chartSalesCustomersOrders.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartSalesCustomersOrders.Series.Clear();
            chartSalesCustomersOrders.ChartAreas[0].AxisX.LabelStyle.Font = new Font("微軟正黑體", 10, FontStyle.Bold);
            var series = chartSalesCustomersOrders.Series.Add("Customers");
            series.ChartType = SeriesChartType.Bar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "客戶名稱: 銷售量"
            }
            foreach (DataRow row in dt.Rows)
            {
                string customer = row["customer_name"].ToString();
                int count = Convert.ToInt32(row["o_count"]);
                int idx = series.Points.AddXY(customer, count);
                series.Points[idx].LegendText = $"{customer}: {count}"; // 設定標籤格式為 "客戶名稱: 銷售量"
            }

            for (int i = 0; i < chartSalesCustomersOrders.Series[0].Points.Count; i++)
            {
                if (i == highlightedBarIndex)
                {
                    chartSalesCustomersOrders.Series[0].Points[i].Color = Color.LightSteelBlue; // 高亮色
                }
                else
                {
                    chartSalesCustomersOrders.Series[0].Points[i].Color = Color.Sienna; // 原色
                }

            }
        }


        // tabPage3, 最下面長條圖 -- 取得某月的銷售數據
        private DataTable GetMonthlyRevenueDetails(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DataTable dataTable = new DataTable();

            DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份

            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";
            var customerFilter = selectedCustomers.Count > 0 ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")" : "";

            string query = $@"
                select c.name as customer_name, p.name as product_name, sum(o.amount) as total_amount, sum(p.price*o.amount) as revenue
                from orders o
                    left join products p on o.product_id = p.product_id
                    left join customers c on o.customer_id = c.customer_id
                    left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth 
                      and u.username = @username
                      {productFilter}
                      {customerFilter}
                group by o.customer_id, o.product_id
                order by o.customer_id, o.product_id;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadMonthlyRevenueDetailsChart(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DataTable dt = GetMonthlyRevenueDetails(username, selectedProducts, selectedCustomers);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartMonthlyRevenuePerProduct.Titles.Clear();
                chartMonthlyRevenuePerProduct.Titles.Add("無資料顯示!!");
                chartMonthlyRevenuePerProduct.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartMonthlyRevenuePerProduct.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartMonthlyRevenuePerProduct.Series.Clear();
                return;
            }
            chartMonthlyRevenuePerProduct.Titles.Clear();
            chartMonthlyRevenuePerProduct.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}銷售數據"); // 設定圖表標題
            chartMonthlyRevenuePerProduct.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartMonthlyRevenuePerProduct.Series.Clear();
            var series = chartMonthlyRevenuePerProduct.Series.Add("Revenue");
            series.ChartType = SeriesChartType.Column;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "$ #VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9); // 設定標籤字體樣式
            chartMonthlyRevenuePerProduct.ChartAreas[0].AxisX.LabelStyle.Font = new Font("微軟正黑體", 10, FontStyle.Bold);
            foreach (DataRow row in dt.Rows)
            {
                string customer = row["customer_name"].ToString();
                string product = row["product_name"].ToString();
                int amount = Convert.ToInt32(row["total_amount"]);
                decimal revenueValue = Convert.ToDecimal(row["revenue"]);
                //Debug.WriteLine($"Customer: {customer}-{product}, Amount: {amount}, Revenue: {revenueValue}");

                int idx = series.Points.AddXY($"{customer} - {product}", revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{customer} - {product}: {revenueValue:N0}"; // 設定圖例文本
            }

            int maxIndex = 0;
            double maxValue = double.MinValue;
            for (int i = 0; i < series.Points.Count; i++)
            {
                if (series.Points[i].YValues[0] > maxValue)
                {
                    var val = series.Points[i].YValues[0];
                    if (val > maxValue)
                    {
                        maxValue = val;
                        maxIndex = i;
                    }
                }
            }
            series.Points[maxIndex].Color = Color.RoyalBlue; // 將最大值的點設置為紅色
            series.Points[maxIndex].Font = new Font("微軟正黑體", 10, FontStyle.Bold);

            for (int i = 0; i < chartMonthlyRevenuePerProduct.Series[0].Points.Count; i++)
            {
                if (i == highlightedRowIndex)
                {
                    chartMonthlyRevenuePerProduct.Series[0].Points[i].Color = Color.Orange; // 高亮色
                }
                else
                {
                    chartMonthlyRevenuePerProduct.Series[0].Points[i].Color = Color.CornflowerBlue; // 原色
                }

            }

        }



        // tabPage3, 左上的dataGridView -- 訂單的日期/客戶/商品
        private void GetMonthlyOrdersRowData(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {

            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份
            string productFilter = selectedProducts.Count > 0
                ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")"
                : "";
            string customerFilter = selectedCustomers.Count > 0
                ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")"
                : "";

            string query = $@"
                SELECT DISTINCT DATE(o.order_date) as 日期, 
                SUBSTRING(c.name,3,1) as 客戶,
                p.name as 產品,
                concat(o.amount,' * $',p.price,' = $',o.amount*p.price) as 總價
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE u.username = @username
                AND YEAR(o.order_date) = @searchYear
                AND MONTH(o.order_date) = @searchMonth
                {productFilter}
                {customerFilter}
             ;";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            dataGridViewMonthlyOrders.DataSource = dataTable; // 將查詢結果綁定到 DataGridView
            dataGridViewMonthlyOrders.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 自動調整列寬
            dataGridViewMonthlyOrders.ColumnHeadersDefaultCellStyle.Font = new Font("微軟正黑體", 9F);
            dataGridViewMonthlyOrders.DefaultCellStyle.Font = new Font("微軟正黑體", 9);

        }

        // tabPage3, 趨勢圖表 -- 點擊產品跳出產品數量趨勢圖

        private void InitTrendChart()
        {
            chartSalesRevenueTrend.Series.Clear();
            chartSalesRevenueTrend.Titles.Clear();
            chartSalesRevenueTrend.Titles.Add("請點選任意表格的標籤查看趨勢圖");
            chartSalesRevenueTrend.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold);
            chartSalesRevenueTrend.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
            chartSalesRevenueTrend.Visible = true;
        }

        // 將選到的圖表高亮
        private void HighlightPanel(Panel panelToHighlight)
        {
            panelBar.BackColor = Color.Transparent; // 恢復其他面板背景色
            panelColumn.BackColor = Color.Transparent; // 恢復其他面板背景色
            panelPie.BackColor = Color.Transparent; // 恢復其他面板背景色
            panelToHighlight.BackColor = Color.Brown; // 高亮選中的面板

        }



        private void buttonReset_Click(object sender, EventArgs e)
        {
            List<string> checkedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> checkedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            InitTrendChart();
            RefreshCharts_salesPage(checkedProducts, checkedCustomers);

        }

        private void chartSalesProducts_MouseClick(object sender, MouseEventArgs e)
        {
            var hit = chartSalesProducts.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                List<string> selectedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
                List<string> selectedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
                string product = chartSalesProducts.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadSalesProductsTrend(product, username);
                highlightedPieIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedBarIndex = -1;
                highlightedRowIndex = -1;
                HighlightPanel(panelPie);

                LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(username, selectedProducts, selectedCustomers);
                chartSalesOrdersTrend.Visible = false;
                chartSalesProductTrend.Visible = true;
                chartSalesRevenueTrend.Visible = false;
            }
        }

        public void LoadSalesProductsTrend(string productName, string username)
        {

            chartSalesProductTrend.Series.Clear();
            chartSalesProductTrend.Titles.Clear();
            chartSalesProductTrend.Titles.Add($"{productName} 銷售趨勢圖");
            chartSalesProductTrend.Titles[0].Font = new Font("微軟正黑體", 13, FontStyle.Bold);
            chartSalesProductTrend.Titles[0].ForeColor = Color.DarkRed; // 設定標題字體顏色
            chartSalesProductTrend.Series.Add("ProductsCount");
            chartSalesProductTrend.Series["ProductsCount"].ChartType = SeriesChartType.Line;
            chartSalesProductTrend.Series["ProductsCount"].IsValueShownAsLabel = true;
            chartSalesProductTrend.Series["ProductsCount"].Font = new Font("微軟正黑體", 9);
            chartSalesProductTrend.Series["ProductsCount"].Color = Color.RoyalBlue;

            chartSalesProductTrend.Series["ProductsCount"].BorderWidth = 2;  // 線條粗一點比較明顯
            chartSalesProductTrend.Series["ProductsCount"].MarkerStyle = MarkerStyle.Circle; // 點上加圓圈

            // 補齊月份（假設2024-12起到現在）
            var months = GetAllMonths("2024-12", DateTime.Today.ToString("yyyy-MM"));

            // 查詢
            Dictionary<string, int> monthData = new Dictionary<string, int>();
            string query = @"
                SELECT DATE_FORMAT(o.order_date, '%Y-%m') as month, IFNULL(SUM(o.amount),0) as products_count
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE p.name = @productName AND u.username = @username
                GROUP BY month
                ORDER BY month;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@productName", productName);
                cmd.Parameters.AddWithValue("@username", username);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string month = reader.GetString(0);
                        int productsCount = reader.IsDBNull(1) ? 0 : Convert.ToInt32(reader.GetDecimal(1));
                        monthData[month] = productsCount;
                    }
                }
            }
            // 補0
            foreach (var m in months)
                chartSalesProductTrend.Series["ProductsCount"].Points.AddXY(m, monthData.ContainsKey(m) ? monthData[m] : 0);
        }

        private List<string> GetAllMonths(string start, string end)
        {
            var result = new List<string>();
            var startDt = DateTime.Parse(start + "-01");
            var endDt = DateTime.Parse(end + "-01");
            while (startDt <= endDt)
            {
                result.Add(startDt.ToString("yyyy-MM"));
                startDt = startDt.AddMonths(1);
            }
            return result;
        }

        // tabPage3, 趨勢圖表 -- 點擊客戶訂單跳出訂單趨勢圖
        private void chartSalesCustomersOrders_MouseClick(object sender, MouseEventArgs e)
        {
            List<string> selectedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            var hit = chartSalesCustomersOrders.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                string customerName = chartSalesCustomersOrders.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadCustomerOrderTrend(customerName, username);
                highlightedBarIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedRowIndex = -1;
                highlightedPieIndex = -1;
                HighlightPanel(panelBar);


                LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(username, selectedProducts, selectedCustomers);
                chartSalesProductTrend.Visible = false;
                chartSalesOrdersTrend.Visible = true;
                chartSalesRevenueTrend.Visible = false;
            }
        }


        public void LoadCustomerOrderTrend(string customerName, string username)
        {
            chartSalesOrdersTrend.Series.Clear();
            chartSalesOrdersTrend.Titles.Clear();
            chartSalesOrdersTrend.Titles.Add($"{customerName} 訂單趨勢圖");
            chartSalesOrdersTrend.Titles[0].Font = new Font("微軟正黑體", 13, FontStyle.Bold);
            chartSalesOrdersTrend.Titles[0].ForeColor = Color.DarkRed; // 設定標題字體顏色
            chartSalesOrdersTrend.Series.Add("OrdersCount");
            chartSalesOrdersTrend.Series["OrdersCount"].ChartType = SeriesChartType.Line;
            chartSalesOrdersTrend.Series["OrdersCount"].IsValueShownAsLabel = true;
            chartSalesOrdersTrend.Series["OrdersCount"].Font = new Font("微軟正黑體", 9);
            chartSalesOrdersTrend.Series["OrdersCount"].Color = Color.RoyalBlue;
            var months = GetAllMonths("2024-12", DateTime.Today.ToString("yyyy-MM"));

            chartSalesOrdersTrend.Series["OrdersCount"].BorderWidth = 2;  // 線條粗一點比較明顯
            chartSalesOrdersTrend.Series["OrdersCount"].MarkerStyle = MarkerStyle.Circle; // 點上加圓圈
            Dictionary<string, int> monthData = new Dictionary<string, int>();

            string query = @"
                SELECT DATE_FORMAT(o.order_date, '%Y-%m') as month, 
                ifnull(COUNT(*),0) as order_count
                FROM orders o
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE c.name = @customerName AND u.username = @username
                GROUP BY month
                ORDER BY month;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@customerName", customerName);
                cmd.Parameters.AddWithValue("@username", username);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string month = reader.GetString(0);
                        int orderCount = reader.GetInt32(1);
                        monthData[month] = orderCount;
                    }
                }
            }
            foreach (var m in months)
                chartSalesOrdersTrend.Series["OrdersCount"].Points.AddXY(m, monthData.ContainsKey(m) ? monthData[m] : 0);
        }

        // tabPage3, 趨勢圖表 -- 點擊產品營收跳出趨勢圖

        private void chartMonthlyRevenuePerProduct_MouseClick(object sender, MouseEventArgs e)
        {
            List<string> selectedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            var hit = chartMonthlyRevenuePerProduct.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                string customerProduct = chartMonthlyRevenuePerProduct.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadRevenueTrend(customerProduct, username);
                highlightedRowIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedBarIndex = -1;
                highlightedPieIndex = -1;
                HighlightPanel(panelColumn);

                LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(username, selectedProducts, selectedCustomers);
                chartSalesOrdersTrend.Visible = false;
                chartSalesProductTrend.Visible = false;
                chartSalesRevenueTrend.Visible = true;
            }
        }

        private void LoadRevenueTrend(string customerProduct, string username)
        {
            chartSalesRevenueTrend.Series.Clear();
            chartSalesRevenueTrend.Titles.Clear();
            chartSalesRevenueTrend.Titles.Add($"{customerProduct} 營收趨勢圖");
            chartSalesRevenueTrend.Titles[0].Font = new Font("微軟正黑體", 13, FontStyle.Bold);
            chartSalesRevenueTrend.Titles[0].ForeColor = Color.DarkRed; // 設定標題字體顏色
            chartSalesRevenueTrend.Series.Add("Revenue");
            chartSalesRevenueTrend.Series["Revenue"].ChartType = SeriesChartType.Line;
            chartSalesRevenueTrend.Series["Revenue"].IsValueShownAsLabel = true;
            chartSalesRevenueTrend.Series["Revenue"].Font = new Font("微軟正黑體", 9);
            chartSalesRevenueTrend.Series["Revenue"].Color = Color.RoyalBlue;

            chartSalesRevenueTrend.Series["Revenue"].BorderWidth = 2;  // 線條粗一點比較明顯
            chartSalesRevenueTrend.Series["Revenue"].MarkerStyle = MarkerStyle.Circle; // 點上加圓圈

            string[] parts = customerProduct.Split(new[] { " - " }, StringSplitOptions.None);
            string customerName = parts[0].Trim();
            string productName = parts[1].Trim();

            var months = GetAllMonths("2024-12", DateTime.Today.ToString("yyyy-MM"));
            Dictionary<string, decimal> monthData = new Dictionary<string, decimal>();

            string query = $@"
                SELECT DATE_FORMAT(o.order_date, '%Y-%m') as month, 
                IFNULL(SUM(o.amount * p.price),0) as revenue
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE c.name = @customerName 
                  AND p.name = @productName
                  and u.username = @username
                GROUP BY month
                ORDER BY month;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@customerName", customerName);
                cmd.Parameters.AddWithValue("@productName", productName);
                cmd.Parameters.AddWithValue("@username", username);
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string month = reader.GetString(0);
                        decimal revenueValue = reader.IsDBNull(1) ? 0 : reader.GetDecimal(1);
                        monthData[month] = revenueValue;
                    }
                }
            }
            foreach (var m in months)
            {
                decimal value = monthData.ContainsKey(m) ? monthData[m] : 0;
                int idx = chartSalesRevenueTrend.Series["Revenue"].Points.AddXY(m, value);
                // 設定標籤格式為 "$10,000"
                chartSalesRevenueTrend.Series["Revenue"].Points[idx].Label = $"${value:N0}";
            }

        }





        // tabPage3, 中間的兩個panel -- 計算成長率(本月和上個月比較)
        private void UpdateSalesMonthlyGrowthLabel(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DateTime selectedMonth = dateTimePickerChart.Value.Date;
            DateTime prevMonth = selectedMonth.AddMonths(-1);

            // 取得營收
            decimal currentMonthRevenue = GetRevenueForMonth_salesPage(username, selectedMonth, selectedProducts, selectedCustomers);
            decimal prevMonthRevenue = GetRevenueForMonth_salesPage(username, prevMonth, selectedProducts, selectedCustomers);

            // 如果prevMonthRevenue為0（表示第一個月或前一月沒資料），直接顯示 "-"
            if (prevMonthRevenue == 0)
            {
                labelSalesRevenueGrowth.Text = "營收月成長率：-";
                pictureBoxSalesRevenueArrow.Image = null;
                labelSalesRevenueGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal revenueGrowthRate = (currentMonthRevenue - prevMonthRevenue) / prevMonthRevenue;
                labelSalesRevenueGrowth.Text = $"營收月成長率：{Math.Abs(revenueGrowthRate):P1}";
                pictureBoxSalesRevenueArrow.Image = revenueGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelSalesRevenueGrowth.ForeColor = revenueGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelSalesRevenueLastMonth.Text = $"上月營收：${prevMonthRevenue:N0}"; // 顯示上月營收
            labelSalesRevenueCurrentMonth.Text = $"本月營收：${currentMonthRevenue:N0}"; // 顯示本月營收

            // 訂單月成長率
            int currentMonthOrders = GetOrderCountForMonth_salesPage(username, selectedMonth, selectedProducts, selectedCustomers);
            int prevMonthOrders = GetOrderCountForMonth_salesPage(username, prevMonth, selectedProducts, selectedCustomers);

            if (prevMonthOrders == 0)
            {
                labelSalesOrderGrowth.Text = "訂單月成長率：-";
                pictureBoxSalesOrdersArrow.Image = null;
                labelSalesOrderGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal orderGrowthRate = (currentMonthOrders - prevMonthOrders) / (decimal)prevMonthOrders;
                labelSalesOrderGrowth.Text = $"訂單月成長率：{Math.Abs(orderGrowthRate):P1}";
                pictureBoxSalesOrdersArrow.Image = orderGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelSalesOrderGrowth.ForeColor = orderGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelSalesOrderLastMonth.Text = $"上月訂單量：{prevMonthOrders} 筆"; // 顯示上月訂單量
            labelSalesOrderCurrentMonth.Text = $"本月訂單量：{currentMonthOrders} 筆"; // 顯示本月訂單量
        }


        private decimal GetRevenueForMonth_salesPage(string username, DateTime month, List<string> selectedProducts, List<string> selectedCustomers)
        {
            decimal revenue = 0;
            string productFilter = selectedProducts.Count > 0
                ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")"
                : "";
            string customerFilter = selectedCustomers.Count > 0
                ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")"
                : "";

            string query = $@"
                SELECT SUM(o.amount * p.price) as total_revenue
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE YEAR(o.order_date) = @year
                  AND MONTH(o.order_date) = @month
                  AND u.username = @username
                  {productFilter}
                  {customerFilter}";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@year", month.Year);
                cmd.Parameters.AddWithValue("@month", month.Month);

                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);

                var result = cmd.ExecuteScalar();
                revenue = result != DBNull.Value && result != null ? Convert.ToDecimal(result) : 0m;
            }
            return revenue;
        }



        private int GetOrderCountForMonth_salesPage(string username, DateTime month, List<string> selectedProducts, List<string> selectedCustomers)
        {
            int count = 0;
            string productFilter = selectedProducts.Count > 0
                ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")"
                : "";
            string customerFilter = selectedCustomers.Count > 0
                ? "AND c.name IN (" + string.Join(",", selectedCustomers.Select((_, i) => $"@c{i}")) + ")"
                : "";

            string query = $@"
                SELECT COUNT(*) as order_count
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE YEAR(o.order_date) = @year
                  AND MONTH(o.order_date) = @month
                  AND u.username = @username
                  {productFilter}
                  {customerFilter}";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@year", month.Year);
                cmd.Parameters.AddWithValue("@month", month.Month);

                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedCustomers.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedCustomers[i]);

                var result = cmd.ExecuteScalar();
                count = result != DBNull.Value && result != null ? Convert.ToInt32(result) : 0;
            }
            return count;
        }

        // tabPage3, 中間的panel -- 負責動態顯示checkListBox所勾選的內容
        private void ShowSalesCurrentFilters(List<string> products, List<string> customers)
        {
            string customerText = customers.Count == 0 ? "全部" :
                (customers.Count > 2 ? string.Join(", ", customers.Take(2)) + $"... (共{customers.Count}位)" : string.Join(" , ", customers));
            string customerFullText = customers.Count == 0 ? "全部" : string.Join(" , ", customers);

            labelSalesRevenueCustomer.Text = "客戶: " + customerText;
            toolTip1.SetToolTip(labelSalesRevenueCustomer, "客戶：" + customerFullText);
            labelSalesOrderCustomer.Text = "客戶: " + customerText;
            toolTip1.SetToolTip(labelSalesOrderCustomer, "客戶：" + customerFullText);

            string productText = products.Count == 0 ? "全部" :
                (products.Count > 2 ? string.Join(", ", products.Take(2)) + $"... (共{products.Count}種)" : string.Join(" , ", products));
            string productFullText = products.Count == 0 ? "全部" : string.Join(" , ", products);

            labelSalesRevenueProduct.Text = "商品: " + productText;
            toolTip1.SetToolTip(labelSalesRevenueProduct, "商品：" + productFullText);
            labelSalesOrderProduct.Text = "商品: " + productText;
            toolTip1.SetToolTip(labelSalesOrderProduct, "商品：" + productFullText);
        }



        // dateTimePicker 時間改的話就呼叫
        private void dateTimePickerChart_ValueChanged(object sender, EventArgs e)
        {
            // 取得當前選擇的年月
            //int year = dateTimePickerChart.Value.Year;
            //int month = dateTimePickerChart.Value.Month;
            // 取得當前已勾選的產品、客戶
            List<string> checkedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> checkedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();

            // 三個圖表都重畫（依你的需求）
            RefreshCharts_salesPage(checkedProducts, checkedCustomers);

            // 若有顯示本月總覽 Label，也可以重抓
            //LoadMonthlyTotalLabel(_username);
            UpdateSalesMonthlyGrowthLabel(username, checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(username, checkedProducts, checkedCustomers);
        }

        // 若有勾選，就加入List
        private void checkedListBoxProducts_sales_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerChart.Value.Year;
            int month = dateTimePickerChart.Value.Month;
            // 預測本次操作後的產品勾選清單
            List<string> checkedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            string currentProduct = checkedListBoxProducts_sales.Items[e.Index].ToString();
            if (e.NewValue == CheckState.Checked)
            {
                if (!checkedProducts.Contains(currentProduct))
                    checkedProducts.Add(currentProduct);
            }
            else
            {
                if (checkedProducts.Contains(currentProduct))
                    checkedProducts.Remove(currentProduct);
            }

            // 客戶用目前 CheckedItems，不用加 e，因為客戶沒變
            List<string> checkedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();

            // 傳進查詢方法
            RefreshCharts_salesPage(checkedProducts, checkedCustomers);
            UpdateSalesMonthlyGrowthLabel(username, checkedProducts, checkedCustomers);
            ShowSalesCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(username, checkedProducts, checkedCustomers);
        }

        private void checkedListBoxCustomers_sales_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerChart.Value.Year;
            int month = dateTimePickerChart.Value.Month;
            List<string> checkedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            string currentCustomer = checkedListBoxCustomers_sales.Items[e.Index].ToString();
            if (e.NewValue == CheckState.Checked)
            {
                if (!checkedCustomers.Contains(currentCustomer))
                    checkedCustomers.Add(currentCustomer);
            }
            else
            {
                if (checkedCustomers.Contains(currentCustomer))
                    checkedCustomers.Remove(currentCustomer);
            }
            // 產品用目前 CheckedItems
            List<string> checkedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();

            RefreshCharts_salesPage(checkedProducts, checkedCustomers);
            UpdateSalesMonthlyGrowthLabel(username, checkedProducts, checkedCustomers);
            ShowSalesCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(username, checkedProducts, checkedCustomers);
        }

        // 統一重畫圖表
        private void RefreshCharts_salesPage(List<string> checkedProducts, List<string> checkedCustomers)
        {
            highlightedPieIndex = -1;
            highlightedBarIndex = -1;
            highlightedRowIndex = -1;
            InitTrendChart();
            panelBar.BackColor = Color.Transparent; // 恢復其他面板背景色
            panelColumn.BackColor = Color.Transparent; // 恢復其他面板背景色
            panelPie.BackColor = Color.Transparent; // 恢復其他面板背景色
            LoadMonthlyProductsChart(username, checkedProducts, checkedCustomers);
            LoadMonthlyOrdersChartByCustomer(username, checkedProducts, checkedCustomers);
            LoadMonthlyRevenueDetailsChart(username, checkedProducts, checkedCustomers);
        }
    }
}
