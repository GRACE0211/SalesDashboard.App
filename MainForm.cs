

using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ZstdSharp.Unsafe;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace SalesDashboard
{
    public partial class MainForm : Form
    {
        private string _username;
        private string _role;
        private string connectionString = "server=localhost;user=root;password=0905656870;database=sales_dashboard;port=3306;SslMode=None";
        private int highlightedPieIndex = -1; // 用於記錄高亮顯示的行索引
        private int highlightedBarIndex = -1; // 用於記錄高亮顯示的行索引
        private int highlightedRowIndex = -1; // 用於記錄高亮顯示的行索引
        //private ToolTip toolTip2 = new ToolTip();


        // 登出標誌，預設為 false
        // { get;  private set; }: 這個屬性只能在 MainForm 類別內部被設置，外部只能讀取它的值
        public bool LogoutFlag { get; private set; } = false;
        public MainForm(string username, string role)
        {
            InitializeComponent();
            this.LogoutFlag = false; // 初始化登出標誌為 false
            _username = username;
            _role = role;
            ApplyFilter_dataMgmtPage();
            dateTimePickerSearch_AdminStart.Value = new DateTime(2024, 12, 1);
            dateTimePickerSearch_AdminEnd.Value = DateTime.Now;
            dateTimePickerSearch_SalesStart.Value = new DateTime(2024, 12, 1);
            dateTimePickerSearch_SalesEnd.Value = DateTime.Now;
            dateTimePickerChart.Value = new DateTime(2024, 12, 1);
            dateTimePickerChart.MaxDate = DateTime.Now; // 設定日期選擇器的最大日期為今天
            tabControl1.TabPages.Remove(dataMgmtPage);
            dateTimePickerAdmin_region.Value = new DateTime(2024, 12, 1);


            if (_role == "sales")
            {
                tabControl1.TabPages.Remove(AdminSearchPage); // 移除 AdminSearchPage (原本就存在SalesSearchPage所以不用Add)
                tabControl1.TabPages.Remove(AdminChartPage_SalesNProduct); // 移除 AdminChartPage_SalesNProduct
                tabControl1.TabPages.Remove(AdminChartPage_Ttl); // 移除 AdminChartPage_Ttl
                tabControl1.TabPages.Remove(AdminChartPage_region); // 移除 AdminChartPage_region
            }
            else if (_role == "admin")
            {
                tabControl1.TabPages.Remove(SalesSearchPage); // 移除 SalesSearchPage (原本就存在AdminSearchPage所以不用Add)
                tabControl1.TabPages.Remove(SalesChartPage_Monthly); // 移除 SalesChartPage_Monthly (原本就存在SalesChartPage_Ttl所以不用Add)
                tabControl1.TabPages.Remove(SalesChartPage_Ttl); // 移除 SalesChartPage_Ttl (原本就存在SalesSearchPage所以不用Add)
            }
            // 右上角顯示目前登入者
            lblCurrentUser.Text = $"目前登入: {_username}";

            // 初始化產品下拉選單
            comboBoxProducts.Items.Clear();
            comboBoxProducts.Items.Add("全部");
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                string sql = "SELECT name FROM products ORDER BY name;";
                using (MySqlCommand cmd = new MySqlCommand(sql, connection))
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    // 抓出下一列資料
                    while (reader.Read())
                    {
                        comboBoxProducts.Items.Add(reader["name"].ToString());
                    }
                }
            }
            comboBoxProducts.SelectedIndex = 0; // 預設選擇 "全部"


            // 初始化客戶下拉選單

            comboBoxCustomers.Items.Clear();
            comboBoxCustomers.Items.Add("全部");
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                string sql = "SELECT name FROM customers ORDER BY name;";
                using (MySqlCommand cmd = new MySqlCommand(sql, connection))
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        comboBoxCustomers.Items.Add(reader["name"].ToString());
                    }
                }
            }
            comboBoxCustomers.SelectedIndex = 0; // 預設選擇 "全部"


        }

        // 判斷選到哪個tabPage -> 當選擇的頁面改變時，根據選擇的頁面載入相應的內容
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> selectedProducts_sales = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers_sales.CheckedItems.Cast<string>().ToList();
            List<string> selectedProducts_admin = checkedListBoxProducts_admin_monthlyPage.CheckedItems.Cast<string>().ToList();
            List<string> selectedRegion_admin = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();
            List<string> selectedRProducts_admin = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();
            List<string> selectedRRegion_admin = checkedListBoxRegion_region_admin.CheckedItems.Cast<string>().ToList();

            // 當選擇的頁面改變時，根據選擇的頁面載入相應的客戶下拉選單

            if (tabControl1.SelectedTab == dataMgmtPage)
            {
                ApplyFilter_dataMgmtPage();
            }
            else if (tabControl1.SelectedTab == AdminSearchPage)
            {
                ApplyFilter_adminSearchPage();
            }
            else if (tabControl1.SelectedTab == SalesSearchPage)
            {
                ApplyFilter_salesSearchPage();
            }
            else if (tabControl1.SelectedTab == SalesChartPage_Monthly)
            {
                // 銷售人員
                // 月份篩選

                LoadMonthlyProductsChart(_username, selectedProducts_sales, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(_username, selectedProducts_sales, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(_username, selectedProducts_sales, selectedCustomers);
                InitTrendChart();
                //LoadMonthlyTotalLabel(_username);

            }
            else if (tabControl1.SelectedTab == SalesChartPage_Ttl)
            {
                // 銷售人員
                // 總計
                //LoadKPIDoughnutChart(_username);
                LoadSalesTtlKPI(_username);
                LoadProductsCountTopThree(_username);
                LoadCustomerOrdersCountTopThree(_username);
                LoadTotalOrdersCountChart(_username);
                LoadTotalRevenueTtlChart(_username);
                LoadBiggestLabel(_username);
                LoadProductsGroupBySales(_username);
                LoadCustomersGroupBySales(_username);

            }
            else if (tabControl1.SelectedTab == AdminChartPage_SalesNProduct)
            {
                // 管理員
                // 月份篩選
                LoadAdminDataGridView();
                LoadAdminGetMonthlySalesRevenue();
                LoadMonthlyProductsSelling();
                LoadAdminMonthlyLabel();
                UpdateAdminMonthlyGrowthLabel_Sales();
                LoadAdminSalesMonthlyOrdersChart();
                UpdateAdminMonthlyGrowthLabel_Product(selectedProducts_admin);
                LoadAdminMonthlyKPI();
                LoadAdminSalesMonthlyRevenueLine();
                LoadAdminTtlRevenueChart_sales();
                LoadAdminTtlRevenueChart_product();

                LoadProductRevenueLineChart(selectedProducts_admin);
            }
            else if (tabControl1.SelectedTab == AdminChartPage_Ttl)
            {
                // 管理員
                // 總計
                LoadAdminBiggestLabel();
                LoadAdminTtlKPI();
                LoadAdminTtlSalesRevenuePie();
                LoadAdminProductRevenuePie();
                LoadAdminQuarterRevenueChart();
                LoadAdminSalesMonthlyRevenuePie();
                LoadAdminTtlRevenueChart();
            }
            else if (tabControl1.SelectedTab == AdminChartPage_region)
            {
                LoadCustomerRegionMthRadarChart(selectedRegion_admin);
                LoadCustomerRegionTtlRadarChart();
                LoadMonthlyRevenueByRegionChart();
                LoadMonthlyTop3RevenueByRegionChart(selectedRProducts_admin, selectedRRegion_admin);
                LoadAdminRegionTtlLabel();
                LoadRegionTtlRevenuePctChart();
            }

        }




        // 按登出鍵
        private void LogoutButton_Click(object sender, EventArgs e)
        {
            this.LogoutFlag = true; // 設置登出標誌
            this.Close(); // 關閉主畫面
        }

        private void ApplyFilter_dataMgmtPage()
        {
            string baseQuery = @"select
                o.id,p.name as product_name,c.name as customer_name,o.amount,DATE(o.order_date) as order_date, u.username
                from orders o 
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id"; // 基礎查詢語句
            string query;
            if (_role == "sales")
            {
                query = baseQuery + " where u.username = @username;"; // 如果是銷售人員，則只顯示自己的訂單
            }
            else
            {
                query = baseQuery + ";"; // 如果是管理員，則顯示所有訂單
            }
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                if (_role == "sales")
                {
                    cmd.Parameters.AddWithValue("@username", _username); // 添加參數以防止 SQL 注入
                }
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView1.DataSource = dataTable; // 將查詢結果綁定到 DataGridView
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 自動調整列寬

            }
        }

        /*
         ------------------------------------ SalesSearchPage -- tabPage1 ------------------------------------
         */
        private void ApplyFilter_salesSearchPage()
        {
            // 客戶條件
            List<string> customers = new List<string>();
            foreach (var item in checkedListBoxCustomer_salesSearchPage.CheckedItems)
            {
                customers.Add($"'{item.ToString()}'");
            }

            // 商品條件
            List<string> products = new List<string>();
            foreach (var item in checkedListBoxProduct_salesSearchPage.CheckedItems)
            {
                products.Add($"'{item.ToString()}'");
            }

            // 時間與 SQL 查詢條件組合
            DateTime startDate = dateTimePickerSearch_SalesStart.Value.Date;
            DateTime endDate = dateTimePickerSearch_SalesEnd.Value.Date;
            string dateFilter = "Date(o.order_date) BETWEEN @startDate AND @endDate";
            string loginUser = "u.username = @username";
            string customerFilter = customers.Count > 0 ? $"c.name IN ({string.Join(",", customers)})" : "1=1";
            string productFilter = products.Count > 0 ? $"p.name IN ({string.Join(",", products)})" : "1=1";

            string query = $@"
                SELECT 
                    p.name as product_name,
                    c.name as customer_name,
                    DATE(o.order_date) as order_date,
                    sum(o.amount) as total_amount,
                    CONCAT('$',sum(p.price*o.amount)) as revenue,
                    u.username as sales_name
                FROM orders o 
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE {loginUser} AND {customerFilter} AND {productFilter} AND {dateFilter}
                GROUP BY o.product_id, o.customer_id, o.sales_user_id, o.order_date 
                ORDER BY o.customer_id, o.sales_user_id;";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", _username);
                cmd.Parameters.AddWithValue("@startDate", startDate);
                cmd.Parameters.AddWithValue("@endDate", endDate);

                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView3.DataSource = dataTable;
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }



        private void checkedListBoxCustomer_salesSearchPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 因為 ItemCheck 還沒更新 CheckedItems，所以要用 BeginInvoke 確保資料更新後再跑篩選
            this.BeginInvoke((MethodInvoker)delegate {
                ApplyFilter_salesSearchPage();
            });
        }

        private void checkedListBoxProduct_salesSearchPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                ApplyFilter_salesSearchPage();
            });
        }

        private void dateTimePickerSearch_SalesStart_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void dateTimePickerSearch_SalesEnd_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }


        // ------------------ AdminSearchPage -- tabPage2 ------------------


        private void ApplyFilter_adminSearchPage()
        {
            List<string> region = new List<string>();
            foreach (var item in checkedListBoxRegion_adminSearchPage.CheckedItems)
            {
                region.Add($"'{item.ToString()}'");
            }

            List<string> products = new List<string>();
            foreach (var item in checkedListBoxProduct_adminSearchPage.CheckedItems)
            {
                products.Add($"'{item.ToString()}'");
            }

            List<string> salesUsers = new List<string>();
            foreach (var item in checkedListBoxSales_adminSearchPage.CheckedItems)
            {
                salesUsers.Add($"'{item.ToString()}'");
            }

            DateTime startDate = dateTimePickerSearch_AdminStart.Value.Date;
            DateTime endDate = dateTimePickerSearch_AdminEnd.Value.Date;
            string dateFilter = "Date(o.order_date) BETWEEN  @startDate AND @endDate";
            string salesUserFilter = salesUsers.Count > 0 ? $"u.username IN ({string.Join(",", salesUsers)})" : "1=1";
            string regionFilter = region.Count > 0 ? $" SUBSTRING_INDEX(c.address,'市',1) IN ({string.Join(",", region)})" : "1=1";
            string productFilter = products.Count > 0 ? $"p.name IN ({string.Join(",", products)})" : "1=1";
            string query = $@"SELECT 
                p.name as product_name,
                SUBSTRING_INDEX(c.address,'市',1) as region,
                DATE(o.order_date) as order_date,
                sum(o.amount) as total_amount,
                CONCAT('$',sum(p.price*o.amount)) as revenue,
                u.username as sales_name
                FROM orders o 
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where {salesUserFilter} AND {regionFilter} AND {productFilter} AND {dateFilter}
                group by o.product_id,o.customer_id,o.sales_user_id,o.order_date 
                order by o.customer_id,o.sales_user_id;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@startDate", startDate);
                cmd.Parameters.AddWithValue("@endDate", endDate);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView2.DataSource = dataTable; // 將查詢結果綁定到 DataGridView
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 自動調整列寬
            }
        }

        private void checkedListBoxRegion_adminSearchPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                ApplyFilter_adminSearchPage();
            });
        }

        private void checkedListBoxProduct_adminSearchPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                ApplyFilter_adminSearchPage();
            });
        }

        private void checkedListBoxSales_adminSearchPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate {
                ApplyFilter_adminSearchPage();
            });
        }
        private void dateTimePickerSearch_AdminStart_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void dateTimePickerSearch_AdminEnd_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
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

        //private void LoadMonthlyTotalLabel(string username)
        //{
        //    DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
        //    DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份
        //    string query = $@"
        //        select 
        //            sum(o.amount*p.price) as ttl_revenue,
        //            count(*) as orders_cnt,
        //            o.sales_user_id,
        //            date_format(o.order_date,'%Y-%m') as o_date
        //        from orders o
        //        left join products p on o.product_id = p.product_id
        //        left join users u on o.sales_user_id = u.user_id
        //        where 
        //            year(o.order_date) = @searchYear 
        //            and month(o.order_date) = @searchMonth
        //            and u.username = @username
        //        group by o.sales_user_id,o_date;";

        //    decimal monthlyRevenue = 0;
        //    int monthlyOrders = 0;
        //    using (MySqlConnection connection = new MySqlConnection(connectionString))
        //    {
        //        connection.Open();
        //        MySqlCommand cmd = new MySqlCommand(query, connection);
        //        cmd.Parameters.AddWithValue("@username", username);
        //        cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
        //        cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數

        //        using (var reader = cmd.ExecuteReader())
        //        {
        //            if (reader.Read())
        //            {
        //                monthlyRevenue = reader["ttl_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["ttl_revenue"]) : 0m;
        //                monthlyOrders = reader["orders_cnt"] != DBNull.Value ? Convert.ToInt32(reader["orders_cnt"]) : 0;
        //            }


        //        }
        //    }

        //    labelMonthlyRevenue.Text = $"{dateTimePickerChart.Value.ToString("yyyy年MM月")}總營業額 - ${monthlyRevenue:N0}";
        //    labelMonthlyOrders.Text = $"{dateTimePickerChart.Value.ToString("yyyy年MM月")}總訂單量 - {monthlyOrders} 筆";
        //}

        // 




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
            chartSalesOrdersTrend.Series.Clear();
            chartSalesOrdersTrend.Titles.Clear();
            chartSalesOrdersTrend.Titles.Add("請點選任意表格的標籤查看趨勢圖");
            chartSalesOrdersTrend.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold);
            chartSalesOrdersTrend.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
            chartSalesOrdersTrend.Visible = true;
            chartSalesProductTrend.Series.Clear();
            chartSalesProductTrend.Titles.Clear();
            chartSalesProductTrend.Titles.Add("請點選任意表格的標籤查看趨勢圖");
            chartSalesProductTrend.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold);
            chartSalesProductTrend.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
            chartSalesProductTrend.Visible = true;
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
                LoadSalesProductsTrend(product, _username);
                highlightedPieIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedBarIndex = -1;
                highlightedRowIndex = -1;
                HighlightPanel(panelPie);

                LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(_username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(_username, selectedProducts, selectedCustomers);
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
                LoadCustomerOrderTrend(customerName, _username);
                highlightedBarIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedRowIndex = -1;
                highlightedPieIndex = -1;
                HighlightPanel(panelBar);
                

                LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(_username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(_username, selectedProducts, selectedCustomers);
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
                LoadRevenueTrend(customerProduct, _username);
                highlightedRowIndex = hit.PointIndex; // 記錄被點到的 index
                highlightedBarIndex = -1;
                highlightedPieIndex = -1;
                HighlightPanel(panelColumn);
                
                LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers); LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(_username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(_username, selectedProducts, selectedCustomers);
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

        

        // MainForm_Load 設定預設內容, for panel 的 ToolTip, 滑鼠移入會顯示完整內容
        private void MainForm_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBoxProduct_salesSearchPage.Items.Count; i++)
            {
                checkedListBoxProduct_salesSearchPage.SetItemChecked(i, true); // 預設全選產品
            }
            for (int i = 0; i < checkedListBoxCustomer_salesSearchPage.Items.Count; i++)
            {
                checkedListBoxCustomer_salesSearchPage.SetItemChecked(i, true); // 預設全選產品
            }
            for (int i = 0; i < checkedListBoxSales_adminSearchPage.Items.Count; i++)
            {
                checkedListBoxSales_adminSearchPage.SetItemChecked(i, true); // 預設全選產品
            }
            for (int i = 0; i < checkedListBoxProduct_adminSearchPage.Items.Count; i++)
            {
                checkedListBoxProduct_adminSearchPage.SetItemChecked(i, true); // 預設全選產品
            }
            for (int i = 0; i < checkedListBoxRegion_adminSearchPage.Items.Count; i++)
            {
                checkedListBoxRegion_adminSearchPage.SetItemChecked(i, true); // 預設全選產品
            }

            toolTip2.SetToolTip(labelSalesOrderCustomer, "客戶：全部");
            toolTip2.SetToolTip(labelSalesOrderProduct, "商品：全部");
            toolTip2.SetToolTip(labelSalesRevenueCustomer, "客戶：全部");
            toolTip2.SetToolTip(labelSalesRevenueProduct, "商品：全部");

            toolTipAdmin.SetToolTip(labelAdminProductOrder, "目前篩選商品：全部");
            toolTipAdmin.SetToolTip(labelAdminProductRevenue, "目前篩選商品：全部");

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

            // 初始化 checkedListBoxProducts_admin 為全選
            for (int i = 0; i < checkedListBoxProducts_admin_monthlyPage.Items.Count; i++)
            {
                checkedListBoxProducts_admin_monthlyPage.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedProducts_admin = checkedListBoxProducts_admin_monthlyPage.CheckedItems.Cast<string>().ToList();
            LoadProductRevenueLineChart(checkedProducts_admin);

            for (int i = 0; i < checkedListBoxRegion_product_admin.Items.Count; i++)
            {
                checkedListBoxRegion_product_admin.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedRegion_product = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();


            for (int i = 0; i < checkedListBoxRegion_region_admin.Items.Count; i++)
            {
                checkedListBoxRegion_region_admin.SetItemChecked(i, true); // 預設全選產品
            }
            var checkedRegion_region = checkedListBoxRegion_region_admin.CheckedItems.Cast<string>().ToList();

        }

        // tabPage3, 中間的panel -- 負責動態顯示checkListBox所勾選的內容
        private void ShowSalesCurrentFilters(List<string> products, List<string> customers)
        {
            string customerText = customers.Count == 0 ? "全部" :
                (customers.Count > 2 ? string.Join(", ", customers.Take(2)) + $"... (共{customers.Count}位)" : string.Join(" , ", customers));
            string customerFullText = customers.Count == 0 ? "全部" : string.Join(" , ", customers);

            labelSalesRevenueCustomer.Text = "客戶: " + customerText;
            toolTip2.SetToolTip(labelSalesRevenueCustomer, "客戶：" + customerFullText);
            labelSalesOrderCustomer.Text = "客戶: " + customerText;
            toolTip2.SetToolTip(labelSalesOrderCustomer, "客戶：" + customerFullText);

            string productText = products.Count == 0 ? "全部" :
                (products.Count > 2 ? string.Join(", ", products.Take(2)) + $"... (共{products.Count}種)" : string.Join(" , ", products));
            string productFullText = products.Count == 0 ? "全部" : string.Join(" , ", products);

            labelSalesRevenueProduct.Text = "商品: " + productText;
            toolTip2.SetToolTip(labelSalesRevenueProduct, "商品：" + productFullText);
            labelSalesOrderProduct.Text = "商品: " + productText;
            toolTip2.SetToolTip(labelSalesOrderProduct, "商品：" + productFullText);
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
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
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
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            ShowSalesCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
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
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            ShowSalesCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
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
            LoadMonthlyProductsChart(_username, checkedProducts, checkedCustomers);
            LoadMonthlyOrdersChartByCustomer(_username, checkedProducts, checkedCustomers);
            LoadMonthlyRevenueDetailsChart(_username, checkedProducts, checkedCustomers);
        }


        // ------------------ SalesChartPage_Ttl -- tabPage4 ------------------------------------


        // tabPage4, 左上兩個label以及兩個甜甜圈圖 -- 取得該銷售員總銷售收入
        //private void LoadKPIDoughnutChart(string username)
        //{
        //    // KPI 目標
        //    const decimal targetRevenue = 100000m; // 10萬元
        //    const int targetOrder = 50;

        //    // 一起查兩個 KPI
        //    string query = @"
        //        SELECT 
        //            IFNULL(SUM(o.amount * p.price), 0) AS total_revenue,
        //            COUNT(*) AS total_orders
        //        FROM orders o
        //        LEFT JOIN products p ON o.product_id = p.product_id
        //        WHERE o.sales_user_id = (SELECT user_id FROM users WHERE username = @username);";

        //    decimal totalRevenue = 0;
        //    int totalOrders = 0;

        //    using (MySqlConnection connection = new MySqlConnection(connectionString))
        //    {
        //        connection.Open();
        //        MySqlCommand cmd = new MySqlCommand(query, connection);
        //        cmd.Parameters.AddWithValue("@username", username);

        //        using (var reader = cmd.ExecuteReader())
        //        {
        //            if (reader.Read())
        //            {
        //                totalRevenue = reader["total_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["total_revenue"]) : 0m;
        //                totalOrders = reader["total_orders"] != DBNull.Value ? Convert.ToInt32(reader["total_orders"]) : 0;
        //            }
        //        }
        //    }

        //    // 設定 Label
        //    labelTtlRevenue.Text = $"${totalRevenue:N0}";
        //    labelTtlOrders.Text = $"{totalOrders} 筆";

        //    decimal revenueAchievementRate = targetRevenue == 0 ? 0 : totalRevenue / targetRevenue;
        //    string percentageText = $"{revenueAchievementRate:P0}"; // 百分比格式

        //    // --- Doughnut for Revenue ---
        //    chartTtlRevenueKPI.Series.Clear();
        //    var seriesRev = chartTtlRevenueKPI.Series.Add("RevenueKPI");
        //    seriesRev.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
        //    seriesRev.Points.AddXY($"{totalRevenue:P0}", totalRevenue);
        //    seriesRev.Points.AddXY("未達成", Math.Max(0, targetRevenue - totalRevenue));
        //    seriesRev.IsValueShownAsLabel = true;
        //    foreach (var pt in seriesRev.Points)
        //    {
        //        if (pt.AxisLabel == "未達成")
        //        {
        //            pt.Label = ""; // 空白區塊不顯示標籤
        //        }
        //        else
        //        {
        //            pt.Label = "#VALY"; // 顯示數值標籤，格式化為千分位
        //        }
        //    }
        //    seriesRev.Label = percentageText; // 設定標籤為百分比格式
        //    seriesRev["PieStartAngle"] = "270"; // 設定圓餅圖的起始角度為 270 度，這樣第一個區塊會從頂部開始
        //    seriesRev.Points[0].Color = Color.Orange; // 設定達成區塊的顏色為橙色
        //    seriesRev.Points[1].Color = Color.FromArgb(51, 47, 41); // 設定未達成區塊的顏色為淺灰色

        //    // 標題
        //    chartTtlRevenueKPI.Titles.Clear();
        //    chartTtlRevenueKPI.Titles.Add("銷售金額達成率");
        //    chartTtlRevenueKPI.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);

        //    decimal OrdersAchievementRate = targetOrder == 0 ? 0 : (decimal)totalOrders / targetOrder;
        //    string OrdersPercentageText = $"{OrdersAchievementRate:P0}"; // 百分比格式

        //    // --- Doughnut for Orders ---
        //    chartTtlOrdersKPI.Series.Clear();
        //    var seriesOrd = chartTtlOrdersKPI.Series.Add("OrderKPI");
        //    seriesOrd.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
        //    seriesOrd.Points.AddXY($"{totalOrders}", totalOrders);
        //    seriesOrd.Points.AddXY("未達成", Math.Max(0, targetOrder - totalOrders));
        //    seriesOrd.IsValueShownAsLabel = true;
        //    foreach (var pt in seriesOrd.Points)
        //    {
        //        if (pt.AxisLabel == "未達成")
        //        {
        //            pt.Label = ""; // 空白區塊不顯示標籤
        //        }
        //        else
        //        {
        //            pt.Label = "#VALY"; // 顯示數值標籤，格式化為千分位
        //        }
        //    }

        //    seriesOrd.Label = OrdersPercentageText; // 設定標籤為百分比格式
        //    seriesOrd["PieStartAngle"] = "270"; // 設定圓餅圖的起始角度為 270 度，這樣第一個區塊會從頂部開始
        //    seriesOrd.Points[0].Color = Color.Orange; // 設定達成區塊的顏色為橙色
        //    seriesOrd.Points[1].Color = Color.FromArgb(51, 47, 41);


        //    labelTtlRevenuePatio.Text = $"({totalRevenue:N0} / {targetRevenue:N0})";
        //    labelTtlOrdersPatio.Text = $"({totalOrders} / {targetOrder})";
        //    labelTtlRevenuePct.Text = $"{revenueAchievementRate:P0}"; // 更新 Label 顯示達成率
        //    labelTtlOrdersPct.Text = $"{OrdersAchievementRate:P0}"; // 更新 Label 顯示達成率
        //    // 標題
        //    chartTtlOrdersKPI.Titles.Clear();
        //    chartTtlOrdersKPI.Titles.Add("訂單數量達成率");
        //    chartTtlOrdersKPI.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
        //}

        // tabPage4, 左上兩個panel -- KPI進度條
        private void LoadSalesTtlKPI(string username)
        {
            const decimal targetRevenue = 100000m; // 10萬元
            const int targetOrder = 50;

            // 一起查兩個 KPI
            string query = @"
                SELECT 
                    IFNULL(SUM(o.amount * p.price), 0) AS total_revenue,
                    COUNT(*) AS total_orders
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                WHERE o.sales_user_id = (SELECT user_id FROM users WHERE username = @username);";

            decimal totalRevenue = 0;
            int totalOrders = 0;

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        totalRevenue = reader["total_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["total_revenue"]) : 0m;
                        totalOrders = reader["total_orders"] != DBNull.Value ? Convert.ToInt32(reader["total_orders"]) : 0;
                    }
                }
            }

            progressBarSalesTtlOrders.Minimum = 0;
            progressBarSalesTtlOrders.Maximum = 50;
            progressBarSalesTtlOrders.Value = totalOrders; // 更新訂單進度條

            progressBarSalesTtlRevenue.Minimum = 0;
            progressBarSalesTtlRevenue.Maximum = 100000;
            progressBarSalesTtlRevenue.Value = (int)(totalRevenue); // 更新銷售金額進度條

            decimal revenueAchievementRate = targetRevenue == 0 ? 0 : totalRevenue / targetRevenue;
            string percentageText = $"{revenueAchievementRate:P0}"; // 百分比格式

            decimal OrdersAchievementRate = targetOrder == 0 ? 0 : (decimal)totalOrders / targetOrder;
            string OrdersPercentageText = $"{OrdersAchievementRate:P0}"; // 百分比格式


            this.labelSalesTtlOrdersPct.BringToFront();
            this.labelSalesTtlRevenuePct.BringToFront();
            labelSalesTtlRevenue.Text = $"${totalRevenue:N0}";
            labelSalesTtlOrders.Text = $"{totalOrders} 筆";
            labelSalesTtlRevenuePct.Text = $"{revenueAchievementRate:P0}"; // 更新 Label 顯示達成率
            labelSalesTtlOrdersPct.Text = $"{OrdersAchievementRate:P0}"; // 更新 Label 顯示達成率
            labelSalesTtlRevenuePatio.Text = $"({totalRevenue:N0} / {targetRevenue:N0})";
            labelSalesTtlOrdersPatio.Text = $"({totalOrders} / {targetOrder})";
        }

        // tabPage4, 左中右圓餅圖 -- 業務自己的產品銷售占比
        private DataTable GetProductsGroupBySalesData(string username)
        {
            string query = $@"
                select sum(o.amount*p.price) as ttl_revenue,p.name
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where u.username = @username
                group by p.product_id;"; // 按銷售量降序排列
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadProductsGroupBySales(string username)
        {
            DataTable dt = GetProductsGroupBySalesData(username);

            // 先算總營收
            decimal totalRevenue = 0m;
            foreach (DataRow row in dt.Rows)
            {
                totalRevenue += Convert.ToDecimal(row["ttl_revenue"]);
            }

            chartProductPie.Titles.Clear();
            chartProductPie.Titles.Add("產品銷售占比");
            chartProductPie.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartProductPie.Series.Clear();

            var series = chartProductPie.Series.Add("Products");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);

            foreach (DataRow row in dt.Rows)
            {
                string product = row["name"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["ttl_revenue"]);
                double percent = totalRevenue == 0 ? 0 : (double)(revenueValue / totalRevenue) * 100;

                // 加進 pie chart
                int idx = series.Points.AddXY(product, revenueValue);

                // 圓餅圖標籤格式：50%
                series.Points[idx].Label = $"{percent:F1}%";

                // 圖例也可以加百分比
                series.Points[idx].LegendText = $"{product}";

                series.Points[idx].ToolTip = $"{product} - {percent:F1}% (${revenueValue:N0})"; // Tooltip 顯示詳細
            }
        }

        // tabPage4, 左中左圓餅圖 -- 業務自己的客戶銷售占比
        private DataTable GetCustomersGroupBySalesData(string username)
        {
            string query = $@"
                select sum(o.amount*p.price) as ttl_revenue,c.name
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where u.username = @username
                group by c.customer_id;"; // 按銷售量降序排列
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadCustomersGroupBySales(string username)
        {
            DataTable dt = GetCustomersGroupBySalesData(username);
            // 先算總營收
            decimal totalRevenue = 0m;
            foreach (DataRow row in dt.Rows)
            {
                totalRevenue += Convert.ToDecimal(row["ttl_revenue"]);
            }
            chartCustomerPie.Titles.Clear();
            chartCustomerPie.Titles.Add("客戶銷售占比");
            chartCustomerPie.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartCustomerPie.Series.Clear();
            var series = chartCustomerPie.Series.Add("Customers");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
            foreach (DataRow row in dt.Rows)
            {
                string customer = row["name"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["ttl_revenue"]);
                double percent = totalRevenue == 0 ? 0 : (double)(revenueValue / totalRevenue) * 100;
                // 加進 pie chart
                int idx = series.Points.AddXY(customer, revenueValue);
                // 圓餅圖標籤格式：50%
                series.Points[idx].Label = $"{percent:F1}%";
                // 圖例也可以加百分比
                series.Points[idx].LegendText = $"{customer}";
                series.Points[idx].ToolTip = $"{customer} - {percent:F1}% (${revenueValue:N0})"; // Tooltip 顯示詳細
            }
        }

        // tabPage4, 右下長條圖 -- 取得產品銷售量前3名
        private DataTable GetProductsCountTopThree(string username)
        {
            string query = $@"
                select p.name as product_name, sum(o.amount) as total_amount
                from orders o
                left join products p on o.product_id = p.product_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by o.product_id
                order by total_amount desc
                limit 3;"; // 限制只取前三名
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        // tabPage4, 右下兩個橫條圖 -- 列出客戶訂單量/產品銷量前三名
        private void LoadProductsCountTopThree(string username)
        {
            DataTable dt = GetProductsCountTopThree(username);
            chartProductsCountTopThree.Titles.Clear();
            chartProductsCountTopThree.Titles.Add("產品總銷售量前3名"); // 設定圖表標題
            chartProductsCountTopThree.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartProductsCountTopThree.Series.Clear();
            var series = chartProductsCountTopThree.Series.Add("Products");
            series.ChartType = SeriesChartType.Bar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "產品名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }
            foreach (DataRow row in dt.Rows)
            {
                string product = row["product_name"].ToString();
                int amount = Convert.ToInt32(row["total_amount"]);
                int idx = series.Points.AddXY(product, amount);
                series.Points[idx].LegendText = $"{product}: {amount}"; // 設定圖例文本
            }
        }

        // tabPage4, 右下長條圖 -- 取得客戶訂單數量前3名
        private DataTable GetCustomerOrdersCountTopThree(string username)
        {
            string query = $@"
                select c.name as customer_name, count(o.customer_id) as total_orders
                from orders o
                left join customers c on o.customer_id = c.customer_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by o.customer_id
                order by total_orders desc
                limit 3;"; // 限制只取前三名
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadCustomerOrdersCountTopThree(string username)
        {
            DataTable dt = GetCustomerOrdersCountTopThree(username);
            chartCustomerOrdersCountTopThree.Titles.Clear();
            chartCustomerOrdersCountTopThree.Titles.Add("客戶總訂單數量前3名"); // 設定圖表標題
            chartCustomerOrdersCountTopThree.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartCustomerOrdersCountTopThree.Series.Clear();
            var series = chartCustomerOrdersCountTopThree.Series.Add("Customers");
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
                int orderCount = Convert.ToInt32(row["total_orders"]);
                int idx = series.Points.AddXY(customer, orderCount);
                series.Points[idx].LegendText = $"{customer}: {orderCount}"; // 設定圖例文本
            }
        }

        // tabPage4, 左下長條圖 -- 取得每月訂單數量
        private DataTable GetTotalOrdersCountData(string username)
        {
            string query = $@"
                select count(*) as monthly_orders, date_format(order_date,'%Y-%m') as months
                from orders o
                where o.sales_user_id = (select user_id from users where username = @username)
                group by months;";
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadTotalOrdersCountChart(string username)
        {
            DataTable dt = GetTotalOrdersCountData(username);
            chartTtlOrdersCount.Titles.Clear();
            chartTtlOrdersCount.Titles.Add("每月訂單數量"); // 設定圖表標題
            chartTtlOrdersCount.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartTtlOrdersCount.Series.Clear();
            var series = chartTtlOrdersCount.Series.Add("Orders");
            series.ChartType = SeriesChartType.Column;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string month = row["months"].ToString();
                int orderCount = Convert.ToInt32(row["monthly_orders"]);
                int idx = series.Points.AddXY(month, orderCount);
                series.Points[idx].Label = $"{orderCount}筆"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {orderCount}筆"; // 設定圖例文本
            }
        }

        //tabPage4, 右上折線圖 -- 列出每月銷售收入
        private DataTable GetTotalRevenueTtlData(string username)
        {
            string query = $@"
                select sum(o.amount*p.price) as monthly_revenue,date_format(order_date,'%Y-%m') as months
                from orders o
                left join products p on o.product_id = p.product_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by months;";
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadTotalRevenueTtlChart(string username)
        {
            DataTable dt = GetTotalRevenueTtlData(username);
            chartMonthlyRevenueTtl.Titles.Clear();
            chartMonthlyRevenueTtl.Titles.Add("每月銷售收入總計"); // 設定圖表標題
            chartMonthlyRevenueTtl.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartMonthlyRevenueTtl.Series.Clear();
            var series = chartMonthlyRevenueTtl.Series.Add("Revenue");
            series.ChartType = SeriesChartType.Line;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "$ #VALY"; // 標籤格式為 "客戶名稱: 銷售額"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            series.BorderWidth = 2;  // 線條粗一點比較明顯
            series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈

            foreach (DataRow row in dt.Rows)
            {
                string month = row["months"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["monthly_revenue"]);
                int idx = series.Points.AddXY(month, revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {revenueValue:N0}"; // 設定圖例文本
            }
        }


        // tabPage4, 右下三個label -- 取得該銷售員的最大產品銷售量、最大客戶訂單數量、最大銷售收入
        private void LoadBiggestLabel(string username)
        {
            string Pquery = $@"
                select p.name as product_name, sum(o.amount) as total_amount
                from orders o
                left join products p on o.product_id = p.product_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by o.product_id
                order by total_amount desc
                limit 1;";
            string Cquery = $@"
                select c.name as customer_name, sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by o.customer_id
                order by ttl_revenue desc
                limit 1;";
            string Rquery = $@"
                select p.name as product_name,sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                where o.sales_user_id = (select user_id from users where username = @username)
                group by p.product_id,o.sales_user_id
                order by ttl_revenue desc limit 1;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmdP = new MySqlCommand(Pquery, connection);
                cmdP.Parameters.AddWithValue("@username", username);
                MySqlCommand cmdC = new MySqlCommand(Cquery, connection);
                cmdC.Parameters.AddWithValue("@username", username);
                MySqlCommand cmdR = new MySqlCommand(Rquery, connection);
                cmdR.Parameters.AddWithValue("@username", username);
                // 取得產品銷售量最大值
                using (var reader = cmdP.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string totalAmount = reader["total_amount"].ToString();
                        labelBiggestSelling.Text = $"{productName} ({totalAmount}件)";
                    }
                    else
                    {
                        labelBiggestSelling.Text = "無";
                    }
                }
                // 取得客戶訂單數量最大值
                using (var reader = cmdC.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string customerName = reader["customer_name"].ToString();
                        labelBiggestCustomer.Text = $"{customerName}";
                    }
                    else
                    {
                        labelBiggestCustomer.Text = "無";
                    }
                }
                // 取得銷售收入最大值
                using (var reader = cmdR.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string totalRevenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelBiggestRevenue.Text = $"{productName} - ${totalRevenue}";
                    }
                    else
                    {
                        labelBiggestRevenue.Text = "無";
                    }
                }
            }
        }




        // ------------------ AdminChartPage_Sales/Product -- tabPage5 ------------------------------------

        // tabPage5, 左下DataGridView -- 取得所有訂單資料
        private void LoadAdminDataGridView()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份

            string query = $@"
                select u.username as sales,
                p.name as product_name,
                sum(o.amount) as amount,
                sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o.product_id,o.sales_user_id
                order by o.sales_user_id;"; // 取得所有訂單資料

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                dataGridViewAdminMonthlyOrders.DataSource = dataTable; // 將查詢結果綁定到 DataGridView
                dataGridViewAdminMonthlyOrders.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridViewAdminMonthlyOrders.DefaultCellStyle.Font = new Font("微軟正黑體", 10);
            }
        }


        // tabPage5 -> tabpage業務, 左下圓餅圖 -- 業務銷售佔比PIE
        private DataTable AdminGetMonthlySalesRevenue()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份

            string query = $@"
                select u.username as sales_user,sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o.sales_user_id
                order by o.sales_user_id;"; // 按月份分組

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable

            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminGetMonthlySalesRevenue()
        {
            DataTable dt = AdminGetMonthlySalesRevenue();
            decimal totalRevenue = 0m;
            foreach (DataRow row in dt.Rows)
            {
                totalRevenue += Convert.ToDecimal(row["ttl_revenue"]);
            }

            chartAdminSalesMonthlyRevenue.Series.Clear();
            chartAdminSalesMonthlyRevenue.Titles.Clear();
            chartAdminSalesMonthlyRevenue.Titles.Add($"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}業務銷售占比"); // 設定圖表標題
            chartAdminSalesMonthlyRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            var series = chartAdminSalesMonthlyRevenue.Series.Add("SalesUsers");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            series.Font = new Font("微軟正黑體", 10, FontStyle.Bold);
            foreach (DataRow row in dt.Rows)
            {
                string salesUser = row["sales_user"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["ttl_revenue"]);
                double percent = totalRevenue == 0 ? 0 : (double)(revenueValue / totalRevenue) * 100;
                // 加進 pie chart
                int idx = series.Points.AddXY(salesUser, revenueValue);
                // 圓餅圖標籤格式：50%
                series.Points[idx].Label = $"{percent:F1}%";
                // 圖例也可以加百分比
                series.Points[idx].LegendText = $"{salesUser}: {percent:F1}%";

                series.Points[idx].ToolTip = $"#VALX - ${revenueValue:N0}";

            }
        }

        // tabPage5 -> tabpage產品, 左下bar -- 某院的產品銷售數量排名
        private DataTable GetMonthlyProductsSellingData()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份

            string query = $@"
                select p.name as product_name, sum(o.amount) as o_count
                from orders o
                left join products p on o.product_id = p.product_id
                where year(o.order_date) = @searchYear 
                 and month(o.order_date) = @searchMonth
                group by o.product_id
                order by o_count asc;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數

                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }


        private void LoadMonthlyProductsSelling()
        {

            DataTable dt = GetMonthlyProductsSellingData();
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartAdminMonthlyProducts.Titles.Clear();
                chartAdminMonthlyProducts.Titles.Add("無資料顯示!!");
                chartAdminMonthlyProducts.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartAdminMonthlyProducts.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartAdminMonthlyProducts.Series.Clear();
                return;
            }
            chartAdminMonthlyProducts.Titles.Clear();
            chartAdminMonthlyProducts.Titles.Add($"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}產品銷售量"); // 設定圖表標題
            chartAdminMonthlyProducts.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartAdminMonthlyProducts.Series.Clear();
            var series = chartAdminMonthlyProducts.Series.Add("Products");
            series.ChartType = SeriesChartType.Bar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "產品名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }
            foreach (DataRow row in dt.Rows)
            {
                string product = row["product_name"].ToString();
                int count = Convert.ToInt32(row["o_count"]);
                int idx = series.Points.AddXY(product, count);
                series.Points[idx].LegendText = $"{product}: {count}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }
        }


        // tabPage5 -> 最左邊的三個label -- 每月營收、業務營收王、產品銷售王 
        private void LoadAdminMonthlyLabel()
        {
            DateTime searchMonth = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份

            // 載入月份總營收
            string Tquery = $@"
                select sum(o.amount*p.price) as ttl_revenue,
                date_format(o.order_date,'%Y-%m') as o_date
                from orders o
                left join products p on o.product_id = p.product_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o_date
                order by o_date;";
            string Squery = $@"
                select sum(o.amount*p.price) as ttl_revenue,
                date_format(o.order_date,'%Y-%m') as o_date,
                u.username as user_name
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o_date,o.sales_user_id
                order by ttl_revenue desc limit 1;";
            string Pquery = $@"
                select sum(o.amount*p.price) as ttl_revenue,
                date_format(o.order_date,'%Y-%m') as o_date,
                p.name as product_name
                from orders o
                left join products p on o.product_id = p.product_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o_date,p.product_id
                order by ttl_revenue desc limit 1;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmdT = new MySqlCommand(Tquery, connection);
                cmdT.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmdT.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                MySqlCommand cmdS = new MySqlCommand(Squery, connection);
                cmdS.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmdS.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                MySqlCommand cmdP = new MySqlCommand(Pquery, connection);
                cmdP.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmdP.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                // 取得每月銷售收入
                using (var reader = cmdT.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string totalRevenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminMTtlRevenue.Text = $"$ {totalRevenue}";
                    }
                    else
                    {
                        labelAdminMTtlRevenue.Text = "無";
                    }
                }
                // 取得業務營業額最大值
                using (var reader = cmdS.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string userName = reader["user_name"].ToString();
                        string revenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminMBestSales.Text = $"{userName}";
                        labelAdminMBestSalesAmt.Text = $"$ {revenue}";
                    }
                    else
                    {
                        labelAdminMBestSales.Text = "無";
                        labelAdminMBestSalesAmt.Text = "無";
                    }
                }
                // 取得產品銷售收入最大值
                using (var reader = cmdP.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string totalRevenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminMBestSelling.Text = $"{productName}";
                        labelAdminMBestSellingAmt.Text = $"${totalRevenue}";
                    }
                    else
                    {
                        labelAdminMBestSelling.Text = "無";
                        labelAdminMBestSellingAmt.Text = "無";
                    }
                }
            }

            labelAdminRevenueTitle.Text = $"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}總營收:";
            labelAdminSalesTitle.Text = $"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}營收王:";
            labelAdminProductTitle.Text = $"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}銷售王:";
            labelAdminSalesGrowthTitleC.Text = $"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}營收";
            // 上個月營收標題
            string prevMonth = dateTimePickerAdmin_SP.Value.AddMonths(-1).ToString("yyyy年MM月");
            labelAdminSalesGrowthTitleP.Text = $"{prevMonth}營收";
        }

        // tabPage5 -> tabpage業務, 上左成長率panel -- 各個sales的銷售成長率
        private void UpdateAdminMonthlyGrowthLabel_Sales()
        {
            DateTime selectedMonth = dateTimePickerAdmin_SP.Value.Date;
            DateTime prevMonth = selectedMonth.AddMonths(-1);

            // 假設這三個帳號是業務
            string[] sales = { "salesA", "salesB", "salesC" };

            // 儲存 label/arrow 的對應
            Label[] currentLabels = { labelSalesACurrentMonth, labelSalesBCurrentMonth, labelSalesCCurrentMonth };
            Label[] lastLabels = { labelSalesALastMonth, labelSalesBLastMonth, labelSalesCLastMonth };
            Label[] growthLabels = { labelSalesAGrowth, labelSalesBGrowth, labelSalesCGrowth };
            PictureBox[] arrows = { pictureBoxARevenueArrow, pictureBoxBRevenueArrow, pictureBoxCRevenueArrow };

            for (int i = 0; i < sales.Length; i++)
            {
                decimal currentRevenue = GetSalesRevenueForMonth_adminPage(sales[i], selectedMonth);
                decimal prevRevenue = GetSalesRevenueForMonth_adminPage(sales[i], prevMonth);

                currentLabels[i].Text = $"${currentRevenue:N0}";
                lastLabels[i].Text = $"${prevRevenue:N0}";

                if (prevRevenue == 0)
                {
                    growthLabels[i].Text = " - %";
                    arrows[i].Image = null;
                    growthLabels[i].ForeColor = Color.Black;
                }
                else
                {
                    decimal growth = (currentRevenue - prevRevenue) / prevRevenue;
                    growthLabels[i].Text = $"{Math.Abs(growth):P1}";
                    arrows[i].Image = growth >= 0
                        ? global::SalesDashboard.Properties.Resources.up
                        : global::SalesDashboard.Properties.Resources.down;
                    growthLabels[i].ForeColor = growth >= 0 ? Color.Green : Color.Red;
                }
            }
        }

        // tabPage5 -> for tabpage業務, 上左成長率panel -- 計算sales某月的總營收
        private decimal GetSalesRevenueForMonth_adminPage(string username, DateTime month)
        {
            decimal revenue = 0m;
            string query = $@"
                select sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear 
                    and month(o.order_date) = @searchMonth
                    and u.username = @username;
    ";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@searchYear", month.Year);
                cmd.Parameters.AddWithValue("@searchMonth", month.Month);
                cmd.Parameters.AddWithValue("@username", username);

                var result = cmd.ExecuteScalar();
                revenue = result != DBNull.Value && result != null ? Convert.ToDecimal(result) : 0m;
            }
            return revenue;
        }

        // tabPage5 -> tabpage業務, 上右的bar -- 計算sales的訂單量排名
        private DataTable GetSalesMonthlyOrdersData()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_SP.Value.Date; // 取得選擇的月份
            string query = $@"
                select count(o.sales_user_id) as o_count,u.username
                from orders o
                left join users u on o.sales_user_id = u.user_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                group by o.sales_user_id
                order by o.sales_user_id desc;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminSalesMonthlyOrdersChart()
        {
            DataTable dt = GetSalesMonthlyOrdersData();
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartAdminSalesMonthlyOrders.Titles.Clear();
                chartAdminSalesMonthlyOrders.Titles.Add("無資料顯示!!");
                chartAdminSalesMonthlyOrders.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartAdminSalesMonthlyOrders.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartAdminSalesMonthlyOrders.Series.Clear();
                return;
            }
            chartAdminSalesMonthlyOrders.Titles.Clear();
            chartAdminSalesMonthlyOrders.Titles.Add($"{dateTimePickerAdmin_SP.Value.ToString("yyyy年MM月")}業務訂單數量"); // 設定圖表標題
            chartAdminSalesMonthlyOrders.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartAdminSalesMonthlyOrders.Series.Clear();
            var series = chartAdminSalesMonthlyOrders.Series.Add("Sales");
            series.ChartType = SeriesChartType.Bar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "產品名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }
            foreach (DataRow row in dt.Rows)
            {
                string salesUser = row["username"].ToString();
                int count = Convert.ToInt32(row["o_count"]);
                int idx = series.Points.AddXY(salesUser, count);
                series.Points[idx].LegendText = $"{salesUser}: {count}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }
        }

        // tabPage5 -> tabpage產品, 上面的成長率panel -- 各個product的銷售/訂單成長率
        private void UpdateAdminMonthlyGrowthLabel_Product(List<string> selectedProducts)
        {
            DateTime selectedMonth = dateTimePickerAdmin_SP.Value.Date;
            DateTime prevMonth = selectedMonth.AddMonths(-1);

            // 取得營收
            decimal currentMonthRevenue = GetProductRevenueForMonth_adminPage(selectedMonth, selectedProducts);
            decimal prevMonthRevenue = GetProductRevenueForMonth_adminPage(prevMonth, selectedProducts);

            // 如果prevMonthRevenue為0（表示第一個月或前一月沒資料），直接顯示 "-"
            if (prevMonthRevenue == 0)
            {
                labelAdminProductsRevenueGrowth.Text = "營收月成長率:-";
                pictureBoxAdminRevenueArrow.Image = null;
                labelAdminProductsRevenueGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal revenueGrowthRate = (currentMonthRevenue - prevMonthRevenue) / prevMonthRevenue;
                labelAdminProductsRevenueGrowth.Text = $"營收月成長率:{Math.Abs(revenueGrowthRate):P1}";
                pictureBoxAdminRevenueArrow.Image = revenueGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelAdminProductsRevenueGrowth.ForeColor = revenueGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelAdminRevenueLastMonth.Text = $"上月營收: ${prevMonthRevenue:N0} 元"; // 顯示上月營收
            labelAdminRevenueCurrentMonth.Text = $"本月營收: ${currentMonthRevenue:N0} 元"; // 顯示本月營收

            // 訂單月成長率
            int currentMonthOrders = GetProductAmountForMonth_adminPage(selectedMonth, selectedProducts);
            int prevMonthOrders = GetProductAmountForMonth_adminPage(prevMonth, selectedProducts);

            if (prevMonthOrders == 0)
            {
                labelAdminProductsOrderGrowth.Text = "銷售月成長率:-";
                pictureBoxAdminOrdersArrow.Image = null;
                labelAdminProductsOrderGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal orderGrowthRate = (currentMonthOrders - prevMonthOrders) / (decimal)prevMonthOrders;
                labelAdminProductsOrderGrowth.Text = $"產品銷售月成長率:{Math.Abs(orderGrowthRate):P1}";
                pictureBoxAdminOrdersArrow.Image = orderGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelAdminProductsOrderGrowth.ForeColor = orderGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelAdminOrderLastMonth.Text = $"上月銷售量: {prevMonthOrders} 件"; // 顯示上月銷售量
            labelAdminOrderCurrentMonth.Text = $"本月銷售量: {currentMonthOrders} 件"; // 顯示本月銷售量
        }

        // for tabPage5, productTabPage的成長率panel -- 計算某月產品的訂單量
        private decimal GetProductRevenueForMonth_adminPage(DateTime month, List<string> selectedProducts)
        {
            decimal revenue = 0;
            string productFilter = selectedProducts.Count > 0
                ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")"
                : "";

            string query = $@"
                SELECT SUM(o.amount * p.price) as total_revenue
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                WHERE YEAR(o.order_date) = @year
                  AND MONTH(o.order_date) = @month
                  {productFilter}";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@year", month.Year);
                cmd.Parameters.AddWithValue("@month", month.Month);

                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);

                var result = cmd.ExecuteScalar();
                revenue = result != DBNull.Value && result != null ? Convert.ToDecimal(result) : 0m;
            }
            return revenue;
        }


        // for tabPage5, productTabPage的成長率panel -- 計算某月產品的訂單量
        private int GetProductAmountForMonth_adminPage(DateTime month, List<string> selectedProducts)
        {
            int count = 0;
            string productFilter = selectedProducts.Count > 0
                ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")"
                : "";

            string query = $@"
                SELECT sum(o.amount) as order_count
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                WHERE YEAR(o.order_date) = @year
                  AND MONTH(o.order_date) = @month
                {productFilter}";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@year", month.Year);
                cmd.Parameters.AddWithValue("@month", month.Month);

                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);

                var result = cmd.ExecuteScalar();
                count = result != DBNull.Value && result != null ? Convert.ToInt32(result) : 0;
            }
            return count;
        }

        // tabPage5 -> tabpage產品, 中間折線圖 -- 每月產品營收趨勢
        private DataTable GetMonthlyProductRevenueData(List<string> selectedProducts)
        {
            DataTable dataTable = new DataTable();
            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";

            string query = $@"
                SELECT 
                    ym_list.ym,
                    p.name AS product_name,
                    IFNULL(SUM(o.amount * p.price), 0) AS total_revenue
                FROM
                    (
                    SELECT '2024-12' AS ym UNION ALL
                    SELECT '2025-01' UNION ALL
                    SELECT '2025-02' UNION ALL
                    SELECT '2025-03' UNION ALL
                    SELECT '2025-04' UNION ALL
                    SELECT '2025-05' UNION ALL
                    SELECT '2025-06' 
                    ) ym_list
                CROSS JOIN products p
                LEFT JOIN orders o ON o.product_id = p.product_id AND DATE_FORMAT(o.order_date, '%Y-%m') = ym_list.ym
                WHERE 1=1
                {productFilter}
                GROUP BY ym_list.ym, p.name
                ORDER BY ym_list.ym, p.name desc;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        // 勾選哪些產品就秀出哪些產品的折線圖
        private void LoadProductRevenueLineChart(List<string> selectedProducts)
        {
            DataTable dt = GetMonthlyProductRevenueData(selectedProducts);  // 傳參數
            if (dt.Rows.Count == 0)
            {
                chartAdminMonthlyProductRevenue.Titles.Clear();
                chartAdminMonthlyProductRevenue.Titles.Add("無資料顯示!!");
                chartAdminMonthlyProductRevenue.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold);
                chartAdminMonthlyProductRevenue.Titles[0].ForeColor = Color.Firebrick;
                chartAdminMonthlyProductRevenue.Series.Clear();
                return;
            }

            chartAdminMonthlyProductRevenue.Series.Clear();
            chartAdminMonthlyProductRevenue.Titles.Clear();
            chartAdminMonthlyProductRevenue.Titles.Add("產品月營收百分比趨勢圖");
            chartAdminMonthlyProductRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);

            // 所有月份
            var months = dt.AsEnumerable()
                .Select(r => r.Field<string>("ym"))
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            // 產品
            var products = dt.AsEnumerable()
                .Select(r => r.Field<string>("product_name"))
                .Distinct()
                .ToList();

            // 先建好所有產品的 Series
            foreach (var product in products)
            {
                var series = chartAdminMonthlyProductRevenue.Series.Add(product);
                series.ChartType = SeriesChartType.StackedColumn100;
                series.IsValueShownAsLabel = true;
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
            }

            // 對每個月，計算總額，再把每個產品的比例算出來
            foreach (var month in months)
            {
                // 這個月的總額
                decimal total = dt.AsEnumerable()
                    .Where(r => r.Field<string>("ym") == month)
                    .Sum(r => r.Field<decimal>("total_revenue"));

                foreach (var product in products)
                {
                    var row = dt.AsEnumerable()
                        .FirstOrDefault(r => r.Field<string>("ym") == month && r.Field<string>("product_name") == product);
                    decimal value = row != null ? row.Field<decimal>("total_revenue") : 0;
                    // 百分比
                    double pct = (total == 0) ? 0 : (double)value / (double)total * 100;
                    int idx = chartAdminMonthlyProductRevenue.Series[product].Points.AddXY(month, value);
                    if(pct < 4)
                    {
                        chartAdminMonthlyProductRevenue.Series[product].Points[idx].Label = " ";
                    }
                    else
                    {
                        chartAdminMonthlyProductRevenue.Series[product].Points[idx].Label = $"{pct:F1}%";
                    }
                    
                    chartAdminMonthlyProductRevenue.Series[product].Points[idx].ToolTip = $"{product}: {pct:F1}%  (${value:N0})";
                }
            }

            // X軸、Y軸
            //chartAdminMonthlyProductRevenue.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminMonthlyProductRevenue.ChartAreas[0].AxisY.Title = "百分比 (%)";
        }


        private void checkedListBoxProducts_adminMonthlyPage_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerAdmin_SP.Value.Year;
            int month = dateTimePickerAdmin_SP.Value.Month;
            // 預測本次操作後的產品勾選清單
            List<string> checkedProducts = checkedListBoxProducts_admin_monthlyPage.CheckedItems.Cast<string>().ToList();
            string currentProduct = checkedListBoxProducts_admin_monthlyPage.Items[e.Index].ToString();
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

            // 傳進查詢方法
            ShowAdminCurrentFilters(checkedProducts);
            UpdateAdminMonthlyGrowthLabel_Product(checkedProducts);
            LoadProductRevenueLineChart(checkedProducts);
        }

        // tabPage5, productTabPage的成長率panel -- 負責動態顯示checkListBox所勾選的內容
        private void ShowAdminCurrentFilters(List<string> products)
        {

            string productText = products.Count == 0 ? "全部" :
                (products.Count > 2 ? string.Join(", ", products.Take(2)) + $"... (共{products.Count}種)" : string.Join(" , ", products));
            string productFullText = products.Count == 0 ? "全部" : string.Join(" , ", products);

            labelAdminProductOrder.Text = "商品: " + productText;
            toolTipAdmin.SetToolTip(labelAdminProductOrder, "篩選商品：" + productFullText);
            labelAdminProductRevenue.Text = "商品: " + productText;
            toolTipAdmin.SetToolTip(labelAdminProductRevenue, "篩選商品：" + productFullText);
        }


        // tabPage5, adminTabPage的progressBar, KPI -- 每月營收目標達成率
        private void LoadAdminMonthlyKPI()
        {
            const decimal targetRevenue = 50000m; // 50萬元
            int year = dateTimePickerAdmin_SP.Value.Year;
            int month = dateTimePickerAdmin_SP.Value.Month;

            string query = $@"
                select sum(o.amount * p.price) as total_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                where year(o.order_date) = @year and month(o.order_date) = @month; 
                ";

            decimal totalRevenue = 0;
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@year", year);
                cmd.Parameters.AddWithValue("@month", month);

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        totalRevenue = reader["total_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["total_revenue"]) : 0m;
                    }
                }
            }

            progressBarAdminMonthlyRevenue.Minimum = 0;
            progressBarAdminMonthlyRevenue.Maximum = (int)targetRevenue;  //這樣寫比較直觀
            int totalRevenueValue = (int)totalRevenue;

            // 防呆避免溢位
            progressBarAdminMonthlyRevenue.Value = totalRevenueValue > progressBarAdminMonthlyRevenue.Maximum
                ? progressBarAdminMonthlyRevenue.Maximum
                : (totalRevenueValue < progressBarAdminMonthlyRevenue.Minimum ? progressBarAdminMonthlyRevenue.Minimum : totalRevenueValue);

            decimal revenueAchievementRate = targetRevenue == 0 ? 0 : totalRevenue / targetRevenue;
            string percentageText = $"{revenueAchievementRate:P0}"; // 百分比格式

            labelAdminMonthlyRevenuePct.Text = percentageText; // 例如「120%」
            labelAdminMonthlyRevenuePatio.Text = $"({totalRevenue:N0} / {targetRevenue:N0})";
        }


        private void dateTimePickerAdmin_ValueChanged(object sender, EventArgs e)
        {
            // 取得當前已勾選的產品、客戶
            List<string> checkedProducts = checkedListBoxProducts_sales.CheckedItems.Cast<string>().ToList();

            LoadAdminDataGridView();
            LoadAdminGetMonthlySalesRevenue();
            LoadAdminMonthlyLabel();
            LoadMonthlyProductsSelling();
            UpdateAdminMonthlyGrowthLabel_Sales();
            LoadAdminSalesMonthlyOrdersChart();
            UpdateAdminMonthlyGrowthLabel_Product(checkedProducts);
            LoadAdminMonthlyKPI();
        }



        // tabPage5,業務tabPage, 中間折線圖 -- 每月業務員營收趨勢
        private void LoadAdminSalesMonthlyRevenueLine()
        {

            DataTable dt = GetAdminSalesMonthlyRevenueData();
            // 2. 取得所有月份和所有業務員
            var months = dt.AsEnumerable().Select(r => r.Field<string>("ym")).Distinct().OrderBy(x => x).ToList();
            var users = dt.AsEnumerable().Select(r => r.Field<string>("username")).Distinct().ToList();

            // 3. 清空舊資料
            chartAdminMonthlyTtlRevenue.Series.Clear();
            chartAdminMonthlyTtlRevenue.Titles.Clear();
            chartAdminMonthlyTtlRevenue.Titles.Add("業務員月營收趨勢圖"); // 設定圖表標題
            chartAdminMonthlyTtlRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式

            // 4. 為每個業務員建立一個 Line Series
            foreach (var user in users)
            {
                var series = chartAdminMonthlyTtlRevenue.Series.Add(user);
                series.ChartType = SeriesChartType.Line;
                series.BorderWidth = 2;  // 線條粗一點比較明顯
                series.IsValueShownAsLabel = true;
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
                series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈
            }

            // 5. 填入每個月的數據
            foreach (var month in months)
            {
                foreach (var user in users)
                {
                    var found = dt.AsEnumerable().FirstOrDefault(r => r.Field<string>("ym") == month && r.Field<string>("username") == user);
                    decimal value = found != null ? found.Field<decimal>("total_revenue") : 0;
                    int idx = chartAdminMonthlyTtlRevenue.Series[user].Points.AddXY(month, value);

                    // 設定 tooltip 格式為 "sales_ - $xx,xxx"
                    chartAdminMonthlyTtlRevenue.Series[user].Points[idx].ToolTip = $"{user} - ${value:N0}";
                    // 如果你還想設定 label 格式
                    chartAdminMonthlyTtlRevenue.Series[user].Points[idx].Label = $"${value:N0}";
                }
            }

            // 6. (選) X軸標籤旋轉角度
            chartAdminMonthlyTtlRevenue.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminMonthlyTtlRevenue.ChartAreas[0].AxisX.Title = "月份";
            chartAdminMonthlyTtlRevenue.ChartAreas[0].AxisY.Title = "營收";

        }


        // tabPage5,業務tabPage, 右下折線圖 -- 每月營收趨勢
        private void LoadAdminTtlRevenueChart_sales()
        {
            DataTable dt = GetAdminTtlRevenueData();
            chartAdminTtlRevenue_sales.Series.Clear();
            chartAdminTtlRevenue_sales.Titles.Clear();
            chartAdminTtlRevenue_sales.ChartAreas[0].AxisX.Title = "月份";
            chartAdminTtlRevenue_sales.ChartAreas[0].AxisY.Title = "營收(元)";
            chartAdminTtlRevenue_sales.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminTtlRevenue_sales.Titles.Add("每月營業額"); // 設定圖表標題
            chartAdminTtlRevenue_sales.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartAdminTtlRevenue_sales.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Line;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            series.BorderWidth = 2;  // 線條粗一點比較明顯
            series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈
            foreach (DataRow row in dt.Rows)
            {
                string month = row["ym"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(month, revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {revenueValue:N0}"; // 設定圖例文本
            }


        }

        // tabPage5,產品tabPage, 右下折線圖 -- 每月營收趨勢
        private void LoadAdminTtlRevenueChart_product()
        {
            DataTable dt = GetAdminTtlRevenueData();
            chartAdminTtlRevenue_product.Series.Clear();
            chartAdminTtlRevenue_product.Titles.Clear();
            chartAdminTtlRevenue_product.ChartAreas[0].AxisX.Title = "月份";
            chartAdminTtlRevenue_product.ChartAreas[0].AxisY.Title = "營收(元)";
            chartAdminTtlRevenue_product.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminTtlRevenue_product.Titles.Add("每月營業額"); // 設定圖表標題
            chartAdminTtlRevenue_product.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartAdminTtlRevenue_product.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Line;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            series.BorderWidth = 2;  // 線條粗一點比較明顯
            series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈

            foreach (DataRow row in dt.Rows)
            {
                string month = row["ym"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(month, revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {revenueValue:N0}"; // 設定圖例文本
            }


        }

        // tabPage5, 左下3個按鈕 -- 切換到業務員的月報表頁面
        private void SwitchToSalesPage(string username)
        {
            MessageBox.Show($"切換到 {username} 的月報表頁面", "切換頁面", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // 彈出一個 Dialog，裡面放 SalesMonthlyUserControl
            var dialog = new Form();
            dialog.Text = $"{username} 的月報表";
            dialog.Width = 1118;
            dialog.Height = 750;
            dialog.StartPosition = FormStartPosition.CenterScreen; // 彈出視窗在螢幕中央
            dialog.FormBorderStyle = FormBorderStyle.Sizable;
            dialog.MaximizeBox = false;
            dialog.MinimizeBox = false;

            // 新建你剛剛寫的 UserControl
            // (假設你有建構式：public SalesMonthlyUserControl(string username) )
            var salesUC = new SalesMonthlyUserControl(username); // 這裡"sales"你可以根據實際需求傳遞角色

            salesUC.Dock = DockStyle.Fill;
            dialog.Controls.Add(salesUC);

            dialog.ShowDialog(); // 這一行會「彈出」視窗
        }

        // 呼叫, 因為 salesTabPage的method都是在登入時傳入username參數, 所以這邊要在觸發按鈕時傳入username參數
        private void buttonSalesA_Click(object sender, EventArgs e)
        {
            SwitchToSalesPage("salesA");
        }
        private void buttonSalesB_Click(object sender, EventArgs e)
        {
            SwitchToSalesPage("salesB");
        }
        private void buttonSalesC_Click(object sender, EventArgs e)
        {
            SwitchToSalesPage("salesC");
        }


        // ------------------ SalesChartPage_Ttl -- tabPage6 ------------------------------------

        // tabPage6, 兩個label,progressBar -- 總營業額、總訂單量、KPI
        private void LoadAdminTtlKPI()
        {
            const decimal targetRevenue = 500000m; // 50萬元
            const int targetOrder = 100;

            string Tquery = $@"
                select sum(o.amount * p.price) as total_revenue,
                count(*) as total_orders
                from orders o
                left join products p on o.product_id = p.product_id;
                ";

            decimal totalRevenue = 0;
            int totalOrders = 0;
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmdT = new MySqlCommand(Tquery, connection);

                using (var reader = cmdT.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        totalRevenue = reader["total_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["total_revenue"]) : 0m;
                        totalOrders = reader["total_orders"] != DBNull.Value ? Convert.ToInt32(reader["total_orders"]) : 0;

                    }
                }
            }

            progressBarAdminTtlOrders.Minimum = 0;
            progressBarAdminTtlOrders.Maximum = 100;
            progressBarAdminTtlOrders.Value = totalOrders; // 更新訂單進度條

            progressBarAdminTtlRevenue.Minimum = 0;
            progressBarAdminTtlRevenue.Maximum = 500000;
            progressBarAdminTtlRevenue.Value = (int)(totalRevenue); // 更新銷售金額進度條

            decimal revenueAchievementRate = targetRevenue == 0 ? 0 : totalRevenue / targetRevenue;
            string percentageText = $"{revenueAchievementRate:P0}"; // 百分比格式

            decimal OrdersAchievementRate = targetOrder == 0 ? 0 : (decimal)totalOrders / targetOrder;
            string OrdersPercentageText = $"{OrdersAchievementRate:P0}"; // 百分比格式


            labelAdminTotalRevenue.Text = $"${totalRevenue:N0}";
            labelAdminTotalOrders.Text = $"{totalOrders} 筆";
            labelAdminTtlRevenuePct.Text = $"{revenueAchievementRate:P0}"; // 更新 Label 顯示達成率
            labelAdminTtlOrdersPct.Text = $"{OrdersAchievementRate:P0}"; // 更新 Label 顯示達成率
            labelAdminTtlRevenuePatio.Text = $"({totalRevenue:N0} / {targetRevenue:N0})";
            labelAdminTtlOrdersPatio.Text = $"({totalOrders} / {targetOrder})";

        }


        // tabPage6, 六個label -- 取得管理員總銷售收入、訂單數量、最佳銷售員、產品銷售量最大值、客戶訂單數量最大值、銷售收入最大值
        private void LoadAdminBiggestLabel()
        {

            string Squery = $@"
                select sum(o.amount*p.price) as total_amount,u.username as username
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                group by o.sales_user_id
                order by total_amount desc
                limit 1;";
            string Pquery = $@"
                select p.name as product_name, sum(o.amount) as total_amount
                from orders o
                left join products p on o.product_id = p.product_id
                group by o.product_id
                order by total_amount desc
                limit 1;";
            string Cquery = $@"
                select c.name as customer_name, sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                group by o.customer_id
                order by ttl_revenue desc
                limit 1;";
            string Rquery = $@"
                select p.name as product_name,sum(o.amount*p.price) as ttl_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                group by p.product_id
                order by ttl_revenue desc
                limit 1;";



            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmdS = new MySqlCommand(Squery, connection);
                MySqlCommand cmdP = new MySqlCommand(Pquery, connection);
                MySqlCommand cmdC = new MySqlCommand(Cquery, connection);
                MySqlCommand cmdR = new MySqlCommand(Rquery, connection);



                using (var reader = cmdS.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string userName = reader["username"].ToString();
                        string totalAmount = Convert.ToDecimal(reader["total_amount"]).ToString("N0");
                        labelAdminTtlBestSales.Text = $"{userName} - ${totalAmount}";
                    }
                }

                // 取得產品銷售量最大值
                using (var reader = cmdP.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string totalAmount = reader["total_amount"].ToString();
                        labelAdminTtlBiggestSelling.Text = $"{productName} ({totalAmount}件)";
                    }
                    else
                    {
                        labelAdminTtlBiggestSelling.Text = "無";
                    }
                }
                // 取得客戶訂單數量最大值
                using (var reader = cmdC.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string customerName = reader["customer_name"].ToString();
                        string totalAmount = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminBiggestCustomer.Text = $"{customerName} - ${totalAmount}";
                    }
                    else
                    {
                        labelAdminBiggestCustomer.Text = "無";
                    }
                }
                // 取得銷售收入最大值
                using (var reader = cmdR.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string productRevenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminBiggestRevenue.Text = $"{productName} - ${productRevenue}";
                    }
                    else
                    {
                        labelAdminBiggestRevenue.Text = "無";
                    }
                }
            }
        }

        // tabPage6, 左下兩個pie -- 取得產品銷售占比、業務銷售占比
        private DataTable GetSalesRevenueData()
        {
            string query = $@"
                select sum(o.amount*p.price) as ttl_revenue,u.username as sales_user
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                group by u.username;";
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminTtlSalesRevenuePie()
        {
            DataTable dt = GetSalesRevenueData();
            // 先算總營收
            decimal totalRevenue = 0m;
            foreach (DataRow row in dt.Rows)
            {
                totalRevenue += Convert.ToDecimal(row["ttl_revenue"]);
            }
            chartAdminSalesRevenuePct.Titles.Clear();
            chartAdminSalesRevenuePct.Titles.Add("業務銷售占比");
            chartAdminSalesRevenuePct.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartAdminSalesRevenuePct.Series.Clear();
            var series = chartAdminSalesRevenuePct.Series.Add("SalesUsers");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            series.Font = new Font("微軟正黑體",10, FontStyle.Bold);
            foreach (DataRow row in dt.Rows)
            {
                string salesUser = row["sales_user"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["ttl_revenue"]);
                double percent = totalRevenue == 0 ? 0 : (double)(revenueValue / totalRevenue) * 100;
                // 加進 pie chart
                int idx = series.Points.AddXY(salesUser, revenueValue);
                // 圓餅圖標籤格式：50%
                series.Points[idx].Label = $"{salesUser} - {percent:F1}%";
                // 圖例也可以加百分比
                series.Points[idx].LegendText = $"{salesUser} - {percent:F1}%";

                series.Points[idx].ToolTip = $"#VALX - ${revenueValue:N0}";
            }
        }

        private DataTable GetProductRevenueData()
        {
            string query = $@"
                select sum(o.amount*p.price) as ttl_revenue,p.name as product_name
                from orders o
                left join products p on o.product_id = p.product_id
                group by p.product_id
                order by ttl_revenue desc;"; // 按產品分組
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminProductRevenuePie()
        {
            DataTable dt = GetProductRevenueData();
            // 先算總營收
            decimal totalRevenue = 0m;
            foreach (DataRow row in dt.Rows)
            {
                totalRevenue += Convert.ToDecimal(row["ttl_revenue"]);
            }
            chartAdminProductRevenuePct.Titles.Clear();
            chartAdminProductRevenuePct.Titles.Add("產品銷售占比");
            chartAdminProductRevenuePct.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartAdminProductRevenuePct.Series.Clear();
            var series = chartAdminProductRevenuePct.Series.Add("Products");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true;
            series.Font = new Font("微軟正黑體", 10, FontStyle.Bold);
            foreach (DataRow row in dt.Rows)
            {
                string productName = row["product_name"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["ttl_revenue"]);
                double percent = totalRevenue == 0 ? 0 : (double)(revenueValue / totalRevenue) * 100;
                // 加進 pie chart
                int idx = series.Points.AddXY(productName, revenueValue);
                // 圓餅圖標籤格式：50%
                series.Points[idx].Label = $"{percent:F1}%";
                // 圖例也可以加百分比
                series.Points[idx].LegendText = $"{productName}";
                series.Points[idx].AxisLabel = productName;

                series.Points[idx].ToolTip = $"#VALX - ${revenueValue:N0}";
            }

        }

        // tabPage6, 中間長條圖 -- 取得季度營業額
        private DataTable GetAdminQuarterRevenue()
        {
            string query = $@"
                select sum(o.amount*p.price) as quarterly_revenue, 
                concat(year(order_date),'-Q',quarter(order_date)) as quarter
                from orders o
                left join products p on o.product_id = p.product_id
                group by quarter;"; // 按季度分組
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminQuarterRevenueChart()
        {
            DataTable dt = GetAdminQuarterRevenue();
            chartAdminQtrRevenue.Titles.Clear();
            chartAdminQtrRevenue.Titles.Add("季度營業額"); // 設定圖表標題
            chartAdminQtrRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            chartAdminQtrRevenue.Series.Clear();
            var series = chartAdminQtrRevenue.Series.Add("Revenue");
            series.ChartType = SeriesChartType.Column;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Label = "#VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string quarter = row["quarter"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["quarterly_revenue"]);
                int idx = series.Points.AddXY(quarter, revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{quarter}: {revenueValue:N0}"; // 設定圖例文本
            }

        }


        // tabPage6, 右下長條圖 -- 每月營收,業務占比

        private DataTable GetAdminSalesMonthlyRevenueData()
        {
            string query = $@"
               select date_format(order_date,""%Y-%m"") as ym,
                u.username as username,
                sum(o.amount*p.price)as total_revenue
               from orders o
               left join products p on o.product_id = p.product_id
               left join users u on o.sales_user_id = u.user_id
               group by ym,u.username
               order by ym,u.username;"; // 按月份分組
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }


        private void LoadAdminSalesMonthlyRevenuePie()
        {
            DataTable dt = GetAdminSalesMonthlyRevenueData();
            chartAdminMonthlySalesRevenue.Titles.Clear();
            chartAdminMonthlySalesRevenue.Series.Clear();
            chartAdminMonthlySalesRevenue.Titles.Add("每月業務銷售占比");
            chartAdminMonthlySalesRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartAdminMonthlySalesRevenue.ChartAreas[0].AxisX.Title = "月份";
            chartAdminMonthlySalesRevenue.ChartAreas[0].AxisY.Title = "營收佔比(%)";
            chartAdminMonthlySalesRevenue.ChartAreas[0].AxisX.LabelStyle.Angle = -45;


            // 先找出所有月份（X 軸）和所有業務員（系列）
            var months = dt.AsEnumerable()
                .Select(r => r.Field<string>("ym"))
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            var users = dt.AsEnumerable()
                .Select(r => r.Field<string>("username"))
                .Distinct()
                .ToList();

            // 為每個業務員建立一個 Series
            foreach (var user in users)
            {
                var series = chartAdminMonthlySalesRevenue.Series.Add(user);
                series.ChartType = SeriesChartType.StackedColumn100;
                series.IsValueShownAsLabel = true; // 顯示標籤
                series.LabelForeColor = Color.White; // 讓標籤白色比較顯眼
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
                //series.Label = "#PERCENT{P2}";
            }

            // 填資料
            foreach (var month in months)
            {
                // 這個月所有業務的營收總和
                decimal total = users.Sum(user =>
                {
                    var found = dt.AsEnumerable().FirstOrDefault(r =>
                        r.Field<string>("ym") == month &&
                        r.Field<string>("username") == user);
                    return found != null ? found.Field<decimal>("total_revenue") : 0;
                });

                foreach (var user in users)
                {
                    var found = dt.AsEnumerable().FirstOrDefault(r =>
                        r.Field<string>("ym") == month &&
                        r.Field<string>("username") == user);
                    decimal value = found != null ? found.Field<decimal>("total_revenue") : 0;
                    var point = chartAdminMonthlySalesRevenue.Series[user].Points.AddXY(month, value);

                    // 手動加百分比字串
                    string label = (total == 0) ? "0%" : $"{Math.Round(value * 100m / total, 2)}%";
                    chartAdminMonthlySalesRevenue.Series[user].Points.Last().Label = label;
                }
            }


            foreach (var s in chartAdminMonthlySalesRevenue.Series)
            {
                foreach (var pt in s.Points)
                {
                    pt.ToolTip = $"{s.Name} - ${pt.YValues[0]:N0}"; // 設定提示工具
                }
            }
        }

        // tabPage6, 左下折線圖 -- 每月營收趨勢
        private DataTable GetAdminTtlRevenueData()
        {
            string query = $@"
                select date_format(order_date,'%Y-%m') as ym, sum(o.amount*p.price) as total_revenue
                from orders o
                left join products p on o.product_id = p.product_id
                group by ym
                order by ym;"; // 按月份分組
            DataTable dataTable = new DataTable();
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
            }
            return dataTable; // 返回填充好的 DataTable
        }

        private void LoadAdminTtlRevenueChart()
        {
            DataTable dt = GetAdminTtlRevenueData();
            chartAdminTtlRevenue.Series.Clear();
            chartAdminTtlRevenue.Titles.Clear();
            chartAdminTtlRevenue.ChartAreas[0].AxisX.Title = "月份";
            chartAdminTtlRevenue.ChartAreas[0].AxisY.Title = "營收(元)";
            chartAdminTtlRevenue.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminTtlRevenue.Titles.Add("每月營業額"); // 設定圖表標題
            chartAdminTtlRevenue.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartAdminTtlRevenue.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Line;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            series.BorderWidth = 2;  // 線條粗一點比較明顯
            series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈
            foreach (DataRow row in dt.Rows)
            {
                string month = row["ym"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(month, revenueValue);
                series.Points[idx].Label = $"${revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {revenueValue:N0}"; // 設定圖例文本
            }


        }


        // ------------------ AdminChartPage_Region -- tabPage7 ------------------------------------

        // tabPage7, 右下雷達圖 -- 根據地區統計總營收
        private DataTable GetCustomerRegionTtlData()
        {
            string query = @"
                select SUBSTRING_INDEX(c.address,'市',1) as region,
                sum(o.amount*p.price) as total_revenue,
                count(o.id) as total_orders
                from orders o
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                group by region; ";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }

        }

        private void LoadCustomerRegionTtlRadarChart()
        {
            DataTable dt = GetCustomerRegionTtlData();
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartCustomerRegionTtlRader.Titles.Clear();
                chartCustomerRegionTtlRader.Titles.Add("無資料顯示!!");
                chartCustomerRegionTtlRader.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartCustomerRegionTtlRader.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartCustomerRegionTtlRader.Series.Clear();
                return;
            }
            chartCustomerRegionTtlRader.Series.Clear();
            chartCustomerRegionTtlRader.Titles.Clear();
            chartCustomerRegionTtlRader.Titles.Add("各地區總營收"); // 設定圖表標題
            chartCustomerRegionTtlRader.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartCustomerRegionTtlRader.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Radar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string region = row["region"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                decimal ordersCount = Convert.ToDecimal(row["total_orders"]);
                int idx = series.Points.AddXY(region, revenueValue);
                series.Points[idx].Label = $"$ {revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{region}: {revenueValue:N0} / ({ordersCount}筆)"; // 設定圖例文本
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            }

        }

        // tabPage7, 右上雷達圖 -- 根據地區以及篩選的產品統計月營收
        private DataTable GetCustomerRegionMthDataByProduct(List<string> selectedProducts)
        {

            DateTime searchMonth = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";
            string query = $@"
                select SUBSTRING_INDEX(c.address,'市',1) as region,
                sum(o.amount*p.price) as total_revenue,
                count(o.id) as total_orders
                from orders o
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                WHERE year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                {productFilter}
                group by region; 
            ";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }
        }

        private void LoadCustomerRegionMthRadarChart(List<string> selectedProducts)
        {
            DataTable dt = GetCustomerRegionMthDataByProduct(selectedProducts);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartCustomerRegionMthRader.Titles.Clear();
                chartCustomerRegionMthRader.Titles.Add("無資料顯示!!");
                chartCustomerRegionMthRader.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartCustomerRegionMthRader.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartCustomerRegionMthRader.Series.Clear();
                return;
            }
            chartCustomerRegionMthRader.Series.Clear();
            chartCustomerRegionMthRader.Titles.Clear();
            chartCustomerRegionMthRader.Titles.Add($"{dateTimePickerAdmin_region.Value.ToString("yyyy年MM月")}各地區營收"); // 設定圖表標題
            chartCustomerRegionMthRader.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartCustomerRegionMthRader.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Radar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string region = row["region"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                decimal ordersCount = Convert.ToDecimal(row["total_orders"]);
                int idx = series.Points.AddXY(region, revenueValue);
                series.Points[idx].Label = $"$ {revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{region}: {revenueValue:N0} / ({ordersCount}筆)"; // 設定圖例文本
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold); // 設定標籤字體樣式
            }
        }



        // tabPage7, 上面左邊bar -- 根據地區統計總營收前3名(搭配產品以及地區篩選)
        private DataTable GetMonthlyTop3RevenueByRegionData(List<string> selectedProducts, List<string> selectedRegion)
        {
            DateTime searchMonth = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            var productFilter = selectedProducts.Count > 0 ? "AND p.name IN (" + string.Join(",", selectedProducts.Select((_, i) => $"@p{i}")) + ")" : "";
            var regionFilter = selectedRegion.Count > 0 ? "AND SUBSTRING_INDEX(c.address,'市',1) IN (" + string.Join(",", selectedRegion.Select((_, i) => $"@c{i}")) + ")" : "";


            string query = $@"
            SELECT 
                date_format(o.order_date,'%Y-%m') as ym,
                SUBSTRING_INDEX(c.address,'市',1) as region,
                IFNULL(SUM(o.amount * p.price), 0) AS total_revenue
                from orders o
            LEFT JOIN customers c ON o.customer_id = c.customer_id
            left join products p on o.product_id = p.product_id
            where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
            {productFilter}
            {regionFilter}
            GROUP BY ym, region
            ORDER BY ym, total_revenue desc limit 3;
            ";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                for (int i = 0; i < selectedProducts.Count; i++)
                    cmd.Parameters.AddWithValue($"@p{i}", selectedProducts[i]);
                for (int i = 0; i < selectedRegion.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedRegion[i]);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }
        }

        private void LoadMonthlyTop3RevenueByRegionChart(List<string> selectedProducts, List<string> selectedRegion)
        {
            DataTable dt = GetMonthlyTop3RevenueByRegionData(selectedProducts, selectedRegion);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartAdminMonthlyTop3RevenueByRegion.Titles.Clear();
                chartAdminMonthlyTop3RevenueByRegion.Titles.Add("無資料顯示!!");
                chartAdminMonthlyTop3RevenueByRegion.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
                chartAdminMonthlyTop3RevenueByRegion.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartAdminMonthlyTop3RevenueByRegion.Series.Clear();
                return;
            }
            chartAdminMonthlyTop3RevenueByRegion.Titles.Clear();
            chartAdminMonthlyTop3RevenueByRegion.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}營收排名"); // 設定圖表標題
            chartAdminMonthlyTop3RevenueByRegion.Titles[0].Font = new Font("微軟正黑體", 10, FontStyle.Bold); // 設定標題字體樣式
            chartAdminMonthlyTop3RevenueByRegion.Series.Clear();
            var series = chartAdminMonthlyTop3RevenueByRegion.Series.Add("Region");
            series.ChartType = SeriesChartType.Bar;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 10, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataPoint pt in series.Points)
            {
                pt.Label = $"{pt.AxisLabel}: {pt.YValues[0]}"; // 設定標籤格式為 "客戶名稱: 銷售量"
            }
            for (int i = dt.Rows.Count - 1; i >= 0; i--)
            {
                var row = dt.Rows[i];
                string region = row["region"].ToString();
                int revenue = Convert.ToInt32(row["total_revenue"]);
                int idx = series.Points.AddXY(region, revenue);
                series.Points[idx].LegendText = $"{region}: {revenue:N0}";
                series.Font = new Font("微軟正黑體", 10, FontStyle.Bold);
                series.Points[idx].Label = $"${revenue:N0}";
            }


        }

        private DataTable GetProductRevenuePctByRegionData(List<string> selectedRegion)
        {
            DateTime searchMonth = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin_region.Value.Date; // 取得選擇的月份
            var regionFilter = selectedRegion.Count > 0 ? "AND SUBSTRING_INDEX(c.address,'市',1) IN (" + string.Join(",", selectedRegion.Select((_, i) => $"@c{i}")) + ")" : "";
            string query = $@"
                select p.name,
                    sum(o.amount*p.price) as total_revenue
                from orders o
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                where year(o.order_date) = @searchYear and month(o.order_date) = @searchMonth
                    {regionFilter}
                group by p.name;
                ";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month);
                for (int i = 0; i < selectedRegion.Count; i++)
                    cmd.Parameters.AddWithValue($"@c{i}", selectedRegion[i]);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }
        }

        private void LoadProductRevenuePctByRegionChart(List<string> selectedRegion)
        {
            DataTable dt = GetProductRevenuePctByRegionData(selectedRegion);
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartProductRevenuePctByRegion.Titles.Clear();
                chartProductRevenuePctByRegion.Titles.Add("無資料顯示!!");
                chartProductRevenuePctByRegion.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
                chartProductRevenuePctByRegion.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartProductRevenuePctByRegion.Series.Clear();
                return;
            }
            chartProductRevenuePctByRegion.Series.Clear();
            chartProductRevenuePctByRegion.Titles.Clear();
            chartProductRevenuePctByRegion.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}各產品營收占比");
            chartProductRevenuePctByRegion.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartProductRevenuePctByRegion.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string productName = row["name"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(productName, revenueValue);
                series.Points[idx].Label = "#PERCENT{P1}"; // 只顯示百分比
                series.Points[idx].ToolTip = $"{productName} - ${revenueValue:N0}"; // Tooltip 顯示詳細
                series.Points[idx].LegendText = $"{productName}: ${revenueValue:N0}"; // 圖例
            }

        }



        private void dateTimePickerAdmin_region_ValueChanged(object sender, EventArgs e)
        {
            List<string> selectedProducts = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();
            List<string> selectedRegion = checkedListBoxRegion_region_admin.CheckedItems.Cast<string>().ToList();
            LoadProductRevenuePctByRegionChart(selectedRegion);
            LoadCustomerRegionMthRadarChart(selectedProducts);
            LoadMonthlyTop3RevenueByRegionChart(selectedProducts, selectedRegion);
        }

        private void checkedListBoxRegion_region_admin_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 預測本次操作後的產品勾選清單
            List<string> checkedRegion = checkedListBoxRegion_region_admin.CheckedItems.Cast<string>().ToList();
            string currentRegion = checkedListBoxRegion_region_admin.Items[e.Index].ToString();
            if (e.NewValue == CheckState.Checked)
            {
                if (!checkedRegion.Contains(currentRegion))
                    checkedRegion.Add(currentRegion);
            }
            else
            {
                if (checkedRegion.Contains(currentRegion))
                    checkedRegion.Remove(currentRegion);
            }
            List<string> checkedProducts = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();
            // 傳進查詢方法

            LoadProductRevenuePctByRegionChart(checkedRegion);
            LoadMonthlyTop3RevenueByRegionChart(checkedProducts, checkedRegion);
        }
        private void checkedListBoxRegion_product_admin_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 預測本次操作後的產品勾選清單
            List<string> checkedProducts = checkedListBoxRegion_product_admin.CheckedItems.Cast<string>().ToList();
            string currentProduct = checkedListBoxRegion_product_admin.Items[e.Index].ToString();
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
            List<string> checkedRegion = checkedListBoxRegion_region_admin.CheckedItems.Cast<string>().ToList();
            // 傳進查詢方法
            LoadCustomerRegionMthRadarChart(checkedProducts);
            LoadMonthlyTop3RevenueByRegionChart(checkedProducts, checkedRegion);
        }


        // tabPage7, 下面中間折線圖 -- 各地區每月營收趨勢
        private DataTable GetMonthlyRevenueByRegionData()
        {
            string query = @"
            SELECT 
            ym_list.ym,
            SUBSTRING_INDEX(c.address,'市',1) as region,
            IFNULL(SUM(o.amount * p.price), 0) AS total_revenue
            FROM
            (
                SELECT '2024-12' AS ym UNION ALL
                SELECT '2025-01' UNION ALL
                SELECT '2025-02' UNION ALL
                SELECT '2025-03' UNION ALL
                SELECT '2025-04' UNION ALL
                SELECT '2025-05' UNION ALL
                SELECT '2025-06' 
                ) ym_list
            CROSS JOIN customers c
            LEFT JOIN orders o ON o.customer_id = c.customer_id AND DATE_FORMAT(o.order_date, '%Y-%m') = ym_list.ym
            left join products p on o.product_id = p.product_id
            GROUP BY ym_list.ym, region
            ORDER BY ym_list.ym, region asc; 
            ";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }
        }

        private void LoadMonthlyRevenueByRegionChart()
        {
            DataTable dt = GetMonthlyRevenueByRegionData();
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartAdminMonthlyRevenueByRegion.Titles.Clear();
                chartAdminMonthlyRevenueByRegion.Titles.Add("無資料顯示!!");
                chartAdminMonthlyRevenueByRegion.Titles[0].Font = new Font("微軟正黑體", 14, FontStyle.Bold); // 設定標題字體樣式
                chartAdminMonthlyRevenueByRegion.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartAdminMonthlyRevenueByRegion.Series.Clear();
                return;
            }
            chartAdminMonthlyRevenueByRegion.Series.Clear();
            chartAdminMonthlyRevenueByRegion.Titles.Clear();
            chartAdminMonthlyRevenueByRegion.Titles.Add($"各地區每月營收"); // 設定圖表標題
            chartAdminMonthlyRevenueByRegion.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var months = dt.AsEnumerable()
                           .Select(r => r.Field<string>("ym"))
                           .Distinct()
                           .OrderBy(x => x)
                           .ToList();
            var regions = dt.AsEnumerable()
                            .Select(r => r.Field<string>("region"))
                            .Distinct()
                            .ToList();
            foreach (var region in regions)
            {
                var series = chartAdminMonthlyRevenueByRegion.Series.Add(region);
                series.ChartType = SeriesChartType.Line;
                series.BorderWidth = 2;  // 線條粗一點比較明顯
                series.IsValueShownAsLabel = true;
                series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
                series.MarkerStyle = MarkerStyle.Circle; // 點上加圓圈
            }

            // 5. 填入每個月的數據
            foreach (var month in months)
            {
                foreach (var region in regions)
                {
                    var found = dt.AsEnumerable().FirstOrDefault(r => r.Field<string>("ym") == month && r.Field<string>("region") == region);
                    decimal value = found != null ? found.Field<decimal>("total_revenue") : 0;
                    int ptIdx = chartAdminMonthlyRevenueByRegion.Series[region].Points.AddXY(month, value);
                    // 設定 DataPoint 的 ToolTip
                    chartAdminMonthlyRevenueByRegion.Series[region].Points[ptIdx].ToolTip = $"{region} - ${value:N0}";
                    chartAdminMonthlyRevenueByRegion.Series[region].IsValueShownAsLabel = false;
                    chartAdminMonthlyRevenueByRegion.Series[region].SmartLabelStyle.Enabled = true; // 啟用智慧標籤樣式，避免重疊
                }
            }

            // 6. (選) X軸標籤旋轉角度
            chartAdminMonthlyRevenueByRegion.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
            chartAdminMonthlyRevenueByRegion.ChartAreas[0].AxisX.Title = "月份";

        }

        private void LoadAdminRegionTtlLabel()
        {

            // 載入月份總營收
            string Tquery = $@"
                select 
                    SUBSTRING_INDEX(c.address,'市',1) as region,
                    sum(o.amount*p.price) as total_revenue
                from orders o
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                group by region
                order by total_revenue desc limit 1;";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmdT = new MySqlCommand(Tquery, connection);
                // 取得每月銷售收入
                using (var reader = cmdT.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string totalRevenue = Convert.ToDecimal(reader["total_revenue"]).ToString("N0");
                        string region = reader["region"].ToString();
                        labelAdminRRegion.Text = $"{region}市";
                        labelAdminRTtlRevenue.Text = $"(${totalRevenue})";
                    }
                    else
                    {
                        labelAdminRRegion.Text = "-";
                        labelAdminRTtlRevenue.Text = "$ - ";
                    }
                }


            }


        }

        private DataTable GetRegionTtlRevenuePctData()
        {
            string query = $@"
                select SUBSTRING_INDEX(c.address,'市',1) as region,
                sum(o.amount*p.price) as total_revenue
                from orders o
                left join customers c on o.customer_id = c.customer_id
                left join products p on o.product_id = p.product_id
                group by region;
            ";

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable); // 將查詢結果填充到 DataTable
                return dataTable; // 返回填充好的 DataTable
            }
        }

        private void LoadRegionTtlRevenuePctChart()
        {
            DataTable dt = GetRegionTtlRevenuePctData();
            if (dt.Rows.Count == 0)
            {
                // 如果沒有資料，清空圖表並顯示提示
                chartRegionTtlRevenuePct.Titles.Clear();
                chartRegionTtlRevenuePct.Titles.Add("無資料顯示!!");
                chartRegionTtlRevenuePct.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
                chartRegionTtlRevenuePct.Titles[0].ForeColor = Color.Firebrick; // 設定標題字體顏色
                chartRegionTtlRevenuePct.Series.Clear();
                return;
            }
            chartRegionTtlRevenuePct.Series.Clear();
            chartRegionTtlRevenuePct.Titles.Clear();
            chartRegionTtlRevenuePct.Titles.Add($"各地區總營收占比"); // 設定圖表標題
            chartRegionTtlRevenuePct.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold); // 設定標題字體樣式
            var series = chartRegionTtlRevenuePct.Series.Add("Total Revenue");
            series.ChartType = SeriesChartType.Pie;
            series.IsValueShownAsLabel = true; // 顯示數值標籤
            series.Font = new Font("微軟正黑體", 10, FontStyle.Bold); // 設定標籤字體樣式
            foreach (DataRow row in dt.Rows)
            {
                string region = row["region"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(region, revenueValue);
                series.Points[idx].Label = "#PERCENT{P1}"; // 只顯示百分比
                series.Points[idx].ToolTip = $"{region} - ${revenueValue:N0}"; // Tooltip 顯示詳細
                series.Points[idx].LegendText = $"{region}: ${revenueValue:N0}"; // 圖例
            }

        }

        
    }
}