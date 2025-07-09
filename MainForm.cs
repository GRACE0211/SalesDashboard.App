

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

            if (_role == "sales")
            {
                tabControl1.TabPages.Remove(AdminSearchPage); // 移除 AdminSearchPage (原本就存在SalesSearchPage所以不用Add)
                tabControl1.TabPages.Remove(AdminChartPage_SalesNProduct); // 移除 AdminChartPage_SalesNProduct
                tabControl1.TabPages.Remove(AdminChartPage_Ttl); // 移除 AdminChartPage_Ttl
            }
            else if (role == "admin")
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
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> selectedProducts = checkedListBoxProducts.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers.CheckedItems.Cast<string>().ToList();

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

                LoadMonthlyProductsChart(_username, selectedProducts, selectedCustomers);
                LoadMonthlyOrdersChartByCustomer(_username, selectedProducts, selectedCustomers);
                LoadMonthlyRevenueDetailsChart(_username, selectedProducts, selectedCustomers);
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
                LoadAdminSalesMonthiyRevenueChart();
                LoadAdminTtlRevenueChart();
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

        private void ApplyFilter_adminSearchPage()
        {
            List<string> customers = new List<string>();
            if (chkCustomerA_AdminSearch.Checked)
                customers.Add("'客戶A'");
            if (chkCustomerB_AdminSearch.Checked)
                customers.Add("'客戶B'");
            if (chkCustomerC_AdminSearch.Checked)
                customers.Add("'客戶C'");
            if (chkCustomerD_AdminSearch.Checked)
                customers.Add("'客戶D'");
            if (chkCustomerE_AdminSearch.Checked)
                customers.Add("'客戶E'");

            List<string> products = new List<string>();
            if (chkToothpaste_AdminSearch.Checked)
                products.Add("'toothpaste'");
            if (chkToothbrush_AdminSearch.Checked)
                products.Add("'toothbrush'");
            if (chkShampoo_AdminSearch.Checked)
                products.Add("'shampoo'");
            if (chkShaver_AdminSearch.Checked)
                products.Add("'shaver'");
            if (chkComb_AdminSearch.Checked)
                products.Add("'comb'");

            List<string> salesUsers = new List<string>();
            if (chkSalesA_AdminSearch.Checked)
                salesUsers.Add("'salesA'");
            if (chkSalesB_AdminSearch.Checked)
                salesUsers.Add("'salesB'");
            if (chkSalesC_AdminSearch.Checked)
                salesUsers.Add("'salesC'");

            DateTime startDate = dateTimePickerSearch_AdminStart.Value.Date;
            DateTime endDate = dateTimePickerSearch_AdminEnd.Value.Date;
            string dateFilter = "Date(o.order_date) BETWEEN  @startDate AND @endDate";
            string salesUserFilter = salesUsers.Count > 0 ? $"u.username IN ({string.Join(",", salesUsers)})" : "1=1";
            string customerFilter = customers.Count > 0 ? $"c.name IN ({string.Join(",", customers)})" : "1=1";
            string productFilter = products.Count > 0 ? $"p.name IN ({string.Join(",", products)})" : "1=1";
            string query = $@"SELECT 
                p.name as product_name,
                c.name as customer_name,
                DATE(o.order_date) as order_date,
                sum(o.amount) as total_amount,
                CONCAT('$',sum(p.price*o.amount)) as revenue,
                u.username as sales_name
                FROM orders o 
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where {salesUserFilter} AND {customerFilter} AND {productFilter} AND {dateFilter}
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

        private void chkSalesA_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkSalesB_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkSalesC_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkCustomerA_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkCustomerB_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkCustomerC_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkCustomerD_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkCustomerE_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkToothpaste_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkToothbrush_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkShampoo_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkShaver_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void chkComb_AdminSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void dateTimePickerSearch_AdminStart_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void dateTimePickerSearch_AdminEnd_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }

        private void ApplyFilter_salesSearchPage()
        {

            List<string> customers = new List<string>();
            if (chkCustomerA_SalesSearch.Checked)
                customers.Add("'客戶A'");
            if (chkCustomerB_SalesSearch.Checked)
                customers.Add("'客戶B'");
            if (chkCustomerC_SalesSearch.Checked)
                customers.Add("'客戶C'");
            if (chkCustomerD_SalesSearch.Checked)
                customers.Add("'客戶D'");
            if (chkCustomerE_SalesSearch.Checked)
                customers.Add("'客戶E'");

            List<string> products = new List<string>();
            if (chkToothpaste_SalesSearch.Checked)
                products.Add("'toothpaste'");
            if (chkToothbrush_SalesSearch.Checked)
                products.Add("'toothbrush'");
            if (chkShampoo_SalesSearch.Checked)
                products.Add("'shampoo'");
            if (chkShaver_SalesSearch.Checked)
                products.Add("'shaver'");
            if (chkComb_SalesSearch.Checked)
                products.Add("'comb'");

            DateTime startDate = dateTimePickerSearch_SalesStart.Value.Date;
            DateTime endDate = dateTimePickerSearch_SalesEnd.Value.Date;
            string dateFilter = "Date(o.order_date) BETWEEN  @startDate AND @endDate";
            string loginUser = "u.username = @username";
            string customerFilter = customers.Count > 0 ? $"c.name IN ({string.Join(",", customers)})" : "1=1";
            string productFilter = products.Count > 0 ? $"p.name IN ({string.Join(",", products)})" : "1=1";
            string query = $@"SELECT 
                p.name as product_name,
                c.name as customer_name,
                DATE(o.order_date) as order_date,
                sum(o.amount) as total_amount,
                CONCAT('$',sum(p.price*o.amount)) as revenue,
                u.username as sales_name
                FROM orders o 
                left join products p on o.product_id = p.product_id
                left join customers c on o.customer_id = c.customer_id
                left join users u on o.sales_user_id = u.user_id
                where {loginUser} AND {customerFilter} AND {productFilter} AND {dateFilter}
                group by o.product_id,o.customer_id,o.sales_user_id,o.order_date 
                order by o.customer_id,o.sales_user_id;";
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                if (_role == "sales")
                {
                    cmd.Parameters.AddWithValue("@username", _username); // 添加參數以防止 SQL 注入
                }
                cmd.Parameters.AddWithValue("@startDate", startDate);
                cmd.Parameters.AddWithValue("@endDate", endDate);
                MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);
                dataGridView3.DataSource = dataTable; // 將查詢結果綁定到 DataGridView
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // 自動調整列寬
            }
        }

        private void chkCustomerA_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkCustomerB_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkCustomerC_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkCustomerD_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkCustomerE_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkToothpaste_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkToothbrush_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkShampoo_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkShaver_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void chkComb_SalesSearch_CheckedChanged(object sender, EventArgs e)
        {
            ApplyFilter_salesSearchPage();
        }
        private void dateTimePickerSearch_SalesStart_ValueChanged(object sender, EventArgs e)
        {
            ApplyFilter_adminSearchPage();
        }
        private void dateTimePickerSearch_SalesEnd_ValueChanged(object sender, EventArgs e)
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

            chartSalesProducts.Titles.Clear();
            chartSalesProducts.Titles.Add($"{dateTimePickerChart.Value.ToString("yyyy年MM月")}產品銷售量"); // 設定圖表標題
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
                int idx = series.Points.AddXY(product, amount);
                series.Points[idx].LegendText = $"{product}: {amount}"; // 設定標籤格式為 "產品名稱: 銷售量"
            }

            Debug.WriteLine($"chartSalesProducts == null?{chartSalesProducts == null}");
            Debug.WriteLine($"chartSalesProducts.Visible:{chartSalesProducts.Visible}");
            Debug.WriteLine($"chartSalesProducts.Size:{chartSalesProducts.Size}");
            Debug.WriteLine($"chartSalesProducts.Parent == null?" +
                $":{chartSalesProducts.Parent == null}");

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
            series.Label = "#VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9); // 設定標籤字體樣式
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

        private void chartSalesProducts_MouseClick(object sender, MouseEventArgs e)
        {
            var hit = chartSalesProducts.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                string product = chartSalesProducts.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadSalesProductsTrend(product, _username);
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
            chartSalesProductTrend.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartSalesProductTrend.Series.Add("ProductsCount");
            chartSalesProductTrend.Series["ProductsCount"].ChartType = SeriesChartType.Line;
            chartSalesProductTrend.Series["ProductsCount"].IsValueShownAsLabel = true;
            chartSalesProductTrend.Series["ProductsCount"].Font = new Font("微軟正黑體", 9);
            chartSalesProductTrend.Series["ProductsCount"].Color = Color.RoyalBlue;

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
            var hit = chartSalesCustomersOrders.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                string customerName = chartSalesCustomersOrders.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadCustomerOrderTrend(customerName, _username);
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
            chartSalesOrdersTrend.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartSalesOrdersTrend.Series.Add("OrdersCount");
            chartSalesOrdersTrend.Series["OrdersCount"].ChartType = SeriesChartType.Line;
            chartSalesOrdersTrend.Series["OrdersCount"].IsValueShownAsLabel = true;
            chartSalesOrdersTrend.Series["OrdersCount"].Font = new Font("微軟正黑體", 9);
            chartSalesOrdersTrend.Series["OrdersCount"].Color = Color.RoyalBlue;
            var months = GetAllMonths("2024-12", DateTime.Today.ToString("yyyy-MM"));
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
            var hit = chartMonthlyRevenuePerProduct.HitTest(e.X, e.Y);
            if (hit.ChartElementType == ChartElementType.DataPoint)
            {
                string customerProduct = chartMonthlyRevenuePerProduct.Series[0].Points[hit.PointIndex].AxisLabel;
                LoadRevenueTrend(customerProduct, _username);
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
            chartSalesRevenueTrend.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
            chartSalesRevenueTrend.Series.Add("Revenue");
            chartSalesRevenueTrend.Series["Revenue"].ChartType = SeriesChartType.Line;
            chartSalesRevenueTrend.Series["Revenue"].IsValueShownAsLabel = true;
            chartSalesRevenueTrend.Series["Revenue"].Font = new Font("微軟正黑體", 9);
            chartSalesRevenueTrend.Series["Revenue"].Color = Color.RoyalBlue;

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
                chartSalesRevenueTrend.Series["Revenue"].Points.AddXY(m, monthData.ContainsKey(m) ? monthData[m] : 0);
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
                labelRevenueGrowth.Text = "營收月成長率：-";
                pictureBoxRevenueArrow.Image = null;
                labelRevenueGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal revenueGrowthRate = (currentMonthRevenue - prevMonthRevenue) / prevMonthRevenue;
                labelRevenueGrowth.Text = $"營收月成長率：{Math.Abs(revenueGrowthRate):P1}";
                pictureBoxRevenueArrow.Image = revenueGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelRevenueGrowth.ForeColor = revenueGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelRevenueLastMonth.Text = $"上月營收：${prevMonthRevenue:N0}"; // 顯示上月營收
            labelRevenueCurrentMonth.Text = $"本月營收：${currentMonthRevenue:N0}"; // 顯示本月營收

            // 訂單月成長率
            int currentMonthOrders = GetOrderCountForMonth_salesPage(username, selectedMonth, selectedProducts, selectedCustomers);
            int prevMonthOrders = GetOrderCountForMonth_salesPage(username, prevMonth, selectedProducts, selectedCustomers);

            if (prevMonthOrders == 0)
            {
                labelOrderGrowth.Text = "訂單月成長率：-";
                pictureBoxOrdersArrow.Image = null;
                labelOrderGrowth.ForeColor = Color.Black;
            }
            else
            {
                decimal orderGrowthRate = (currentMonthOrders - prevMonthOrders) / (decimal)prevMonthOrders;
                labelOrderGrowth.Text = $"訂單月成長率：{Math.Abs(orderGrowthRate):P1}";
                pictureBoxOrdersArrow.Image = orderGrowthRate >= 0
                    ? global::SalesDashboard.Properties.Resources.arrow_up
                    : global::SalesDashboard.Properties.Resources.arrow_down;
                labelOrderGrowth.ForeColor = orderGrowthRate >= 0 ? Color.Green : Color.Red;
            }

            labelOrderLastMonth.Text = $"上月訂單量：{prevMonthOrders} 筆"; // 顯示上月訂單量
            labelOrderCurrentMonth.Text = $"本月訂單量：{currentMonthOrders} 筆"; // 顯示本月訂單量
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
            toolTip2.SetToolTip(labelOrderCustomer, "客戶：全部");
            toolTip2.SetToolTip(labelOrderProduct, "商品：全部");
            toolTip2.SetToolTip(labelRevenueCustomer, "客戶：全部");
            toolTip2.SetToolTip(labelRevenueProduct, "商品：全部");
        }

        // tabPage3, 中間的panel -- 負責動態顯示checkListBox所勾選的內容
        private void ShowCurrentFilters(List<string> products, List<string> customers)
        {
            string customerText = customers.Count == 0 ? "全部" :
                (customers.Count > 2 ? string.Join(", ", customers.Take(2)) + $"... (共{customers.Count}位)" : string.Join(" , ", customers));
            string customerFullText = customers.Count == 0 ? "全部" : string.Join(" , ", customers);

            labelRevenueCustomer.Text = "客戶: " + customerText;
            toolTip2.SetToolTip(labelRevenueCustomer, "客戶：" + customerFullText);
            labelOrderCustomer.Text = "客戶: " + customerText;
            toolTip2.SetToolTip(labelOrderCustomer, "客戶：" + customerFullText);

            string productText = products.Count == 0 ? "全部" :
                (products.Count > 2 ? string.Join(", ", products.Take(2)) + $"... (共{products.Count}種)" : string.Join(" , ", products));
            string productFullText = products.Count == 0 ? "全部" : string.Join(" , ", products);

            labelRevenueProduct.Text = "商品: " + productText;
            toolTip2.SetToolTip(labelRevenueProduct, "商品：" + productFullText);
            labelOrderProduct.Text = "商品: " + productText;
            toolTip2.SetToolTip(labelOrderProduct, "商品：" + productFullText);
        }

        private void LoadMonthlyRevenueKPI(string username)
        {
            string query = $@"
                SELECT COUNT(*) as order_count
                FROM orders o
                LEFT JOIN products p ON o.product_id = p.product_id
                LEFT JOIN customers c ON o.customer_id = c.customer_id
                LEFT JOIN users u ON o.sales_user_id = u.user_id
                WHERE YEAR(o.order_date) = @year
                  AND MONTH(o.order_date) = @month
                  AND u.username = @username";
        }


        // dateTimePicker 時間改的話就呼叫
        private void dateTimePickerChart_ValueChanged(object sender, EventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerChart.Value.Year;
            int month = dateTimePickerChart.Value.Month;
            // 取得當前已勾選的產品、客戶
            List<string> checkedProducts = checkedListBoxProducts.CheckedItems.Cast<string>().ToList();
            List<string> checkedCustomers = checkedListBoxCustomers.CheckedItems.Cast<string>().ToList();

            // 三個圖表都重畫（依你的需求）
            RefreshCharts(checkedProducts, checkedCustomers);

            // 若有顯示本月總覽 Label，也可以重抓
            //LoadMonthlyTotalLabel(_username);
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
        }

        // 若有勾選，就加入List
        private void checkedListBoxProducts_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerChart.Value.Year;
            int month = dateTimePickerChart.Value.Month;
            // 預測本次操作後的產品勾選清單
            List<string> checkedProducts = checkedListBoxProducts.CheckedItems.Cast<string>().ToList();
            string currentProduct = checkedListBoxProducts.Items[e.Index].ToString();
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
            List<string> checkedCustomers = checkedListBoxCustomers.CheckedItems.Cast<string>().ToList();

            // 傳進查詢方法
            RefreshCharts(checkedProducts, checkedCustomers);
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            ShowCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
        }

        private void checkedListBoxCustomers_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // 取得當前選擇的年月
            int year = dateTimePickerChart.Value.Year;
            int month = dateTimePickerChart.Value.Month;
            List<string> checkedCustomers = checkedListBoxCustomers.CheckedItems.Cast<string>().ToList();
            string currentCustomer = checkedListBoxCustomers.Items[e.Index].ToString();
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
            List<string> checkedProducts = checkedListBoxProducts.CheckedItems.Cast<string>().ToList();

            RefreshCharts(checkedProducts, checkedCustomers);
            UpdateSalesMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
            ShowCurrentFilters(checkedProducts, checkedCustomers);
            GetMonthlyOrdersRowData(_username, checkedProducts, checkedCustomers);
        }

        // 統一重畫圖表
        private void RefreshCharts(List<string> checkedProducts, List<string> checkedCustomers)
        {
            LoadMonthlyProductsChart(_username, checkedProducts, checkedCustomers);
            LoadMonthlyOrdersChartByCustomer(_username, checkedProducts, checkedCustomers);
            LoadMonthlyRevenueDetailsChart(_username, checkedProducts, checkedCustomers);
        }


        // ------------------ SalesChartPage_Ttl -- tabPage4 ------------------


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

        // tabPage4, 左中兩個圓餅圖 -- 客戶/產品銷售占比
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
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold);

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
                series.Points[idx].LegendText = $"{product}: {percent:F1}%";
            }
        }

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
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold);
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
                series.Points[idx].LegendText = $"{customer}: {percent:F1}%";
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
            series.Label = "#VALY"; // 標籤格式為 "客戶名稱: 銷售量"
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
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
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
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

            }
            series.ToolTip = $"#VALX: $#VALY";
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
            series.Font = new Font("微軟正黑體", 8, FontStyle.Bold);
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
                series.Points[idx].LegendText = $"{productName}: {percent:F1}%";
                series.Points[idx].AxisLabel = productName;
            }

            series.ToolTip = $"#VALX: $#VALY";
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


        // tabPage6, 右上長條圖 -- 每月營收,業務占比

        private DataTable GetAdminSalesMonthlyRevenueData()
        {
            string query = $@"
               select date_format(order_date,""%Y-%m"") as ym,u.username, sum(o.amount*p.price)as total_revenue
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


        private void LoadAdminSalesMonthiyRevenueChart()
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
                    pt.ToolTip = $"{s.Name} : {pt.YValues[0]:N0}元"; // 設定提示工具
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
            foreach (DataRow row in dt.Rows)
            {
                string month = row["ym"].ToString();
                decimal revenueValue = Convert.ToDecimal(row["total_revenue"]);
                int idx = series.Points.AddXY(month, revenueValue);
                series.Points[idx].Label = $"{revenueValue:N0}"; // 設定標籤格式為 "產品名稱: 銷售量"
                series.Points[idx].LegendText = $"{month}: {revenueValue:N0}"; // 設定圖例文本
            }


        }

        // tabPage5, 左下DataGridView -- 取得所有訂單資料
        private void LoadAdminDataGridView()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin.Value.Date; // 取得選擇的月份

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

        
        // tabPage5 -> tabpage業務, 業務銷售佔比PIE
        private DataTable AdminGetMonthlySalesRevenue()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin.Value.Date; // 取得選擇的月份

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
            chartAdminSalesMonthlyRevenue.Titles.Add($"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}業務銷售占比"); // 設定圖表標題
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

            }
            series.ToolTip = $"#VALX: $#VALY";
        }


        // tabPage5, <產品> Bar, 計算每月各產品銷售量

        private DataTable GetMonthlyProductsSellingData()
        {
            DataTable dataTable = new DataTable();
            DateTime searchMonth = dateTimePickerAdmin.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin.Value.Date; // 取得選擇的月份

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
            chartAdminMonthlyProducts.Titles.Add($"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}產品銷售量"); // 設定圖表標題
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


        // tabPage5, 三個label -- 每月營收、業務營收王、產品銷售王 
        private void LoadAdminMonthlyLabel()
        {
            DateTime searchMonth = dateTimePickerAdmin.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerAdmin.Value.Date; // 取得選擇的月份

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
                        labelAdminMBestSales.Text = $"{userName} - $ {revenue}";
                    }
                    else
                    {
                        labelAdminMBestSales.Text = "無";
                    }
                }
                // 取得產品銷售收入最大值
                using (var reader = cmdP.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        string productName = reader["product_name"].ToString();
                        string totalRevenue = Convert.ToDecimal(reader["ttl_revenue"]).ToString("N0");
                        labelAdminMBestSelling.Text = $"{productName} - ${totalRevenue}";
                    }
                    else
                    {
                        labelAdminMBestSelling.Text = "無";
                    }
                }
            }

            labelAdminRevenueTitle.Text = $"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}總營收:";
            labelAdminSalesTitle.Text = $"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}營收王:";
            labelAdminProductTitle.Text = $"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}銷售王:";
            labelAdminSalesGrowthTitleC.Text = $"{dateTimePickerAdmin.Value.ToString("yyyy年MM月")}營收";
            string prevMonth = dateTimePickerAdmin.Value.AddMonths(-1).ToString("yyyy年MM月");
            labelAdminSalesGrowthTitleP.Text = $"{prevMonth}營收";
        }

        // tabPage5, 成長率panel -- 各個sales的銷售成長率
        private void UpdateAdminMonthlyGrowthLabel_Sales()
        {
            DateTime selectedMonth = dateTimePickerAdmin.Value.Date;
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
                        ? global::SalesDashboard.Properties.Resources.arrow_up
                        : global::SalesDashboard.Properties.Resources.arrow_down;
                    growthLabels[i].ForeColor = growth >= 0 ? Color.Green : Color.Red;
                }
            }
        }

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


        private void dateTimePickerAdmin_ValueChanged(object sender, EventArgs e)
        {
            LoadAdminDataGridView();
            LoadAdminGetMonthlySalesRevenue();
            LoadAdminMonthlyLabel();
            LoadMonthlyProductsSelling();
            UpdateAdminMonthlyGrowthLabel_Sales();
        }
    }
}
