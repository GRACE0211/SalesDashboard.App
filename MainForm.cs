using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ZstdSharp.Unsafe;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace SalesDashboard
{
    public partial class MainForm : Form
    {
        private string _username;
        private string _role;
        private string connectionString = "server=localhost;user=root;password=0905656870;database=sales_dashboard;port=3306;SslMode=None";

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

            if (_role == "sales")
            {
                tabControl1.TabPages.Remove(AdminSearchPage); // 移除 AdminSearchPage (原本就存在SalesSearchPage所以不用Add)
            }
            else if (role == "admin")
            {
                tabControl1.TabPages.Remove(SalesSearchPage); // 移除 SalesSearchPage (原本就存在AdminSearchPage所以不用Add)
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
                LoadMonthlyTotalLabel(_username);

            }
            else if (tabControl1.SelectedTab == SalesChartPage_Ttl)
            {
                // 銷售人員
                // 總計
                LoadKPIDoughnutChart(_username);
                LoadProductsCountTopThree(_username);
                LoadCustomerOrdersCountTopThree(_username);
                LoadTotalOrdersCountChart(_username);
                LoadTotalRevenueTtlChart(_username);
                LoadBiggestLabel(_username);
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
        // 圓餅圖 -- 計算某月某產品的總銷售量
        private DataTable GetMonthlyProductsData(string username,List<string>selectedProducts,List<string>selectedCustomers)
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
            DataTable dt = GetMonthlyProductsData(username,selectedProducts,selectedCustomers);
            if(dt.Rows.Count == 0)
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

        // 圓餅圖 -- 計算某月某客戶的訂單數量
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


        // 長條圖 -- 取得某月的銷售數據
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
            series.Font = new Font("微軟正黑體", 9, FontStyle.Bold); // 設定標籤字體樣式
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
        }

        private void LoadMonthlyTotalLabel(string username)
        {
            DateTime searchMonth = dateTimePickerChart.Value.Date; // 取得選擇的月份
            DateTime searchYear = dateTimePickerChart.Value.Date; // 取得選擇的月份
            string query = $@"
                select 
                    sum(o.amount*p.price) as ttl_revenue,
                    count(*) as orders_cnt,
                    o.sales_user_id,
                    date_format(o.order_date,'%Y-%m') as o_date
                from orders o
                left join products p on o.product_id = p.product_id
                left join users u on o.sales_user_id = u.user_id
                where 
                    year(o.order_date) = @searchYear 
                    and month(o.order_date) = @searchMonth
                    and u.username = @username
                group by o.sales_user_id,o_date;";

            decimal monthlyRevenue = 0;
            int monthlyOrders = 0;
            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                MySqlCommand cmd = new MySqlCommand(query, connection);
                cmd.Parameters.AddWithValue("@username", username);
                cmd.Parameters.AddWithValue("@searchMonth", searchMonth.Month); // 添加月份參數
                cmd.Parameters.AddWithValue("@searchYear", searchYear.Year); // 添加年份參數

                using (var reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        monthlyRevenue = reader["ttl_revenue"] != DBNull.Value ? Convert.ToDecimal(reader["ttl_revenue"]) : 0m;
                        monthlyOrders = reader["orders_cnt"] != DBNull.Value ? Convert.ToInt32(reader["orders_cnt"]) : 0;
                    }


                }
            }

            labelMonthlyRevenue.Text = $"{dateTimePickerChart.Value.ToString("yyyy年MM月")}總營業額 - ${monthlyRevenue:N0}";
            labelMonthlyOrders.Text = $"{dateTimePickerChart.Value.ToString("yyyy年MM月")}總訂單量 - {monthlyOrders} 筆";
        }

        private decimal GetRevenueForMonth(string username, DateTime month, List<string> selectedProducts, List<string> selectedCustomers)
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



        private int GetOrderCountForMonth(string username, DateTime month, List<string> selectedProducts, List<string> selectedCustomers)
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

        
        private void UpdateMonthlyGrowthLabel(string username, List<string> selectedProducts, List<string> selectedCustomers)
        {
            DateTime selectedMonth = dateTimePickerChart.Value.Date;
            DateTime prevMonth = selectedMonth.AddMonths(-1);

            // 營收
            decimal currentMonthRevenue = GetRevenueForMonth(username, selectedMonth, selectedProducts, selectedCustomers);
            decimal prevMonthRevenue = GetRevenueForMonth(username, prevMonth, selectedProducts, selectedCustomers);
            decimal revenueGrowthRate = 0;
            if (prevMonthRevenue > 0)
                revenueGrowthRate = (currentMonthRevenue - prevMonthRevenue) / prevMonthRevenue;
            else if (currentMonthRevenue > 0)
                revenueGrowthRate = 1;
            if (revenueGrowthRate >= 0)
            {
                pictureBoxRevenueArrow.Image = global::SalesDashboard.Properties.Resources.arrow_up;
                labelRevenueGrowth.ForeColor = Color.Green;
            }
            else
            {
                pictureBoxRevenueArrow.Image = global::SalesDashboard.Properties.Resources.arrow_down;
                labelRevenueGrowth.ForeColor = Color.Red;
            }
            labelRevenueGrowth.Text = $"營收月成長率：{Math.Abs(revenueGrowthRate):P1}";

            Color revenueColor = revenueGrowthRate >= 0 ? Color.Green : Color.Red;

            labelRevenueGrowth.ForeColor = revenueColor;

            // 訂單
            int currentMonthOrders = GetOrderCountForMonth(username, selectedMonth, selectedProducts, selectedCustomers);
            int prevMonthOrders = GetOrderCountForMonth(username, prevMonth, selectedProducts, selectedCustomers);
            decimal orderGrowthRate = 0;
            if (prevMonthOrders > 0)
                orderGrowthRate = (currentMonthOrders - prevMonthOrders) / (decimal)prevMonthOrders;
            else if (currentMonthOrders > 0)
                orderGrowthRate = 1;
            if (orderGrowthRate >= 0)
            {
                pictureBoxOrdersArrow.Image = global::SalesDashboard.Properties.Resources.arrow_up;
                labelOrderGrowth.ForeColor = Color.Green;
            }
            else
            {
                pictureBoxOrdersArrow.Image = global::SalesDashboard.Properties.Resources.arrow_down;
                labelOrderGrowth.ForeColor = Color.Red;
            }
            labelOrderGrowth.Text = $"訂單月成長率：{Math.Abs(orderGrowthRate):P1}";

            Color orderColor = orderGrowthRate >= 0 ? Color.Green : Color.Red;
            labelOrderGrowth.ForeColor = orderColor;
        }

        

        private void dateTimePickerChart_ValueChanged(object sender, EventArgs e)
        {
            // 取得當前已勾選的產品、客戶
            List<string> selectedProducts = checkedListBoxProducts.CheckedItems.Cast<string>().ToList();
            List<string> selectedCustomers = checkedListBoxCustomers.CheckedItems.Cast<string>().ToList();

            // 三個圖表都重畫（依你的需求）
            RefreshCharts(selectedProducts, selectedCustomers);

            // 若有顯示本月總覽 Label，也可以重抓
            LoadMonthlyTotalLabel(_username);
            UpdateMonthlyGrowthLabel(_username, selectedProducts, selectedCustomers);
        }


        private void checkedListBoxProducts_ItemCheck(object sender, ItemCheckEventArgs e)
        {
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
            UpdateMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
        }

        private void checkedListBoxCustomers_ItemCheck(object sender, ItemCheckEventArgs e)
        {
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
            UpdateMonthlyGrowthLabel(_username, checkedProducts, checkedCustomers);
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
        private void LoadKPIDoughnutChart(string username)
        {
            // KPI 目標
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

            // 設定 Label
            labelTtlRevenue.Text = $"${totalRevenue:N0}";
            labelTtlOrders.Text = $"{totalOrders} 筆";

            decimal revenueAchievementRate = targetRevenue == 0 ? 0 : totalRevenue / targetRevenue;
            string percentageText = $"{revenueAchievementRate:P0}"; // 百分比格式

            // --- Doughnut for Revenue ---
            chartTtlRevenueKPI.Series.Clear();
            var seriesRev = chartTtlRevenueKPI.Series.Add("RevenueKPI");
            seriesRev.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            seriesRev.Points.AddXY($"{totalRevenue:P0}", totalRevenue);
            seriesRev.Points.AddXY("未達成", Math.Max(0,targetRevenue - totalRevenue));
            seriesRev.IsValueShownAsLabel = true;
            foreach(var pt in seriesRev.Points)
            {
                if(pt.AxisLabel == "未達成")
                {
                    pt.Label = ""; // 空白區塊不顯示標籤
                }
                else
                {
                    pt.Label = "#VALY"; // 顯示數值標籤，格式化為千分位
                }
            }
            seriesRev.Label = percentageText; // 設定標籤為百分比格式
            seriesRev["PieStartAngle"]= "270"; // 設定圓餅圖的起始角度為 270 度，這樣第一個區塊會從頂部開始
            seriesRev.Points[0].Color = Color.Orange; // 設定達成區塊的顏色為橙色
            seriesRev.Points[1].Color = Color.FromArgb(51, 47, 41); // 設定未達成區塊的顏色為淺灰色

            // 標題
            chartTtlRevenueKPI.Titles.Clear();
            chartTtlRevenueKPI.Titles.Add("銷售金額達成率");
            chartTtlRevenueKPI.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);

            decimal OrdersAchievementRate = targetOrder == 0 ? 0 : (decimal)totalOrders / targetOrder;
            string OrdersPercentageText = $"{OrdersAchievementRate:P0}"; // 百分比格式

            // --- Doughnut for Orders ---
            chartTtlOrdersKPI.Series.Clear();
            var seriesOrd = chartTtlOrdersKPI.Series.Add("OrderKPI");
            seriesOrd.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Doughnut;
            seriesOrd.Points.AddXY($"{totalOrders}", totalOrders);
            seriesOrd.Points.AddXY("未達成", Math.Max(0, targetOrder - totalOrders));
            seriesOrd.IsValueShownAsLabel = true;
            foreach(var pt in seriesOrd.Points)
            {
                if(pt.AxisLabel == "未達成")
                {
                    pt.Label = ""; // 空白區塊不顯示標籤
                }
                else
                {
                    pt.Label = "#VALY"; // 顯示數值標籤，格式化為千分位
                }
            }

            seriesOrd.Label = OrdersPercentageText; // 設定標籤為百分比格式
            seriesOrd["PieStartAngle"] = "270"; // 設定圓餅圖的起始角度為 270 度，這樣第一個區塊會從頂部開始
            seriesOrd.Points[0].Color = Color.Orange; // 設定達成區塊的顏色為橙色
            seriesOrd.Points[1].Color = Color.FromArgb(51, 47, 41);


            labelTtlRevenuePatio.Text = $"({totalRevenue:N0} / {targetRevenue:N0})";
            labelTtlOrdersPatio.Text = $"({totalOrders} / {targetOrder})"; 
            labelTtlRevenuePct.Text = $"{revenueAchievementRate:P0}"; // 更新 Label 顯示達成率
            labelTtlOrdersPct.Text = $"{OrdersAchievementRate:P0}"; // 更新 Label 顯示達成率
            // 標題
            chartTtlOrdersKPI.Titles.Clear();
            chartTtlOrdersKPI.Titles.Add("訂單數量達成率");
            chartTtlOrdersKPI.Titles[0].Font = new Font("微軟正黑體", 12, FontStyle.Bold);
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

        //tabPage4, 右上長條圖 -- 列出每月銷售收入
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
                    if(reader.Read())
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
                using(var reader = cmdR.ExecuteReader())
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
    }
}
