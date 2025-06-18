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

            if (_role == "sales")
            {
                tabControl1.TabPages.Remove(AdminSearchPage); // 移除 AdminSearchPage (原本就存在SalesSearchPage所以不用Add)
                tabControl1.TabPages.Remove(AdminStatChartPage);
            }
            else if (role == "admin")
            {
                tabControl1.TabPages.Remove(SalesSearchPage); // 移除 SalesSearchPage (原本就存在AdminSearchPage所以不用Add)
                tabControl1.TabPages.Remove(SalesStatChartPage);
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
    }
}
