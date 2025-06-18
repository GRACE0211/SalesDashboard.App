using System;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using MySql.Data.MySqlClient;

namespace SalesDashboard
{
    public partial class LoginForm : Form
    {
        public string LoginUsername { get; private set; }
        public string LoginRole { get; private set; }
        public LoginForm()
        {
            InitializeComponent();
        }

        // 密碼hash（SHA256）
        private string HashPassword(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password));
                StringBuilder sb = new StringBuilder();
                foreach (var b in bytes)
                    sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        private void LoginButton_Click(object sender, EventArgs e)
        {
            string connectionString = "server=127.0.0.1;user=root;password=0905656870;database=sales_dashboard;port=3306;";

        string username = UsernameTextBox.Text.Trim();
            string password = PasswordTextBox.Text; // 不要Trim密碼

            // hash 用戶輸入的密碼
            string hash = HashPassword(password);

            using (MySqlConnection connection = new MySqlConnection(connectionString))
            {
                connection.Open();
                // 假設用 username 登入（也可以用 user_id）
                string sql = "SELECT * FROM users WHERE username = @username LIMIT 1";
                using (var cmd = new MySqlCommand(sql, connection))
                {
                    cmd.Parameters.AddWithValue("@username", username);
                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            string password_hash = reader["password_hash"].ToString();
                            string role = reader["role"].ToString();

                            if (password_hash == hash)
                            {
                                // 登入成功
                                MessageBox.Show("登入成功!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                // 你可以把 role、user_id等資訊傳給 MainForm
                                LoginUsername = username;
                                LoginRole = role;
                                MainForm mainForm = new MainForm(username, role); // 假設有這建構子
                                this.DialogResult = DialogResult.OK;
                                this.Close();

                            }
                            else
                            {
                                MessageBox.Show("帳號或密碼錯誤!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("帳號或密碼錯誤!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            PasswordTextBox.PasswordChar = checkBox1.Checked ? '\0' : '*'; // 顯示或隱藏密碼
        }
    }
}
