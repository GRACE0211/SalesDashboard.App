using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SalesDashboard
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 在這裡作流程控制
            while (true)
            {
                
                using (LoginForm loginForm = new LoginForm())
                {
                    // 顯示loginForm
                    var result = loginForm.ShowDialog();

                    /* 
                     * loginForm 有設定登入成功的話 this.DialogResult = DialogResult.OK;
                     * 所以如果沒有登入成功(按叉叉關閉視窗)
                     */
                    if (result != DialogResult.OK)
                    {
                        break;
                    }

                    // 執行到這裡表示登入成功，LoginForm已被關閉
                    string username = loginForm.LoginUsername;
                    string role = loginForm.LoginRole;
                    // 在這裡可以使用 username 和 role 進行後續操作(MainForm顯示目前登入帳號)
                    using (MainForm mainForm = new MainForm(username, role))
                    {
                        // 顯示mainForm
                        mainForm.ShowDialog();
                        if (!mainForm.LogoutFlag)
                        {
                            // 按 X 關閉主畫面，退出應用程式
                            break;
                        }

                        // 如果是登出，登出標誌為 true，繼續while迴圈，重開LoginForm
                    }

                }
            }
            


               
        }
    }
}
