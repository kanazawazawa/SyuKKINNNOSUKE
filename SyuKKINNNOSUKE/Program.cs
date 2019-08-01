using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Text;
using System.Threading;
namespace SyuKKINNNOSUKE
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var EnvironmentVarCompanycd = "SYUKKINNNOSUKE_COMPANYCD";
            var EnvironmentVarLogincd = "SYUKKINNNOSUKE_LOGINCD";
            var EnvironmentVarPassword = "SYUKKINNNOSUKE_PASS";
            // コマンドライン引数で clear が指定された場合、ログイン情報をリセットする
            if (args.Length != 0 && (args[0] == "clear" || args[0] == "Clear"))
            {
                // ユーザー環境変数をクリア
                Environment.SetEnvironmentVariable(EnvironmentVarCompanycd, string.Empty, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarLogincd, string.Empty, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarPassword, string.Empty, EnvironmentVariableTarget.User);
                Environment.Exit(0);
            }
            var companycd = Environment.GetEnvironmentVariable(EnvironmentVarCompanycd, EnvironmentVariableTarget.User);
            var logincd = Environment.GetEnvironmentVariable(EnvironmentVarLogincd, EnvironmentVariableTarget.User);
            var password = Environment.GetEnvironmentVariable(EnvironmentVarPassword, EnvironmentVariableTarget.User);
            if (string.IsNullOrEmpty(companycd) || string.IsNullOrEmpty(logincd) || string.IsNullOrEmpty(password))
            {
                Console.WriteLine("お客様IDを入力してください");
                companycd = Console.ReadLine();
                Console.WriteLine("ログインIDを入力してください");
                logincd = Console.ReadLine();
                Console.WriteLine("パスワードを入力してください");
                ConsoleColor origBG = Console.BackgroundColor;
                ConsoleColor origFG = Console.ForegroundColor;
                Console.BackgroundColor = ConsoleColor.DarkBlue;
                Console.ForegroundColor = ConsoleColor.DarkBlue;
                password = Console.ReadLine();
                Console.BackgroundColor = origBG;
                Console.ForegroundColor = origFG;
                // ユーザー環境変数に保存
                Environment.SetEnvironmentVariable(EnvironmentVarCompanycd, companycd, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarLogincd, logincd, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarPassword, password, EnvironmentVariableTarget.User);
            }

            using (IWebDriver webDriver = new ChromeDriver(Environment.CurrentDirectory))
            {
                webDriver.Url = @"https://www.e4628.jp/";
                IWebElement elementCompanycd = webDriver.FindElement(By.CssSelector("#y_companycd"));
                IWebElement elementLogincd = webDriver.FindElement(By.CssSelector("#y_logincd"));
                IWebElement elementPassword = webDriver.FindElement(By.CssSelector("#password"));
                //Thread.Sleep(TimeSpan.FromSeconds(1));
                elementCompanycd.SendKeys(companycd);
                elementLogincd.SendKeys(logincd);
                elementPassword.SendKeys(password);
                elementPassword.Submit();
                Thread.Sleep(TimeSpan.FromSeconds(1));
                // 出社押す
                // 出社、退社押してない状態
                // 出社 #tr_submit_form > table > tbody > tr > td:nth-child(1) > button
                // 退社 #tr_submit_form > table > tbody > tr > td:nth-child(2) > button
                // 出社済み、退社押してない状態
                // 退社 #tr_submit_form > table > tbody > tr > td:nth-child(3) > button
                try
                {
                    // xPath の方がいいのでは？$x('//button[contains(text(), "出社")]')
                    IWebElement elementSyussya = webDriver.FindElement(By.CssSelector("#tr_submit_form > table > tbody > tr > td:nth-child(1) > button"));
                    elementSyussya.Click();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
                // ブラウザを閉じる
                webDriver.Quit();
            }
        }
    }
}