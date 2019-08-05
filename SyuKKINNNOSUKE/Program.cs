using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Text;
using System.Threading;
using System.Security.Cryptography;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;

namespace SyuKKINNNOSUKE
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var EnvironmentVarCompanycd = "SYUKKINNNOSUKE_COMPANYCD";
            var EnvironmentVarLogincd = "SYUKKINNNOSUKE_LOGINCD";
            var EnvironmentVarPassword = "SYUKKINNNOSUKE_PASS";
            var EnvironmentVarKey = "SYUKKINNNOSUKE_KEY";

            // コマンドライン引数で clear が指定された場合、ログイン情報をリセットする
            if (args.Length != 0 && (args[0] == "clear" || args[0] == "Clear"))
            {
                // ユーザー環境変数をクリア
                Environment.SetEnvironmentVariable(EnvironmentVarCompanycd, string.Empty, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarLogincd, string.Empty, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarPassword, string.Empty, EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarKey, string.Empty, EnvironmentVariableTarget.User);
                Environment.Exit(0);
            }
            var companycd = Environment.GetEnvironmentVariable(EnvironmentVarCompanycd, EnvironmentVariableTarget.User);
            var logincd = Environment.GetEnvironmentVariable(EnvironmentVarLogincd, EnvironmentVariableTarget.User);
            var password = Environment.GetEnvironmentVariable(EnvironmentVarPassword, EnvironmentVariableTarget.User);
            var key = Environment.GetEnvironmentVariable(EnvironmentVarKey, EnvironmentVariableTarget.User);


            // ログイン情報が保存されていない場合、入力してもらう
            if (string.IsNullOrEmpty(companycd) || string.IsNullOrEmpty(logincd) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(key))
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

                // 読めなくする
                key = Guid.NewGuid().ToString("N").Substring(0, password.Length);
                Environment.SetEnvironmentVariable(EnvironmentVarPassword, ToUnreadable(password, key), EnvironmentVariableTarget.User);
                Environment.SetEnvironmentVariable(EnvironmentVarKey, key, EnvironmentVariableTarget.User);

                if (NotExistsStartupFile())
                {
                    Console.WriteLine("スタートアップに登録しますか？(y/(any))");
                    var res = Console.ReadLine();
                    if (res.ToLower() == "y")
                    {
                        SetStartUpFile();
                    }
                }
            }
            else
            {
                // 読めるようにする
                password = ToCanRead(password, key);
            }

            // chromedriver.exe の更新をどうするか。
            //using (IWebDriver webDriver = new ChromeDriver(Environment.CurrentDirectory))
            using (IWebDriver webDriver = new ChromeDriver())
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


                // コマンドライン引数で shutdown が指定された場合、退社を押してシャットダウンする
                if (args.Length != 0 && (args[0] == "shutdown"))
                {
                    try
                    {
                        IWebElement elementTaisya = webDriver.FindElement(By.XPath("//button[contains(text(),'退社')]"));
                        elementTaisya.Click();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }

                    Thread.Sleep(TimeSpan.FromSeconds(5));

                    System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();

                    psi.FileName = "shutdown.exe";
                    psi.Arguments = "-s -f";

                    System.Diagnostics.Process p = System.Diagnostics.Process.Start(psi);
                }
                else
                {
                    // 出社押す
                    // 出社、退社押してない状態
                    // 出社 #tr_submit_form > table > tbody > tr > td:nth-child(1) > button
                    // 退社 #tr_submit_form > table > tbody > tr > td:nth-child(2) > button
                    // 出社済み、退社押してない状態
                    // 退社 #tr_submit_form > table > tbody > tr > td:nth-child(3) > button
                    try
                    {
                        // xPath の方がいいのでは？IWebElement elementSyussya = webDriver.FindElement(By.XPath("//button[contains(text(),'出社')]"));
                        IWebElement elementSyussya = webDriver.FindElement(By.CssSelector("#tr_submit_form > table > tbody > tr > td:nth-child(1) > button"));
                        elementSyussya.Click();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }

                // ブラウザを閉じる
                webDriver.Quit();
            }
        }

        // 読めなくする
        private static string ToUnreadable(string targetCharacter, string xorKey)
        {
            byte[] targetCharacterConvertedByte = Encoding.ASCII.GetBytes(targetCharacter);
            byte[] xorKeyConvertedByte = Encoding.ASCII.GetBytes(xorKey);

            int j = 0;
            string targetCharacterToUnreadable = string.Empty;
            for (int i = 0; i < targetCharacterConvertedByte.Length; i++)
            {
                if (j < xorKeyConvertedByte.Length)
                {
                    j++;
                }
                else
                {
                    j = 1;
                }

                targetCharacterConvertedByte[i] = (byte)(targetCharacterConvertedByte[i] ^ xorKeyConvertedByte[j - 1]);
                string ValueConvertedHex = Convert.ToString(targetCharacterConvertedByte[i], 16).PadLeft(2, '0');
                targetCharacterToUnreadable += ValueConvertedHex;
            }
            return targetCharacterToUnreadable;
        }

        // 読めなくしたものを読めるようにする
        private static string ToCanRead(string targetCharacter, string xorKey)
        {
            byte[] targetCharacterConvertedByte = new byte[targetCharacter.Length / 2];
            byte[] xorKeyConvertedByte = Encoding.ASCII.GetBytes(xorKey);

            for (int i = 0; i < targetCharacter.Length / 2; i++)
            {
                int ValueConvertedDecimalNumber = Convert.ToInt32(targetCharacter.Substring(i * 2, 2), 16);
                targetCharacterConvertedByte[i] = byte.Parse(ValueConvertedDecimalNumber.ToString());
            }

            int j = 0;
            for (int i = 0; i < targetCharacterConvertedByte.Length; i++)
            {
                if (j < xorKeyConvertedByte.Length)
                {
                    j++;
                }
                else
                {
                    j = 1;
                }
                targetCharacterConvertedByte[i] = (byte)(targetCharacterConvertedByte[i] ^ xorKeyConvertedByte[j - 1]);
            }
            return Encoding.ASCII.GetString(targetCharacterConvertedByte);
        }

        private static void SetStartUpFile()
        {
            var startUpFilePath = StartupFilePath();
            var applicationName = ApplicationName();
            var exePath = $@"{Directory.GetCurrentDirectory()}\{applicationName}.exe";

            object shell = null;
            object shortcut = null;
            try
            {
                // 72C24DD5-D70A-438B-8A42-98424B88AFB8(Windwos Script Hostを使用するオブジェクトの生成)
                Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"));
                shell = Activator.CreateInstance(t);

                // ショートカットの作成
                shortcut = t.InvokeMember("CreateShortcut",
                    BindingFlags.InvokeMethod, null, shell,
                    new object[] { startUpFilePath });

                // ショートカットの実行パス設定
                t.InvokeMember("TargetPath",
                    BindingFlags.SetProperty, null, shortcut,
                    new object[] { exePath });

                // ショートカットファイルのアイコン設定
                t.InvokeMember("IconLocation",
                    BindingFlags.SetProperty, null, shortcut,
                    new object[] { exePath + ",0" });

                // 保存
                t.InvokeMember("Save",
                    BindingFlags.InvokeMethod,
                    null, shortcut, null);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                if (shortcut != null)
                {
                    Marshal.FinalReleaseComObject(shortcut);
                    shortcut = null;
                }

                if (shell != null)
                {
                    Marshal.FinalReleaseComObject(shell);
                    shell = null;
                }
            }
        }

        private static bool NotExistsStartupFile()
        {
            return !File.Exists(StartupFilePath());
        }

        private static string StartupFilePath()
        {
            var applicationName = ApplicationName();
            var startUpFileName = $"{applicationName}.lnk";
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), startUpFileName);
        }

        private static string ApplicationName()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var attribute = Attribute.GetCustomAttribute(assembly, typeof(AssemblyTitleAttribute)) as AssemblyTitleAttribute;
            return attribute.Title;
        }

    }
}