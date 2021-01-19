using ClosedXML.Excel;
using Microsoft.Win32;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using System.IO.Compression;
using System.Windows.Forms;
using OpenQA.Selenium.Support.UI;
using ClosedXML;

namespace Trendlyne_Automation
{
    public class Business_Execution
    {
        IWebDriver driver;
        XLWorkbook wb = new XLWorkbook();
        DataTable dt = new DataTable();
        WebDriverWait wait;
        public void Setup()
        {           
            if (!File.Exists($@"C:\Users\{Environment.UserName}\Desktop\chromedriver.exe"))
            {
                InstalChromeDriver();
            }
            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddArguments(new List<string>() { "headless", "disable-gpu" });
            var chromeDriverService = ChromeDriverService.CreateDefaultService($@"C:\Users\{Environment.UserName}\Desktop");
            chromeDriverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(chromeDriverService, chromeOptions);
            driver.Manage().Window.Maximize();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
            ICapabilities capabilities = ((RemoteWebDriver)driver).Capabilities;
            string driverversion = Convert.ToString((capabilities.GetCapability("chrome") as Dictionary<string, object>)["chromedriverVersion"]).Split('(')[0].Trim();
            CheckChromeVersion(driverversion);
            driver.Navigate().GoToUrl("https://trendlyne.com/fundamentals/screen/raw/create/1/");  
        }

        public void Execute()
        {
            try
            {
                Setup();
                Thread.Sleep(100);
                wait.Until(drv => drv.FindElement(By.ClassName("mobile-banner-4")));
                driver.FindElement(By.ClassName("mobile-banner-4")).FindElements(By.TagName("div"))[1].FindElement(By.TagName("button")).Click();
                driver.FindElement(By.Id("id_login")).SendKeys(Properties.Settings.Default.UserName);
                driver.FindElement(By.Id("id_password")).SendKeys(Properties.Settings.Default.Password);
                driver.FindElements(By.ClassName("tl-btn-blue"))[0].Click();
                driver.Navigate().GoToUrl("https://trendlyne.com/fundamentals/screen/raw/create/1/");
                AddHeader();
                string InvalidQueryList = "Company;URL";
                foreach (DataColumn col in dt.Columns)
                {
                    if (!InvalidQueryList.Contains(col.Caption))
                    {
                        Querywithdata(col.ToString());
                    }

                }
                driver.Quit();
                var worksheet = wb.Worksheets.Add("Result");
                worksheet.Cell(1, 1).InsertTable(dt);
                wb.SaveAs($@"C:\Users\{Environment.UserName}\Downloads\Result{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx");
                MessageBox.Show("Completed", "Completed");
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            
        }

        public void CheckChromeVersion(string driverversion)
        {
            driverversion = driverversion.Split('.')[0];
            string path = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
            string Chromeversion = FileVersionInfo.GetVersionInfo(path).FileVersion.ToString().Split('.')[0].Trim();
            if(driverversion != Chromeversion)
            {
                InstalChromeDriver();
            }
        }

        public void InstalChromeDriver()
        {
            string path = (string)Registry.GetValue(@"HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", "", null);
            string Chromeversion = FileVersionInfo.GetVersionInfo(path).FileVersion.ToString().Split('.')[0].Trim();
            string url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" + Chromeversion.Split('.')[0];
            WebClient webClient = new WebClient();
            string Chromever = webClient.DownloadString(url);
            string url_file = "https://chromedriver.storage.googleapis.com/" + Convert.ToString(Chromever) + "/" + "chromedriver_win32.zip";
            byte[] data = webClient.DownloadData(url_file);
            MemoryStream stream = new MemoryStream(data, true);
            ZipArchive archive = new ZipArchive(stream);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    entry.ExtractToFile($@"C:\Users\{Environment.UserName}\Desktop\" + entry.Name, true);
                }
            }
        }

        public void AddHeader()
        {
            dt.Columns.Add("Company");
            dt.Columns.Add("URL");
            dt.Columns.Add("Institutional holding current Qtr %");
            dt.Columns.Add("Promoter holding latest %");
            dt.Columns.Add("Promoter holding change QoQ %");
            dt.Columns.Add("Promoter pledge change QoQ %");
            dt.Columns.Add("Promoter holding pledge percentage % Qtr");
            dt.Columns.Add("MF holding change QoQ %");
            dt.Columns.Add("MF holding current Qtr %");
            dt.Columns.Add("FII holding change QoQ %");
            dt.Columns.Add("FII holding current Qtr %");
            dt.Columns.Add("Institutional holding change QoQ %");

        }
        public void Querywithdata(string Query)
        {
            if (Query != "Company")
            {
                
                if(Query.Contains("QoQ"))
                {
                    string newquery = Query + " < -5";
                    QueryExecute(newquery);
                    double old_value = -5;
                    for (double i = -5; i <= 20; i = i+ 0.5)
                    {
                        try
                        {
                            string val = driver.FindElement(By.ClassName("alert-warning")).Text;
                            if(val.Contains("Only the first 200 results are being shown"))
                                ErrorQuery(Query, old_value, i);
                        }
                        catch
                        {
                            newquery = Query + " < " + i + " AND " + Query + " > " + old_value;
                            QueryExecute(newquery);
                            old_value = i;
                        }
                    }
                    for (double i = 20; i <= 100; i = i + 2)
                    {
                        old_value = LimitedError(Query, old_value, i);                       
                    }
                }
                else
                {
                    string newquery = Query + " < 0";
                    QueryExecute(newquery);
                    double old_value = 0;
                    for (double i = 2; i <= 100; i = i + 2)
                    {
                        old_value = LimitedError(Query, old_value, i);    
                    }
                }
                
            }
        }

        public double LimitedError(string Query,double old_value,double i)
        {
            try
            {
                string val = driver.FindElement(By.ClassName("alert-warning")).Text;
                if (val.Contains("Only the first 200 results are being shown"))
                    ErrorQuery(Query, old_value, i);
            }
            catch
            {
                string newquery = Query + " < " + i + " AND " + Query + " > " + old_value;
                QueryExecute(newquery);
                old_value = i;
            }
            return old_value;
        }

        public void ErrorQuery(string Query,double start,double end)
        {
            double old_value = start;
            if (Query.Contains("QoQ") && end<=20)
            {
                for(double i=start;i<=end;i=i+0.1)
                {
                    string newquery = Query + " < " + i + " AND " + Query + " > " + old_value;
                    QueryExecute(newquery);
                    old_value = i;
                }
            }
            else
            {
                for (double i = start; i <= end; i = i + 1)
                {
                    string newquery = Query + " < " + i + " AND " + Query + " > " + old_value;
                    QueryExecute(newquery);
                    old_value = i;
                }
            }
        }

        public void QueryExecute(string Query)
        {
            if(Query!="Company")
            {
                var js = (IJavaScriptExecutor)driver;
                wait.Until(drv => drv.FindElement(By.Id("ScreenSQLquerytextbox")));
                driver.FindElement(By.Id("ScreenSQLquerytextbox")).Clear();
                driver.FindElement(By.Id("ScreenSQLquerytextbox")).SendKeys(Query);
                Thread.Sleep(100);
                js.ExecuteScript("window.scrollTo(0,0)");
                wait.Until(drv => drv.FindElement(By.Id("screener_intro5")));
                driver.FindElement(By.Id("screener_intro5")).Click();
                Thread.Sleep(100);
                try
                {
                    wait.Until(drv => drv.FindElement(By.Name("DataTables_Table_0_length")));
                    driver.FindElement(By.Name("DataTables_Table_0_length")).SendKeys("100");
                }
                catch
                {
                    try
                    {
                        js.ExecuteScript("window.scrollTo(0,0)");
                        js.ExecuteScript("window.scrollTo(0,100)");
                        wait.Until(drv => drv.FindElement(By.Name("DataTables_Table_0_length")));
                        driver.FindElement(By.Name("DataTables_Table_0_length")).SendKeys("100");
                    }
                    catch
                    { }
                }
                
                int rowcount = Convert.ToInt32(driver.FindElement(By.Id("DataTables_Table_0_info")).Text.Split(' ')[5]);
                
                int k = 0;
                if (rowcount >= 101)
                    k = 1;
                for (int j = 0; j <= k; j++)
                {
                    wait.Until(drv => drv.FindElement(By.Id("DataTables_Table_0_info")));
                    rowcount = Convert.ToInt32(driver.FindElement(By.Id("DataTables_Table_0_info")).Text.Split(' ')[3]);
                    if(rowcount >=101)
                        rowcount = rowcount - 100;
                    for (int i = 0; i < rowcount; i++)
                    {
                        string asd = Convert.ToString(js.ExecuteScript($"return document.getElementsByClassName('dataTables_scrollBody')[0].getElementsByTagName('tbody')[0].getElementsByTagName('tr')[{i}].innerText"));
                        Thread.Sleep(10);
                        string GetCompany = asd.Split('\r')[0];
                        string[] Splitasd = asd.Split('\t');
                        bool flag = false;
                        string dataquery = Query.Split('<')[0].Trim();
                        foreach (DataRow dat in dt.Rows)
                        {
                            string exitcomp = Convert.ToString(dat["Company"]);
                            if (exitcomp == GetCompany)
                            {
                                dat[dataquery] = Splitasd[3];
                                flag = true;
                                break;
                            }
                        }
                        if (!flag)
                        {
                            DataRow dataRow = dt.NewRow();
                            dataRow["Company"] = GetCompany;
                            dataRow[dataquery] = Splitasd[3];
                            dt.Rows.Add(dataRow);
                        }
                    }
                    driver.FindElement(By.ClassName("next")).FindElement(By.TagName("a")).Click();
                }
            }
            
        }
    }
}
