using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CPM
{
    public partial class Form1 : Form
    {
        // 초기 메인페이지 URL
        private static string MAIN_URL = "https://www.coupang.com/";
        // 작업쓰래드 
        private Thread wokerThread;
        // 웹드라이버
        IWebDriver driver;
        // 웹엘레먼트
        IWebElement element;
        // ADB
        private static string cmd = Application.StartupPath + @"\adb\adb.exe";
        // PC UserAgent List
        List<string> userAgentList = new List<string>();
        // 시스템컨피그
        Dictionary<string, string> systemConfig = new Dictionary<string, string>();
        // 라이센스
        string license = System.IO.File.ReadAllLines(Application.StartupPath + @"\Config\License.txt")[0];
        // 라이센스키
        string LICENSE_KEY = "rg9gDHJtjfpuJ4FZ";

        public Form1()
        {
            InitializeComponent();
            SetConfig();
            Init();
           
        }

        private void Init()
        {
            label1.Text = GetExternalIPAddress();

            // 라이센스복호화
            license = AESDecrypt128(license, LICENSE_KEY);
        }

        // 컨피그 세팅
        private void SetConfig()
        {
            //글로벌 세팅
            string[] line = System.IO.File.ReadAllLines(Application.StartupPath + @"\Config\System_Config.txt");
            if (line.Length > 0)
            {
                for (int i = 0; i < line.Length; i++)
                {
                    if (!line[i].StartsWith("#") && line[i].Trim().Length > 0)
                    {
                        string[] c = line[i].Split('=');
                        systemConfig.Add(c[0], c[1]);
                    }
                }
            }

            // 유저에이전트
            line = System.IO.File.ReadAllLines(Application.StartupPath + @"\Config\User_Agent_List.txt");
            if (line.Length > 0)
            {
                for (int i = 0; i < line.Length; i++)
                {
                    userAgentList.Add(line[i]);
                }
            }
        }

        private IWebDriver MakeDriver()
        {
            return MakeDriver(false, "Mozilla/5.0 (iPad; CPU OS 6_0 like Mac OS X) AppleWebKit/536.26 (KHTML, like Gecko) Version/6.0 Mobile/10A5355d Safari/8536.25");
        }
        private IWebDriver MakeDriver(bool isHide, string userAgent)
        {
            ChromeOptions cOptions = new ChromeOptions();
            cOptions.AddArguments("disable-infobars");
            cOptions.AddArguments("--js-flags=--expose-gc");
            cOptions.AddArguments("--enable-precise-memory-info");
            cOptions.AddArguments("--disable-popup-blocking");
            cOptions.AddArguments("--disable-default-apps");
            cOptions.AddArguments("--window-size=1280,900");
            cOptions.AddArguments("--incognito");

            if (isHide)
            {
                cOptions.AddArguments("headless");
            }


            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;

            cOptions.AddArgument("--user-agent=" + userAgent);

            // 셀레니움실행
            IWebDriver driver = new ChromeDriver(chromeDriverService, cOptions);
            //driver.Manage().Window.Maximize();
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);

            return driver;
        }

        // 폼이 로드되었을 때
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        // 시작버튼 클릭
        private void button1_Click(object sender, EventArgs e)
        {
            Button button = ((Button)sender);
            if (button.Text.Equals("START"))
            {
                button.Text = "PENDING";
                button.BackColor = Color.FromArgb(128, 128, 128);
                button.Enabled = false;

                // 작업쓰래드 시작
                wokerThread = new Thread(() => DoWork());
                wokerThread.Start();
            }
            else
            {
                button.Text = "START";
                button.BackColor = Color.FromArgb(47, 153, 39);

                new Thread(() => {
                    wokerThread.Abort();
                    CloseBrowser();
                }).Start();
            }
        }

        private string getProperty(string key)
        {
            return systemConfig[key];
        }

        private int getIntProperty(string key)
        {
            return int.Parse(systemConfig[key]);
        }


        private string[] getRangeStringProperty(string key)
        {
            string[] r = new string[2];
            r[0] = systemConfig[key].Split('-')[0];
            r[1] = systemConfig[key].Split('-')[1];
            return r;
        }

        private int[] getRangeProperty(string key)
        {
            int[] r = new int[2];
            r[0] = int.Parse(systemConfig[key].Split('-')[0]);
            r[1] = int.Parse(systemConfig[key].Split('-')[1]);
            return r;
        }

        private int getRandomRangeProperty(string key)
        {
            int[] r = new int[2];
            r[0] = int.Parse(systemConfig[key].Split('-')[0]);
            r[1] = int.Parse(systemConfig[key].Split('-')[1]);
            return new Random().Next(r[0], r[1]);
        }


        private void Scroll(int c)
        {
            // 스크롤 다운icon-sprite icon-gender-f
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            for (int i = 0; i < c; i++)
            {
                js.ExecuteScript("window.scrollBy(0, 200)", "");
                Thread.Sleep(500);
            }
        }

        private void ScrollTo(int c)
        {
            // 스크롤 다운icon-sprite icon-gender-f
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript(string.Format("window.scrollBy(0, {0})", c));
            Thread.Sleep(500);
        }

        private void Scroll(string script, int c)
        {
            // 스크롤 다운icon-sprite icon-gender-f
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            for (int i = 0; i < c; i++)
            {
                js.ExecuteScript(script);
                Thread.Sleep(500);
            }
        }

        private void ExecuteJS(string script)
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            js.ExecuteScript(script);
        }


        private static void WaitForVisible(IWebDriver driver, By by, int seconds)
        {            
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(seconds));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));// instead of id u can use cssSelector or xpath of ur element.
        }

        // 백그라운드 작업
        private void DoWork()
        {
            // 라이센스 체크
            isValidLicense();

            // 드라이버 생성
            driver = MakeDriver(false, userAgentList[new Random().Next(userAgentList.Count)]);
            // 루프횟수
            int searchLoopCnt  = getIntProperty("search.loop.cnt");
            // 검색수행시간
            string[] searchAvaliableHm = getRangeStringProperty("search.available.hm");

            for (int i = 0; i < searchLoopCnt; i++)
            {

                Console.WriteLine("===========================[ {0} / {1} ]=============================", i, searchLoopCnt);
                Console.WriteLine(string.Format("[INFO] 프로그램 동작시간 체크"));
                Console.WriteLine(string.Format("[INFO] 현재시간(Hm) : {0}", DateTime.Now.ToString("HHmm")));

                // 프로그램이 동작해야하는 시간인지 체크
                while (!(DateTime.Now.ToString("HHmm").CompareTo(searchAvaliableHm[0]) >= 0
                && DateTime.Now.ToString("HHmm").CompareTo(searchAvaliableHm[1]) <= 0)) 
                {
                    Console.WriteLine("[INFO] 프로그램 동작 시간이 아닙니다.");
                    Thread.Sleep(TimeSpan.FromMinutes(1));
                }

                // 페이지 이동
                Console.WriteLine(string.Format("[INFO] 메인페이지 이동"));
                driver.Navigate().GoToUrl(MAIN_URL);
                Thread.Sleep(TimeSpan.FromSeconds(2));

                // 정지 가능하도록 처리
                this.Invoke(new Action(delegate ()
                {
                    button1.Text = "STOP";
                    button1.Enabled = true;
                    button1.BackColor = Color.FromArgb(255, 50, 50);
                }));

                // 검색창클릭
                Console.WriteLine(string.Format("[INFO] 검색창 클릭"));
                element = driver.FindElement(By.CssSelector("#headerSearchKeyword"));
                element.Click();

                // 검색어입력           
                Console.WriteLine(string.Format("[INFO] 검색어::{0}", getProperty("search.keyword")));
                SendKey(element, getProperty("search.keyword"), 250);

                // 검색어 입력 후 대기 시간
                Console.WriteLine(string.Format("[INFO] 검색어 입력 후 대기"));
                Thread.Sleep(TimeSpan.FromSeconds(getIntProperty("search.keyword.input.wait")));

                // 검색버튼 클릭
                Console.WriteLine(string.Format("[INFO] 검색 버튼 클릭"));
                driver.FindElement(By.CssSelector("#headerSearchBtn")).Click();

                // 검색 후 대기 시간 ( 설정값)                
                int searchAfterWait = getIntProperty("search.after.wait");
                Console.WriteLine(string.Format("[INFO] 검색 후 대기 시간::{0}", searchAfterWait));
                Thread.Sleep(TimeSpan.FromSeconds(searchAfterWait));

                // 상품리스트가 보일때 까지 대기
                WaitForVisible(driver, By.CssSelector(".search-query-result"), 10);

                // 새창이아닌 현재창에 상품페이지가 뜨도록 타겟을 변경
                ExecuteJS("$('.search-product-link').attr('target', '_self');");

                // 상품검색
                Console.WriteLine(string.Format("[INFO] 상품 리스트 중 랜덤 클릭"));
                List<IWebElement> searchProduct = driver.FindElements(By.CssSelector(".search-product:not(.search-product__ad-badge)")).ToList();
                element = searchProduct[new Random().Next(searchProduct.Count)];
                element.Click();

                // 상품이 보일때 까지 대기
                WaitForVisible(driver, By.CssSelector("#headerSearchKeyword"), 10);

                // 상품 클릭 후 대기
                Console.WriteLine(string.Format("[INFO] 상품 클릭 후 대기"));
                Thread.Sleep(TimeSpan.FromSeconds(getIntProperty("product.click.after.wait")));

                // 스크롤 처리 ( 설정값)            
                int productScrollRange = getRandomRangeProperty("product.scroll.range");
                Console.WriteLine(string.Format("[INFO] 스크롤 처리::{0}", productScrollRange));
                Scroll(productScrollRange);

                // 스크롤 후 페이지 체류 (설정값)
                int productVisitRange = getRandomRangeProperty("product.visit.range");
                Console.WriteLine(string.Format("[INFO] 체류::{0}", productVisitRange));
                Thread.Sleep(TimeSpan.FromSeconds(productVisitRange));

                // 뒤로가기
                Console.WriteLine(string.Format("[INFO] 뒤로가기"));
                ExecuteJS("window.history.go(-1)");

                // 검색창이 보일때까지 대기
                WaitForVisible(driver, By.CssSelector("#headerSearchKeyword"), 10);
                Thread.Sleep(TimeSpan.FromSeconds(3));

                // 새창이아닌 현재창에 상품페이지가 뜨도록 타겟을 변경
                ExecuteJS("$('.search-product-link').attr('target', '_self');");

                // 1행 1열의 광고클릭
                Console.WriteLine(string.Format("[INFO] 광고 클릭"));
                element = driver.FindElement(By.CssSelector(".search-product__ad-badge"));

                if (element == null)
                {
                    Thread.Sleep(TimeSpan.FromSeconds(getRandomRangeProperty("ad.not.visible.range")));
                    continue;
                }
                else
                {
                    element.Click();
                }

                // 상품이 보일때 까지 대기
                WaitForVisible(driver, By.CssSelector("#headerSearchKeyword"), 10);
                Thread.Sleep(TimeSpan.FromSeconds(3));

                // 상품 클릭 후 대기
                Console.WriteLine(string.Format("[INFO] 상품 클릭 후 대기"));
                Thread.Sleep(TimeSpan.FromSeconds(getIntProperty("product.click.after.wait")));

                // 스크롤 처리 ( 설정값)            
                productScrollRange = getRandomRangeProperty("product.scroll.range");
                Console.WriteLine(string.Format("[INFO] 스크롤 처리::{0}", productScrollRange));
                Scroll(productScrollRange);

                // 스크롤 후 페이지 체류 (설정값)
                productVisitRange = getRandomRangeProperty("product.visit.range");
                Console.WriteLine(string.Format("[INFO] 체류::{0}", productVisitRange));
                Thread.Sleep(TimeSpan.FromSeconds(productVisitRange));

                // 쿠키삭제
                Console.WriteLine(string.Format("[INFO] 쿠키 삭제"));
                DeleteCookie();

                // 아이피변경
                Console.WriteLine(string.Format("[INFO] 아이피 변경"));
                ChangeIP();
            }

            // 브라우저 종료처리
            Console.WriteLine(string.Format("[INFO] 루프횟수 도달하여 종료처리"));
            CloseBrowser();
        }

        // 쿠키 삭제
        private void DeleteCookie()
        {
            driver.Manage().Cookies.DeleteAllCookies();
        }

        private void SendKey(IWebElement element, string keyword, int delay)
        {
            char[] charArray = keyword.ToCharArray();
            foreach (char c in charArray)
            {
                element.SendKeys(c.ToString());
                Thread.Sleep(delay);
            }
        }

        public static string GetExternalIPAddress()
        {
            string url = "https://ip.pe.kr/";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";

            string resResult = string.Empty;
            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            {
                StreamReader readerPost = new StreamReader(response.GetResponseStream(), System.Text.Encoding.UTF8, true);
                resResult = readerPost.ReadToEnd();
            }

            Regex regexp = new Regex(@"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}");
            string IP = regexp.Matches(resResult)[0].ToString();


            //int ingNO = resResult.IndexOf("조회 IP");
            //string varTemp = resResult.Substring(ingNO, 50);
            //string realIP = Parsing(Parsing(varTemp, "Current IP Address: ", 1), "</body>", 0).Trim();
            return IP;

        }


        // 아이피 변경
        private void ChangeIP()
        {
            Console.WriteLine("[INFO] 데이터 OFF");
            DisAbleData();
            Thread.Sleep(3000);

            Console.WriteLine("[INFO] 데이터 ON");
            EnAbleData();
            Thread.Sleep(3000);

            label1.Text = GetExternalIPAddress();
        }

        // 데이터활성
        public static void EnAbleData()
        {             
            Process process = new Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo();
            process.StartInfo.FileName = cmd;
            process.StartInfo.Arguments = " shell svc data enable";
            process.StartInfo.UseShellExecute = false; 
            process.StartInfo.RedirectStandardInput = true; 
            process.StartInfo.RedirectStandardOutput = true;  
            process.StartInfo.RedirectStandardError = true; 
            process.StartInfo.CreateNoWindow = true; 
            process.Start();
            process.WaitForExit();

        }

        // 데이터비활성
        public static void DisAbleData()
        {
            Process process = new Process();
            process.StartInfo = new System.Diagnostics.ProcessStartInfo();
            process.StartInfo.FileName = cmd; 
            process.StartInfo.Arguments = " shell svc data disable";
            process.StartInfo.UseShellExecute = false; 
            process.StartInfo.RedirectStandardInput = true; 
            process.StartInfo.RedirectStandardOutput = true; 
            process.StartInfo.RedirectStandardError = true; 
            process.StartInfo.CreateNoWindow = true;
            process.Start();
            process.WaitForExit();

        }

        //  브라우저 종료
        private void CloseBrowser()
        {

            if (driver != null)
            {
                try
                {
                    this.Invoke(new Action(delegate ()
                    {
                        button1.Text = "START";
                        button1.BackColor = Color.FromArgb(47, 153, 39);
                    }));

                    driver.Quit();                    
                }
                catch (Exception ex)
                {
                    if (ex is InvalidOperationException || ex is NoSuchWindowException)
                    {
                        ProcessKillByName("chromedriver");
                    }
                }
            }
            else
            {
                ProcessKillByName("chromedriver");
            }
        }


        public static void ProcessKillByName(string name)
        {
            Process[] processList = Process.GetProcessesByName(name);
            if (processList.Length > 0)
            {
                for (int i = 0; i < processList.Length; i++)
                {
                    processList[i].Kill();
                }
            }
        }

        public List<string> GetMacAddr() {

            List<string> list = new List<string>();

            foreach (NetworkInterface networkInterface in NetworkInterface.GetAllNetworkInterfaces())
            {
                list.Add(networkInterface.GetPhysicalAddress().ToString());
            }

            return list;
        }

        public void isValidLicense() {
                    
            if (license.Length == 0) {
                throw new Exception("License Not Vaild!");
            }
            string[] licenseArr = license.Split('^');
            if (licenseArr.Length != 3) {
                throw new Exception("License Not Vaild!");
            }
            if (!licenseArr[0].Equals("CPM")) 
            {
                throw new Exception("License Not Vaild!");
            }
            if (!GetMacAddr().Contains( licenseArr[1]))
            {
                throw new Exception("License Not Vaild!");
            }
            if (licenseArr[2].CompareTo(DateTime.Now.ToString("yyyyMMddHHmmss")) < 0)
            {
                throw new Exception("License Not Vaild!");
            }            
        }

        //AE_S128 복호화
        public String AESDecrypt128(String Input, String key)
        {
            RijndaelManaged RijndaelCipher = new RijndaelManaged();

            byte[] EncryptedData = Convert.FromBase64String(Input);
            byte[] Salt = Encoding.ASCII.GetBytes(key.Length.ToString());

            PasswordDeriveBytes SecretKey = new PasswordDeriveBytes(key, Salt);
            ICryptoTransform Decryptor = RijndaelCipher.CreateDecryptor(SecretKey.GetBytes(32), SecretKey.GetBytes(16));
            MemoryStream memoryStream = new MemoryStream(EncryptedData);
            CryptoStream cryptoStream = new CryptoStream(memoryStream, Decryptor, CryptoStreamMode.Read);

            byte[] PlainText = new byte[EncryptedData.Length];

            int DecryptedCount = cryptoStream.Read(PlainText, 0, PlainText.Length);

            memoryStream.Close();
            cryptoStream.Close();

            string DecryptedData = Encoding.Unicode.GetString(PlainText, 0, DecryptedCount);

            return DecryptedData;
        }

       
    }
}
