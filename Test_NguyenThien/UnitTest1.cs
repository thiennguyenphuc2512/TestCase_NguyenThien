using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelDataReader;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using NUnit.Framework;
using OpenQA.Selenium.Support.UI;



namespace Test_NguyenThien
{
    [TestFixture]
    public class Tests
    {
        private IWebDriver driver;
        private DataTable testData;
        private string filePath = "E:\\Huflit\\BDTKPM\\TestCase_NguyenThien.xlsx";

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            testData = ReadTestData(filePath);
        }

        public DataTable ReadTestData(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                // In danh sách sheet
                foreach (DataTable tbl in result.Tables) // Đổi 'table' thành 'tbl'
                {
                    Console.WriteLine("- " + tbl.TableName);
                }

                var table = result.Tables["TestCase_NguyenThien"];

                if (table == null)
                {
                    throw new Exception("Sheet 'TestCase_NguyenThien' không tồn tại trong file Excel.");
                }

                // In danh sách cột
                Console.WriteLine("Danh sách cột trong Excel:");
                foreach (DataColumn column in table.Columns)
                {
                    Console.WriteLine("- " + column.ColumnName);
                }

                return table;
            }
        }


        public void WriteTestResult(string testTitle, string actualResult)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial; // Hoặc LicenseContext.NonCommercial nếu bạn sử dụng bản miễn phí
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets["TestCase_NguyenThien"];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        if (worksheet.Cells[row, 2].Text.Trim() == testTitle) // Tìm test case trong cột 2 (TestTitle)
                        {
                            string expectedUrl = worksheet.Cells[row, 5].Text.Trim(); // ExpectedResults ở cột 5

                            worksheet.Cells[row, 6].Value = actualResult; // Ghi ActualResult vào cột 6
                            worksheet.Cells[row, 7].Value = actualResult == expectedUrl ? "Pass" : "Fail"; // Ghi Status vào cột 7
                            break;
                        }

                    }
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi khi ghi file Excel: " + ex.Message);
            }
        }


        [Test]
        public void RegisterTest_Success()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testData = row["TestData"].ToString().Trim();
                string expectedUrl = row["ExpectedResults"].ToString().Trim(); // Đọc expectedUrl từ Excel

                if (testTitle == "Đăng ký thành công")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000/register");

                    // Tách dữ liệu test từ cột TestData
                    var dataLines = testData.Split('\n');
                    string email = dataLines[0].Split(":")[1].Trim();
                    string password = dataLines[1].Split(":")[1].Trim();
                    string name = dataLines[2].Split(":")[1].Trim();
                    string phone = dataLines[3].Split(":")[1].Trim();

                    // Nhập dữ liệu vào form
                    driver.FindElement(By.Id("email")).SendKeys(email);
                    driver.FindElement(By.Id("password")).SendKeys(password);
                    driver.FindElement(By.Name("name")).SendKeys(name);
                    driver.FindElement(By.Name("phone")).SendKeys(phone);

                    // Click vào checkbox đồng ý điều khoản
                    IWebElement checkBox = driver.FindElement(By.CssSelector("input.ant-checkbox-input"));
                    if (!checkBox.Selected) checkBox.Click();

                    // Click vào nút "Create Account"
                    driver.FindElement(By.CssSelector("button.my-2.ml-8.hover\\:scale-105.btn.btn-primary")).Click();

                    // Chờ trang chuyển hướng đến login (tối đa 5 giây)
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    wait.Until(drv => drv.Url == expectedUrl);

                    // Kiểm tra kết quả test
                    string actualResult = driver.Url;

                    // Ghi kết quả vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    // Xác nhận kết quả test
                    Assert.AreEqual(expectedUrl, actualResult);

                    Console.WriteLine("Đã chuyển hướng thành công đến trang đăng nhập.");
                }
            }
        }
        [Test]
        public void RegisterTest_Fail()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testData = row["TestData"].ToString().Trim();
                string expectedUrl = row["ExpectedResults"].ToString().Trim(); // Đọc expectedUrl từ Excel

                if (testTitle == "Đăng ký thiếu gmail")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000/register");

                    // Tách dữ liệu test từ cột TestData
                    var dataLines = testData.Split('\n');
                    string email = dataLines[0].Split(":")[1].Trim();
                    string password = dataLines[1].Split(":")[1].Trim();
                    string name = dataLines[2].Split(":")[1].Trim();
                    string phone = dataLines[3].Split(":")[1].Trim();

                    // Nhập dữ liệu vào form
                    driver.FindElement(By.Id("email")).SendKeys(email);
                    driver.FindElement(By.Id("password")).SendKeys(password);
                    driver.FindElement(By.Name("name")).SendKeys(name);
                    driver.FindElement(By.Name("phone")).SendKeys(phone);

                    // Click vào checkbox đồng ý điều khoản
                    IWebElement checkBox = driver.FindElement(By.CssSelector("input.ant-checkbox-input"));
                    if (!checkBox.Selected) checkBox.Click();

                    // Click vào nút "Create Account"
                    driver.FindElement(By.CssSelector("button.my-2.ml-8.hover\\:scale-105.btn.btn-primary")).Click();

                    // Chờ xem có bị chuyển hướng không (tối đa 5 giây)
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    bool isRedirected;
                    try
                    {
                        wait.Until(drv => drv.Url != expectedUrl);
                        isRedirected = true;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        isRedirected = false; // Timeout nghĩa là trang vẫn ở register
                    }

                    // Kiểm tra kết quả
                    string actualResult = driver.Url;
                    bool isTestPass = !isRedirected; // Nếu KHÔNG chuyển hướng thì PASS

                    // 📝 Ghi kết quả vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    if (isTestPass)
                    {
                        Console.WriteLine("✅ Đăng ký thất bại như mong đợi, vẫn ở trang đăng ký.");
                        WriteTestResult(testTitle, actualResult); // Ghi kết quả vào file Excel
                        return; // Dừng test case ở đây, không dùng Assert.Pass()
                    }
                    else
                    {
                        Console.WriteLine("❌ Đã bị chuyển hướng, test case failed.");
                        Assert.Fail("Test case failed: Đã chuyển hướng khỏi http://localhost:3000/register");
                    }

                }
            }
        }
        [Test]
        public void LanguageSwitchTest_EN()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testSteps = row["Test Steps"].ToString().Trim(); // Đọc Test Steps từ Excel
                string testData = row["TestData"].ToString().Trim();  // Đọc TestData từ Excel
                string expectedResults = row["ExpectedResults"].ToString().Trim();  // Đọc ExpectedResults từ Excel

                if (testTitle == "Chuyển đổi ngôn ngữ  tiếng anh")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000");  // Đảm bảo trang web được truy cập
                    Thread.Sleep(5000);

                    // Tìm phần tử để chuyển ngôn ngữ
                    IWebElement languageSwitcher = driver.FindElement(By.XPath("//p[contains(text(), 'EN')]")); // Tìm nút "EN"
                    languageSwitcher.Click();  // Click để chuyển sang ngôn ngữ tiếng Anh
                    Thread.Sleep(5000);

                    // Đợi cho giao diện chuyển ngôn ngữ, ví dụ tìm phần tử có chữ "Booking" sau khi chuyển ngôn ngữ
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    wait.Until(drv => drv.FindElement(By.XPath("//li[contains(text(), 'Booking')]")));  // Đảm bảo chữ "Booking" xuất hiện trên trang
                    Thread.Sleep(5000);

                    // Lấy phần tử "Booking" từ trang và so sánh
                    string actualResult = driver.FindElement(By.XPath("//li[contains(text(), 'Booking')]")).Text.Trim();

                    // Ghi kết quả vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    // So sánh với ExpectedResults
                    bool isPass = actualResult.Equals(expectedResults, StringComparison.OrdinalIgnoreCase);

                    // Cập nhật kết quả Pass/Fail vào cột Status
                    string status = isPass ? "Pass" : "Fail";

                    // Ghi kết quả thực tế (ActualResult) và trạng thái (Status) vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    // Xác nhận kết quả test
                    Assert.AreEqual(expectedResults, actualResult);  // So sánh ActualResult với ExpectedResults

                    Console.WriteLine($"Giao diện trang web đã chuyển sang Tiếng Anh với chữ '{expectedResults}'. Kết quả: {status}");
                }
            }
        }
        [Test]
        public void LanguageSwitchTest_VIE()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testSteps = row["Test Steps"].ToString().Trim(); // Đọc Test Steps từ Excel
                string testData = row["TestData"].ToString().Trim();  // Đọc TestData từ Excel
                string expectedResults = row["ExpectedResults"].ToString().Trim();  // Đọc ExpectedResults từ Excel

                if (testTitle == "Chuyển đổi ngôn ngữ  tiếng việt")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000");  // Đảm bảo trang web được truy cập
                    Thread.Sleep(5000);

                    // Tìm phần tử để chuyển ngôn ngữ
                    IWebElement languageSwitcher = driver.FindElement(By.XPath("//p[contains(text(), 'VIE')]")); // Tìm nút "EN"
                    languageSwitcher.Click();  // Click để chuyển sang ngôn ngữ tiếng Anh
                    Thread.Sleep(5000);

                    // Đợi cho giao diện chuyển ngôn ngữ, ví dụ tìm phần tử có chữ "Booking" sau khi chuyển ngôn ngữ
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    wait.Until(drv => drv.FindElement(By.XPath("//li[contains(text(), 'Đặt phòng')]")));  // Đảm bảo chữ "Booking" xuất hiện trên trang
                    Thread.Sleep(5000);

                    // Lấy phần tử "Booking" từ trang và so sánh
                    string actualResult = driver.FindElement(By.XPath("//li[contains(text(), 'Đặt phòng')]")).Text.Trim();

                    // Ghi kết quả vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    // So sánh với ExpectedResults
                    bool isPass = actualResult.Equals(expectedResults, StringComparison.OrdinalIgnoreCase);

                    // Cập nhật kết quả Pass/Fail vào cột Status
                    string status = isPass ? "Pass" : "Fail";

                    // Ghi kết quả thực tế (ActualResult) và trạng thái (Status) vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    // Xác nhận kết quả test
                    Assert.AreEqual(expectedResults, actualResult);  // So sánh ActualResult với ExpectedResults

                    Console.WriteLine($"Giao diện trang web đã chuyển sang Tiếng Anh với chữ '{expectedResults}'. Kết quả: {status}");
                }
            }
        }
        [Test]
        public void LoginTest_Success()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testData = row["TestData"].ToString().Trim();
                string expectedUrl = row["ExpectedResults"].ToString().Trim(); // URL mong đợi sau khi login thành công

                if (testTitle == "Đăng nhập với đúng với tài khoản đăng ký")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000/login");

                    // Tách dữ liệu test từ cột TestData
                    var dataLines = testData.Split('\n');
                    if (dataLines.Length < 2)
                    {
                        Assert.Fail("Dữ liệu test không hợp lệ.");
                        return;
                    }

                    string email = dataLines[0].Split(":")[1].Trim();
                    string password = dataLines[1].Split(":")[1].Trim();

                    // Nhập dữ liệu vào form đăng nhập
                    driver.FindElement(By.Id("email")).SendKeys(email);
                    driver.FindElement(By.Id("password")).SendKeys(password);

                    // Click vào nút "Đăng nhập"
                    driver.FindElement(By.CssSelector("button.btn-primary")).Click();

                    // Chờ trang chuyển hướng
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    bool isRedirected;
                    try
                    {
                        wait.Until(drv => drv.Url == expectedUrl);
                        isRedirected = true;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        isRedirected = false; // Nếu trang không chuyển hướng, nghĩa là test thất bại
                    }

                    // Lấy kết quả thực tế
                    string actualResult = driver.Url;
                    bool isTestPass = isRedirected; // Nếu chuyển hướng đúng, test PASS

                    // 📝 Ghi kết quả vào file Excel
                    WriteTestResult(testTitle, actualResult);

                    if (isTestPass)
                    {
                        Console.WriteLine("✅ Đăng nhập thành công, đúng như mong đợi.");
                        return;
                    }
                    else
                    {
                        Console.WriteLine("❌ Đăng nhập thất bại, không chuyển hướng đúng.");
                        Assert.Fail($"Test case failed: Không chuyển hướng đến {expectedUrl}");
                    }
                }
            }
        }

        [Test]
        public void LoginTest_Fail_NoEmail()
        {
            if (testData == null)
            {
                Assert.Fail("Không thể đọc dữ liệu từ file Excel.");
                return;
            }

            foreach (DataRow row in testData.Rows)
            {
                string testTitle = row["TestTitle"].ToString().Trim();
                string testData = row["TestData"].ToString().Trim();
                string expectedErrorMessage = row["ExpectedResults"].ToString().Trim(); // Lấy ExpectedResult từ Excel

                if (testTitle == "Đăng nhập bỏ trống tài khoản")
                {
                    driver.Navigate().GoToUrl("http://localhost:3000/login");

                    // Tách dữ liệu test từ cột TestData
                    var dataLines = testData.Split('\n');
                    if (dataLines.Length < 1)
                    {
                        Assert.Fail("Dữ liệu test không hợp lệ.");
                        return;
                    }

                    string password = dataLines[0].Split(":")[1].Trim(); // Lấy password, bỏ trống email

                    // Chỉ nhập password, không nhập email
                    driver.FindElement(By.Id("password")).SendKeys(password);

                    // Click vào nút "Đăng nhập"
                    driver.FindElement(By.CssSelector("button.btn-primary")).Click();

                    // 🕵️ Chờ xem có thông báo lỗi xuất hiện không
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    string actualErrorMessage = "";

                    try
                    {
                        var errorMessageElement = wait.Until(drv => drv.FindElement(By.CssSelector(".text-\\[15px\\]")));
                        actualErrorMessage = errorMessageElement.Text;
                    }
                    catch (WebDriverTimeoutException)
                    {
                        actualErrorMessage = "Không có thông báo lỗi!";
                    }

                    // 🔍 So sánh ExpectedResult và ActualResult
                    bool isTestPass = actualErrorMessage.Trim() == expectedErrorMessage.Trim();

                    // 📝 Ghi vào file Excel
                    row["ActualResult"] = actualErrorMessage;  // Ghi kết quả thực tế
                    row["Status"] = isTestPass ? "Pass" : "Fail"; // Ghi Pass/Fail

                    if (isTestPass)
                    {
                        Console.WriteLine($"✅ Test Passed: Thông báo lỗi đúng mong đợi! ({actualErrorMessage})");
                        Assert.IsTrue(true, "Test Passed.");
                    }
                    else
                    {
                        Console.WriteLine($"❌ Test Failed: Thông báo lỗi không khớp. Expected: '{expectedErrorMessage}', Actual: '{actualErrorMessage}'");
                        Assert.Fail("Test case failed.");
                    }
                }
            }
        }








        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }
    }
}
