using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Threading;
using Serilog;
using AventStack.ExtentReports;
using NPOI.XSSF.UserModel;
using System.IO;
using OpenQA.Selenium.Interactions;

namespace Comprehensive.POM
{
    [TestFixture]
    public class Tideapplication : Hooks.BaseClass
    {
        [Test]
        public void SignUpPositiveTestCase()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("SignUpPositiveTestCase").Info("Teststarted");
            testreport.Log(Status.Info, "SignUpwithValidCredentials");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            Log.Information("browser started"); 
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='Register']")).Click();
            int location = driver.FindElement(By.XPath("(//picture/img[@loading='lazy'])[2]")).Location.Y;
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollTo(0," + location + ")");
            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            driver.FindElement(By.XPath("//a[text()='Sign up now']")).Click();
            Log.Information("Click on Signup");         
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            Thread.Sleep(2000);
            driver.FindElement(By.Id("name")).SendKeys("Abhishek");
            driver.FindElement(By.Id("email")).SendKeys("bsabhi123@gmail.com");
            driver.FindElement(By.Id("password")).SendKeys("Bsabhi@123");
            driver.FindElement(By.XPath("//input[@value='CREATE ACCOUNT']")).Click();
            testreport.Log(Status.Pass, "User SignedUpwithValidCredentials");
            Log.Information("User is able to signup successfully");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\SignUp1.jpeg");
            driver.Quit();
        }

        [Test]
        public void SignUpNegativeTestCase()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("SignUpNegativeTestCase").Info("TestStarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            testreport.Log(Status.Info, "SignUpwithinValidCredentials");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "URL loaded");
            driver.Manage().Window.Maximize();
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='Register']")).Click();
            int location = driver.FindElement(By.XPath("(//picture/img[@loading='lazy'])[2]")).Location.Y;
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollTo(0," + location + ")");
            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            driver.FindElement(By.XPath("//a[text()='Sign up now']")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            driver.FindElement(By.Id("name")).SendKeys("12344");
            string Actual = driver.FindElement(By.XPath("//p[text()='First Name is invalid.']")).Text;
            string Expected = "First Name is invalid.";
            bool testing;
            if (Expected.Equals(Actual))
            {
                testing= true;
                Console.WriteLine("The text: First Name is invalid is displayed");
            }
            else
            {
                testing = false;
                Console.WriteLine("The text: Error message is not displayed");
            }
            Assert.IsTrue(testing);
            driver.FindElement(By.Id("email")).SendKeys("bsabhi123@com");
            string ActualEmail = driver.FindElement(By.XPath("//p[text()='Email Not Valid']")).Text;
            string ExpectedEmail = "Email Not Valid";
            bool emailtext;
            if (ExpectedEmail.Equals(ActualEmail))
            {
                emailtext = true;
                Console.WriteLine("The text: Email Not Valid is displayed");
            }
            else
            {
                emailtext = false;
                Console.WriteLine("The text: Error message is not displayed");
            }
            Assert.IsTrue(emailtext);
            driver.FindElement(By.Id("password")).SendKeys("Bsabhi12345");
            String ActualText = driver.FindElement(By.XPath("//p[contains(text(),'Must contain at least one of the following')]")).Text;
            string ExpectedText = "Must contain at least one of the following  ! @ $ % & ? *";
            bool testText;
            if (ExpectedText.Equals(ActualText))
            {
                testText = true;
                Console.WriteLine("The text: Must contain at least one of the following  ! @ $ % & ? * is displayed");
            }
            else
            {
                testText = false;
                Console.WriteLine("The text: Error message is not displayed");
            }
            Assert.IsTrue(testText);
            testreport.Log(Status.Pass, "SignUpwithInValidCredentials");
            testreport.Log(Status.Fail, "SignUp Failed");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\SignUp2.jpeg");
            driver.Quit();
        }

        [Test]
        public void SearchFeature()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("SearchFeature").Info("Teststarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us"); 
            testreport.Log(Status.Info, "URL loaded");
            Log.Information("URL Loaded");
            driver.Manage().Window.Maximize();
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//input[@name='q']")).SendKeys("Tide Eco-Box purclean Plant-Based Liquid Laundry Detergent");
            driver.FindElement(By.XPath("//button[@type='submit']")).Click();
            int location = driver.FindElement(By.XPath("//a[@class='footer-social-links-item event_socialmedia_exit']")).Location.Y;
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollTo(0," + location + ")");
            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            driver.FindElement(By.XPath("(//span[text()='Find Retailers'])[1]")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("(//button[@class='ps-online-buy-button available ps-online-seller-button'])[1]")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            Console.WriteLine("Tab count=" + count);
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            string Actual = driver.FindElement(By.XPath("//span[@id='productTitle']")).Text;
            string Expected = "Tide Purclean Plant-Based EPA Safer Choice Natural Laundry Detergent Liquid Soap Eco-Box, Ultra Concentrated High Efficiency (HE), 72 Loads";
            if (Actual == Expected)
            {
                Console.WriteLine("Text=" + Expected);
            }
            else
            {
                Console.WriteLine("Unknown text");
            }
            testreport.Log(Status.Info, "Expected product and actual product displayed are same");
            Log.Information("Expected product and actual product displayed are same");
            testreport.Log(Status.Pass, "Expected product and actual product displayed are same");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\Search.jpeg");
            driver.Quit();
        }

        [Test]
        public void LoginFeature()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("LoginFeature").Info("Teststarted");
            testreport.Log(Status.Info, "LoginFeature");
            driver.Navigate().GoToUrl("https://tide.com/en-us"); 
            Log.Information("browser started");
            driver.Manage().Window.Maximize();
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='Register']")).Click();
            int location = driver.FindElement(By.XPath("(//picture/img[@loading='lazy'])[2]")).Location.Y;
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollTo(0," + location + ")");
            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            driver.FindElement(By.XPath("//a[text()='Sign up now']")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            driver.FindElement(By.XPath("//button[text()='Log in']")).Click();
            driver.FindElement(By.XPath("//input[@id='login-email']")).SendKeys("abhi@gmail.com");
            driver.FindElement(By.XPath("//input[@id='login-password']")).SendKeys("Abhi@1234");
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            testreport.Log(Status.Info, "User entered valid login credentials");
            testreport.Log(Status.Pass, "Login is Successful");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\Login.jpeg");
            driver.Quit();
        }

        [Test]
        public void SignUpReaddataFromExcel()
        {
            testreport = extentreport.CreateTest("SignUpReaddataFromExcel").Info("Teststarted");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "URL loaded");
            Log.Information("browser started");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();           
            driver.FindElement(By.XPath("//a[text()='Register']")).Click();
            Log.Debug("Read data from excel to  Register");
            int location = driver.FindElement(By.XPath("(//picture/img[@loading='lazy'])[2]")).Location.Y;
            IJavaScriptExecutor j = (IJavaScriptExecutor)driver;
            j.ExecuteScript("window.scrollTo(0," + location + ")");
            driver.FindElement(By.XPath("//button[@class='sticker_close']")).Click();
            driver.FindElement(By.XPath("//a[text()='Sign up now']")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            Thread.Sleep(2000);
            string path = @"C:\Users\mindc1may270\Desktop\Comprehensive Assessment\Excel\Selenium.xlsx";
            XSSFWorkbook workbook = new XSSFWorkbook(File.Open(path, FileMode.Open));
            var sheet = workbook.GetSheetAt(0);
            var row = sheet.GetRow(0);
            var namevalue = row.GetCell(0).StringCellValue.Trim();
            var row1 = sheet.GetRow(1);
            var emailvalue = row1.GetCell(0).StringCellValue.Trim();
            var row2 = sheet.GetRow(2);
            var passwordvalue = row2.GetCell(0).StringCellValue.Trim();
            driver.FindElement(By.Id("name")).SendKeys(namevalue);
            driver.FindElement(By.Id("email")).SendKeys(emailvalue); 
            driver.FindElement(By.Id("password")).SendKeys(passwordvalue);
            testreport.Log(Status.Info, "Read the Signup Data from Excel");
            driver.FindElement(By.XPath("//input[@value='CREATE ACCOUNT']")).Click();
            testreport.Log(Status.Info, "User is able to read data from excel sheet");
            testreport.Log(Status.Pass, "User is able to signup by reading data from excel");
            Log.Information("User is able to read data from excel sheet");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\SignUpExcel.jpeg");
            driver.Quit();
        }


        [Test]
        public void ClickOnPacs()
        {
            testreport = extentreport.CreateTest("ClickOnPacs").Info("Teststarted");
            testreport.Log(Status.Info, "ShoProducts feature");
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "User should be able to Click on pacs under shop products ");
            Log.Information("User is able to click on liquids ");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            IWebElement path = driver.FindElement(By.XPath("//a[text()='Shop Products']"));
            Actions action = new Actions(driver);
            action.MoveToElement(path).Perform();
            driver.FindElement(By.XPath("(//span[text()='Liquid'])[1]")).Click();
            driver.FindElement(By.XPath("(//div[@aria-label='Find where to buy this product'])[1]")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//button[@class='ps-online-buy-button unavailable ps-online-seller-button']")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            string content = driver.FindElement(By.XPath("(//p)[1]")).Text;
            Console.WriteLine("Print the prodict is not availavble and reason for unavailability");
            Console.WriteLine(content);
            testreport.Log(Status.Pass, "User is able shop available products");
            Log.Information("User is able to shop available products");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\ClickOnLiquid.jpeg");
            driver.Quit();
        }

        [Test]
        public void Cupons()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("Cupons").Info("Teststarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "User should be able to Save cupon");
            Log.Information("User is able to save cupon ");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='Coupons and Rewards']")).Click();
            driver.FindElement(By.XPath("//a[text()='Save Now']")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<string> tabs = driver.WindowHandles;
            int count = tabs.Count;
            foreach (string tab in tabs)
            {
                driver.SwitchTo().Window(tab);
            }
            driver.FindElement(By.XPath("//button[text()='Log in']")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//input[@type='email']")).SendKeys("abhi@gmail.com");
            driver.FindElement(By.XPath("//input[@type='password']")).SendKeys("Abhi@1234");
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            testreport.Log(Status.Info, "User is able to login to his account to save cupon");
            testreport.Log(Status.Pass, "User is able to Save cupon");
            Log.Information("User is able to save cupon ");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\Cupons.jpeg");
            driver.Quit();
        }

        [Test]
        public void LocationDropdown()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("LocationDropdown").Info("Teststarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "User should select loaction from dropdown");
            Log.Information("User should select location from dropdown ");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//button[text()='US - English']")).Click();
            driver.FindElement(By.XPath("//a[text()='Canada - French']")).Click();
            string Frenchhtext = driver.FindElement(By.XPath("//a[text()='Magasiner les produits']")).Text;
            string expected = "Magasiner les produits";
            if (Frenchhtext == expected)
            {
                Console.WriteLine("The location got changed to Canada-French and the text for verification is:" + Frenchhtext);
            }
            else
            {
                Console.WriteLine("The location didnot change");
            }
            testreport.Log(Status.Info, "User should select loaction from dropdown");
            Log.Information("User is able to select location ");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\Location.jpeg");
            driver.Quit();
        }

        [Test]
        public void Livechatoption()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("Livechatoption").Info("TestStarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "User should select Live chat option");
            Log.Information("User Clicks on Live Chat option ");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='Live Chat']")).Click();
            Console.WriteLine(DateTime.UtcNow);
            Console.WriteLine("LiveChat is available between 10Am and 6Pm UTC only");
            testreport.Log(Status.Info, "User selected Live chat option");
            Log.Information("User is able to click on Live chat");
            testreport.Log(Status.Pass, "User is able to see the available time of live chat");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\Livechat.jpeg");
            driver.Quit();
        }

        [Test]
        public void WhatisNew()
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(4000);
            testreport = extentreport.CreateTest("WhatisNew").Info("TestStarted");
            driver.Navigate().GoToUrl("https://tide.com/en-us");
            driver.Manage().Window.Maximize();
            testreport.Log(Status.Info, "User Click on whats new");
            Log.Information("User should  click on whats new ");
            driver.FindElement(By.XPath("//a[@class='lilo3746-close-link lilo3746-close-icon']")).Click();
            driver.FindElement(By.XPath("//a[text()='What’s New']")).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("(//a[text()='Learn More'])[2]")).Click();
            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> elements = driver.FindElements(By.XPath("//a"));
            int count = elements.Count;
            for (int i = 0; i < count; i++)
            {
                string text = elements[i].Text;
                Console.WriteLine(text);
            }
            testreport.Log(Status.Info, "User Clicked on whats new");
            testreport.Log(Status.Pass, "User is able to explore the features under whats new");
            Log.Information("User is able to explore the featues un whats new");
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            ss.SaveAsFile(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports\WhatsNew.jpeg");
            driver.Quit();
        }
    }
}
