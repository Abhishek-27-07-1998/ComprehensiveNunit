using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;

namespace Comprehensive.Hooks
{
    public class BaseClass
    {
       public static IWebDriver driver;
       public static ExtentReports extentreport;
       public static ExtentTest testreport;

        [OneTimeSetUp]
        public void ExtentStart()
        {
            extentreport = new ExtentReports();
            var htmlReport = new ExtentHtmlReporter(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports and Screenshot\index.html");
            extentreport.AttachReporter(htmlReport);
            LoggingLevelSwitch loggingLevel = new LoggingLevelSwitch(LogEventLevel.Debug);
            Log.Logger = new LoggerConfiguration().MinimumLevel.ControlledBy(levelSwitch: loggingLevel).WriteTo.File(@"C:\Users\mindc1may270\source\repos\Comprehensive\Reports and Screenshot\logs", outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}").CreateLogger();
        }

        [SetUp]
        public void Startup()
        {
            driver = new ChromeDriver(); 
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            extentreport.Flush();
           
        }

    }
}
