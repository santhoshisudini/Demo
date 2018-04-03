using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Chrome;
namespace DemoProject
{
    [TestFixture]
    public class UnitTest1
    {
       IWebDriver driver;

        [SetUp]
        public void startBrowser()
        {
            driver = new ChromeDriver();
        }

        [Test]
        public void test()
        {
            driver.Url = "https://www.google.com";
         string title=  driver.Title;
        }

        [TearDown]
        public void closeBrowser()
        {
            driver.Close();
        }

    }
   }

