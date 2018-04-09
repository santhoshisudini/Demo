using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
namespace DemoProject
{
    [TestFixture]
    public class UnitTest1
    {
        IWebDriver driver;
        [SetUp]
        public void Initialize()
        {
            driver = new ChromeDriver();
        }
        [Test,Order(1)]
        public void CheckpageTitle()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            Assert.AreEqual(driver.Title,"Microsoft Dynamics 365: Intelligent Business Applications");
        }
        [Test,Order(2)]
        public void CheckpageStatus()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/heloo";
            Assert.AreEqual(driver.Title,"Page Not Found | Microsoft Dynamics 365");
        }

        [TearDown]
        public void EndTest()
        {
            driver.Close();
        }
    }
   }

