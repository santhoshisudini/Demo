using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Linq;

namespace DemoProject
{
    [TestFixture]
    public class UnitTest1
    {
        private static readonly log4net.ILog log =
            log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        IWebDriver driver;
        [SetUp]
        public void Initialize()
        {
            driver = new ChromeDriver();
        }
        
        [Test,Order(1)]
        public void CheckpageStatus()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            Assert.AreNotEqual(driver.Title,"Page Not Found | Microsoft Dynamics 365");
        }
        [Test,Order(2)]
        public void VerifyH1Tag()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            int H1Tagcount;
            List<IWebElement> H1Tags = driver.FindElements(By.TagName("h1")).ToList();
            H1Tagcount = H1Tags.Count;
            Assert.AreEqual(H1Tagcount, 1);
        }
        [Test]
        public void VerfiyMetaProperties()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            var MetaTitle = driver.FindElement(By.XPath("//meta[@name='title']@content")).Text;
            Assert.NotNull(MetaTitle);
            var MetaDescription = driver.FindElement(By.XPath("//meta[@name='description']@content")).Text;
            Assert.NotNull(MetaDescription);
            var MetaKeywords = driver.FindElement(By.XPath("//meta[@name='description']@content")).Text;
            Assert.NotNull(MetaKeywords);
        }

        [Test]
        public void VerifyClcid()
        {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            List<IWebElement> H1Tags = driver.FindElements(By.TagName("a")).ToList();
            foreach(IWebElement H1Tag in H1Tags)
            {
                if (H1Tag.GetAttribute("href").Contains("https://go.microsoft.com/fwlink/"))
                {
                    string value;
                    value = H1Tag.GetAttribute("href");
                    Assert.IsTrue(value.Contains("clcid"),value);
                }
            }

        }

        [Test]
        public void FindBrokenLinks()
        {

        }
        [Test,Explicit]
        public void Logging() {
            driver.Url = "https://dynamics.microsoft.com/en-us/";
            log.Info("Browser Launched");
        } 

        [TearDown]
        public void EndTest()
        {
            driver.Close();
        }
    }
   }

