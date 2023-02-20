using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using NUnit.Framework;

namespace ExtractReport
{
    public class ExportExtract
    {
        //browser driver
        IWebDriver webDriver = new ChromeDriver();

        [SetUp]
        public void setup()
        {
            //goto site 
            webDriver.Navigate().GoToUrl("https://wrmstorm-dev-mt-utility-ng.azurewebsites.net/default.aspx");
            webDriver.Manage().Window.Maximize();
        }
        [Test]
        public void TEST()
        {
            LoginPage();
            ExportPage();
            CustomizedReportExtract();
            Thread.Sleep(2000);
            checkboxChoose();
            ExportFileButton();
            Thread.Sleep(4000);
        }
        public void LoginPage()
        {
            // Identify Login 
            IWebElement LoginPage = webDriver.FindElement(By.Id("AdvanceSearch_ASPxButton1_CD"));
            LoginPage.Click();

            // login
            var Emailusername = webDriver.FindElement(By.Id("username"));
            var Passwordinput = webDriver.FindElement(By.Id("password"));
            Emailusername.SendKeys("nakorn_ton@wrmsoftware.com");
            Passwordinput.SendKeys("Multitenant01@");
       
            

            // login submit
            IWebElement LoginButton = webDriver.FindElement(By.Id("btn_VendorLogin"));
            LoginButton.Click();

            // Click Choose OC 
            IWebElement SelectOC = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_gluOC_B-1"));
            SelectOC.Click();
            Thread.Sleep(2000);

            var NewEngland = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_gluOC_DDD_gv_tccell3_0"));
            Thread.Sleep(2000);
            NewEngland.Click();

            // Click Choose OC and Event
            IWebElement SelectEvent = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_gluEvents_B-1"));
            SelectEvent.Click();
            Thread.Sleep(2000);

            var selectEleEvent = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_gluEvents_DDD_gv_tccell3_0"));
            Thread.Sleep(2000);
            selectEleEvent.Click();
            Thread.Sleep(2000);
            
        }

        public void ExportPage()
        {
            // menu report 
            IWebElement MenureportButton = webDriver.FindElement(By.Id("TopPanel_ASPxMenu1_DXI1_"));
            MenureportButton.Click();
            Thread.Sleep(2000);

            // submenu Customized Report Extract
            IWebElement MenuCustomizedExtract = webDriver.FindElement(By.LinkText("Customized Report Extract"));
            MenuCustomizedExtract.Click();
            Thread.Sleep(2000);
            
        }

        public void CustomizedReportExtract()
        {
   
            //input Description
            var Desinput = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_mmDescription_I"));
            Desinput.SendKeys("TestAutomatedByPare");
            Thread.Sleep(2000);

            //dropdown select data source
            IWebElement DataSource = webDriver.FindElement(By.Name("ctl00$ASPxPanel2$ContentPlaceHolder1$cpnMain$cbDataSource"));
            DataSource.Click();
            Thread.Sleep(2000);

            //select Timesheet
            IWebElement LodgingData = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_cbDataSource_DDD_L_LBI0T0"));
            LodgingData.Click();
            Thread.Sleep(2000);

            // INPUT DATE
            IWebElement DateStart = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_deStartDate_I"));
            DateStart.SendKeys("02/13/2023");
            IWebElement DateEnd = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_deEndDate_I"));
            DateEnd.SendKeys("02/18/2023");


        }


        public void checkboxChoose(){

            //add field
            IWebElement AddAllButton = webDriver.FindElement(By.CssSelector("#ASPxPanel2_ContentPlaceHolder1_cpnMain_cpnDataSource_btnMoveAllItemsToRight_CD"));
            AddAllButton.Click();
            Thread.Sleep(2000);
            
            // Retrieve Data
            IWebElement RetrieveData = webDriver.FindElement(By.CssSelector("#ASPxPanel2_ContentPlaceHolder1_cpnMain_btnRefresh"));
            RetrieveData.Click();
            Thread.Sleep(5000);
            
        }

        public void ExportFileButton()
        {
            // Export
            IWebElement ExportButton = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_btnExport_CD"));
            ExportButton.Click();
            Thread.Sleep(2000);
            
            //csv
            IWebElement ExportCSV = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_popUpExport_mnExport_DXI2_T"));
            ExportCSV.Click();
            Thread.Sleep(2000);
            


        }

        [TearDown]
        public void clear()
        {
            webDriver.Quit();
        }



    }
}