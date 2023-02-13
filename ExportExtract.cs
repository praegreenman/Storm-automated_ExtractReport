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

            var selectEleEvent = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_gluEvents_DDD_gv_tccell1_0"));
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
            //dropdown select report 
            IWebElement Report = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_cbReport_B-1"));
            Report.Click();
            Thread.Sleep(2000);

            //select workforce :
            IWebElement ListWorkforce = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_cbReport_DDD_L_LBI1T0"));
            ListWorkforce.Click();
            Thread.Sleep(5000);


            //input Description
            var Desinput = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_mmDescription_I"));
            Desinput.SendKeys("TestAutomatedByPare");
            Thread.Sleep(2000);

            //dropdown select data source
            IWebElement DataSource = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_cbDataSource_B-1"));
            DataSource.Click();
            Thread.Sleep(2000);

            //select lodging
            IWebElement LodgingData = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_cbDataSource_DDD_L_LBI2T0"));
            LodgingData.Click();
            Thread.Sleep(2000);


        }


        public void checkboxChoose(){

            //check box 
            IWebElement AssignLocation = webDriver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='Assigned_Location'])[1]/preceding::span[2]"));
            AssignLocation.Click();
            Thread.Sleep(1000);
            IWebElement Crewsheet = webDriver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='Crew_Sheet'])[1]/preceding::span[2]"));
            Crewsheet.Click();
            Thread.Sleep(1000);
            IWebElement CrewsheetName = webDriver.FindElement(By.XPath("(.//*[normalize-space(text()) and normalize-space(.)='Crew_Sheet_Name'])[1]/preceding::span[2]"));
            CrewsheetName.Click();
            Thread.Sleep(2000);

            IWebElement HotelName = webDriver.FindElement(By.XPath(".//*/text()[normalize-space(.)='Lodging_Assigned_Date']/parent::*"));
            HotelName.Click();
            Thread.Sleep(1000);

            IWebElement LodgingDate = webDriver.FindElement(By.XPath(".//*/text()[normalize-space(.)='Operating_Company']/parent::*"));
            LodgingDate.Click();
            Thread.Sleep(1000);

            IWebElement OC = webDriver.FindElement(By.XPath(".//*/text()[normalize-space(.)='Operating_Company']/parent::*"));
            OC.Click();
            Thread.Sleep(1000);
            IWebElement region = webDriver.FindElement(By.XPath(".//*/text()[normalize-space(.)='Region']/parent::*"));
            region.Click();
            Thread.Sleep(2000);

            //add field
            IWebElement AddButton = webDriver.FindElement(By.CssSelector("#ASPxPanel2_ContentPlaceHolder1_cpnMain_cpnDataSource_btnMoveSelectedItemsToRight_CD"));
            AddButton.Click();
            Thread.Sleep(2000);
            

            // save Report
            IWebElement SaveButton = webDriver.FindElement(By.CssSelector("#ASPxPanel2_ContentPlaceHolder1_cpnMain_btnSave_CD"));
            SaveButton.Click();
            Thread.Sleep(2000);

            //edit reportname
            var NameReport = "paretestt";
            var NameReportinput = webDriver.FindElement(By.Id("ASPxPanel2_ContentPlaceHolder1_cpnMain_pucReportName_tbPucReportName_I"));
            NameReportinput.SendKeys(NameReport);

            

            // comfirm save 
            IWebElement SaveConfirmButton = webDriver.FindElement(By.CssSelector("#ASPxPanel2_ContentPlaceHolder1_cpnMain_pucReportName_btnSavePucReportName_CD"));
            SaveConfirmButton.Click();
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
            


        }

        [TearDown]
        public void clear()
        {
            webDriver.Quit();
        }



    }
}