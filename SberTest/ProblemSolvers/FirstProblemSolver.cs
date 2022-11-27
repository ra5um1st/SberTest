using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SberTest
{
    internal class FirstProblemSolver : ISolver
    {
        private readonly string _sberMegaMarketUrl = "https://sbermegamarket.ru/";
        private readonly string _productName;
        private readonly int _count;

        public FirstProblemSolver(string productName, int count)
        {
            _productName = productName;
            _count = count;
        }

        public void Solve()
        {
            using (var driver = new ChromeDriver())
            {
                driver.Navigate().GoToUrl(_sberMegaMarketUrl);

                CloseRewardIfExists(driver);
                CloseRegionModal(driver);

                var productInfos = SelectProductInfos(driver).Take(_count);
                ExportToExcel(productInfos);
            }
        }

        private void CloseRegionModal(IWebDriver driver)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));

            try
            {
                var element = wait.Until(d => d.FindElement(By.XPath("//*[contains(@class, \"modal-win-profile-address-confirm\")]//*[@class=\"close-button\"]")));
                element.Click();
            }
            catch (Exception) { }
        }

        private void CloseRewardIfExists(IWebDriver driver) =>
            ClickOnChildElementIfExists(driver, "reward", "close");

        private void ClickOnChildElementIfExists(IWebDriver driver, string mainClassName, string childClassName) => driver
            .FindElements(By.ClassName(mainClassName))
            .FirstOrDefault()
            ?.FindElement(By.ClassName(childClassName))
            .Click();

        private IEnumerable<(string Name, int Price)> SelectProductInfos(IWebDriver driver)
        {
            var searchFieldInput = driver.FindElement(By.ClassName("search-field-input"));
            searchFieldInput.SendKeys(_productName);

            var searchButton = driver.FindElement(By.ClassName("header-search-form__send-button"));
            searchButton.Submit();

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(StaleElementReferenceException));

            try
            {
                var f = wait.Until(d =>
                {
                    var productItem = d.FindElement(By.XPath("//div[contains(@class, \"ddl_product\")]"));
                    return !string.IsNullOrEmpty(SelectTitle(productItem));
                });
            }
            catch (Exception) { }

            var productElements = driver.FindElements(By.XPath("//div[contains(@class, \"ddl_product\")]"));
            var productInfos = productElements.Select(item => (Name: SelectTitle(item), Price: SelectPrice(item)));

            return productInfos;
        }

        private void ExportToExcel(IEnumerable<(string Name, int Price)> productInfos)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application()
                {
                    DisplayAlerts = false
                };

                workbook = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Cells[1, 1] = "Наименование";
                worksheet.Cells[1, 2] = "Цена";

                var row = 2;

                foreach (var productInfo in productInfos)
                {
                    worksheet.Cells[row, 1] = productInfo.Name;
                    worksheet.Cells[row, 2] = productInfo.Price;
                    row++;
                }

                ((Excel.Range)worksheet.Columns[1]).AutoFit();
                ((Excel.Range)worksheet.Columns[2]).AutoFit();

                var filepath = Environment.CurrentDirectory + "/ProductInfos";
                workbook.SaveAs(filepath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workbook.Close();
                excelApp.Quit();
            }
            finally
            {
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }

        private int SelectPrice(IWebElement item) => int.Parse(SelectCurrencyPrice(item), NumberStyles.Currency);

        private string SelectCurrencyPrice(IWebElement item) => item
            .FindElement(By.XPath(".//*[@class=\"item-price\"]//span"))
            .Text;

        private string SelectTitle(IWebElement item) => item
            .FindElement(By.XPath(".//a[@class=\"ddl_product_link\"]"))
            .GetAttribute("title");
    }
}