using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Reflection;

namespace WebScraper
{
    class Program
    {

        private static Microsoft.Office.Interop.Excel.Workbook mWorkBook;
        private static Microsoft.Office.Interop.Excel.Sheets mWorkSheets;
        private static Microsoft.Office.Interop.Excel.Worksheet mWSheet1;
        private static Microsoft.Office.Interop.Excel.Application oXL;

        static void Main(string[] args)
        {

            var optionsChrome = new ChromeOptions();
            optionsChrome.AddArgument("--headless");

            var driver = new ChromeDriver(optionsChrome);
            driver.Navigate().GoToUrl("https://fulltime.thefa.com/ff/DivisionDetails?divisionid=7044651&leagueid=3956158&seasonid=659468196");

            Thread.Sleep(1000);

            var titleSection = driver.FindElementById("ff-division-table-obj");
            var tableValues = titleSection.Text;

            var footballLeagueData = new List<FootballLeagueRow>();

            driver.FindElement(By.XPath("//*[@id=\"coachmark-modal-block\"]/div/div/div[2]/div/div[2]")).Click();

            for (int i = 1; i <= 22; i++)
            {
                var dataRow = new FootballLeagueRow();
                for (int j = 1; j <= 11; j++)
                {
                    if (j == 10) continue;
                    var xPathExpression = string.Format("//*[@id=\"ff-division-table-obj\"]/tbody/tr[{0}]/td[{1}]", i, j);
                    var dataItem = driver.FindElement(By.XPath(xPathExpression)).Text;

                    switch (j)
                    {
                        case 1:
                            dataRow.POS = dataItem;
                            break;
                        case 2:
                            dataRow.Team = dataItem;
                            break;
                        case 3:
                            dataRow.PLD = dataItem;
                            break;
                        case 4:
                            dataRow.W = dataItem;
                            break;
                        case 5:
                            dataRow.D = dataItem;
                            break;
                        case 6:
                            dataRow.L = dataItem;
                            break;
                        case 7:
                            dataRow.GF = dataItem;
                            break;
                        case 8:
                            dataRow.GA = dataItem;
                            break;
                        case 9:
                            dataRow.GD = dataItem;
                            break;
                        case 11:
                            dataRow.PTS = dataItem;
                            break;
                    }
                }

                footballLeagueData.Add(dataRow);
            }

            driver.Quit();
            driver.Dispose();

            string path = @"C:\Users\jackw\Documents\WebScraper\mybook.xlsx";
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.DisplayAlerts = false;

            mWorkBook = oXL.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Get all the sheets in the workbook
            mWorkSheets = mWorkBook.Worksheets;
            //Get the allready exists sheet
            mWSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)mWorkSheets.get_Item("Sheet1");
            Microsoft.Office.Interop.Excel.Range range = mWSheet1.UsedRange;

            //Add data to excel sheet
            int rowCount = 2;
            foreach (var item in footballLeagueData)
            {
                mWSheet1.Cells[rowCount, 1] = item.POS;
                mWSheet1.Cells[rowCount, 2] = item.Team;
                mWSheet1.Cells[rowCount, 3] = item.PLD;
                mWSheet1.Cells[rowCount, 4] = item.W;
                mWSheet1.Cells[rowCount, 5] = item.D;
                mWSheet1.Cells[rowCount, 6] = item.L;
                mWSheet1.Cells[rowCount, 7] = item.GF;
                mWSheet1.Cells[rowCount, 8] = item.GA;
                mWSheet1.Cells[rowCount, 9] = item.GD;
                mWSheet1.Cells[rowCount, 10] = item.PTS;
                rowCount++;
            }

            mWorkBook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
            Missing.Value, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
            mWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
            mWSheet1 = null;
            mWorkBook = null;
            oXL.Quit();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

        }
    }
    public class FootballLeagueRow
    {
        public string POS { get; set; }
        public string Team { get; set; }
        public string PLD { get; set; }
        public string W { get; set; }
        public string D { get; set; }
        public string L { get; set; }
        public string GF { get; set; }
        public string GA { get; set; }
        public string GD { get; set; }

        public string PTS { get; set; }
    }


}
