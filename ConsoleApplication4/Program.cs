using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;

/// <summary>
/// The Bulk Deals Sensex namespace.
/// Author: Onkesh Bansal
/// Created Date: 10th October 2017
/// Last Modified Date: 9th April 2020
/// License: MIT
/// </summary>
namespace BulkDealsSensex
{
    /// <summary>
    /// Class Program.
    /// </summary>
    internal static class Program
    {
        /// <summary>
        /// Days count that are left.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        /// <param name="excludeWeekends">The exclude weekends.</param>
        /// <param name="excludeDates">The exclude dates.</param>
        /// <returns>System.Int32.</returns>
        public static int DaysLeft(DateTime startDate, DateTime endDate, Boolean excludeWeekends = true, List<DateTime> excludeDates = null)
        {
            int count = 0;
            for (DateTime index = startDate; index < endDate; index = index.AddDays(1))
            {
                if (excludeWeekends && index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    bool excluded = false; ;
                    for (int i = 0; i < excludeDates.Count; i++)
                    {
                        if (index.Date.CompareTo(excludeDates[i].Date) == 0)
                        {
                            excluded = true;
                            break;
                        }
                    }

                    if (!excluded)
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        /// <summary>
        /// Trades the days.
        /// </summary>
        /// <param name="startDate">The start date.</param>
        /// <param name="endDate">The end date.</param>
        /// <returns>List&lt;DateTime&gt;.</returns>
        public static List<DateTime> TradeDays(DateTime startDate, DateTime endDate)
        {
            List<DateTime> result = new List<DateTime>();

            for (DateTime index = startDate; index <= endDate; index = index.AddDays(1))
            {
                if (index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    result.Add(index);
                }
            }

            return result;
        }

        /// <summary>
        /// Changes the date in site.
        /// </summary>
        /// <param name="driver">The selenium web driver.</param>
        /// <param name="datevalue">The datevalue.</param>
        private static void ChangeDateInSite(IWebDriver driver, DateTime datevalue)
        {
            SelectElement dd = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlDay']")));
            SelectElement mm = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlMonth']")));
            SelectElement yyyy = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlYear']")));

            IWebElement goBtn = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnGo"));

            String dy = datevalue.Day.ToString();
            String mn = datevalue.Month.ToString();
            String yy = datevalue.Year.ToString();

            if (Convert.ToInt32(dy) < 10)
                dd.SelectByText("0" + dy);
            else
                dd.SelectByText(dy);
            mm.SelectByIndex(Convert.ToInt32(mn));
            yyyy.SelectByText(yy);
            goBtn.Click();
        }

        /// <summary>
        /// Changes the string to date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns>DateTime.</returns>
        private static DateTime ChangeStringToDate(string date)
        {
            //date format = "dd-MM-yyyy";
            DateTime dt = DateTime.ParseExact(date, "dd-MM-yyyy", CultureInfo.InvariantCulture);
            return dt;
        }

        /// <summary>
        /// Enters the data in excel.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="worksheet">The worksheet.</param>
        private static void EnterDataInExcel(DataContainer data, ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Deal Date";
            worksheet.Cells[1, 2].Value = "Company";
            worksheet.Cells[1, 3].Value = "Client Name";
            worksheet.Cells[1, 4].Value = "Deal Type";
            worksheet.Cells[1, 5].Value = "Qty (000's)";
            worksheet.Cells[1, 6].Value = "Trade Price";
            worksheet.Cells[1, 7].Value = "Value (Rs.in Lakhs)";
            worksheet.Cells[1, 8].Value = "Close Price";

            IList<string> rowTD = data.GetColumns();

            int i = 1, j = 0, col;
            while (j < rowTD.Count)
            {
                col = 0;
                worksheet.Cells[i + 1, col + 1].Value = DateTime.Parse(rowTD[j]);
                worksheet.Cells[i + 1, col + 1].Style.Numberformat.Format = "dd-mmm-yy";

                worksheet.Cells[i + 1, col + 2].Value = rowTD[j + 1];
                worksheet.Cells[i + 1, col + 3].Value = rowTD[j + 2];
                worksheet.Cells[i + 1, col + 4].Value = rowTD[j + 3];
                worksheet.Cells[i + 1, col + 5].Value = Math.Round(Convert.ToDecimal(rowTD[j + 4]), 2);
                worksheet.Cells[i + 1, col + 6].Value = Math.Round(Convert.ToDecimal(rowTD[j + 5]), 2);
                worksheet.Cells[i + 1, col + 7].Value = Math.Round(Convert.ToDecimal(rowTD[j + 6]), 2);
                worksheet.Cells[i + 1, col + 8].Value = Math.Round(Convert.ToDecimal(rowTD[j + 7]), 2);
                i++;
                j += 8;
            }

            //Ok now format the values;
            using (var range = worksheet.Cells[1, 1, 1, 8])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.DarkBlue);
                range.Style.Font.Color.SetColor(Color.White);
            }

            worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells
        }

        /// <summary>
        /// Fetches the data between start and end date.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="strtDate">The STRT date.</param>
        /// <param name="endDate">The end date.</param>
        /// <returns>DataContainer.</returns>
        private static DataContainer FetchDataBetweenStartAndEndDate(IWebDriver driver, string strtDate, string endDate)
        {
            DateTime d1 = ChangeStringToDate(strtDate);
            DateTime d2 = ChangeStringToDate(endDate);
            List<DateTime> tradeDays = TradeDays(d1, d2);
            IList<IWebElement> tableRowBSE = null;
            DataContainer data = new DataContainer();

            foreach (DateTime date in tradeDays)
            {
                SetDateAndWait(driver, date);
                IWebElement tableElement = driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_GrdGridViewBulkDeals']"));
                if (tableRowBSE == null)
                {
                    tableRowBSE = tableElement.FindElements(By.TagName("tr"));
                    data = new DataContainer(tableRowBSE);
                    foreach (IWebElement row in tableRowBSE)
                    {
                        IList<IWebElement> rowTD = row.FindElements(By.TagName("td"));
                        foreach (IWebElement col in rowTD)
                        {
                            data.SetColumns(col);
                        }
                    }
                }
                else
                {
                    IList<IWebElement> tableRowTemp = tableElement.FindElements(By.TagName("tr"));
                    foreach (IWebElement row in tableRowTemp)
                    {
                        data.AddRows(row);
                    }

                    foreach (IWebElement row in tableRowTemp)
                    {
                        IList<IWebElement> rowTD = row.FindElements(By.TagName("td"));
                        foreach (IWebElement col in rowTD)
                        {
                            data.SetColumns(col);
                        }
                    }
                }
            }
            return data;
        }

        /// <summary>
        /// Sets the date and wait.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="date">The date.</param>
        private static void SetDateAndWait(IWebDriver driver, DateTime date)
        {
            ChangeDateInSite(driver, date);
            Console.WriteLine();
            Console.WriteLine("Date set to " + date.ToShortDateString() + " and waiting...");
            System.Threading.Thread.Sleep(2000);
        }

        /// <summary>
        /// Finals the method to get data.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="URL">The URL.</param>
        /// <param name="dir">The dir.</param>
        /// <param name="strtDate">The STRT date.</param>
        /// <param name="endDate">The end date.</param>
        /// <param name="exchange">The exchange.</param>
        private static void FinalMethodToGetData(IWebDriver driver, string URL, DirectoryInfo dir, string strtDate, string endDate, string exchange)
        {
            DataContainer dc = CreateDataForExcel(driver, URL, strtDate, endDate, exchange);
            OutputDataToExcel(dc, dir, exchange);
        }

        private static DataContainer CreateDataForExcel(IWebDriver driver, string URL, string strtDate, string endDate, string exchange)
        {
            driver.Navigate().Refresh();
            driver.Url = URL;
            Console.WriteLine("------------------------------------Fetching data for " + exchange.ToUpper() + "-----------------------------------------");
            DataContainer dc = FetchDataBetweenStartAndEndDate(driver, strtDate, endDate);
            return dc;
        }

        /// <summary>
        /// Finals the method to get data.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="URL">The URL.</param>
        /// <param name="dir">The dir.</param>
        /// <param name="strtDate">The STRT date.</param>
        /// <param name="endDate">The end date.</param>
        /// <param name="exchange">The exchange.</param>
        private static void FinalMethodToGetDataForBoth(IWebDriver driver, string URL_BASE, DirectoryInfo dir, string strtDate, string endDate)
        {
            DataContainer data_bse = CreateDataForExcel(driver, URL_BASE + "BSE", strtDate, endDate, "bse");
            DataContainer data_nse = CreateDataForExcel(driver, URL_BASE + "NSE", strtDate, endDate, "nse");

            OutputBothDataToSameExcel(data_bse, data_nse, dir);
        }

        /// <summary>
        /// Outputs the data to excel.
        /// </summary>
        /// <param name="data">The data.</param>
        /// <param name="outDir">The out dir.</param>
        /// <param name="fileName">Name of the file.</param>
        private static void OutputDataToExcel(DataContainer data, DirectoryInfo outDir, string exchange)
        {
            var fileName = exchange.ToUpper();
            FileInfo newFile = CheckIfFileExists(outDir, fileName);

            using (var package = new ExcelPackage(newFile))
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = null;
                worksheet = package.Workbook.Worksheets.Add("Bulk Deals " + fileName);

                //Add the headers
                EnterDataInExcel(data, worksheet);

                // set some document properties
                package.Workbook.Properties.Title = "Financial Data";
                package.Workbook.Properties.Author = "Onkesh Bansal";
                package.Workbook.Properties.Comments = "All data for bulk deals performed on the respective exchanges";

                // set some extended property values
                package.Workbook.Properties.Company = "Flat 221, SLS Serinity, Bangalore";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Onkesh Bansal");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
                // save our new workbook and we are done!
                package.Save();
            }

            Console.WriteLine();
            Console.WriteLine("Excel file created , you can find the file at " + outDir + "\\" + fileName + ".xlsx");
        }

        /// <summary>
        /// Method to export data for both exchanges to a single file
        /// </summary>
        /// <param name="data_bse"></param>
        /// <param name="data_nse"></param>
        /// <param name="outDir"></param>
        private static void OutputBothDataToSameExcel(DataContainer data_bse, DataContainer data_nse, DirectoryInfo outDir)
        {
            var fileName = "BSE_NSE_BulkDeals";
            FileInfo newFile = CheckIfFileExists(outDir, fileName);

            using (var package = new ExcelPackage(newFile))
            {
                // Add worksheets to the empty workbook
                package.Workbook.Worksheets.Add("BSE");
                package.Workbook.Worksheets.Add("NSE");

                //Add the headers
                EnterDataInExcel(data_bse, package.Workbook.Worksheets["BSE"]);
                EnterDataInExcel(data_nse, package.Workbook.Worksheets["NSE"]);

                // set some document properties
                package.Workbook.Properties.Title = "Financial Data";
                package.Workbook.Properties.Author = "Onkesh Bansal";
                package.Workbook.Properties.Comments = "All data for bulk deals performed on the respective exchanges";

                // set some extended property values
                package.Workbook.Properties.Company = "Flat 221, SLS Serinity, Bangalore";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Onkesh Bansal");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
                // save our new workbook and we are done!
                package.Save();
            }

            Console.WriteLine();
            Console.WriteLine("Excel file created , you can find the file at " + outDir + "\\" + fileName + ".xlsx");
        }

        /// <summary>
        /// Checks if the xlsx file exists in the provided directory
        /// </summary>
        /// <param name="outDir"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static FileInfo CheckIfFileExists(DirectoryInfo outDir, string fileName)
        {
            var newFile = new FileInfo(outDir.FullName + @"\" + fileName + ".xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(outDir.FullName + @"\" + fileName + ".xlsx");
            }

            return newFile;
        }

        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        private static void Main(string[] args)
        {
            try
            {
                // make chrome run in headless mode
                ChromeOptions option = new ChromeOptions();
                option.AddArgument("--headless");
                IWebDriver driver = new ChromeDriver(option);
                DataContainer data = new DataContainer();

                string URL = "http://anandrathi.accordfintech.com/Equity/BulkDeals.aspx?id=22&Option=&EXCHG=";
                string URLBSE = URL + "BSE";
                string URLNSE = URL + "NSE";
                string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                DirectoryInfo dir = new DirectoryInfo(documentsFolder);
                //Taking arguments from command-line
                string strtDate, endDate, exchange; //start date is always lesser then end date as start date will be in the past

                if (args.Length == 3)
                {
                    strtDate = args[0];
                    endDate = args[1]; //start date is always lesser then end date as start date will be in the past
                    exchange = args[2]; // BSE || NSE || Both
                }
                else
                {
                    //Fallback to generic arguments incase of missing commandline arguments
                    strtDate = DateTime.Today.AddDays(-2).ToString("dd-MM-yyyy");
                    endDate = DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy"); //start date is always lesser then end date as start date will be in the past
                    exchange = "both";
                }

                Console.Clear();
                Console.WriteLine("********************************Getting data from ANAND RATHI Web Table**********************************************");
                Console.WriteLine();

                if(args.Length == 0)
                {
                    System.Threading.Thread.Sleep(500);
                    Console.WriteLine("No Arguments Provided!!");
                    System.Threading.Thread.Sleep(500);
                    Console.WriteLine("Options: BulkDealsSensex.exe START-DATE END-DATE bse or nse or both");
                    System.Threading.Thread.Sleep(500);
                    Console.WriteLine("Command: BulkDealsSensex.exe dd-MM-yyyy dd-MM-yyyy both ");
                    System.Threading.Thread.Sleep(500);
                    Console.WriteLine("Falling back.......");
                    System.Threading.Thread.Sleep(500);

                }

                if (exchange.ToLower() == "bse")
                {
                    FinalMethodToGetData(driver, URLBSE, dir, strtDate, endDate, exchange);
                }
                else if (exchange.ToLower() == "nse")
                {
                    FinalMethodToGetData(driver, URLNSE, dir, strtDate, endDate, exchange);
                }
                else if (exchange.ToLower() == "both")
                {
                    FinalMethodToGetDataForBoth(driver, URL, dir, strtDate, endDate);
                }
                driver.Quit();
            }
#pragma warning disable CA1031 // Do not catch general exception types
            catch (Exception e)
#pragma warning restore CA1031 // Do not catch general exception types
            {
                Console.WriteLine("!!!ERROR ENCOUNTERED!!!:::");
                Console.WriteLine(e.Message);
                Console.WriteLine("----------------------------");
                Console.WriteLine(e.InnerException);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                Console.WriteLine();
                Console.WriteLine("******************************************Press any key to exit.........********************************************");
                Console.Read();
            }
        }
    }
}