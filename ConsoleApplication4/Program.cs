using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using OpenQA.Selenium.Support.UI;
using System.Globalization;

/// <summary>
/// The Bulk Deals Sensex namespace.
/// Author: Onkesh Bansal
/// Date: 10th October 2017
/// License: MIT
/// </summary>
namespace BulkDealsSensex
{

    /// <summary>
    /// Class DataContainer.
    /// </summary>
    class DataContainer
    {
        private List<IWebElement> rows;
        private List<string> columns;

        /// <summary>
        /// Initializes a new instance of the <see cref="DataContainer"/> class.
        /// </summary>
        public DataContainer()
        {
            rows = new List<IWebElement>();
            columns = new List<string>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataContainer"/> class.
        /// </summary>
        /// <param name="tableRows">The table rows.</param>
        public DataContainer(IList<IWebElement> tableRows)
        {
            rows = new List<IWebElement>(tableRows);
            columns = new List<string>();
        }

        /// <summary>
        /// Adds the rows.
        /// </summary>
        /// <param name="rowData">The row data.</param>
        public void addRows(IWebElement rowData)
        {
            this.rows.Add(rowData);
        }

        /// <summary>
        /// Sets the columns.
        /// </summary>
        /// <param name="columnData">The column data.</param>
        public void setColumns(IWebElement columnData)
        {
            if(columnData.Text != "No Records Found.")
                this.columns.Add(columnData.Text);
         }

        /// <summary>
        /// Gets the rows.
        /// </summary>
        /// <returns>List&lt;IWebElement&gt;.</returns>
        public List<IWebElement> getRows()
        {
            return this.rows;
        }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        /// <returns>List&lt;System.String&gt;.</returns>
        public List<string> getColumns()
        {
            return this.columns;
        }
    }

    /// <summary>
    /// Class Program.
    /// </summary>
    static class Program
    {
        /// <summary>
        /// Defines the entry point of the application.
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void Main(string[] args)
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
                string strtDate = args[0];
                string endDate =  args[1]; //start date is always lesser then end date as start date will be in the past
                string exchange = args[2]; // BSE || NSE || Both

                Console.Clear();
                Console.WriteLine("********************************Getting data from ANAND RATHI Web Table**********************************************");

                if (exchange.ToLower() == "bse")
                {
                    finalMethodToGetData(driver, URLBSE, dir, strtDate, endDate, exchange);
                }
                else if (exchange.ToLower() == "nse")
                {
                    finalMethodToGetData(driver, URLNSE, dir, strtDate, endDate, exchange);
                }
                else
                {
                    finalMethodToGetData(driver, URLBSE, dir, strtDate, endDate, "bse");
                    finalMethodToGetData(driver, URLNSE, dir, strtDate, endDate, "nse");
                }
                driver.Quit();

            }
            catch (Exception e)
            {
                Console.WriteLine("!!!ERROR ENCOUNTERED!!!:::");
                Console.WriteLine(e.Message);
                Console.WriteLine("----------------------------");
                Console.WriteLine(e.InnerException);
                Console.WriteLine(e.StackTrace);
            }
            finally
            {
                Console.WriteLine("******************************************Press any key to exit.........********************************************");
                Console.Read();
            }
            
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
        private static void finalMethodToGetData(IWebDriver driver, string URL, DirectoryInfo dir, string strtDate, string endDate, string exchange)
        {
            driver.Navigate().Refresh();
            driver.Url = URL;
            Console.WriteLine("------------------------------------Fetching data for "+exchange.ToUpper()+"-----------------------------------------");
            DataContainer bse = new DataContainer();
            bse = fetchDataBetweenStartAndEndDate(driver, strtDate, endDate);
            outputDataToExcel(driver, bse, dir, exchange.ToUpper());
        }

        /// <summary>
        /// Fetches the data between start and end date.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="strtDate">The STRT date.</param>
        /// <param name="endDate">The end date.</param>
        /// <returns>DataContainer.</returns>
        static DataContainer fetchDataBetweenStartAndEndDate(IWebDriver driver, string strtDate, string endDate)
        {
            DateTime d1 = changeStringToDate(strtDate);
            DateTime d2 = changeStringToDate(endDate);
            List<DateTime> tradeDays = TradeDays(d1, d2);
            IWebElement tableElement = null;
            IList<IWebElement> tableRowBSE = null;
            IList<IWebElement> tableRowTemp = null;

            DataContainer data = new DataContainer();

            foreach (DateTime date in tradeDays)
            {
                setDateAndWait(driver, date);
                tableElement = driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_GrdGridViewBulkDeals']"));
                if (tableRowBSE == null)
                {
                    tableRowBSE = tableElement.FindElements(By.TagName("tr"));
                    data = new DataContainer(tableRowBSE);
                    foreach (IWebElement row in tableRowBSE)
                    {
                        IList<IWebElement> rowTD = row.FindElements(By.TagName("td"));
                        foreach (IWebElement col in rowTD)
                        {
                            data.setColumns(col);
                        }
                    }
                }
                else
                {
                    tableRowTemp = tableElement.FindElements(By.TagName("tr"));
                    foreach (IWebElement row in tableRowTemp)
                    {
                        data.addRows(row);
                    }

                    foreach (IWebElement row in tableRowTemp)
                    {
                        IList<IWebElement> rowTD = row.FindElements(By.TagName("td"));
                        foreach (IWebElement col in rowTD)
                        {
                            data.setColumns(col);
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
        static void setDateAndWait(IWebDriver driver, DateTime date)
        {
            changeDateInSite(driver, date);
            Console.WriteLine("Date set to "+date.ToShortDateString()+" and waiting...");
            System.Threading.Thread.Sleep(2000);
        }

        /// <summary>
        /// Changes the string to date.
        /// </summary>
        /// <param name="date">The date.</param>
        /// <returns>DateTime.</returns>
        static DateTime changeStringToDate(string date)
        {
            //date format = "dd-MM-yyyy";
            DateTime dt = DateTime.ParseExact(date, "dd-MM-yyyy", CultureInfo.InvariantCulture);
            return dt;
        }

        /// <summary>
        /// Changes the date in site.
        /// </summary>
        /// <param name="driver">The selenium web driver.</param>
        /// <param name="datevalue">The datevalue.</param>
        static void changeDateInSite(IWebDriver driver, DateTime datevalue)
        {
            SelectElement dd = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlDay']")));
            SelectElement mm = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlMonth']")));
            SelectElement yyyy = new SelectElement(driver.FindElement(By.XPath(".//*[@id='ctl00_ContentPlaceHolder1_DateUsrCtl1_ddlYear']")));

            IWebElement goBtn = driver.FindElement(By.Id("ctl00_ContentPlaceHolder1_btnGo"));

            String dy = datevalue.Day.ToString();
            String mn = datevalue.Month.ToString();
            String yy = datevalue.Year.ToString();

            if(Convert.ToInt32(dy) < 10)
                dd.SelectByText("0"+dy);
            else
                dd.SelectByText(dy);
            mm.SelectByIndex(Convert.ToInt32(mn));
            yyyy.SelectByText(yy);
            goBtn.Click();
        }

        /// <summary>
        /// Outputs the data to excel.
        /// </summary>
        /// <param name="driver">Selenium WebDriver Object.</param>
        /// <param name="data">The data.</param>
        /// <param name="outDir">The out dir.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="flag">if set to <c>true</c> [flag].</param>
        static void outputDataToExcel(IWebDriver driver, DataContainer data, DirectoryInfo outDir, string fileName, bool flag = true)
        {
            var newFile = new FileInfo(outDir.FullName + @"\"+fileName+".xlsx");
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(outDir.FullName + @"\" + fileName + ".xlsx");
            }

            using (var package = new ExcelPackage(newFile))
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = null;
                if (flag)
                {
                    worksheet = package.Workbook.Worksheets.Add("Bulk Deals BSE");
                }
                else if(!flag)
                {
                    worksheet = package.Workbook.Worksheets.Add("Bulk Deals NSE");
                }

                //Add the headers
                EnterDataInExcel(data, worksheet);

                // set some document properties
                package.Workbook.Properties.Title = "Financial Data";
                package.Workbook.Properties.Author = "Onkesh Bansal";
                package.Workbook.Properties.Comments = "All data for bulk deals";

                // set some extended property values
                package.Workbook.Properties.Company = "Flat 221, SLS Serinity, Bangalore";

                // set some custom property values
                package.Workbook.Properties.SetCustomPropertyValue("Checked by", "Onkesh Bansal");
                package.Workbook.Properties.SetCustomPropertyValue("AssemblyName", "EPPlus");
                // save our new workbook and we are done!
                package.Save();
            }

            Console.WriteLine("Excel file created , you can find the file at "+outDir+"\\"+fileName+".xlsx");
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

            IList<IWebElement> tableRow = data.getRows();
            IList<string> rowTD = data.getColumns();

            int i = 1, j = 0, col;
            while(j < rowTD.Count)
            {
                col = 0;
                worksheet.Cells[i + 1, col + 1].Value = rowTD[j];
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
        /// Cleanups the excel.
        /// </summary>
        /// <param name="xlApp">The xl application.</param>
        /// <param name="xlWorkbook">The xl workbook.</param>
        /// <param name="xlWorksheet">The xl worksheet.</param>
        /// <param name="xlRange">The xl range.</param>
        private static void cleanupExcel(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel._Worksheet xlWorksheet, Excel.Range xlRange)
        {
            //add useful things here!
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

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

            for (DateTime index = startDate; index <=endDate; index = index.AddDays(1))
            {
                if (index.DayOfWeek != DayOfWeek.Sunday && index.DayOfWeek != DayOfWeek.Saturday)
                {
                    result.Add(index);
                }
            }

            return result;
        }
    }
}