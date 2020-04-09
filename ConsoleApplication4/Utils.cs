using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BulkDealsSensex
{
    class Utils
    {
        /// <summary>
        /// This is a utility method. Cleans up any stale instances of excel in the background
        /// </summary>
        /// <param name="xlApp">The xl application.</param>
        /// <param name="xlWorkbook">The xl workbook.</param>
        /// <param name="xlWorksheet">The xl worksheet.</param>
        /// <param name="xlRange">The xl range.</param>
        public static void CleanupExcel(Excel.Application xlApp, Excel.Workbook xlWorkbook, Excel._Worksheet xlWorksheet, Excel.Range xlRange)
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
    }
}
