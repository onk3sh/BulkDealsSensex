using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkDealsSensex
{
    /// <summary>
    /// Class DataContainer.
    /// </summary>
    internal class DataContainer
    {
        private List<string> columns;
        private List<IWebElement> rows;

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
        public void AddRows(IWebElement rowData)
        {
            this.rows.Add(rowData);
        }

        /// <summary>
        /// Gets the columns.
        /// </summary>
        /// <returns>List&lt;System.String&gt;.</returns>
        public List<string> GetColumns()
        {
            return this.columns;
        }

        /// <summary>
        /// Gets the rows.
        /// </summary>
        /// <returns>List&lt;IWebElement&gt;.</returns>
        public List<IWebElement> GetRows()
        {
            return this.rows;
        }

        /// <summary>
        /// Sets the columns.
        /// </summary>
        /// <param name="columnData">The column data.</param>
        public void SetColumns(IWebElement columnData)
        {
            if (columnData.Text != "No Records Found.")
                this.columns.Add(columnData.Text);
        }
    }
}
