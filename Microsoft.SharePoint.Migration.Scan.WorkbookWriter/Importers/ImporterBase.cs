using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Diagnostics;
using System.Linq;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Importers
{
    abstract class ImporterBase : IDisposable
    {
        private static readonly string TraceCategory = "ImporterBase";
        private SharedStringTablePart _sharedStringPart;

        /// <summary>
        /// Creates a new instance of the FileImporter class
        /// </summary>
        /// <param name="targetWorkbookPath">Target workbook instance</param>
        public ImporterBase(SpreadsheetDocument targetWorkbook)
        {
            this.TargetWorkbook = targetWorkbook;
            this._sharedStringPart = null;
        }

        #region properties

        protected SpreadsheetDocument TargetWorkbook { get; set; }

        protected SharedStringTable StringTable
        {
            get
            {
                if (this._sharedStringPart == null)
                {
                    // Get the SharedStringTablePart. If it does not exist, create a new one.
                    if (this.TargetWorkbook.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        this._sharedStringPart = this.TargetWorkbook.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        this._sharedStringPart = this.TargetWorkbook.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }

                    // If the part does not contain a SharedStringTable, create one.
                    if (this._sharedStringPart.SharedStringTable == null)
                    {
                        this._sharedStringPart.SharedStringTable = new SharedStringTable();
                    }
                }

                return this._sharedStringPart.SharedStringTable;
            }
        }

        #endregion

        #region abstract Import

        /// <summary>
        /// Orchestrates the import from the source csv to the target workbook
        /// </summary>
        public abstract void Import();

        #endregion

        #region GetSheet

        /// <summary>
        /// Gets the sheet to which we will write based on the supplied sheet name
        /// </summary>
        /// <returns></returns>
        protected Sheet GetSheet(string sheetName)
        {
            Sheet sheet = null;
            try
            {
                sheet = this.TargetWorkbook.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().First(s => s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception err)
            {
                Trace.WriteLine(string.Format("Could not locate the worksheet '{0}' in the target workbook.", sheetName), TraceCategory);
                Trace.TraceError(err.ToString());
                throw;
            }

            return sheet;
        }

        #endregion

        #region GetRow

        /// <summary>
        /// Ensures a row by index
        /// </summary>
        /// <param name="sheetData">The sheet to contain the row</param>
        /// <param name="rowIndex">The index of the row</param>
        /// <returns>The new row</returns>
        protected Row GetRow(SheetData sheetData, uint rowIndex)
        {
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            return row;
        }

        #endregion

        #region GetCell

        /// <summary>
        /// Ensures the required cell exists in the row
        /// </summary>
        /// <param name="sheetPart">The sheet in which the cell will reside</param>
        /// <param name="row">The row to contain the cell</param>
        /// <param name="cellReference">Cell reference for the location such as "A1" or "C438"</param>
        /// <returns>A new or existing cell</returns>
        protected Cell GetCell(WorksheetPart sheetPart, Row row, string cellReference)
        {
            // If there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value.Equals(cellReference, StringComparison.OrdinalIgnoreCase)).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value.Equals(cellReference, StringComparison.OrdinalIgnoreCase)).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };

                if (refCell == null)
                {
                    row.Append(newCell);
                }
                else
                {
                    row.InsertBefore(newCell, refCell);
                }

                return newCell;
            }
        }

        #endregion

        #region InsertSharedStringItem

        /// <summary>
        /// Ensures the supplied text appears in the shared string table and returns the index
        /// </summary>
        /// <param name="text">The text to add to the string table</param>
        /// <returns>Index of the string in the shared table</returns>
        protected int InsertSharedStringItem(string text)
        {
            int i = 0;

            foreach (SharedStringItem item in this.StringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            this.StringTable.AppendChild(new SharedStringItem(new Text(text)));

            return i;
        }

        #endregion

        #region WriteToCell

        /// <summary>
        /// Writes a value to the supplied cell using
        /// </summary>
        /// <param name="cell">Cell to which we write</param>
        /// <param name="fieldType">Data type of the cell value</param>
        /// <param name="rawSourceValue">The raw string value to be written to the cell</param>
        protected void WriteToCell(Cell cell, CellValues fieldType, string rawSourceValue)
        {
            cell.DataType = new EnumValue<CellValues>(fieldType);

            if (fieldType == CellValues.SharedString)
            {
                var sharedStringIndex = this.InsertSharedStringItem(rawSourceValue);
                cell.CellValue = new CellValue(sharedStringIndex.ToString());
            }
            else
            {
                cell.CellValue = new CellValue(rawSourceValue);
            }
        }

        #endregion

        #region GetCellAddress

        /// <summary>
        /// Calculates the cell address from the supplied columnIndex and rowIndex
        /// </summary>
        /// <param name="columnIndex">One based index of the column</param>
        /// <param name="rowIndex">One based index of the row</param>
        /// <returns>The formatted cell address such as AA2 or V432</returns>
        protected string GetCellAddress(uint columnIndex, uint rowIndex)
        {
            if (columnIndex < 1)
            {
                throw new ArgumentOutOfRangeException("columnIndex is one based and cannot be less than one.");
            }

            if (rowIndex < 1)
            {
                throw new ArgumentOutOfRangeException("rowIndex is one based and cannot be less than one.");
            }

            // after: https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
            // this builds out a cell address based on the repeated column lettering used in excel
            // such as A23 or AA1 or ACR43 from an integer index
            byte c;
            string s = "";
            while (columnIndex > 0)
            {
                c = (byte)((columnIndex - 1) % 26);
                s = Char.ConvertFromUtf32(c + 65) + s;
                columnIndex = (columnIndex - c) / 26;
            }

            // append the row index to the calculated string column address
            return string.Format("{0}{1}", s, rowIndex);
        }

        #endregion

        #region Dispose

        /// <summary>
        /// Ensure we always save the workbook once complete
        /// </summary>
        public void Dispose()
        {
            if (this.TargetWorkbook != null)
            {
                this.TargetWorkbook.Save();
            }
        }

        #endregion
    }
}
