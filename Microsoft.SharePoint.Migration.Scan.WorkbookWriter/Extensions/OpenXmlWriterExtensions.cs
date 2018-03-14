using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Extensions
{
    /// <summary>
    /// Extension methods for working with <see cref="OpenXmlWriter"/>
    /// </summary>
    public static class OpenXmlWriterExt
    {
        /// <summary>
        /// Writes data to a cell
        /// </summary>
        /// <param name="writer">this OpenXmlWriter writer instance</param>
        /// <param name="cellAddress">The cell address</param>
        /// <param name="cellValue">The cell value</param>
        /// <param name="dataType">The data type for the cell</param>
        public static void WriteCell(this OpenXmlWriter writer, string cellAddress, string cellValue, CellValues dataType)
        {
            // default to a standard cell value
            OpenXmlElement value = new CellValue(cellValue);
            bool testBool;
            long testLong;

            // fix up some values
            switch (dataType)
            {
                // we are handling all strings as inline for performance reasons
                case CellValues.SharedString:
                case CellValues.InlineString:

                    dataType = CellValues.InlineString;
                    value = new InlineString(new Text(cellValue));
                    break;

                case CellValues.Date:

                    // write the value as a string to the sheet
                    dataType = CellValues.String;
                    break;

                case CellValues.Number:

                    // this is a safety check as we sometimes get bad values such as N/A in columns
                    if (!long.TryParse(cellValue, out testLong))
                    {
                        // we default to writing it as a string to be safe
                        dataType = CellValues.InlineString;
                        value = new InlineString(new Text(cellValue));
                    }
                    break;

                case CellValues.Boolean:

                    // this is a safety check as we sometimes get bad values such as N/A in columns
                    if (bool.TryParse(cellValue, out testBool))
                    {
                        value = new CellValue(testBool ? "1" : "0");
                    }
                    else
                    {
                        // we default to writing it as a string to be safe
                        dataType = CellValues.InlineString;
                        value = new InlineString(new Text(cellValue));
                    }
                    break;
            }

            // write cell xml to the writer
            writer.WriteStartElement(new Cell() { DataType = dataType, CellReference = cellAddress });
            writer.WriteElement(value);
            writer.WriteEndElement();
        }

        /// <summary>
        /// Starts a row
        /// </summary>
        /// <param name="writer">this OpenXmlWriter writer instance</param>
        /// <param name="rowIndex">The desired row index</param>
        /// <returns>A disposable that will write the row end tag when disposed</returns>
        public static RowEnder StartRow(this OpenXmlWriter writer, uint rowIndex)
        {
            writer.WriteStartElement(new Row() { RowIndex = rowIndex });

            return new RowEnder(writer);
        }
    }

    /// <summary>
    /// Allows use of the using() pattern to wrap rows and ensure they are closed
    /// </summary>
    public class RowEnder : IDisposable
    {
        private OpenXmlWriter _writer;

        /// <summary>
        /// Creates a new instance of the RowEnder class
        /// </summary>
        /// <param name="writer"></param>
        public RowEnder(OpenXmlWriter writer)
        {
            _writer = writer;
        }

        public void Dispose()
        {
            _writer.WriteEndElement();
        }
    }
}
