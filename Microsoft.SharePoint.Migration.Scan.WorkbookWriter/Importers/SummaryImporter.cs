using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Diagnostics;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Importers
{
    /// <summary>
    /// handles the special summary page import 
    /// </summary>
    class SummaryImporter : ImporterBase
    {
        private static readonly string TraceCategory = "SummaryImporter";
        private static readonly string SummaryWorksheetName = "Summary";

        public SummaryImporter(SpreadsheetDocument targetWorkbook, string summaryFilePath) : base(targetWorkbook)
        {
            this.SummaryFilePath = summaryFilePath;
        }

        private string SummaryFilePath { get; set; }

        #region Import

        public override void Import()
        {
            Trace.Indent();

            // get target sheet and associated OpenXML parts
            var sheet = this.GetSheet(SummaryImporter.SummaryWorksheetName);
            var sheetPart = (WorksheetPart)this.TargetWorkbook.WorkbookPart.GetPartById(sheet.Id.Value);
            var sheetData = sheetPart.Worksheet.GetFirstChild<SheetData>();

            // start on the first row of the new worksheet
            uint currentTargetRowPointer = 1;

            Trace.WriteLine(string.Format("Importing summary file {0}.", this.SummaryFilePath), TraceCategory);

            // open csv, if it doesn't exist let the code in ImportManager catch and log that exception as it is fatal to this FileImporter
            using (var csvReader = new TextFieldParser(this.SummaryFilePath))
            {
                csvReader.Delimiters = new[] { "," };
                csvReader.TextFieldType = FieldType.Delimited;
                csvReader.HasFieldsEnclosedInQuotes = true;

                while (!csvReader.EndOfData)
                {
                    try
                    {
                        // this file has no headers, we just write what we find in the csv to the target
                        var fields = csvReader.ReadFields();

                        // add some spacing for "sections"
                        if (fields.Length == 1)
                        {
                            currentTargetRowPointer++;
                        }

                        // get the row to which we will write
                        var row = this.GetRow(sheetData, currentTargetRowPointer);

                        for (var i = 0; i < fields.Length; i++)
                        {
                            // we are adding 2 because we always want to start in column B for formatting reasons and need to also
                            // convert from zero based array to one based cell address
                            var cell = this.GetCell(sheetPart, row, this.GetCellAddress((uint)(i + 2), currentTargetRowPointer));
                            this.WriteToCell(cell, CellValues.SharedString, fields[i]);
                        }
                    }
                    catch (Exception err)
                    {
                        // we will continue from this error
                        Trace.TraceWarning("Error writing row {0} (csv line number: {1}) from summary file.", currentTargetRowPointer, csvReader.LineNumber);
                        Trace.TraceError(err.ToString());
                    }
                    finally
                    {
                        currentTargetRowPointer++;
                    }
                }
            }

            Trace.WriteLine("Completed Import of Summary file", TraceCategory);
            Trace.Unindent();
        }

        #endregion
    }
}
