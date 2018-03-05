using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.FastTrack.SMATWorkbookGenerator.Exceptions;
using Microsoft.FastTrack.SMATWorkbookGenerator.Extensions;
using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Importers
{
    /// <summary>
    /// Handles importing the individual detail csv files into the workbook
    /// </summary>
    /// <remarks>
    /// This class makes some assumptions about the file it is importing
    /// 1) all rows are written into the sheet as-is
    /// 2) it makes its best guess as to the value type in each column
    /// 3) the sheet title, possibly manipulated/shortened, will be set to the filename without extension
    /// </remarks>
    class SheetImporter : ImporterBase
    {
        private static readonly string TraceCategory = "SheetImporter";
        private Action<ISMATGeneratorProgressInfo> _progressCallback;

        /// <summary>
        /// Creates a new instance of the <see cref="GenericImporter"/> class
        /// </summary>
        /// <param name="targetWorkbook">Workbook intowhich the new sheet will be imported</param>
        /// <param name="filePath">The absolute file path to the file to import</param>
        /// <param name="progressCallback">A callback used to report progress within this importer</param>
        public SheetImporter(SpreadsheetDocument targetWorkbook, ProcessingRecord fileInfo, Action<ISMATGeneratorProgressInfo> progressCallback) : base(targetWorkbook)
        {
            this.FileInfo = fileInfo;
            this._progressCallback = progressCallback;
        }

        /// <summary>
        /// The absolute file path to the file to import 
        /// </summary>
        public ProcessingRecord FileInfo { get; protected set; }

        /// <summary>
        /// Conducts the import of the file to the workbook
        /// </summary>
        public override void Import()
        {
            Trace.Indent();
            Trace.WriteLine(string.Format("Importing file {0}", this.FileInfo.AbsolutePath), TraceCategory);

            if (!File.Exists(this.FileInfo.AbsolutePath))
            {
                Trace.TraceWarning("Importing file {0}", this.FileInfo.AbsolutePath);
                return;
            }

            var csvSourcePath = File.OpenRead(this.FileInfo.AbsolutePath);

            // get the existing sheets collection
            var sheets = this.TargetWorkbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();

            // used as the name for the sheet and import tracking display
            var sheetName = this.GetSheetName(sheets);

            // we need to check to see if a sheet with this name already exists
            if (sheets.Cast<Sheet>().Where(s => s.Name.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase)).Any())
            {
                throw new TargetSheetExistsException(sheetName);
            }

            var totalLineCount = File.ReadAllLines(this.FileInfo.AbsolutePath).LongCount();
            _progressCallback(new SMATGeneratorProgressInfo()
            {
                CurrentFileMax = totalLineCount,
                CurrentFileName = this.FileInfo.Filename,
                CurrentFilePosition = 0,
                TotalFilesMax = 0,
                TotalFilesPosition = 0,
            });

            // create the new worksheet
            var worksheetPart = this.TargetWorkbook.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheet = new Sheet() { Id = this.TargetWorkbook.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = sheets.NextAvailableSheetId(), Name = sheetName };
            sheets.Append(sheet);

            using (var csvReader = new TextFieldParser(csvSourcePath))
            {
                csvReader.Delimiters = new[] { "," };
                csvReader.TextFieldType = FieldType.Delimited;
                csvReader.HasFieldsEnclosedInQuotes = true;
                csvReader.TrimWhiteSpace = true;

                // used to track the current row
                uint currentTargetRowPointer = 1;
                var typeEnumMapping = new List<CellValues>();
                string cellAddress;
                string[] fields;

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    while (!csvReader.EndOfData)
                    {
                        fields = csvReader.ReadFields();

                        try
                        {
                            // line 2 means "fields" actually has the values from row 1 but the pointer has been advanced in LineNumber
                            if (csvReader.LineNumber == 2)
                            {
                                using (writer.StartRow(currentTargetRowPointer))
                                {
                                    // the original template sheets had two empty columns, so we match that
                                    writer.WriteCell(string.Format("B{0}", currentTargetRowPointer), "Remediation", CellValues.InlineString);

                                    for (var i = 0; i < fields.Length; i++)
                                    {
                                        // the +3 here is derived from a need to +1 for the 1 based columns and +2 as we skip the first two columns by design
                                        cellAddress = this.GetCellAddress((uint)(i + 3), currentTargetRowPointer);
                                        writer.WriteCell(cellAddress, fields[i], CellValues.InlineString);
                                    }
                                }
                            }
                            else
                            {
                                // line 3 means "fields" actually has the values from row 2 but the pointer has been advanced in LineNumber
                                if (csvReader.LineNumber == 3)
                                {
                                    // we need to use the first row of data to determine as best we can the column types
                                    for (var i = 0; i < fields.Length; i++)
                                    {
                                        typeEnumMapping.Add(this.GetCellDataType(fields[i]));
                                    }
                                }

                                using (writer.StartRow(currentTargetRowPointer))
                                {
                                    for (var i = 0; i < fields.Length; i++)
                                    {
                                        // the +3 here is derived from a need to +1 for the 1 based columns and +2 as we skip the first two columns by design
                                        cellAddress = this.GetCellAddress((uint)(i + 3), currentTargetRowPointer);
                                        writer.WriteCell(cellAddress, fields[i], typeEnumMapping[i]);
                                    }
                                }
                            }
                        }
                        catch (Exception err)
                        {
                            Trace.TraceError(string.Format("Error writing row {0} (csv line number: {1}) for target sheet {2}. Error: {3}", currentTargetRowPointer, csvReader.LineNumber, sheetName, err));
                        }
                        finally
                        {
                            // move to the next row
                            currentTargetRowPointer++;
                        }

                        if (csvReader.LineNumber > 0)
                        {
                            _progressCallback(new SMATGeneratorProgressInfo()
                            {
                                CurrentFileMax = totalLineCount,
                                CurrentFileName = this.FileInfo.Filename,
                                CurrentFilePosition = currentTargetRowPointer,
                                TotalFilesMax = 0,
                                TotalFilesPosition = 0,
                            });
                        }
                    }

                    writer.WriteEndElement(); // end sheet data
                    writer.WriteEndElement(); // end worksheet
                }
            }

            // always end with full progress reported
            _progressCallback(new SMATGeneratorProgressInfo()
            {
                CurrentFileMax = totalLineCount,
                CurrentFileName = this.FileInfo.Filename,
                CurrentFilePosition = totalLineCount,
                TotalFilesMax = 0,
                TotalFilesPosition = 0,
            });

            Trace.Unindent();
        }

        #region GetCellDataType

        /// <summary>
        /// Tries to determine the CellValues type vased on the supplied value
        /// </summary>
        /// <param name="candidateValue">The raw cell value</param>
        /// <returns>The determined CellValues</returns>
        private CellValues GetCellDataType(string candidateValue)
        {
            bool b;
            if (bool.TryParse(candidateValue, out b))
            {
                return CellValues.Boolean;
            }

            decimal d;
            if (decimal.TryParse(candidateValue, out d))
            {
                return CellValues.Number;
            }

            DateTime dt;
            if (DateTime.TryParse(candidateValue, out dt))
            {
                return CellValues.Date;
            }

            return CellValues.InlineString;
        }

        #endregion

        #region GetSheetName

        /// <summary>
        /// Method to avoid sheet name collisions for generic imports
        /// </summary>
        /// <param name="sheets">Collection of sheets</param>
        /// <returns>The sheet name that should be used</returns>
        private string GetSheetName(Sheets sheets)
        {
            // this logic is here to get around limitations in excel sheet name length and file name collision.
            // if we trim our filenames to 30 chars we end up with names that are identical for sheets, which is dissallowed

            // replace the long prefix string if present and get a name of the right length
            var protoName = Regex.Replace(Path.GetFileNameWithoutExtension(this.FileInfo.AbsolutePath), "^FullTrustSolution_", "FTS_");
            protoName = Regex.Replace(protoName, "-detail$", "");
            var sheetName = sheets.MakeSafeSheetName(protoName);

            var counter = 0;
            while (sheets.SheetExists(sheetName))
            {
                // this obtuse line of code "knows" that the correct sheetName length is set above by "MakeSafeSheetName"
                // so we keep using that as a reference and trimming so we don't exceed that length with the counter we are appending
                sheetName = sheetName.Substring(0, sheetName.Length - counter.ToString().Length) + counter.ToString();
                counter++;
            }

            return sheetName;
        }

        #endregion
    }
}
