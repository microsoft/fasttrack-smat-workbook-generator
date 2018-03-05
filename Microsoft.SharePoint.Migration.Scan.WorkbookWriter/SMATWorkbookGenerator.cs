using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Linq;
using Microsoft.FastTrack.SMATWorkbookGenerator.Importers;
using System.Collections.Generic;

namespace Microsoft.FastTrack.SMATWorkbookGenerator
{
    /// <summary>
    /// Encapsulates the functionality to generate a SMAT workbook from a set of scanners and the supplied template
    /// </summary>
    public class SMATWorkbookGenerator : IDisposable
    {
        private static readonly string TraceCategory = "Microsoft.FastTrack.SMATWorkbookGenerator";
        private static readonly string BuiltInTemplateNameFormat = "Microsoft.FastTrack.SMATWorkbookGenerator.Templates.{0}.xltm";

        /// <summary>
        /// Creates a new instance of the <see cref="SMATWorkbookGenerator"/> class
        /// </summary>
        /// <param name="scanners">The set of scanners that executed during the scan. Provides the information we need for the import.</param>
        /// <param name="progressCallback">Callback funtion used to report progress</param>
        /// <param name="templatePath">Optional path to a template file than may exist</param>
        public SMATWorkbookGenerator(Settings settings)
        {
            Settings = settings;

            // define a callback here to avoid lots of if tests later
            ProgressCallback = settings.ProgressCallback == null ? (ISMATGeneratorProgressInfo msg) => { } : settings.ProgressCallback;
        }

        #region properties

        protected Settings Settings { get; set; }
        protected Action<ISMATGeneratorProgressInfo> ProgressCallback { get; set; }

        private string _outputPath;

        public string OutputPath
        {
            get
            {
                if (string.IsNullOrEmpty(_outputPath))
                {
                    _outputPath = GetOutputPath();
                }

                return _outputPath;
            }
        }

        #endregion

        #region Run

        /// <summary>
        /// Executes the actual workbook creation and import
        /// </summary>
        public void Run()
        {
            Trace.WriteLine("Beginning Import", TraceCategory);

            using (var targetWorkbook = CreateTargetWorkbook())
            {
                ProcessSummaryFile(targetWorkbook);
                ProcessDetailFiles(targetWorkbook);
            }

            Trace.WriteLine("Completed Import", TraceCategory);
        }

        #endregion

        #region ProcessSummaryFile

        /// <summary>
        /// Processes the summary file and loads it into the workbook. This is handled as a special case
        /// </summary>
        /// <param name="targetWorkbook">The workbook to which the summary report sheet will be added</param>
        private void ProcessSummaryFile(SpreadsheetDocument targetWorkbook)
        {
            Trace.Indent();

            try
            {
                Trace.WriteLine("Creating and executing SummaryImporter", TraceCategory);

                using (var summaryImporter = new SummaryImporter(targetWorkbook, Path.Combine(Settings.RootFolder, Constants.FinalSummaryReportCsv)))
                {
                    summaryImporter.Import();
                }

                Trace.WriteLine("Successfully executed SummaryImporter", TraceCategory);
            }
            catch (Exception err)
            {
                Trace.TraceError("Error in ImportManager.ProcessSummaryFile, Summary import failed.");
                Trace.TraceError(err.ToString());
            }
            finally
            {
                Trace.Unindent();
            }
        }

        #endregion

        #region ProcessDetailFiles

        private void ProcessDetailFiles(SpreadsheetDocument targetWorkbook)
        {
            // load all the files from the output 
            var files = Directory.GetFiles(Path.Combine(Settings.RootFolder, Constants.ReportFolderName))
                .Aggregate(new List<ProcessingRecord>(), (list, s) =>
                {
                    list.Add(new ProcessingRecord()
                    {
                        AbsolutePath = s,
                        Filename = Path.GetFileName(s),
                        Processed = false
                    });

                    return list;
                }).OrderBy(pr => pr.Filename).ToList();

            // this will take care of the SiteAssessmentReport.csv that is in the root folder alongside the summary report
            files.Insert(0, new ProcessingRecord()
            {
                AbsolutePath = Path.Combine(Settings.RootFolder, Constants.FinalSiteReportCsv),
                Filename = Path.GetFileName(Path.Combine(Settings.RootFolder, Constants.FinalSiteReportCsv)),
                Processed = false
            });

            // track total number of files and which one we are currently processing
            var counter = 1;
            var total = files.Count();

            try
            {
                var importerTop = Console.CursorTop;
                var importerLeft = Console.CursorLeft;

                foreach (var file in files)
                {
                    ProgressCallback(new SMATGeneratorProgressInfo()
                    {
                        TotalFilesMax = total,
                        TotalFilesPosition = counter,
                    });

                    try
                    {
                        Trace.WriteLine(string.Format("Processing import for source file {0}", file.AbsolutePath), TraceCategory);

                        using (var importer = new SheetImporter(targetWorkbook, file, ProgressCallback))
                        {
                            importer.Import();
                        }

                        file.Processed = true;

                        Trace.WriteLine(string.Format("Processed import for source file {0}", file.AbsolutePath), TraceCategory);
                    }
                    catch (Exception err)
                    {
                        Trace.TraceError(err.ToString());
                    }
                    finally
                    {
                        counter++;
                    }
                }
            }
            catch (Exception err)
            {
                Trace.TraceError("Error in ImportManager.ProcessDetailFiles, could not continue.");
                Trace.TraceError(err.ToString());
            }
            finally
            {
                // report full progress at the end for consistency
                ProgressCallback(new SMATGeneratorProgressInfo()
                {
                    TotalFilesMax = total,
                    TotalFilesPosition = total,
                });
            }
        }

        #endregion

        #region CreateTargetWorkbook

        /// <summary>
        /// Creates the workbook to which all sheets will be added, using the specified template file
        /// </summary>
        /// <returns>The newly created <see cref="SpreadsheetDocument"/></returns>
        private SpreadsheetDocument CreateTargetWorkbook()
        {
            Trace.Indent();

            try
            {
                SpreadsheetDocument targetWorkbook = null;
                SpreadsheetDocument workbook = null;

                try
                {
                    // we will either use a template from a path they specify OR one of the built in templates
                    if (!string.IsNullOrEmpty(Settings.TemplateFilePath) && File.Exists(Settings.TemplateFilePath))
                    {
                        Trace.WriteLine(string.Format("Creating target workbook from template at path ({0})", Settings.TemplateFilePath), TraceCategory);

                        workbook = SpreadsheetDocument.CreateFromTemplate(Settings.TemplateFilePath);
                        workbook.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
                        targetWorkbook = (SpreadsheetDocument)workbook.SaveAs(OutputPath);
                    }
                    else
                    {
                        Trace.WriteLine(string.Format("Creating target workbook from compiled template ({0})", Settings.TemplateName), TraceCategory);

                        // this is done to juggle limitations in open xml creating a file from a stream where the stream is readonly
                        var resourceName = string.Format(BuiltInTemplateNameFormat, Settings.TemplateName);

                        if (Assembly.GetExecutingAssembly().GetManifestResourceNames().Contains(resourceName, StringComparer.OrdinalIgnoreCase))
                        {
                            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                            {
                                workbook = SpreadsheetDocument.Open(stream, false);
                                var temp = workbook.SaveAs(OutputPath);
                                workbook.Close();
                                temp.Close(); // this allows us to reopen the file below
                            }

                            targetWorkbook = SpreadsheetDocument.Open(OutputPath, true);
                            targetWorkbook.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);
                            targetWorkbook.Save();
                        }
                        else
                        {
                            var message = string.Format("Could not locate embedded resource template {0}.", resourceName);
                            Trace.TraceError(message);
                            throw new Exception(message);
                        }
                    }

                    Trace.WriteLine(string.Format("Created target workbook at {0}", OutputPath), TraceCategory);
                }
                finally
                {
                    if (workbook != null)
                    {
                        workbook.Dispose();
                    }
                }

                return targetWorkbook;
            }
            catch (Exception err)
            {
                Trace.TraceError("Error creating output workbook from template");
                Trace.TraceError(err.ToString());
                throw;
            }
            finally
            {
                Trace.Unindent();
            }
        }

        #endregion

        #region GetOutputPath

        /// <summary>
        /// Generates an output path for the workbook file
        /// </summary>
        /// <returns>Relative string path</returns>
        private string GetOutputPath()
        {
            var outputRoot = Path.Combine(Settings.OutputFolderPath, Constants.WorkbookOutputFolder);

            // ensure we have our workbook output directory
            Directory.CreateDirectory(outputRoot);

            // used to generate output paths
            Func<int, string> _getOutputPath = (iteration) =>
            {
                var s = iteration > 0 ? "_" + iteration : "";
                return Path.Combine(outputRoot, string.Format(Constants.WorkbookFilenameFormat, s));
            };

            var outputWorkbookPath = _getOutputPath(0);
            var counter = 0;

            // just loop until we get a filename that doesn't exist to which we can write.
            while (File.Exists(outputWorkbookPath))
            {
                outputWorkbookPath = _getOutputPath(++counter);
            }

            return Path.GetFullPath(outputWorkbookPath);
        }

        #endregion

        #region Dispose

        public void Dispose()
        {
        }

        #endregion
    }

    class ProcessingRecord
    {
        public string AbsolutePath { get; set; }
        public string Filename { get; set; }
        public bool Processed { get; set; }
    }
}


