using Microsoft.FastTrack.SMATWorkbookGenerator.Commandline;
using Microsoft.FastTrack.SMATWorkbookGenerator.Localization;
using System;
using System.IO;

namespace Microsoft.FastTrack.SMATWorkbookGenerator
{
    /// <summary>
    /// command line application to run the workbook generation
    /// </summary>
    class Program
    {
        // used to store the positions for console display
        private static int overallTop = -1;
        private static int overallLeft = -1;
        private static int fileTop = -1;
        private static int fileLeft = -1;
        private static string _mask = null;

        /// <summary>
        /// Main entry point
        /// </summary>
        /// <param name="args">Commandline arguments</param>
        static void Main(string[] args)
        {
            // process our command line
            var arguments = new Args(args);

            // if user wants help we just display that here and exit
            if (arguments.HelpFlag)
            {
                arguments.WriteHelp(Strings.AppDisplayTitle, Strings.AppDisplayDesc, Console.Out);
                return;
            }

            // setup our display
            Console.Clear();
            Console.CursorVisible = false;
            Console.WriteLine(Strings.AppDisplayTitle);
            Console.WriteLine();

            // get our cursor positions for displaying progress on the console
            overallTop = Console.CursorTop;
            overallLeft = Console.CursorLeft;
            Console.WriteLine();
            fileTop = Console.CursorTop;
            fileLeft = Console.CursorLeft;

            string outputPath = "";

            // validate the arguments
            if (ValidateArguments(arguments))
            {
                var settings = arguments.ToSettings();
                settings.ProgressCallback = ProgressReporter;

                // create our generator from the supplied settings
                using (var generator = new SMATWorkbookGenerator(settings))
                {
                    // grab our computed output path and run the generator
                    outputPath = generator.OutputPath;
                    generator.Run();
                }

                // output some final messages and return the console settings to what they were
                Console.WriteLine();
                var fg = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine(Strings.ImportCompleteMsgFormat, outputPath);
                Console.ForegroundColor = fg;
            }

            Console.WriteLine();
            Console.CursorVisible = true;
            Console.WriteLine(Strings.PressAnyKeyToExit);
            Console.ReadKey();
        }

        #region ValidateArguments

        /// <summary>
        /// Validates the arguments supplied by the user
        /// </summary>
        /// <param name="arguments">The set of user supplied arguments</param>
        /// <returns>True if the arguments are valid; otherwise, false</returns>
        private static bool ValidateArguments(Args arguments)
        {
            // ensure the input directory exists
            var path = Path.Combine(arguments.SMATFolderPath, Constants.ReportFolderName);
            if (!Directory.Exists(path))
            {
                Console.WriteLine(Strings.SourceFolderDoesNotExist, path);
                return false;
            }

            // ensure the summary file exists in the input folder path
            var summaryFilePath = Path.Combine(arguments.SMATFolderPath, Constants.FinalSummaryReportCsv);
            if (!File.Exists(summaryFilePath))
            {
                Console.WriteLine(Strings.SummaryFileDoesNotExist);
                return false;
            }

            // ensure the template file specified exists if one was specified
            if (!string.IsNullOrEmpty(arguments.TemplateFilePath) && !File.Exists(arguments.TemplateFilePath))
            {
                Console.WriteLine(Strings.TemplateFileDoesNotExist, arguments.TemplateFilePath);
                return false;
            }

            return true;
        }

        #endregion

        #region ProgressReporter

        /// <summary>
        /// Reports progress to the screen based on the supplied info object
        /// </summary>
        /// <param name="info">Contains details on progress from the <see cref="SMATWorkbookGenerator"/></param>
        public static void ProgressReporter(ISMATGeneratorProgressInfo info)
        {
            int left = Console.CursorLeft;
            int top = Console.CursorTop;

            if (string.IsNullOrEmpty(info.CurrentFileName))
            {
                // we have switched files so update line 1 and clear line 2
                Console.SetCursorPosition(overallLeft, overallTop);
                Console.Write("Importing file {0} of {1}.", info.TotalFilesPosition, info.TotalFilesMax);
                Console.SetCursorPosition(fileLeft, fileTop);
                Console.Write(Mask);
            }
            else
            {
                // we are updating a file so update line 2
                Console.SetCursorPosition(fileLeft, fileTop);
                Console.Write(Mask);
                Console.SetCursorPosition(fileLeft, fileTop);
                Console.Write("Importing {0}...{1}%", info.CurrentFileName, info.CurrentFilePercentage);
            }

            // reset cursor
            if (left > 0 || top > 0)
            {
                Console.SetCursorPosition(left, top);
            }
        }

        /// <summary>
        /// A mask used to blank out lines as needed to keep the output clean and consistent
        /// </summary>
        private static string Mask
        {
            get
            {
                if (_mask == null)
                {
                    for (var i = 0; i < Console.WindowWidth; ++i)
                    {
                        _mask += " ";
                    }
                }

                return _mask;
            }
        }

        #endregion
    }
}
