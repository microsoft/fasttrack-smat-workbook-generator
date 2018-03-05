namespace Microsoft.FastTrack.SMATWorkbookGenerator
{
    public static class Constants
    {
        /// <summary>
        /// Folder containing detail reports from SMAT
        /// </summary>
        public readonly static string ReportFolderName = "ScannerReports";

        /// <summary>
        /// Filename for final site report
        /// </summary>
        public readonly static string FinalSiteReportCsv = "SiteAssessmentReport.csv";

        /// <summary>
        /// Default output folder for generated workbook
        /// </summary>
        public readonly static string WorkbookOutputFolder = "Workbook";

        /// <summary>
        /// Filename format for generated workbook
        /// </summary>
        public readonly static string WorkbookFilenameFormat = "SMATWorkbook{0}.xlsm";

        /// <summary>
        /// Summary report filename withing SMAT rootfolder
        /// </summary>
        public readonly static string FinalSummaryReportCsv = "SummaryReport.csv";
    }
}
