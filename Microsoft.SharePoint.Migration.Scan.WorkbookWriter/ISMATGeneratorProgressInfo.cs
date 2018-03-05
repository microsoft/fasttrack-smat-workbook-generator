namespace Microsoft.FastTrack.SMATWorkbookGenerator
{
    /// <summary>
    /// Defines the progress info we will report to the progress callback
    /// </summary>
    public interface ISMATGeneratorProgressInfo
    {
        string CurrentFileName { get; }
        long CurrentFilePosition { get; }
        long CurrentFileMax { get; }
        string Message { get; }
        int TotalFilesPosition { get; }
        int TotalFilesMax { get; }
        int CurrentFilePercentage { get; }
    }

    /// <summary>
    /// A class used to report progress info
    /// </summary>
    internal class SMATGeneratorProgressInfo : ISMATGeneratorProgressInfo
    {
        public string CurrentFileName { get; set; }
        public long CurrentFilePosition { get; set; }
        public long CurrentFileMax { get; set; }
        public string Message { get; set; }
        public int TotalFilesPosition { get; set; }
        public int TotalFilesMax { get; set; }

        public SMATGeneratorProgressInfo()
        {
            CurrentFileName = null;
            CurrentFilePosition = -1;
            CurrentFileMax = -1;
            Message = null;
            TotalFilesPosition = -1;
            TotalFilesMax = -1;
        }

        public int CurrentFilePercentage
        {
            get
            {
                return (int)(((decimal)CurrentFilePosition / (decimal)CurrentFileMax) * 100);
            }
        }
    }
}