using System;

namespace Microsoft.FastTrack.SMATWorkbookGenerator
{
    public class Settings
    {
        public string RootFolder { get; set; }
        public string TemplateName { get; set; }
        public string TemplateFilePath { get; set; }
        public string OutputFolderPath { get; set; }
        public Action<ISMATGeneratorProgressInfo> ProgressCallback { get; set; }
    }
}
