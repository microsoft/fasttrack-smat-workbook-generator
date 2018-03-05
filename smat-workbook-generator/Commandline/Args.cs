using System;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Commandline
{
    class Args : ArgsBase
    {
        public Args(string[] args) : base(args) { }

        [ArgProp("-s", "Path to the root folder containing the source SMAT output.")]
        public string SMATFolderPath { get; set; }

        [ArgProp("-t", "Name of one of the included template files to use", "FastTrack")]
        public string TemplateName { get; set; }

        [ArgProp("-tp", "Path to the xltm template used to format the output.", "")]
        public string TemplateFilePath { get; set; }

        [ArgProp("-o", "Path to the output folder where the resulting spreadsheet will be written.", "")]
        public string OutputFolderPath { get; set; }

        //[ArgProp("-c", "Path to the mappings file used to map the sheets into the template.", "smatwb-config.json", Required = false)]
        //public string ConfigurationFilePath { get; set; }

        public Settings ToSettings()
        {
            return new Settings()
            {
                RootFolder = SMATFolderPath,
                TemplateName = TemplateName,
                TemplateFilePath = TemplateFilePath,
                OutputFolderPath = OutputFolderPath
            };
        }
    }
}
