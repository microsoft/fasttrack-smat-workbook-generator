namespace Microsoft.FastTrack.SMATWorkbookGenerator.Commandline
{
    /// <summary>
    /// Represents the arguments supplied to the program
    /// </summary>
    class Args : ArgsBase
    {
        public Args(string[] args) : base(args) { }

        /// <summary>
        /// Gets or sets the path to the root folder containing the source SMAT output
        /// </summary>
        [ArgProp("-s", "Path to the root folder containing the source SMAT output.")]
        public string SMATFolderPath { get; set; }

        /// <summary>
        /// Gets or sets the name of one of the included template files to use
        /// </summary>
        [ArgProp("-t", "Name of one of the included template files to use", "FastTrack")]
        public string TemplateName { get; set; }

        /// <summary>
        /// Gets or sets the path to the external xltm template used to format the output
        /// </summary>
        [ArgProp("-tp", "Path to the external xltm template used to format the output.", "")]
        public string TemplateFilePath { get; set; }

        /// <summary>
        /// Gets or sets the path to the output folder where the resulting spreadsheet will be written
        /// </summary>
        [ArgProp("-o", "Path to the output folder where the resulting spreadsheet will be written.", "")]
        public string OutputFolderPath { get; set; }

        /// <summary>
        /// Converts this argument instance into a <see cref="Settings"/> instance
        /// </summary>
        /// <returns></returns>
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
