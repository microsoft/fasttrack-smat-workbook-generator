# Microsoft FastTrack

## SMAT Workbook Generator

The SMAT workbook generator is a tool that takes the output from [Microsoft SMAT](https://www.microsoft.com/en-us/download/details.aspx?id=53598) and tranforms it into a single wookbook file. Originally this was done as part of our engagements with customers within the FastTrack program, but we are releasing here as OSS as other might find it useful.

## Usage

Please find the [latest release](https://github.com/Microsoft/fasttrack-smat-workbook-generator/releases) and download and unzip the file. It contains the smatwb.exe as well as supporting files and DLLS. You may need to "unblock" the files after downloading depending on your system settings.

This tool runs on the command line and takes several parameters as described below.

### Specify the source folder (required)

The -s argument takes the path to the root of the SMAT output. This is the folder containing the "SummaryReport.csv" file.

`smatwb -s "c:\path\to\SMAT\output"`

### Specify the wokbook output folder

`smatwb -s "c:\path\to\SMAT\output" -o "c:\path\to\write\workbook\"`

### Specify template name

`smatwb -s "c:\path\to\SMAT\output" -t Simple`

### All Options

|Param|Description|Default|
|:----:|--------------------------|---------------|
|-s|Source folder|none, required|
|-o|Output folder|./Workbook|
|-t|Template Name|FastTrack|
|-tp|Path to template|none|

## Custom Template

Some templates are built into the solution, currently these are FastTrack and Simple. The FastTrack template is the default and contains FastTrack formatting and branding. You can optionally switch to the simple template using the `-t Simple` argument. If you would like to create and use your own template you can download either of the existing templates and modify them as needed. Once your template is ready you can use the `-tp "c:\path\tomy\temlate"` argument to use it in the import.


## Issues

If you find issues or have suggestions for enhancements please [submit an issue](https://github.com/Microsoft/fasttrack-smat-workbook-generator/issues) so we can respond and track it. Be sure to include enough information so we can understand what's happening and respond - but also remember that the issues list is public so please don't include any PII or farm specific information when reporting an issue.


## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
