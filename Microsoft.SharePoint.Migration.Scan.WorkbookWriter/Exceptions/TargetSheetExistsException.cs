using System;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Exceptions
{
    /// <summary>
    /// Indicates the sheet we want to create already exists in the template
    /// </summary>
    class TargetSheetExistsException : Exception
    {
        public TargetSheetExistsException(string sheetName) :
            base(string.Format("A sheet with the name '{0}' already exists in the template workbook, this is not supported for SAX import.", sheetName))
        {
        }
    }
}
