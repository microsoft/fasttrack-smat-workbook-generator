using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace Microsoft.FastTrack.SMATWorkbookGenerator.Extensions
{
    public static class SheetsExt
    {
        /// <summary>
        /// Calculates the next available SheetId for this sheets collection
        /// </summary>
        /// <param name="sheets">this Sheets instance</param>
        /// <returns>The next available id </returns>
        public static uint NextAvailableSheetId(this Sheets sheets)
        {
            uint sheetId = 0;
            foreach (var s in sheets.Cast<Sheet>())
            {
                if (s.SheetId >= sheetId)
                {
                    sheetId = s.SheetId + 1;
                }
            }
            return sheetId;
        }

        /// <summary>
        /// Ensures a candidate sheet name is safe for use as a sheet name in Excel
        /// </summary>
        /// <param name="sheets">this Sheets instance</param>
        /// <param name="name">Candidate sheet name</param>
        /// <returns>Safe sheet name</returns>
        public static string MakeSafeSheetName(this Sheets sheets, string name)
        {
            return name.Length < 30 ? name : name.Substring(0, 30);
        }

        /// <summary>
        /// Determines if a sheet already exists by name
        /// </summary>
        /// <param name="sheets">this Sheets instance</param>
        /// <param name="name">Potential sheet name</param>
        /// <returns>True if the sheet exists; otherwise, false</returns>
        public static bool SheetExists(this Sheets sheets, string name)
        {
            return sheets.Cast<Sheet>().Any(s => s.Name.Value.Equals(name, System.StringComparison.OrdinalIgnoreCase));
        }
    }
}

