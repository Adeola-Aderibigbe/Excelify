using Excelify.Models;

namespace Excelify.API.Models
{
    public class ImportSheet(Stream file, int sheetName) : ISheetImport
    {
        public Stream File { get; set; } = file;
        public int SheetName { get; set; } = sheetName;
    }
}
