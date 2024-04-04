using Excelify.Models;


namespace Excelify.Test.Models
{
    internal class ImportEntity : IImportSheet
    {
        public Stream File { get; set; }
    }
}
