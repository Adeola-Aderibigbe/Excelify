using Excelify.Models;

namespace Excelify.API.Models
{
    public class ExportEntity : ISheetExport<TeacherDto>
    {
        public string SheetName { get; set ; }
        public IList<TeacherDto> Entities { get; set ; }
    }
}
