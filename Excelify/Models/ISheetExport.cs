

namespace Excelify.Models
{
    public interface ISheetExport<TEntity> 
    {
       string SheetName { get; set; }
       IList<TEntity> Entities { get; set; }
    }
}
