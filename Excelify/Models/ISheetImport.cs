using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelify.Models
{
    public interface ISheetImport
    {
       Stream File { get; set; }

       int SheetName { get; set; }
    }
}
