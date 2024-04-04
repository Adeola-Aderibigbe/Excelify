using Excelify.Models;
using Excelify.Services;
using Excelify.Test.Models;
using Moq;
using NPOI.XSSF.UserModel;

namespace Excelify.Test
{
    public class ExcelifyServiceTests
    {
        [SetUp]
        public void Setup()
        {
            excelifyService = new Mock<ExcelifyService>().Object;
            excelifyService.SetSheetName(1, ExtensionType.xlsx.ToString());
            XSSFWorkbook workBook = new("C:\\Users\\adeola.aderibigbe\\Documents\\local_git\\Excelify\\Excelify.API\\first_try.xlsx");
            using var ms = new MemoryStream();
             workBook.Write(ms);
            ms.Position = 0;
            IImportSheet importSheet = new ImportEntity() 
            {
                File = ms
            };

        }

        [Test]
        public void ShouldImportSheetSuccessfully()
        {
            excelifyService.SetSheetName(0, ExtensionType.xlsx.ToString());



        }

        private ExcelifyService excelifyService;
    }
}