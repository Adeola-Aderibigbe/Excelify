using Excelify.API.Models;
using Excelify.Services;
using Excelify.Services.Extensions;
using Microsoft.AspNetCore.Mvc;

namespace Excelify.API.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        public ExcelController(ExcelifyFactory excelFactory)
        {
            _excelFactory = excelFactory;      
        }
        [HttpPost("import/sheetName")]
        public  IActionResult ImportExcel(IFormFile sheet,int sheetName = 0)
        {
            var excelService = _excelFactory.CreateService(sheet.ContentType);
            try
            {
                var ms = new MemoryStream();
                sheet.CopyTo(ms);
                ms.Position = 0;

                var teacherDtos = excelService.ImportToEntity<TeacherDto>(new ImportSheet(ms,sheetName));

                if (teacherDtos.Count == 0)
                {
                    return BadRequest("Empty excel sheet");
                }

                var teachers = teacherDtos.Select(s => new Teacher
                {
                    StartDate = s.StartDate,
                    PhoneNumber = s.PhoneNumber,
                    Email = s.Email,
                    FirstName = s.FirstName,
                    Id = s.Id,
                    LastName = s.LastName,
                    Title = s.Title
                });
                return Ok(teachers);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("export")]
        public IActionResult ExportToExcel(string sheetName,string extensionType)
        {
            var excelService = _excelFactory.CreateService(extensionType);
            var teacher = new TeacherDto() 
            {
                Email = "adeolaaderibigbe09@gmail.com",
                PhoneNumber = 0000,
                FirstName = "Ade",
                LastName = "ad",
                Id = 1,
                StartDate = DateTime.Now,
                Title = "Miss"
            };
            var exportEntity = new ExportEntity()
            { 
                Entities = new[] { teacher },
                SheetName = sheetName
            };

            try
            {
                excelService.ExportToBytes(exportEntity)
                    .ToFile($"{sheetName}.{extensionType}");
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        private readonly ExcelifyFactory _excelFactory;
    }
}
