using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OpinityWebApi.Services;

namespace OpinityWebApi.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IConvertExcelToObject _convertExcelToObject;

        public ExcelController(IConvertExcelToObject convertExcelToObject)
        {
            _convertExcelToObject = convertExcelToObject;
        }

        public Dictionary<string, string> ColumnMapping { get; set; } = new Dictionary<string, string>
        {
            {"Voornaam", "firstname"},
            {"Achternaam", "lastname"},
            {"Gebruikersnaam", "username"},
            {"Wachtwoord", "password"}
        };
        
        /// <summary>
        /// Upload an excel file to map it to an object structure.
        /// </summary>
        /// <remarks>Choose a file with an .xlsx or .xls extension</remarks>
        /// <returns></returns>
        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("There was no file selected");
            }

            string[] validFileTypes = { "xlsx", "xls" };
            
            string ext = Path.GetExtension(file.FileName);
            
            bool isValidFile = false;
            foreach (string validFileType in validFileTypes)
            {
                if (ext == "." + validFileType)
                {
                    isValidFile = true;
                    break;
                }
            }

            if (!isValidFile)
            {
                return BadRequest("Format does not match .xlsx and .xls extensions");
            }

            XLWorkbook workbook = new XLWorkbook(file.OpenReadStream());

            var result = _convertExcelToObject.Convert<ConvertedObject>(workbook, ColumnMapping);
            var serialized = JsonConvert.SerializeObject(result, Formatting.Indented);

            return Ok(serialized);
        }
    }
}
