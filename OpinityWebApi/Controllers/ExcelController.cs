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

        // The columnmapping is straightforward, the key is the corresponding value as it appears in Excel,
        // while the value represents the name of the property in the object
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
            // Check if there is a file
            if (file == null || file.Length == 0)
            {
                return BadRequest("There was no file selected");
            }

            // Check if the file matches an .xlsx or .xls extension
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
                return BadRequest("Format does not match .xlsx or .xls extensions");
            }

            // Convert file to XLWorkbook
            XLWorkbook workbook = new XLWorkbook(file.OpenReadStream());

            // Convert workbook to list with object(s)
            var result = _convertExcelToObject.Convert<ConvertedObject>(workbook, ColumnMapping);
            // Convert list with object(s) to json object(s)
            var serialized = JsonConvert.SerializeObject(result, Formatting.Indented);

            return Ok(serialized);
        }
    }
}
