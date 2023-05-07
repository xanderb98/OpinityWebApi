using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection;

namespace OpinityWebApi.Services
{
    public interface IConvertExcelToObject
    {
        List<ConvertedObject> Convert<ConvertedObject>(XLWorkbook excel, Dictionary<string, string> columnMapping);
    }

    public class ConvertExcelToObjectService : IConvertExcelToObject
    {
        public List<ConvertedObject> Convert<ConvertedObject>(XLWorkbook excel, Dictionary<string, string> columnMapping)
        {
            List<ConvertedObject> list = new List<ConvertedObject>();
            IXLWorksheet worksheet = excel.Worksheet(1);

            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row
            foreach (var row in rows)
            {
                ConvertedObject classInstance = (ConvertedObject)Activator.CreateInstance(typeof(ConvertedObject), null);

                foreach (KeyValuePair<string, string> column in columnMapping)
                {
                    var firstRowCells = worksheet.RangeUsed().Row(1).Cells();

                    foreach (var cell in firstRowCells)
                    {
                        if (cell.Value.ToString() == column.Key)
                        {
                            var propertyValue = worksheet.Cell(row.RowNumber(), cell.WorksheetColumn().ColumnNumber()).Value.ToString();

                            MethodInfo methodInfo = typeof(ConvertedObject).GetMethod("AddProperty");
                            object[] parametersArray = new object[] { column.Value, propertyValue };
                            methodInfo.Invoke(classInstance, parametersArray);

                            break;
                        }
                    }

                }
                list.Add(classInstance);
            }

            return list;
        }
    }
}
