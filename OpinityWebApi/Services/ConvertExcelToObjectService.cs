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
            // Create a list of ConvertedObjects
            List<ConvertedObject> list = new List<ConvertedObject>();
            IXLWorksheet worksheet = excel.Worksheet(1);

            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip properties row
            
            // Loop through all the rows except the first one
            foreach (var row in rows)
            {
                // Create an instance of ConvertedObject
                ConvertedObject objectInstance = (ConvertedObject)Activator.CreateInstance(typeof(ConvertedObject), null);

                // Loop through the columnmapping
                foreach (KeyValuePair<string, string> column in columnMapping)
                {
                    var firstRowCells = worksheet.RangeUsed().Row(1).Cells();
                    
                    // Loop through the first row cells
                    foreach (var cell in firstRowCells)
                    {
                        // Check if the cell value equals a mapping column
                        if (cell.Value.ToString() == column.Key)
                        {
                            // Get the value of the property
                            var propertyValue = worksheet.Cell(row.RowNumber(), cell.WorksheetColumn().ColumnNumber()).Value.ToString();

                            // column.Value contains the property name, propertyValue contains the value of the property
                            MethodInfo methodInfo = typeof(ConvertedObject).GetMethod("AddProperty");
                            object[] parametersArray = new object[] { column.Value, propertyValue };
                            methodInfo.Invoke(objectInstance, parametersArray);

                            break;
                        }
                    }

                }
                // Add the object to the list
                list.Add(objectInstance);
            }

            return list;
        }
    }
}
