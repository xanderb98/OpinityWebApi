# OpinityWebApi

Tool to upload an excel file which will be converted to objects

## How to use it?
- Install Visual Studio and ensure that you have the latest version of .NET (7) installed.
- It is easy to change the mapping of the Excel. Go to the ExcelController, and there is a Dictionary which is called ColumnMapping. 
The key is the corresponding value as it appears in Excel, while the value represents the name of the property in the object.
- Run the solution.
- Upload an Excel file on the /api/excel/upload endpoint. The first row of the Excel file has to contain the keys of the ColumnMapping dictionary.
