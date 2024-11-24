using FilePorcessorTemplate.Models;
using ClosedXML.Excel;

namespace FilePorcessorTemplate.Helpers;

public class FileTransformerService
{
     public GenericFile TransformFile(string inputFilePath)
     {
          string fileExtension = Path.GetExtension(inputFilePath);

          return ConvertFileToGeneric(fileExtension, inputFilePath);
     }

     private GenericFile ConvertFileToGeneric(string extension, string inputFilePath)
     {
          using var workbook = new XLWorkbook(inputFilePath);

          var genericFile = new GenericFile();
          var worksheet = workbook.Worksheet(1);
          int rowCount = GetEndOfDataRows(worksheet);
          int colCount = GetEndOfDataColumns(worksheet);

          genericFile.ColumnNames = new List<string>();
          genericFile.ColumnValues = new List<object[]>();

          for (int row = 1; row <= rowCount + 1; row++)
          {
               if (row == 1)
               {
                    var headerData = GetColumnValues(worksheet, row, colCount);

                    for(int header = 0; header < headerData.Count; header++)
                    {
                         genericFile.ColumnNames.Add(headerData[header].ToString());
                    }
               }
               else
               {
                    genericFile.ColumnValues.Add(GetColumnValues(worksheet, row, colCount).ToArray());
               }

          }

          return genericFile;
     }

     private List<object> GetColumnValues(IXLWorksheet worksheet, int rowNumber, int columnDepth)
     {
          var rowData = new List<object>();

          for (int col = 1; col <= columnDepth; col++)
          {
               switch (worksheet.Cell(rowNumber, col).Value.Type)
               {
                    case XLDataType.Boolean:
                         rowData.Add((bool)worksheet.Cell(rowNumber, col).Value);
                         break;

                    case XLDataType.Number:
                         if (int.TryParse(worksheet.Cell(rowNumber, col).Value.ToString(), out int intNumber))
                         {
                              rowData.Add(intNumber);
                         }
                         else if (double.TryParse(worksheet.Cell(rowNumber, col).Value.ToString(), out double doubleNumber))
                         {
                              rowData.Add(doubleNumber);
                         }

                         break;

                    default:
                    case XLDataType.Text:
                         rowData.Add(worksheet.Cell(rowNumber, col).Value.ToString());
                         break;

                    case XLDataType.DateTime:
                         if (DateTime.TryParse(worksheet.Cell(rowNumber, col).Value.ToString(), out DateTime dateValue))
                         {
                              rowData.Add(dateValue);
                         }
                         break;
               }
          }

          return rowData;
     }

     private int GetEndOfDataColumns(IXLWorksheet worksheet)
     {
          for (int col = 1; col <= worksheet.ColumnCount(); col++)
          {
               if (worksheet.Cell(1, col).Value.Type == XLDataType.Blank || worksheet.Cell(1, col).Value.ToString() == string.Empty)
               {
                    return col - 1;
               }
          }

          throw new Exception("No data found");
     }

     private int GetEndOfDataRows(IXLWorksheet worksheet)
     {
          for (int row = 2; row <=  worksheet.RowCount(); row++)
          {
               var rowData = worksheet.Cell(row, 1).Value;
               if (rowData.Type == XLDataType.Blank || rowData.ToString() == string.Empty)
               {
                    return row - 2;
               }
          }

          throw new Exception("No data found");
     }

}