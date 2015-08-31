using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Humanizer;

namespace ExcelParser
{
    public abstract class BaseExcelParser : IExcelParser
    {
        public IList<T> Parse<T>(string fileName) where T : new()
        {
            using (var doc = SpreadsheetDocument.Open(fileName, false))
            {
                var bookPart = doc.WorkbookPart;
                var sheetPart = bookPart.WorksheetParts.First();
                var sheetData = sheetPart.Worksheet.Elements<SheetData>().First();
                var strings = bookPart.GetPartsOfType<SharedStringTablePart>().First().SharedStringTable;

                var list = new List<T>();
                var first = true;
                var columns = new Dictionary<string, string>();

                foreach (var row in sheetData.Elements<Row>())
                {
                    if (first)
                    {
                        first = false;
                        ParseColumnNames(row, strings, columns);
                    }
                    else
                    {
                        list.Add(ParseObject<T>(row, strings, columns));
                    }

                }

                return list;
            }
        }

        private T ParseObject<T>(Row row, SharedStringTable strings, Dictionary<string, string> columns)
            where T : new()
        {
            var obj = new T();

            foreach (var cell in row.Elements<Cell>())
            {
                var column = columns[GetColumnFromCell(cell)];
                var cellValueString = GetValueAsString(cell, strings);
                var property = typeof(T).GetProperty(column);

                object setValue = cellValueString;

                if (property.PropertyType == typeof(int))
                {
                    setValue = int.Parse(cellValueString);
                }
                else if (property.PropertyType == typeof(double))
                {
                    setValue = double.Parse(cellValueString);
                }
                else if (property.PropertyType == typeof(decimal))
                {
                    setValue = decimal.Parse(cellValueString);
                }
                else if (property.PropertyType == typeof(DateTime))
                {
                    setValue = DateTime.FromOADate(double.Parse(cellValueString));
                }

                property.SetValue(obj, setValue);
            }

            return obj;
        }

        private void ParseColumnNames(Row row, SharedStringTable strings, Dictionary<string, string> columns)
        {
            foreach (var cell in row.Elements<Cell>())
            {
                columns.Add(
                    GetColumnFromCell(cell),
                    GetValueAsString(cell, strings).Dehumanize()
                );
            }
        }

        private string GetColumnFromCell(Cell cell)
        {
            var regex = new Regex("[A-Za-z]+");
            var match = regex.Match(cell.CellReference);
            return match.Value;
        }

        private string GetValueAsString(Cell cell, SharedStringTable strings)
        {
            var cv = cell.CellValue.Text;
            var dataType = cell.DataType?.Value;

            if (dataType.HasValue == false) return cv;

            switch (dataType.Value)
            {
                case CellValues.Boolean:
                    if (cv == "0") return "False";
                    else return "True";
                case CellValues.SharedString:
                    return strings.ElementAt(int.Parse(cv)).InnerText;
                default:
                    return cv;
            }
        }
    }
}
