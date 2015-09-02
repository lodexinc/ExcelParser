using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Humanizer;
using System.IO;

namespace ExcelParser
{
    public abstract class BaseExcelParser : IExcelParser
    {
        public List<INameInterceptor> NameInterceptors { get; } = new List<INameInterceptor>();
        public List<IValueInterceptor> ValueInterceptors { get; } = new List<IValueInterceptor>();
        protected IOrderedEnumerable<INameInterceptor> OrderedNameInterceptors =>
            NameInterceptors.OrderBy(e => e.Order);
        protected IOrderedEnumerable<IValueInterceptor> OrderedValueInterceptors =>
            ValueInterceptors.OrderBy(e => e.Order);

        public BaseExcelParser(IEnumerable<INameInterceptor> nameInterceptors, IEnumerable<IValueInterceptor> valueInterceptors)
        {
            this.NameInterceptors.AddRange(nameInterceptors);
            this.ValueInterceptors.AddRange(valueInterceptors);
        }

        public IList<T> Parse<T>(string fileName) where T : new()
        {
            using (var doc = SpreadsheetDocument.Open(fileName, false))
            {
                return ProcessDocument<T>(doc);
            }
        }

        public IList<T> Parse<T>(Stream stream) where T : new()
        {
            using (var doc = SpreadsheetDocument.Open(stream, false))
            {
                return ProcessDocument<T>(doc);
            }
        }

        private IList<T> ProcessDocument<T>(SpreadsheetDocument doc)
            where T : new()
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
        private T ParseObject<T>(Row row, SharedStringTable strings, Dictionary<string, string> columns)
                    where T : new()
        {
            var obj = new T();

            foreach (var cell in row.Elements<Cell>())
            {
                var column = columns[GetColumnFromCell(cell)];
                var originalValue = GetValueAsString(cell, strings);
                var property = typeof(T).GetProperty(column);

                object currentValue = originalValue;

                foreach (var interceptor in OrderedValueInterceptors)
                {
                    currentValue = interceptor.Intercept(property, originalValue, currentValue);
                }

                property.SetValue(obj, currentValue);
            }

            return obj;
        }

        private void ParseColumnNames(Row row, SharedStringTable strings, Dictionary<string, string> columns)
        {
            foreach (var cell in row.Elements<Cell>())
            {
                string columnIdentifier = GetColumnFromCell(cell);
                string columnName = GetValueAsString(cell, strings);

                foreach (var interceptor in OrderedNameInterceptors)
                {
                    columnName = interceptor.Intercept(columnName);
                }

                columns.Add(
                    columnIdentifier,
                    columnName
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
