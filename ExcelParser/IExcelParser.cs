using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Humanizer;

namespace ExcelParser
{
    public interface IExcelParser
    {
        IList<T> Parse<T>(System.IO.Stream stream, string sheetName = null) where T : class, new();
        IList<T> Parse<T>(string fileName, string sheetName = null) where T : class, new();
    }
}