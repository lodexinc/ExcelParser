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
        IList<T> Parse<T>(string fileName) where T : new();
    }
}