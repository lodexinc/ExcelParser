using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    public interface IExcelParser
    {
        T Parse<T>(string fileName) where T : new();
    }

    public class BaseExcelParser 
}
