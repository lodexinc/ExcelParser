using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
    public interface IValueInterceptor : IInterceptor
    {
        object Intercept(PropertyInfo property, string originalValue, object currentValue);
    }
}