using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser
{
    public interface INameInterceptor : IInterceptor
    {
        string Intercept(string name);
    }
}
