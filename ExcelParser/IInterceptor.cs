using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser
{
    public interface IInterceptor
    {
        int Order { get; }
    }
}
