using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class StringToNullValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public string Marker { get; }
        public StringToNullValueInterceptor()
            : this(int.MinValue)
        {
        }

        public StringToNullValueInterceptor(int order)
            : this(order, "NULL")
        {
        }
        public StringToNullValueInterceptor(int order, string marker)
            : base(order)
        {
            Marker = marker;
        }
        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (originalValue == Marker)
                return null;
            else return currentValue;
        }
    }
}
