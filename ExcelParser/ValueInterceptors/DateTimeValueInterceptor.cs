using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class DateTimeValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public DateTimeValueInterceptor()
            : base()
        {
        }
        public DateTimeValueInterceptor(int order)
            : base(order)
        {
        }

        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (property.PropertyType == typeof(DateTime))
            {
                if (currentValue != null)
                    return DateTime.FromOADate(double.Parse(originalValue));
                else return currentValue;
            }
            else if (property.PropertyType == typeof(DateTime?))
            {
                if (currentValue == null) return default(DateTime?);
                else return new DateTime?(DateTime.FromOADate(double.Parse(originalValue)));
            }
            else return currentValue;
        }
    }
}
