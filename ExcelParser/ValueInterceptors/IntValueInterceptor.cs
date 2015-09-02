using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class IntValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public IntValueInterceptor()
            : base()
        {
        }
        public IntValueInterceptor(int order)
            : base(order)
        {
        }

        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (property.PropertyType == typeof(int))
            {
                if (currentValue != null)
                    return int.Parse(originalValue);
                else return currentValue;
            }
            else if (property.PropertyType == typeof(int?))
            {
                if (currentValue == null) return default(int?);
                else return new int?(int.Parse(originalValue));
            }
            else return currentValue;
        }
    }
}
