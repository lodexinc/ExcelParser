using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class DoubleValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public DoubleValueInterceptor()
            : base()
        {
        }
        public DoubleValueInterceptor(int order)
            : base(order)
        {
        }

        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (property.PropertyType == typeof(double))
            {
                if (currentValue != null)
                    return double.Parse(originalValue);
                else return currentValue;
            }
            else if (property.PropertyType == typeof(double?))
            {
                if (currentValue == null) return default(double?);
                else return new double?(double.Parse(originalValue));
            }
            else return currentValue;
        }
    }
}
