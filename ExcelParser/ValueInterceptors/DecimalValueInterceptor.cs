using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class DecimalValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public DecimalValueInterceptor()
            : base()
        {
        }
        public DecimalValueInterceptor(int order)
            : base(order)
        {
        }

        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (property.PropertyType == typeof(decimal))
            {
                if (currentValue != null)
                    return decimal.Parse(originalValue, System.Globalization.NumberStyles.Float);
                else return currentValue;
            }
            else if (property.PropertyType == typeof(decimal?))
            {
                if (currentValue == null) return default(decimal?);
                else return new decimal?(decimal.Parse(originalValue, System.Globalization.NumberStyles.Float));
            }
            else return currentValue;
        }
    }
}
