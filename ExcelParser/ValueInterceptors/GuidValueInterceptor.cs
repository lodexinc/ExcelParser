using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelParser
{
    public class GuidValueInterceptor : BaseInterceptor, IValueInterceptor
    {
        public GuidValueInterceptor()
            : base()
        {
        }
        public GuidValueInterceptor(int order)
            : base(order)
        {
        }

        public object Intercept(PropertyInfo property, string originalValue, object currentValue)
        {
            if (property.PropertyType == typeof(Guid))
            {
                if (currentValue != null)
                    return Guid.Parse(originalValue);
                else return currentValue;
            }
            else if (property.PropertyType == typeof(Guid?))
            {
                if (currentValue == null) return default(Guid?);
                else return new Guid?(Guid.Parse(originalValue));
            }
            else return currentValue;
        }
    }
}
