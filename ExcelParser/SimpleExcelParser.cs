using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser
{
    public class SimpleExcelParser : BaseExcelParser
    {
        public static readonly INameInterceptor[] DefaultNameInterceptors = new INameInterceptor[]
        {
            new DehumanizeNameInterceptor()
        };

        public static readonly IValueInterceptor[] DefaultValueInterceptors = new IValueInterceptor[]
        {
            new IntValueInterceptor(),
            new DecimalValueInterceptor(),
            new DoubleValueInterceptor(),
            new DateTimeValueInterceptor(),
            new GuidValueInterceptor(),
            new StringToNullValueInterceptor()
        };

        public SimpleExcelParser() :
            base(DefaultNameInterceptors, DefaultValueInterceptors)
        {
        }

        public SimpleExcelParser(
            IEnumerable<INameInterceptor> nameInterceptors,
            IEnumerable<IValueInterceptor> valueInterceptors) : this()
        {
            NameInterceptors.AddRange(nameInterceptors);
            ValueInterceptors.AddRange(valueInterceptors);
        }
    }
}
