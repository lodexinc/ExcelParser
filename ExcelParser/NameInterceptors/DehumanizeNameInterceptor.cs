using System;
using System.Collections.Generic;
using System.Linq;
using Humanizer;

namespace ExcelParser
{
    public class DehumanizeNameInterceptor : BaseInterceptor, INameInterceptor
    {
        public DehumanizeNameInterceptor()
        {
        }
        public DehumanizeNameInterceptor(int order)
            : base(order)
        {
        }

        public string Intercept(string name)
        {
            return name.Dehumanize();
        }
    }
}