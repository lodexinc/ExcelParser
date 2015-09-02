using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser
{
    public abstract class BaseInterceptor : IInterceptor
    {
        public BaseInterceptor()
        {
            order = 0;
        }

        public BaseInterceptor(int order)
        {
            this.order = order;
        }

        private int order;
        public int Order
        {
            get
            {
                return order;
            }
        }
    }
}
