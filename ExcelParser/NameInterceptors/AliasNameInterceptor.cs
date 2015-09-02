using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelParser
{
    public class AliasNameInterceptor : BaseInterceptor, INameInterceptor
    {
        private string targetName;
        private string targetAlias;
        public AliasNameInterceptor(string name, string alias)
            : this(name, alias, int.MaxValue)
        {
        }
        public AliasNameInterceptor(string name, string alias, int order)
            : base(order)
        {
            this.targetName = name;
            this.targetAlias = alias;
        }

        public string Intercept(string name)
        {
            return name == targetName ? targetAlias : name;
        }
    }
}
