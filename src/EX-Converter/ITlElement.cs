using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EX_Converter
{
    internal interface ITlElement
    {
        string AttrName { get; }
        string NodeOrder { get; }

        bool NameEquals(ITlElement other);
    }
}
