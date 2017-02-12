using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace npoi.core.Corefx
{
#if netcore
    public interface ICloneable
    {
        object Clone();
    }
#endif
}
