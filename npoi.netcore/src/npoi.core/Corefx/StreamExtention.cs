using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace npoi.core.Corefx
{
    public static class StreamExtention
    {
        public static void Close(this Stream source)
        {
            source.Flush();
        }
    }
}
