using System;
using System.Collections.Generic;
using System.Text;

namespace NPOI.Util
{
    internal class AssertFailedException:Exception //: ApplicationException
    {
        public AssertFailedException(string message)
            : base(message)
        {

        }
    }
}
