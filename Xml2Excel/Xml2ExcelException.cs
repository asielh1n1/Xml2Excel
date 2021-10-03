using System;
using System.Collections.Generic;
using System.Text;

namespace Xml2Excel
{
    class Xml2ExcelException:Exception
    {
        public Xml2ExcelException(string message) : base(message)
        {

        }

        public Xml2ExcelException(string message,Exception e) : base(message,e)
        {

        }

    }
}
