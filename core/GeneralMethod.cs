using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRDEConverterJsonExcel.core
{
    class GeneralMethod
    {
        public static string getProjectDirectory()
        {
            return Environment.CurrentDirectory;
        }

        public static string getTimeStampNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }
    }
}
