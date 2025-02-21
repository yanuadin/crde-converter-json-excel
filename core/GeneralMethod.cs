using System;
using System.Collections.Generic;
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
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName;

            return projectDirectory;
        }

        public static string getTimeStampNow()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmssffff");
        }
    }
}
