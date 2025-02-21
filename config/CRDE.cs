using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRDEConverterJsonExcel.config
{
    class CRDE
    {
        public static string ENDPOINT_REQUEST = "https://crde-etl-uat.mylab.local/api/v1/s1/online";

        public static string[] getAllEndpoint()
        {
            return [
                ENDPOINT_REQUEST
            ];
        }
    }
}
