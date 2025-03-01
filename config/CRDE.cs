using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CRDEConverterJsonExcel.core;
using System.Diagnostics;

namespace CRDEConverterJsonExcel.config
{
    class CRDE
    {
        private JObject config;
        public CRDE() {
            string jsonContent = File.ReadAllText(GeneralMethod.getProjectDirectory() + @"\config\CRDE.json");
            config = JObject.Parse(jsonContent);
        }

        public string getCurrentEnv()
        {
            return config["CURRENT_ENV"].ToObject<string>();
        }

        public JArray getColorCells()
        {
            return config["COLOR_CELLS"].ToObject<JArray>();
        }

        public JObject getEnvironment(string env = "")
        {
            env = env == "" ? getCurrentEnv() : env;

            JObject envConfig = config["ENVIRONMENT"].Children<JObject>().FirstOrDefault(child =>
            {
                foreach (var ch in child)
                {
                    return ch.Key.ToUpper() == env.ToUpper();
                }
                return false;
            });
            
            return envConfig == null ? null : envConfig[env].ToObject<JObject>();
        }
    }
}
