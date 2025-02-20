using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace CRDEConverterJsonExcel.core
{
    class Api
    {
        public static async Task<string> GetApiDataAsync(string url)
        {
            using (HttpClient client = new HttpClient())
            {
                // Send a GET request to the API
                HttpResponseMessage response = await client.GetAsync(url);

                // Ensure the request was successful
                response.EnsureSuccessStatusCode();

                // Read and return the response content as a string
                return await response.Content.ReadAsStringAsync();
            }
        }

        public static async Task<string> PostApiDataAsync(string url, object data)
        {
            using (HttpClient client = new HttpClient())
            {
                // Serialize the data to JSON
                string jsonData = JsonConvert.SerializeObject(data);
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");

                // Send a POST request to the API
                HttpResponseMessage response = await client.PostAsync(url, content);

                // Ensure the request was successful
                response.EnsureSuccessStatusCode();

                // Read and return the response content as a string
                return await response.Content.ReadAsStringAsync();
            }
        }
    }
}
