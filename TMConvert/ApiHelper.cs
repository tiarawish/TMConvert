using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace TMConvert
{
    public static class ApiHelper
    {
        public static HttpClient ApiClient;

        public static void InitializeClient()
        {
            ApiClient = new HttpClient();
            //ApiClient.BaseAddress = new Uri("https://szgwdsfnutmhvnz.weclapp.com/webapp/api/v1/");
            ApiClient.DefaultRequestHeaders.Accept.Clear();
            ApiClient.DefaultRequestHeaders.Add("AuthenticationToken", "7572d2ea-464e-449f-a8e9-39c68ae12fe5");
            //ApiClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("AuthenticationToken", "7572d2ea-464e-449f-a8e9-39c68ae12fe5");
            ApiClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }
    }
}
