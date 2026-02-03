using System;
using System.Net.Http;

namespace ExcelSmartCostAddIn
{
    internal static class EngineClient
    {
        public static string BaseUrl = "http://127.0.0.1:17831";
        public static string Token = "DEV_TOKEN_SET_ME";

        public static HttpClient Create()
        {
            var client = new HttpClient();
            client.BaseAddress = new Uri(BaseUrl);
            client.DefaultRequestHeaders.Add("X-Token", Token);
            return client;
        }
    }
}
