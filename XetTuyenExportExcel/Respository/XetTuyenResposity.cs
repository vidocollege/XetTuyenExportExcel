using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using XetTuyenExportExcel.Models;

namespace XetTuyenExportExcel.Respository
{
    internal class XetTuyenResposity
    {
        ConnectionStringSettingsCollection settings = ConfigurationManager.ConnectionStrings;
        public async Task<string> GetXetTuyen()
        {
            HttpClient client = new HttpClient();
            HttpResponseMessage httpResponse = client.GetAsync(settings[1].ConnectionString).GetAwaiter().GetResult();
            httpResponse.EnsureSuccessStatusCode();
            string responseString = await httpResponse.Content.ReadAsStringAsync();
            return responseString;
        }
    }
}
