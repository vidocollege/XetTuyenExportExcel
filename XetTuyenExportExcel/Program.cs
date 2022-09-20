using Newtonsoft.Json;
using System;
using XetTuyenExportExcel.Respository;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XetTuyenExportExcel.Models;

namespace XetTuyenExportExcel
{
    internal class Program
    {
        
        
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            DataTable hd = new DataTable();
            XetTuyenResposity XetTuyenResposity = new XetTuyenResposity();
            var dataBody = XetTuyenResposity.GetXetTuyen();
            var data = JsonConvert.DeserializeObject<List<Tuyensinh>>(dataBody.Result);
            Console.WriteLine(data);
            
        }

    }
}
