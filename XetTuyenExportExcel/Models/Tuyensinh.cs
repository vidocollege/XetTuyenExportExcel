using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XetTuyenExportExcel.Models
{
    internal class Tuyensinh
    {
        public int Id { get; set; }
        public string Hoten { get; set; }
        public DateTime Ngaysinh { get; set; }
        public string CMND { get; set; }    
        public string DTB12 { get; set; }
        public string DTN_THPT { get; set; }
        public string Truong { get; set; }
        public string SDT { get; set; }
        public string Diachi { get; set; }
        public string BacHocName { get; set; }
        public string NganhName { get; set; }
    }
}
