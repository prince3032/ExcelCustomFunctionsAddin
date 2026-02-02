using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Product.Business.Dtos
{
    public class CalcResponse
    {
        public double Sum { get; set; }
        public string Message { get; set; } = "";
    }
}
