using Product.Business.Dtos;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Product.Business.Services
{
    public class CalculatorService
    {
        public CalcResponse Add(CalcRequest request)
        {
            return new CalcResponse
            {
                Sum = request.A + request.B,
                Message = "OK"
            };
        }
    }
}
