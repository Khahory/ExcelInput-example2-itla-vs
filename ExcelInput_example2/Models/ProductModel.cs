using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using ExcelInput_example2.Models; //agregamos este using

namespace ExcelInput_example2.Models
{
    public class ProductModel
    {
        //variables de la tabla que tenemos en excel
        public string Id { get; set; }
        public string Nombre { get; set; }
        public decimal Precio { get; set; }
        public int Cantidad { get; set; }
    }
}