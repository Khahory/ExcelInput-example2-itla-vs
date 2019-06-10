using ExcelInput_example2.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInput_example2.Controllers
{
    public class ProductController : Controller
    {
        // GET: Product
        public ActionResult Index()
        {

            return View();
        }


        [HttpPost] //metodo para comprobar si el archivo que se introduce termina en xls o xlsx
        public ActionResult Import(HttpPostedFileBase excelFile)
        {
            //saber si coloco un .xls
            if (excelFile == null || excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Please select a excel file";
                return View("Index");
            }
            else
            {   
                //otro if
                if (excelFile.FileName.EndsWith("xls") || excelFile.FileName.EndsWith("xlsx"))
                {
                    //guardar el archivo en la carpeta "Content" (Solo es esta parte del codigo)
                    string path = Server.MapPath("~/Content/" +excelFile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelFile.SaveAs(path);
                    //Fin del guardado

                    //Read data from excel file
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    //hay que crear un modelo para el producto
                    List<ProductModel> listProduct = new List<ProductModel>();
                    for (int row = 3; row <= range.Rows.Count; row++)
                    {
                        ProductModel productModel = new ProductModel(); //supongo que esto es para leer los datos del File Excel
                        productModel.Id = ((Excel.Range)range.Cells[row, 1]).Text; //asignamos los valores al modelo
                        productModel.Nombre = ((Excel.Range)range.Cells[row, 2]).Text;
                        productModel.Precio = decimal.Parse(((Excel.Range)range.Cells[row, 3]).Text);
                        productModel.Cantidad = int.Parse(((Excel.Range)range.Cells[row, 4]).Text);
                        listProduct.Add(productModel);
                    }

                    ViewBag.ListProduct = listProduct;
                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "Please select a excel file";
                    return View("Index");
                }
            }
            
        }
    }
}