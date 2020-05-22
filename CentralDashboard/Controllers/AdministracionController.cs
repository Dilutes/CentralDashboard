using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CentralDashboard.Controllers
{
    public class AdministracionController : AppController
    {
        // GET: Administración
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public FileResult Index(int isa = 0)
        {
            var bd = bdBuilder.GetEntiCorporativa();
            using (ExcelPackage excel = new ExcelPackage())
            {
                var rpt = bd.RPT_DiarioHospitalizacion().ToList();
                var messageBook = excel.Workbook.Worksheets.Add("Hospitalizaciones");

                int i = 1;
                var cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Nombre");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Edad");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Rut");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "PAC PAC Sexo");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Via Admision");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Diagnostico Principal");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "EDIFICIO");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "PISO");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "NOMBRE SALA");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Cama");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "ServicioIngreso");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Fecha Hospitalizacion");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "ESTADO");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Destino Alta");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Profesional Egreso");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "ServicioAlta");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "IQ");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "diaHospi");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MesHospi");
                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "AñoHospi");
                i += 1;

                int j = 3;
                foreach (var item in rpt)
                {
                    i = 0;
                    messageBook.Cells[j, ++i].Value = item.Nombre;
                    messageBook.Cells[j, ++i].Value = item.Edad;
                    messageBook.Cells[j, ++i].Value = item.Rut;
                    messageBook.Cells[j, ++i].Value = item.PAC_PAC_Sexo;
                    messageBook.Cells[j, ++i].Value = item.Via_Admision;
                    messageBook.Cells[j, ++i].Value = item.Diagnostico_Principal;
                    messageBook.Cells[j, ++i].Value = item.EDIFICIO;
                    messageBook.Cells[j, ++i].Value = item.PISO;
                    messageBook.Cells[j, ++i].Value = item.NOMBRE_SALA;
                    messageBook.Cells[j, ++i].Value = item.Cama;
                    messageBook.Cells[j, ++i].Value = item.ServicioIngreso;
                    messageBook.Cells[j, ++i].Value = item.Fecha_Hospitalizacion.ToString("dd/MM/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                    messageBook.Cells[j, ++i].Value = item.ESTADO;
                    messageBook.Cells[j, ++i].Value = item.Destino_Alta;
                    messageBook.Cells[j, ++i].Value = item.Profesional_Egreso;
                    messageBook.Cells[j, ++i].Value = item.ServicioAlta;
                    messageBook.Cells[j, ++i].Value = item.IQ;
                    messageBook.Cells[j, ++i].Value = item.diaHospi;
                    messageBook.Cells[j, ++i].Value = item.MesHospi;
                    messageBook.Cells[j, ++i].Value = item.AñoHospi;
                    j++;
                }

                var stream = new MemoryStream(excel.GetAsByteArray());

                return File(stream, "application/excel", "Hospitalización diaria generado " + DateTime.Now.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + ".xlsx");
            }
        }
    }
}