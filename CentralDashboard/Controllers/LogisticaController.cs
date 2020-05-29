using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CentralDashboard.Controllers
{
    public class LogisticaController : AppController
    {
        //[HttpPost]
        public ActionResult CemCenabast(int anio, bool insumo = false)
        {
            var bdEnti = bdBuilder.GetEntiCorporativa();
            string idUsuario = GetUsuario();
            if (!bdEnti.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 4))
            {
                throw new UnauthorizedAccessException();
            }
            var bd = bdBuilder.GetAbastecimiento();
            var rows = bd.RPT_CEM_CENABAST(anio, insumo);

            using (ExcelPackage excel = new ExcelPackage())
            {
                var messageBook = excel.Workbook.Worksheets.Add("Hospitalizaciones");
                var cellsTemp = messageBook.Cells[2, 1];
                #region cabecera del Reporte
                
                int i = 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CODIGO_CENABAST");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "GLOSA_MEDICAMENTO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CODIGO_HRA");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado ENE");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST ENE");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST ENE");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS ENE");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS ENE");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado FEB");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST FEB");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST FEB");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS FEB");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS FEB");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado MAR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST MAR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST MAR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS MAR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS MAR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado ABR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST ABR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST ABR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS ABR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS ABR");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado MAY");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST MAY");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST MAY");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS MAY");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS MAY");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado JUN");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST JUN");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST JUN");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS JUN");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS JUN");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado JUL");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST JUL");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST JUL");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS JUL");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS JUL");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado AGO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST AGO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST AGO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS AGO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS AGO");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado SEP");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST SEP");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST SEP");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS SEP");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS SEP");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado OCT");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST OCT");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST OCT");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS OCT");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS OCT");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado NOV");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST NOV");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST NOV");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS NOV");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS NOV");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "Programado DIC");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT CENABAST DIC");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO CENABAST DIC");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "CANT OTROS DIC");

                i += 1;
                cellsTemp = messageBook.Cells[2, i];
                CentralDashboard.Clases.ExcelHelper.CeldaTitulo(ref cellsTemp, "MONTO OTROS DIC");
                #endregion

                i += 1;
                int j = 3;

                #region Construccion rpt
                foreach (var item in rows)
                {
                    i = 0;
                    messageBook.Cells[j, ++i].Value = item.CODIGO_CENABAST;
                    messageBook.Cells[j, ++i].Value = item.GLOSA_MEDICAMENTO;
                    messageBook.Cells[j, ++i].Value = item.CODIGO_HRA;
                    messageBook.Cells[j, ++i].Value = item.Programado_ENE;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_ENE;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_ENE;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_ENE;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_ENE;
                    messageBook.Cells[j, ++i].Value = item.Programado_FEB;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_FEB;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_FEB;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_FEB;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_FEB;
                    messageBook.Cells[j, ++i].Value = item.Programado_MAR;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_MAR;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_MAR;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_MAR;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_MAR;
                    messageBook.Cells[j, ++i].Value = item.Programado_ABR;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_ABR;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_ABR;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_ABR;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_ABR;
                    messageBook.Cells[j, ++i].Value = item.Programado_MAY;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_MAY;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_MAY;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_MAY;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_MAY;
                    messageBook.Cells[j, ++i].Value = item.Programado_JUN;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_JUN;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_JUN;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_JUN;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_JUN;
                    messageBook.Cells[j, ++i].Value = item.Programado_JUL;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_JUL;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_JUL;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_JUL;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_JUL;
                    messageBook.Cells[j, ++i].Value = item.Programado_AGO;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_AGO;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_AGO;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_AGO;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_AGO;
                    messageBook.Cells[j, ++i].Value = item.Programado_SEP;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_SEP;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_SEP;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_SEP;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_SEP;
                    messageBook.Cells[j, ++i].Value = item.Programado_OCT;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_OCT;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_OCT;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_OCT;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_OCT;
                    messageBook.Cells[j, ++i].Value = item.Programado_NOV;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_NOV;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_NOV;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_NOV;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_NOV;
                    messageBook.Cells[j, ++i].Value = item.Programado_DIC;
                    messageBook.Cells[j, ++i].Value = item.CANT_CENABAST_DIC;
                    messageBook.Cells[j, ++i].Value = item.MONTO_CENABAST_DIC;
                    messageBook.Cells[j, ++i].Value = item.CANT_OTROS_DIC;
                    messageBook.Cells[j, ++i].Value = item.MONTO_OTROS_DIC;
                    j++;
                }
                #endregion

                var stream = new MemoryStream(excel.GetAsByteArray());
                return File(
                    stream, 
                    "application/excel", 
                    (
                        "CemCenabast Año " + 
                        anio.ToString() + 
                        " Generado en " + 
                        DateTime.Now.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + 
                        ".xlsx"
                    )
                );
            }
        }
    }
}