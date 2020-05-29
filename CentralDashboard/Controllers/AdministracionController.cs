using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CentralDashboard.Models.EntiCorporativa;
using OfficeOpenXml.Style;
using System.Drawing;

namespace CentralDashboard.Controllers
{
    public class AdministracionController : AppController
    {
        // GET: Administración
        public ActionResult Index()
        {
            var bd = bdBuilder.GetEntiCorporativa();
            string idUsuario = GetUsuario();
            ViewBag.ReporteDiarioHosp = bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 1);
            ViewBag.ReporteFuentesRemA08 = bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 2);
            ViewBag.ReporteRemA08 = bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 3);
            ViewBag.ReporteCemCenabast = bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 4);
            return View();
        }
        [HttpPost]
        public FileResult Index(int isa = 0)
        {
            var bd = bdBuilder.GetEntiCorporativa();
            string idUsuario = GetUsuario();
            if (!bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 1))
            {
                throw new UnauthorizedAccessException();
            }
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

        public FileResult getFuenteREMA08(int mes, int anio) {
            var bd = bdBuilder.GetEntiCorporativa();
            string idUsuario = GetUsuario();
            if (!bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 2))
            {
                throw new UnauthorizedAccessException();
            }
            using (ExcelPackage excel = new ExcelPackage())
            {
                List<REM_DatosBase_Result> result = bd.REM_DatosBase(mes, anio).ToList();
                string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                var messageBook = excel.Workbook.Worksheets.Add(meses[Math.Abs(1-mes)] + ' ' + anio);

                //Cabezera
                messageBook.Cells["A1"].Value = "N DAU";
                messageBook.Cells["B1"].Value = "Fecha y Hora de Admisón";
                messageBook.Cells["C1"].Value = "Fecha y Hora de Alta Médica";
                messageBook.Cells["D1"].Value = "Demanda";
                messageBook.Cells["E1"].Value = "Atención Completa";
                messageBook.Cells["F1"].Value = "Sexo";
                messageBook.Cells["G1"].Value = "Grupo etario";
                messageBook.Cells["H1"].Value = "Beneficiario";
                messageBook.Cells["I1"].Value = "Derivación";
                messageBook.Cells["J1"].Value = "Tipo de atención";
                messageBook.Cells["K1"].Value = "Categorización";
                messageBook.Cells["L1"].Value = "Indicación de Hospitalización";
                messageBook.Cells["M1"].Value = "Tiempo de espera de cama";
                messageBook.Cells["N1"].Value = "Tiempo de espera Hosp.Domiciliaria";
                messageBook.Cells["O1"].Value = "Destino del Paciente";
                messageBook.Cells["P1"].Value = "Fallecimiento";

                //Data
                for (int i = 0; i < result.Count(); i++)
                {
                    messageBook.Cells["A" + (i + 2).ToString()].Value = result[i].N_DAU;
                    messageBook.Cells["B" + (i + 2).ToString()].Value = result[i].Fecha_y_Hora_de_Admisón;
                    messageBook.Cells["C" + (i + 2).ToString()].Value = result[i].Fecha_y_Hora_de_Alta_Médica;
                    messageBook.Cells["D" + (i + 2).ToString()].Value = result[i].Demanda;
                    messageBook.Cells["E" + (i + 2).ToString()].Value = result[i].Atención_Completa;
                    messageBook.Cells["F" + (i + 2).ToString()].Value = result[i].Sexo;
                    messageBook.Cells["G" + (i + 2).ToString()].Value = result[i].Grupo_etario;
                    messageBook.Cells["H" + (i + 2).ToString()].Value = result[i].Beneficiario;
                    messageBook.Cells["I" + (i + 2).ToString()].Value = result[i].Derivación;
                    messageBook.Cells["J" + (i + 2).ToString()].Value = result[i].Tipo_de_atención;
                    messageBook.Cells["K" + (i + 2).ToString()].Value = result[i].Categorización;
                    messageBook.Cells["L" + (i + 2).ToString()].Value = result[i].Indicación_de_Hospitalización;
                    messageBook.Cells["M" + (i + 2).ToString()].Value = result[i].Tiempo_de_espera_de_cama;
                    messageBook.Cells["N" + (i + 2).ToString()].Value = result[i].Tiempo_de_espera_Hosp__Domiciliaria;
                    messageBook.Cells["O" + (i + 2).ToString()].Value = result[i].Destino_del_Paciente;
                    messageBook.Cells["P" + (i + 2).ToString()].Value = result[i].Fallecimiento;
                }

                //Formato
                messageBook.Cells["A1:P" + (result.Count() + 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:P" + (result.Count() + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:P" + (result.Count() + 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:P" + (result.Count() + 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:P" + (result.Count() + 1)].AutoFilter = true;
                messageBook.Cells["A2:C" + (result.Count() + 1)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                messageBook.Cells["A1:P1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["A1:P1"].Style.Fill.BackgroundColor.SetColor(Color.Black);
                messageBook.Cells["A1:P1"].Style.Font.Color.SetColor(Color.White);
                messageBook.Cells["A1:P1"].Style.Font.Bold = true;

                messageBook.Cells.AutoFitColumns();

                var stream = new MemoryStream(excel.GetAsByteArray());

                return File(stream, "application/excel", "Fuente_REM_A08_" + (meses[Math.Abs(1 - mes)]).ToString().ToUpper() + '_' + (anio).ToString() + ".xlsx");
            }
        }

        public FileResult getREMA08(int mes, int anio)
        {
            var bd = bdBuilder.GetEntiCorporativa();
            string idUsuario = GetUsuario();
            if (!bd.USR_PermisoSitioWeb.Any(x => x.Usuario == idUsuario && x.IdPaginaSitioWeb == 3))
            {
                throw new UnauthorizedAccessException();
            }
            using (ExcelPackage excel = new ExcelPackage())
            {
                List<REM_SeccionA_Result> resultA = bd.REM_SeccionA(mes, anio).ToList();
                List<REM_SeccionB_Result> resultB = bd.REM_SeccionB(mes, anio).ToList();
                List<REM_SeccionD_Result> resultD = bd.REM_SeccionD(mes, anio).ToList();
                List<REM_SeccionF_Result> resultF = bd.REM_SeccionF(mes, anio).ToList();
                string[] meses = { "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre" };
                var messageBook = excel.Workbook.Worksheets.Add("A08");

                //Formato
                #region Formato
                messageBook.Column(1).Width = GetTrueColumnWidth(44);
                messageBook.Column(2).Width = GetTrueColumnWidth(30.43);
                messageBook.Column(3).Width = GetTrueColumnWidth(13.43);
                messageBook.Column(4).Width = GetTrueColumnWidth(11.71);
                for (int i = 5; i <= 45; i++)
                {
                    messageBook.Column(i).Width = GetTrueColumnWidth(10.71);
                }
                messageBook.Column(44).Width = GetTrueColumnWidth(14.43);
                messageBook.Row(7).Height = 31.50;
                messageBook.Row(8).Height = 31.50;
                messageBook.Row(9).Height = 31.50;
                messageBook.Row(11).Height = 27.75;
                messageBook.Row(15).Height = 31.50;
                messageBook.Row(18).Height = 31.50;
                messageBook.Row(26).Height = 31.50;
                messageBook.Row(38).Height = 31.50;
                messageBook.Cells.Style.Font.SetFromFont(new Font("Verdana", 11));
                //Seccion A
                messageBook.Cells["A5"].Style.Font.Bold = true;
                messageBook.Cells["A6:O6"].Merge = true;
                messageBook.Cells["A6:O6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A6:O6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["A6:O6"].Style.Font.Bold = true;
                messageBook.Cells["A6:O6"].Style.Font.Size = 12;
                messageBook.Cells["A7"].Style.Font.Bold = true;
                messageBook.Cells["A8"].Style.Font.Bold = true;
                //Seccion A.1
                messageBook.Cells["A5"].Style.Font.Size = 8;
                messageBook.Cells["A9:AS14"].Style.Font.Size = 8;
                messageBook.Cells["A9:A11"].Merge = true;
                messageBook.Cells["A9:AS11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A9:AS11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["B12:AS14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                messageBook.Cells["B9:D10"].Merge = true;
                messageBook.Cells["E9:AL9"].Merge = true;
                messageBook.Cells["E10:F10"].Merge = true;
                messageBook.Cells["G10:H10"].Merge = true;
                messageBook.Cells["I10:J10"].Merge = true;
                messageBook.Cells["K10:L10"].Merge = true;
                messageBook.Cells["M10:N10"].Merge = true;
                messageBook.Cells["O10:P10"].Merge = true;
                messageBook.Cells["Q10:R10"].Merge = true;
                messageBook.Cells["S10:T10"].Merge = true;
                messageBook.Cells["U10:V10"].Merge = true;
                messageBook.Cells["W10:X10"].Merge = true;
                messageBook.Cells["Y10:Z10"].Merge = true;
                messageBook.Cells["AA10:AB10"].Merge = true;
                messageBook.Cells["AC10:AD10"].Merge = true;
                messageBook.Cells["AE10:AF10"].Merge = true;
                messageBook.Cells["AG10:AH10"].Merge = true;
                messageBook.Cells["AI10:AJ10"].Merge = true;
                messageBook.Cells["AK10:AL10"].Merge = true;
                messageBook.Cells["AM9:AM11"].Merge = true;
                messageBook.Cells["AN9:AQ9"].Merge = true;
                messageBook.Cells["AN10:AN11"].Merge = true;
                messageBook.Cells["AO10:AO11"].Merge = true;
                messageBook.Cells["AP10:AP11"].Merge = true;
                messageBook.Cells["AQ10:AQ11"].Merge = true;
                messageBook.Cells["AR9:AR11"].Merge = true;
                messageBook.Cells["AS9:AS11"].Merge = true;
                messageBook.Cells["E12:AS14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["E12:AS14"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 204));
                messageBook.Cells["AN14:AS14"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["AN14:AS14"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 192, 192));
                messageBook.Cells["AN9:AS11"].Style.WrapText = true;
                //Seccion B
                messageBook.Cells["A15"].Style.Font.Bold = true;
                messageBook.Cells["A16:AN25"].Style.Font.Size = 8;
                messageBook.Cells["A16:A18"].Merge = true;
                messageBook.Cells["A16:AS18"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A16:AS18"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["A19:A25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["B19:AS25"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                messageBook.Cells["B16:D17"].Merge = true;
                messageBook.Cells["E16:AL16"].Merge = true;
                messageBook.Cells["E17:F17"].Merge = true;
                messageBook.Cells["G17:H17"].Merge = true;
                messageBook.Cells["I17:J17"].Merge = true;
                messageBook.Cells["K17:L17"].Merge = true;
                messageBook.Cells["M17:N17"].Merge = true;
                messageBook.Cells["O17:P17"].Merge = true;
                messageBook.Cells["Q17:R17"].Merge = true;
                messageBook.Cells["S17:T17"].Merge = true;
                messageBook.Cells["U17:V17"].Merge = true;
                messageBook.Cells["W17:X17"].Merge = true;
                messageBook.Cells["Y17:Z17"].Merge = true;
                messageBook.Cells["AA17:AB17"].Merge = true;
                messageBook.Cells["AC17:AD17"].Merge = true;
                messageBook.Cells["AE17:AF17"].Merge = true;
                messageBook.Cells["AG17:AH17"].Merge = true;
                messageBook.Cells["AI17:AJ17"].Merge = true;
                messageBook.Cells["AK17:AL17"].Merge = true;
                messageBook.Cells["AM16:AN17"].Merge = true;
                messageBook.Cells["E19:AN24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["E19:AN24"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 204));
                messageBook.Cells["AM24:AN24"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["AM24:AN24"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 192, 192));
                messageBook.Cells["AM16:AN18"].Style.WrapText = true;
                //Seccion D
                messageBook.Cells["A26"].Style.Font.Bold = true;
                messageBook.Cells["A27:AO37"].Style.Font.Size = 8;
                messageBook.Cells["A27:AO29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A27:AO29"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["A30:B33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A30:B33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["C30:AO37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                messageBook.Cells["A27:B29"].Merge = true;
                messageBook.Cells["A30:B30"].Merge = true;
                messageBook.Cells["A31:A33"].Merge = true;
                messageBook.Cells["C27:E28"].Merge = true;
                messageBook.Cells["A34:B34"].Merge = true;
                messageBook.Cells["A35:B35"].Merge = true;
                messageBook.Cells["A36:B36"].Merge = true;
                messageBook.Cells["A37:B37"].Merge = true;
                messageBook.Cells["F27:AM27"].Merge = true;
                messageBook.Cells["F28:G28"].Merge = true;
                messageBook.Cells["H28:I28"].Merge = true;
                messageBook.Cells["J28:K28"].Merge = true;
                messageBook.Cells["L28:M28"].Merge = true;
                messageBook.Cells["N28:O28"].Merge = true;
                messageBook.Cells["P28:Q28"].Merge = true;
                messageBook.Cells["R28:S28"].Merge = true;
                messageBook.Cells["T28:U28"].Merge = true;
                messageBook.Cells["V28:W28"].Merge = true;
                messageBook.Cells["X28:Y28"].Merge = true;
                messageBook.Cells["Z28:AA28"].Merge = true;
                messageBook.Cells["AB28:AC28"].Merge = true;
                messageBook.Cells["AD28:AE28"].Merge = true;
                messageBook.Cells["AF28:AG28"].Merge = true;
                messageBook.Cells["AH28:AI28"].Merge = true;
                messageBook.Cells["AJ28:AK28"].Merge = true;
                messageBook.Cells["AL28:AM28"].Merge = true;
                messageBook.Cells["AN27:AN29"].Merge = true;
                messageBook.Cells["AO27:AO29"].Merge = true;
                messageBook.Cells["F31:AO37"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["F31:AO37"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 204));
                messageBook.Cells["AO34:AO37"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["AO34:AO37"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(192, 192, 192));
                messageBook.Cells["A31:A33"].Style.WrapText = true;
                messageBook.Cells["AN27:AO29"].Style.WrapText = true;
                //Seccion F
                messageBook.Cells["A38"].Style.Font.Bold = true;
                messageBook.Cells["A39:AN44"].Style.Font.Size = 8;
                messageBook.Cells["A39:AN41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A39:AN41"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                messageBook.Cells["C42:AN44"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                messageBook.Cells["A39:B41"].Merge = true;
                messageBook.Cells["A42:B42"].Merge = true;
                messageBook.Cells["A43:B43"].Merge = true;
                messageBook.Cells["A44:B44"].Merge = true;
                messageBook.Cells["C39:E40"].Merge = true;
                messageBook.Cells["F39:AM39"].Merge = true;
                messageBook.Cells["F40:G40"].Merge = true;
                messageBook.Cells["H40:I40"].Merge = true;
                messageBook.Cells["J40:K40"].Merge = true;
                messageBook.Cells["L40:M40"].Merge = true;
                messageBook.Cells["N40:O40"].Merge = true;
                messageBook.Cells["P40:Q40"].Merge = true;
                messageBook.Cells["R40:S40"].Merge = true;
                messageBook.Cells["T40:U40"].Merge = true;
                messageBook.Cells["V40:W40"].Merge = true;
                messageBook.Cells["X40:Y40"].Merge = true;
                messageBook.Cells["Z40:AA40"].Merge = true;
                messageBook.Cells["AB40:AC40"].Merge = true;
                messageBook.Cells["AD40:AE40"].Merge = true;
                messageBook.Cells["AF40:AG40"].Merge = true;
                messageBook.Cells["AH40:AI40"].Merge = true;
                messageBook.Cells["AJ40:AK40"].Merge = true;
                messageBook.Cells["AL40:AM40"].Merge = true;
                messageBook.Cells["AN39:AN41"].Merge = true;
                messageBook.Cells["F42:AN44"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["F42:AN44"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 204));
                messageBook.Cells["AN39:AN41"].Style.WrapText = true;

                //Bordes
                messageBook.View.ShowGridLines = false;
                //Seccion A.1
                messageBook.Cells["A9:AS14"].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A9:AS14"].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A9:AS14"].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A9:AS14"].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A9:AS14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A9:A11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A12:A14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["B9:AL10"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B9:AL10"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B9:AL10"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B9:AL10"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B11:AL11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["B11:D14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["E11:F14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["G11:H14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["I11:J14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["K11:L14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["M11:N14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["O11:P14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Q11:R14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["S11:T14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["U11:V14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["W11:X14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Y11:Z14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AA11:AB14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AC11:AD14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AE11:AF14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AG11:AH14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AI11:AJ14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AK11:AL14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AM9:AM11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AM12:AM14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN9:AQ9"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN10:AQ11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN12:AQ14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AR9:AR11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AR12:AR14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AS9:AS11"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AS12:AS14"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //Seccion B
                messageBook.Cells["A16:AN25"].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A16:AN25"].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A16:AN25"].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A16:AN25"].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A16:AN25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A16:A18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A19:A25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["B16:AN17"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B16:AN17"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B16:AN17"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B16:AN17"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["B18:AN18"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A25:AN25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["B18:D25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["E18:F25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["G18:H25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["I18:J25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["K18:L25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["M18:N25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["O18:P25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Q18:R25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["S18:T25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["U18:V25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["W18:X25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Y18:Z25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AA18:AB25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AC18:AD25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AE18:AF25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AG18:AH25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AI18:AJ25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AK18:AL25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AM18:AN25"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //Seccion D
                messageBook.Cells["A27:AO37"].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A27:AO37"].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A27:AO37"].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A27:AO37"].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A27:AO37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A27:B29"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A30:B30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A31:A33"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["B31:B33"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A34:B37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["C27:AM28"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C27:AM28"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C27:AM28"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C27:AM28"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C29:AM29"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["C30:AO30"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A34:AO34"].Style.Border.Top.Style = ExcelBorderStyle.Double;
                messageBook.Cells["C29:E37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["F29:G37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["H29:I37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["J29:K37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["L29:M37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["N29:O37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["P29:Q37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["R29:S37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["T29:U37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["V29:W37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["X29:Y37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Z29:AA37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AB29:AC37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AD29:AE37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AF29:AG37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AH29:AI37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AJ29:AK37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AL29:AM37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN27:AN29"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AO27:AO29"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN30:AN37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AO30:AO37"].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //Seccion F
                messageBook.Cells["A39:AN44"].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A39:AN44"].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A39:AN44"].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A39:AN44"].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                messageBook.Cells["A39:B41"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["A42:B44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["C39:AM40"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C39:AM40"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C39:AM40"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C39:AM40"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["C41:AM41"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["C41:E44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["F41:G44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["H41:I44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["J41:K44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["L41:M44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["N41:O44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["P41:Q44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["R41:S44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["T41:U44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["V41:W44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["X41:Y44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["Z41:AA44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AB41:AC44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AD41:AE44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AF41:AG44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AH41:AI44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AJ41:AK44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AL41:AM44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN39:AN41"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                messageBook.Cells["AN42:AN44"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                #endregion

                //Cabezera
                #region Cabezera
                //Seccion A
                messageBook.Cells["A5"].Value = "AÑO: " + anio;
                messageBook.Cells["A6:O6"].Value = "REM-A08.  ATENCIÓN DE URGENCIA";
                messageBook.Cells["A7"].Value = "SECCIÓN A: ATENCIONES REALIZADAS EN UNIDADES DE URGENCIA DE LA RED";
                //Seccion A.1
                messageBook.Cells["A8"].Value = "SECCIÓN A.1: ATENCIONES REALIZADAS EN UNIDADES DE EMERGENCIA HOSPITALARIA DE ALTA Y MEDIANA COMPLEJIDAD (UEH)";
                messageBook.Cells["A9:A11"].Value = "TIPO DE ATENCIÓN";
                messageBook.Cells["B9:D10"].Value = "TOTAL";
                messageBook.Cells["B11"].Value = "Ambos Sexos";
                messageBook.Cells["C11"].Value = "Hombres";
                messageBook.Cells["D11"].Value = "Mujeres";
                messageBook.Cells["E9:AL9"].Value = "GRUPOS DE EDAD (en años)";
                messageBook.Cells["E10:F10"].Value = "0 - 4";
                messageBook.Cells["G10:H10"].Value = "5 - 9";
                messageBook.Cells["I10:J10"].Value = "10 - 14";
                messageBook.Cells["K10:L10"].Value = "15 - 19";
                messageBook.Cells["M10:N10"].Value = "20 - 24";
                messageBook.Cells["O10:P10"].Value = "25 - 29";
                messageBook.Cells["Q10:R10"].Value = "30 - 34";
                messageBook.Cells["S10:T10"].Value = "35 - 39";
                messageBook.Cells["U10:V10"].Value = "40 - 44";
                messageBook.Cells["W10:X10"].Value = "45 - 49";
                messageBook.Cells["Y10:Z10"].Value = "50 - 54";
                messageBook.Cells["AA10:AB10"].Value = "55 - 59";
                messageBook.Cells["AC10:AD10"].Value = "60 - 64";
                messageBook.Cells["AE10:AF10"].Value = "65 - 69";
                messageBook.Cells["AG10:AH10"].Value = "70 - 74";
                messageBook.Cells["AI10:AJ10"].Value = "75 - 79";
                messageBook.Cells["AK10:AL10"].Value = "80 y mas";
                messageBook.Cells["E11"].Value = "Hombres";
                messageBook.Cells["F11"].Value = "Mujeres";
                messageBook.Cells["G11"].Value = "Hombres";
                messageBook.Cells["H11"].Value = "Mujeres";
                messageBook.Cells["I11"].Value = "Hombres";
                messageBook.Cells["J11"].Value = "Mujeres";
                messageBook.Cells["K11"].Value = "Hombres";
                messageBook.Cells["L11"].Value = "Mujeres";
                messageBook.Cells["M11"].Value = "Hombres";
                messageBook.Cells["N11"].Value = "Mujeres";
                messageBook.Cells["O11"].Value = "Hombres";
                messageBook.Cells["P11"].Value = "Mujeres";
                messageBook.Cells["Q11"].Value = "Hombres";
                messageBook.Cells["R11"].Value = "Mujeres";
                messageBook.Cells["S11"].Value = "Hombres";
                messageBook.Cells["T11"].Value = "Mujeres";
                messageBook.Cells["U11"].Value = "Hombres";
                messageBook.Cells["V11"].Value = "Mujeres";
                messageBook.Cells["W11"].Value = "Hombres";
                messageBook.Cells["X11"].Value = "Mujeres";
                messageBook.Cells["Y11"].Value = "Hombres";
                messageBook.Cells["Z11"].Value = "Mujeres";
                messageBook.Cells["AA11"].Value = "Hombres";
                messageBook.Cells["AB11"].Value = "Mujeres";
                messageBook.Cells["AC11"].Value = "Hombres";
                messageBook.Cells["AD11"].Value = "Mujeres";
                messageBook.Cells["AE11"].Value = "Hombres";
                messageBook.Cells["AF11"].Value = "Mujeres";
                messageBook.Cells["AG11"].Value = "Hombres";
                messageBook.Cells["AH11"].Value = "Mujeres";
                messageBook.Cells["AI11"].Value = "Hombres";
                messageBook.Cells["AJ11"].Value = "Mujeres";
                messageBook.Cells["AK11"].Value = "Hombres";
                messageBook.Cells["AL11"].Value = "Mujeres";
                messageBook.Cells["AM9:AM11"].Value = "Beneficiarios";
                messageBook.Cells["AN9:AQ9"].Value = "ORIGEN DE LA PROCEDENCIA (Sólo pacientes derivados de establecimientos de la Red)";
                messageBook.Cells["AN10:AN11"].Value = "SAPU/ SAR / SUR";
                messageBook.Cells["AO10:AO11"].Value = "Hospital Baja Complejidad";
                messageBook.Cells["AP10:AP11"].Value = "Hospital Mediana/Alta Complejidad";
                messageBook.Cells["AQ10:AQ11"].Value = "Otros Estableci-mientos de la  Red";
                messageBook.Cells["AR9:AR11"].Value = "Establecimientos de otra Red";
                messageBook.Cells["AS9:AS11"].Value = "Demanda de Urgencia";
                //Seccion B
                messageBook.Cells["A15"].Value = "SECCIÓN B: CATEGORIZACIÓN DE PACIENTES, PREVIA A LA ATENCIÓN MÉDICA (Establecimientos Alta, Mediana o Baja Complejidad y SAR).";
                messageBook.Cells["A16:A18"].Value = "CATEGORÍAS";
                messageBook.Cells["B16:D17"].Value = "TOTAL";
                messageBook.Cells["B18"].Value = "Ambos Sexos";
                messageBook.Cells["C18"].Value = "Hombres";
                messageBook.Cells["D18"].Value = "Mujeres";
                messageBook.Cells["E16:AL16"].Value = "GRUPOS DE EDAD (en años)";
                messageBook.Cells["E17:F17"].Value = "0 - 4";
                messageBook.Cells["G17:H17"].Value = "5 - 9";
                messageBook.Cells["I17:J17"].Value = "10 - 14";
                messageBook.Cells["K17:L17"].Value = "15 - 19";
                messageBook.Cells["M17:N17"].Value = "20 - 24";
                messageBook.Cells["O17:P17"].Value = "25 - 29";
                messageBook.Cells["Q17:R17"].Value = "30 - 34";
                messageBook.Cells["S17:T17"].Value = "35 - 39";
                messageBook.Cells["U17:V17"].Value = "40 - 44";
                messageBook.Cells["W17:X17"].Value = "45 - 49";
                messageBook.Cells["Y17:Z17"].Value = "50 - 54";
                messageBook.Cells["AA17:AB17"].Value = "55 - 59";
                messageBook.Cells["AC17:AD17"].Value = "60 - 64";
                messageBook.Cells["AE17:AF17"].Value = "65 - 69";
                messageBook.Cells["AG17:AH17"].Value = "70 - 74";
                messageBook.Cells["AI17:AJ17"].Value = "75 - 79";
                messageBook.Cells["AK17:AL17"].Value = "80 y mas";
                messageBook.Cells["E18"].Value = "Hombres";
                messageBook.Cells["F18"].Value = "Mujeres";
                messageBook.Cells["G18"].Value = "Hombres";
                messageBook.Cells["H18"].Value = "Mujeres";
                messageBook.Cells["I18"].Value = "Hombres";
                messageBook.Cells["J18"].Value = "Mujeres";
                messageBook.Cells["K18"].Value = "Hombres";
                messageBook.Cells["L18"].Value = "Mujeres";
                messageBook.Cells["M18"].Value = "Hombres";
                messageBook.Cells["N18"].Value = "Mujeres";
                messageBook.Cells["O18"].Value = "Hombres";
                messageBook.Cells["P18"].Value = "Mujeres";
                messageBook.Cells["Q18"].Value = "Hombres";
                messageBook.Cells["R18"].Value = "Mujeres";
                messageBook.Cells["S18"].Value = "Hombres";
                messageBook.Cells["T18"].Value = "Mujeres";
                messageBook.Cells["U18"].Value = "Hombres";
                messageBook.Cells["V18"].Value = "Mujeres";
                messageBook.Cells["W18"].Value = "Hombres";
                messageBook.Cells["X18"].Value = "Mujeres";
                messageBook.Cells["Y18"].Value = "Hombres";
                messageBook.Cells["Z18"].Value = "Mujeres";
                messageBook.Cells["AA18"].Value = "Hombres";
                messageBook.Cells["AB18"].Value = "Mujeres";
                messageBook.Cells["AC18"].Value = "Hombres";
                messageBook.Cells["AD18"].Value = "Mujeres";
                messageBook.Cells["AE18"].Value = "Hombres";
                messageBook.Cells["AF18"].Value = "Mujeres";
                messageBook.Cells["AG18"].Value = "Hombres";
                messageBook.Cells["AH18"].Value = "Mujeres";
                messageBook.Cells["AI18"].Value = "Hombres";
                messageBook.Cells["AJ18"].Value = "Mujeres";
                messageBook.Cells["AK18"].Value = "Hombres";
                messageBook.Cells["AL18"].Value = "Mujeres";
                messageBook.Cells["AM16:AN17"].Value = "Herramientas de Categorización";
                messageBook.Cells["AM18"].Value = "Discrecional";
                messageBook.Cells["AN18"].Value = "Estructurado (ESI)";
                //Seccion D
                messageBook.Cells["A26"].Value = "SECCIÓN D: PACIENTES CON INDICACIÓN DE HOSPITALIZACIÓN EN ESPERA DE CAMAS EN UEH";
                messageBook.Cells["A27:B29"].Value = "TIPO DE PACIENTES";
                messageBook.Cells["A30:B30"].Value = "TOTAL DE PACIENTES CON INDICACIÓN DE HOSPITALIZACIÓN";
                messageBook.Cells["A31:A33"].Value = "PACIENTES QUE INGRESAN A CAMA HOSPITALARIA SEGÚN TIEMPO DE DEMORA AL INGRESO";
                messageBook.Cells["C27:E28"].Value = "TOTAL";
                messageBook.Cells["C29"].Value = "Ambos Sexos";
                messageBook.Cells["D29"].Value = "Hombres";
                messageBook.Cells["E29"].Value = "Mujeres";
                messageBook.Cells["F27:AM27"].Value = "GRUPOS DE EDAD (en años)";
                messageBook.Cells["F28:G28"].Value = "0 - 4";
                messageBook.Cells["H28:I28"].Value = "5 - 9";
                messageBook.Cells["J28:K28"].Value = "10 - 14";
                messageBook.Cells["L28:M28"].Value = "15 - 19";
                messageBook.Cells["N28:O28"].Value = "20 - 24";
                messageBook.Cells["P28:Q28"].Value = "25 - 29";
                messageBook.Cells["R28:S28"].Value = "30 - 34";
                messageBook.Cells["T28:U28"].Value = "35 - 39";
                messageBook.Cells["V28:W28"].Value = "40 - 44";
                messageBook.Cells["X28:Y28"].Value = "45 - 49";
                messageBook.Cells["Z28:AA28"].Value = "50 - 54";
                messageBook.Cells["AB28:AC28"].Value = "55 - 59";
                messageBook.Cells["AD28:AE28"].Value = "60 - 64";
                messageBook.Cells["AF28:AG28"].Value = "65 - 69";
                messageBook.Cells["AH28:AI28"].Value = "70 - 74";
                messageBook.Cells["AJ28:AK28"].Value = "75 - 79";
                messageBook.Cells["AL28:AM28"].Value = "80 y mas";
                messageBook.Cells["F29"].Value = "Hombres";
                messageBook.Cells["G29"].Value = "Mujeres";
                messageBook.Cells["H29"].Value = "Hombres";
                messageBook.Cells["I29"].Value = "Mujeres";
                messageBook.Cells["J29"].Value = "Hombres";
                messageBook.Cells["K29"].Value = "Mujeres";
                messageBook.Cells["L29"].Value = "Hombres";
                messageBook.Cells["M29"].Value = "Mujeres";
                messageBook.Cells["N29"].Value = "Hombres";
                messageBook.Cells["O29"].Value = "Mujeres";
                messageBook.Cells["P29"].Value = "Hombres";
                messageBook.Cells["Q29"].Value = "Mujeres";
                messageBook.Cells["R29"].Value = "Hombres";
                messageBook.Cells["S29"].Value = "Mujeres";
                messageBook.Cells["T29"].Value = "Hombres";
                messageBook.Cells["U29"].Value = "Mujeres";
                messageBook.Cells["V29"].Value = "Hombres";
                messageBook.Cells["W29"].Value = "Mujeres";
                messageBook.Cells["X29"].Value = "Hombres";
                messageBook.Cells["Y29"].Value = "Mujeres";
                messageBook.Cells["Z29"].Value = "Hombres";
                messageBook.Cells["AA29"].Value = "Mujeres";
                messageBook.Cells["AB29"].Value = "Hombres";
                messageBook.Cells["AC29"].Value = "Mujeres";
                messageBook.Cells["AD29"].Value = "Hombres";
                messageBook.Cells["AE29"].Value = "Mujeres";
                messageBook.Cells["AF29"].Value = "Hombres";
                messageBook.Cells["AG29"].Value = "Mujeres";
                messageBook.Cells["AH29"].Value = "Hombres";
                messageBook.Cells["AI29"].Value = "Mujeres";
                messageBook.Cells["AJ29"].Value = "Hombres";
                messageBook.Cells["AK29"].Value = "Mujeres";
                messageBook.Cells["AL29"].Value = "Hombres";
                messageBook.Cells["AM29"].Value = "Mujeres";
                messageBook.Cells["AN27:AN29"].Value = "Beneficiarios";
                messageBook.Cells["AO27:AO29"].Value = "Hospitalización Domiciliaria";
                //Seccion F
                messageBook.Cells["A38"].Value = "SECCIÓN F: PACIENTES FALLECIDOS EN UEH (Establecimientos Alta, Mediana o Baja Complejidad y SAR) ";
                messageBook.Cells["A39:B41"].Value = "TIPO DE PACIENTES";
                messageBook.Cells["C39:E40"].Value = "TOTAL";
                messageBook.Cells["C41"].Value = "Ambos Sexos";
                messageBook.Cells["D41"].Value = "Hombres";
                messageBook.Cells["E41"].Value = "Mujeres";
                messageBook.Cells["F39:AM39"].Value = "GRUPOS DE EDAD (en años)";
                messageBook.Cells["F40:G40"].Value = "0 - 4";
                messageBook.Cells["H40:I40"].Value = "5 - 9";
                messageBook.Cells["J40:K40"].Value = "10 - 14";
                messageBook.Cells["L40:M40"].Value = "15 - 19";
                messageBook.Cells["N40:O40"].Value = "20 - 24";
                messageBook.Cells["P40:Q40"].Value = "25 - 29";
                messageBook.Cells["R40:S40"].Value = "30 - 34";
                messageBook.Cells["T40:U40"].Value = "35 - 39";
                messageBook.Cells["V40:W40"].Value = "40 - 44";
                messageBook.Cells["X40:Y40"].Value = "45 - 49";
                messageBook.Cells["Z40:AA40"].Value = "50 - 54";
                messageBook.Cells["AB40:AC40"].Value = "55 - 59";
                messageBook.Cells["AD40:AE40"].Value = "60 - 64";
                messageBook.Cells["AF40:AG40"].Value = "65 - 69";
                messageBook.Cells["AH40:AI40"].Value = "70 - 74";
                messageBook.Cells["AJ40:AK40"].Value = "75 - 79";
                messageBook.Cells["AL40:AM40"].Value = "80 y mas";
                messageBook.Cells["AN39:AN41"].Value = "Beneficiarios";
                messageBook.Cells["F41"].Value = "Hombres";
                messageBook.Cells["G41"].Value = "Mujeres";
                messageBook.Cells["H41"].Value = "Hombres";
                messageBook.Cells["I41"].Value = "Mujeres";
                messageBook.Cells["J41"].Value = "Hombres";
                messageBook.Cells["K41"].Value = "Mujeres";
                messageBook.Cells["L41"].Value = "Hombres";
                messageBook.Cells["M41"].Value = "Mujeres";
                messageBook.Cells["N41"].Value = "Hombres";
                messageBook.Cells["O41"].Value = "Mujeres";
                messageBook.Cells["P41"].Value = "Hombres";
                messageBook.Cells["Q41"].Value = "Mujeres";
                messageBook.Cells["R41"].Value = "Hombres";
                messageBook.Cells["S41"].Value = "Mujeres";
                messageBook.Cells["T41"].Value = "Hombres";
                messageBook.Cells["U41"].Value = "Mujeres";
                messageBook.Cells["V41"].Value = "Hombres";
                messageBook.Cells["W41"].Value = "Mujeres";
                messageBook.Cells["X41"].Value = "Hombres";
                messageBook.Cells["Y41"].Value = "Mujeres";
                messageBook.Cells["Z41"].Value = "Hombres";
                messageBook.Cells["AA41"].Value = "Mujeres";
                messageBook.Cells["AB41"].Value = "Hombres";
                messageBook.Cells["AC41"].Value = "Mujeres";
                messageBook.Cells["AD41"].Value = "Hombres";
                messageBook.Cells["AE41"].Value = "Mujeres";
                messageBook.Cells["AF41"].Value = "Hombres";
                messageBook.Cells["AG41"].Value = "Mujeres";
                messageBook.Cells["AH41"].Value = "Hombres";
                messageBook.Cells["AI41"].Value = "Mujeres";
                messageBook.Cells["AJ41"].Value = "Hombres";
                messageBook.Cells["AK41"].Value = "Mujeres";
                messageBook.Cells["AL41"].Value = "Hombres";
                messageBook.Cells["AM41"].Value = "Mujeres";
                messageBook.Cells["AN27:AN29"].Value = "Beneficiarios";
                #endregion

                //Data
                #region Data
                //Seccion A
                for (int i = 0; i < resultA.Count(); i++)
                {
                    messageBook.Cells["A" + (12 + i).ToString()].Value = resultA[i].TIPO_DE_ATENCIÓN;
                    messageBook.Cells["B" + (12 + i).ToString()].Value = resultA[i].TOTAL_Ambos_Sexos;
                    messageBook.Cells["C" + (12 + i).ToString()].Value = resultA[i].TOTAL_Hombres;
                    messageBook.Cells["D" + (12 + i).ToString()].Value = resultA[i].TOTAL_Mujeres;
                    messageBook.Cells["E" + (12 + i).ToString()].Value = resultA[i].C0___4_Hombres;
                    messageBook.Cells["F" + (12 + i).ToString()].Value = resultA[i].C0___4_Mujeres;
                    messageBook.Cells["G" + (12 + i).ToString()].Value = resultA[i].C10___14_Hombres;
                    messageBook.Cells["H" + (12 + i).ToString()].Value = resultA[i].C10___14_Mujeres;
                    messageBook.Cells["I" + (12 + i).ToString()].Value = resultA[i].C15___19_Hombres;
                    messageBook.Cells["J" + (12 + i).ToString()].Value = resultA[i].C15___19_Mujeres;
                    messageBook.Cells["K" + (12 + i).ToString()].Value = resultA[i].C20___24_Hombres;
                    messageBook.Cells["L" + (12 + i).ToString()].Value = resultA[i].C20___24_Mujeres;
                    messageBook.Cells["M" + (12 + i).ToString()].Value = resultA[i].C25___29_Hombres;
                    messageBook.Cells["N" + (12 + i).ToString()].Value = resultA[i].C25___29_Mujeres;
                    messageBook.Cells["O" + (12 + i).ToString()].Value = resultA[i].C30___34_Hombres;
                    messageBook.Cells["P" + (12 + i).ToString()].Value = resultA[i].C30___34_Mujeres;
                    messageBook.Cells["Q" + (12 + i).ToString()].Value = resultA[i].C35___39_Hombres;
                    messageBook.Cells["R" + (12 + i).ToString()].Value = resultA[i].C35___39_Mujeres;
                    messageBook.Cells["S" + (12 + i).ToString()].Value = resultA[i].C40___44_Hombres;
                    messageBook.Cells["T" + (12 + i).ToString()].Value = resultA[i].C40___44_Mujeres;
                    messageBook.Cells["U" + (12 + i).ToString()].Value = resultA[i].C45___49_Hombres;
                    messageBook.Cells["V" + (12 + i).ToString()].Value = resultA[i].C45___49_Mujeres;
                    messageBook.Cells["W" + (12 + i).ToString()].Value = resultA[i].C50___54_Hombres;
                    messageBook.Cells["X" + (12 + i).ToString()].Value = resultA[i].C50___54_Mujeres;
                    messageBook.Cells["Y" + (12 + i).ToString()].Value = resultA[i].C55___59_Hombres;
                    messageBook.Cells["Z" + (12 + i).ToString()].Value = resultA[i].C55___59_Mujeres;
                    messageBook.Cells["AA" + (12 + i).ToString()].Value = resultA[i].C5___9_Hombres;
                    messageBook.Cells["AB" + (12 + i).ToString()].Value = resultA[i].C5___9_Mujeres;
                    messageBook.Cells["AC" + (12 + i).ToString()].Value = resultA[i].C60___64_Hombres;
                    messageBook.Cells["AD" + (12 + i).ToString()].Value = resultA[i].C60___64_Mujeres;
                    messageBook.Cells["AE" + (12 + i).ToString()].Value = resultA[i].C65___69_Hombres;
                    messageBook.Cells["AF" + (12 + i).ToString()].Value = resultA[i].C65___69_Mujeres;
                    messageBook.Cells["AG" + (12 + i).ToString()].Value = resultA[i].C70___74_Hombres;
                    messageBook.Cells["AH" + (12 + i).ToString()].Value = resultA[i].C70___74_Mujeres;
                    messageBook.Cells["AI" + (12 + i).ToString()].Value = resultA[i].C75___79_Hombres;
                    messageBook.Cells["AJ" + (12 + i).ToString()].Value = resultA[i].C75___79_Mujeres;
                    messageBook.Cells["AK" + (12 + i).ToString()].Value = resultA[i].C80_y_mas_Hombres;
                    messageBook.Cells["AL" + (12 + i).ToString()].Value = resultA[i].C80_y_mas_Mujeres;
                    messageBook.Cells["AM" + (12 + i).ToString()].Value = resultA[i].Beneficiarios;
                    messageBook.Cells["AN" + (12 + i).ToString()].Value = resultA[i].SAPU__SAR___SUR;
                    messageBook.Cells["AO" + (12 + i).ToString()].Value = resultA[i].Hospital_Baja_Complejidad;
                    messageBook.Cells["AP" + (12 + i).ToString()].Value = resultA[i].Hospital_Mediana_Alta_Complejidad;
                    messageBook.Cells["AQ" + (12 + i).ToString()].Value = resultA[i].Otros_Establecimientos_de_la__Red;
                    messageBook.Cells["AR" + (12 + i).ToString()].Value = resultA[i].Establecimientos_de_otra_Red;
                    messageBook.Cells["AS" + (12 + i).ToString()].Value = resultA[i].Demanda_de_Urgencia;
                }
                messageBook.Cells["AN14:AS14"].Value = "";

                //Seccion B
                for (int i = 0; i < resultB.Count(); i++)
                {
                    messageBook.Cells["A" + (19 + i).ToString()].Value = resultB[i].CATEGORÍAS;
                    messageBook.Cells["B" + (19 + i).ToString()].Value = resultB[i].TOTAL_Ambos_Sexos;
                    messageBook.Cells["C" + (19 + i).ToString()].Value = resultB[i].TOTAL_Hombres;
                    messageBook.Cells["D" + (19 + i).ToString()].Value = resultB[i].TOTAL_Mujeres;
                    messageBook.Cells["E" + (19 + i).ToString()].Value = resultB[i].C0___4_Hombres;
                    messageBook.Cells["F" + (19 + i).ToString()].Value = resultB[i].C0___4_Mujeres;
                    messageBook.Cells["G" + (19 + i).ToString()].Value = resultB[i].C10___14_Hombres;
                    messageBook.Cells["H" + (19 + i).ToString()].Value = resultB[i].C10___14_Mujeres;
                    messageBook.Cells["I" + (19 + i).ToString()].Value = resultB[i].C15___19_Hombres;
                    messageBook.Cells["J" + (19 + i).ToString()].Value = resultB[i].C15___19_Mujeres;
                    messageBook.Cells["K" + (19 + i).ToString()].Value = resultB[i].C20___24_Hombres;
                    messageBook.Cells["L" + (19 + i).ToString()].Value = resultB[i].C20___24_Mujeres;
                    messageBook.Cells["M" + (19 + i).ToString()].Value = resultB[i].C25___29_Hombres;
                    messageBook.Cells["N" + (19 + i).ToString()].Value = resultB[i].C25___29_Mujeres;
                    messageBook.Cells["O" + (19 + i).ToString()].Value = resultB[i].C30___34_Hombres;
                    messageBook.Cells["P" + (19 + i).ToString()].Value = resultB[i].C30___34_Mujeres;
                    messageBook.Cells["Q" + (19 + i).ToString()].Value = resultB[i].C35___39_Hombres;
                    messageBook.Cells["R" + (19 + i).ToString()].Value = resultB[i].C35___39_Mujeres;
                    messageBook.Cells["S" + (19 + i).ToString()].Value = resultB[i].C40___44_Hombres;
                    messageBook.Cells["T" + (19 + i).ToString()].Value = resultB[i].C40___44_Mujeres;
                    messageBook.Cells["U" + (19 + i).ToString()].Value = resultB[i].C45___49_Hombres;
                    messageBook.Cells["V" + (19 + i).ToString()].Value = resultB[i].C45___49_Mujeres;
                    messageBook.Cells["W" + (19 + i).ToString()].Value = resultB[i].C50___54_Hombres;
                    messageBook.Cells["X" + (19 + i).ToString()].Value = resultB[i].C50___54_Mujeres;
                    messageBook.Cells["Y" + (19 + i).ToString()].Value = resultB[i].C55___59_Hombres;
                    messageBook.Cells["Z" + (19 + i).ToString()].Value = resultB[i].C55___59_Mujeres;
                    messageBook.Cells["AA" + (19 + i).ToString()].Value = resultB[i].C5___9_Hombres;
                    messageBook.Cells["AB" + (19 + i).ToString()].Value = resultB[i].C5___9_Mujeres;
                    messageBook.Cells["AC" + (19 + i).ToString()].Value = resultB[i].C60___64_Hombres;
                    messageBook.Cells["AD" + (19 + i).ToString()].Value = resultB[i].C60___64_Mujeres;
                    messageBook.Cells["AE" + (19 + i).ToString()].Value = resultB[i].C65___69_Hombres;
                    messageBook.Cells["AF" + (19 + i).ToString()].Value = resultB[i].C65___69_Mujeres;
                    messageBook.Cells["AG" + (19 + i).ToString()].Value = resultB[i].C70___74_Hombres;
                    messageBook.Cells["AH" + (19 + i).ToString()].Value = resultB[i].C70___74_Mujeres;
                    messageBook.Cells["AI" + (19 + i).ToString()].Value = resultB[i].C75___79_Hombres;
                    messageBook.Cells["AJ" + (19 + i).ToString()].Value = resultB[i].C75___79_Mujeres;
                    messageBook.Cells["AK" + (19 + i).ToString()].Value = resultB[i].C80_y_mas_Hombres;
                    messageBook.Cells["AL" + (19 + i).ToString()].Value = resultB[i].C80_y_mas_Mujeres;
                    messageBook.Cells["AM" + (19 + i).ToString()].Value = resultB[i].Discrecional;
                    messageBook.Cells["AN" + (19 + i).ToString()].Value = resultB[i].Estructurado__ESI_;
                }

                messageBook.Cells["A25"].Value = "TOTAL";

                for (int x = 2; x <= 40; x++)
                {
                    messageBook.Cells[25, x].Formula = "=SUM(" + messageBook.Cells[19, x].Address + ":" + messageBook.Cells[24, x].Address + ")";
                }

                messageBook.Cells["AM24:AN24"].Value = "";
                messageBook.Cells["AM25"].Value = "";

                //Seccion D
                for (int i = 0; i < resultD.Count(); i++)
                {
                    if (i < 3)
                    {
                        messageBook.Cells["B" + (31 + i).ToString()].Value = resultD[i].TIPO_PACIENTE;
                    }
                    else
                    {
                        messageBook.Cells["A" + (31 + i).ToString()+":B" + (31 + i).ToString()].Value = resultD[i].TIPO_PACIENTE;
                    }
                    messageBook.Cells["C" + (31 + i).ToString()].Value = resultD[i].TOTAL_Ambos_Sexos;
                    messageBook.Cells["D" + (31 + i).ToString()].Value = resultD[i].TOTAL_Hombres;
                    messageBook.Cells["E" + (31 + i).ToString()].Value = resultD[i].TOTAL_Mujeres;
                    messageBook.Cells["F" + (31 + i).ToString()].Value = resultD[i].C0___4_Hombres;
                    messageBook.Cells["G" + (31 + i).ToString()].Value = resultD[i].C0___4_Mujeres;
                    messageBook.Cells["H" + (31 + i).ToString()].Value = resultD[i].C10___14_Hombres;
                    messageBook.Cells["I" + (31 + i).ToString()].Value = resultD[i].C10___14_Mujeres;
                    messageBook.Cells["J" + (31 + i).ToString()].Value = resultD[i].C15___19_Hombres;
                    messageBook.Cells["K" + (31 + i).ToString()].Value = resultD[i].C15___19_Mujeres;
                    messageBook.Cells["L" + (31 + i).ToString()].Value = resultD[i].C20___24_Hombres;
                    messageBook.Cells["M" + (31 + i).ToString()].Value = resultD[i].C20___24_Mujeres;
                    messageBook.Cells["N" + (31 + i).ToString()].Value = resultD[i].C25___29_Hombres;
                    messageBook.Cells["O" + (31 + i).ToString()].Value = resultD[i].C25___29_Mujeres;
                    messageBook.Cells["P" + (31 + i).ToString()].Value = resultD[i].C30___34_Hombres;
                    messageBook.Cells["Q" + (31 + i).ToString()].Value = resultD[i].C30___34_Mujeres;
                    messageBook.Cells["R" + (31 + i).ToString()].Value = resultD[i].C35___39_Hombres;
                    messageBook.Cells["S" + (31 + i).ToString()].Value = resultD[i].C35___39_Mujeres;
                    messageBook.Cells["T" + (31 + i).ToString()].Value = resultD[i].C40___44_Hombres;
                    messageBook.Cells["U" + (31 + i).ToString()].Value = resultD[i].C40___44_Mujeres;
                    messageBook.Cells["V" + (31 + i).ToString()].Value = resultD[i].C45___49_Hombres;
                    messageBook.Cells["W" + (31 + i).ToString()].Value = resultD[i].C45___49_Mujeres;
                    messageBook.Cells["X" + (31 + i).ToString()].Value = resultD[i].C50___54_Hombres;
                    messageBook.Cells["Y" + (31 + i).ToString()].Value = resultD[i].C50___54_Mujeres;
                    messageBook.Cells["Z" + (31 + i).ToString()].Value = resultD[i].C55___59_Hombres;
                    messageBook.Cells["AA" + (31 + i).ToString()].Value = resultD[i].C55___59_Mujeres;
                    messageBook.Cells["AB" + (31 + i).ToString()].Value = resultD[i].C5___9_Hombres;
                    messageBook.Cells["AC" + (31 + i).ToString()].Value = resultD[i].C5___9_Mujeres;
                    messageBook.Cells["AD" + (31 + i).ToString()].Value = resultD[i].C60___64_Hombres;
                    messageBook.Cells["AE" + (31 + i).ToString()].Value = resultD[i].C60___64_Mujeres;
                    messageBook.Cells["AF" + (31 + i).ToString()].Value = resultD[i].C65___69_Hombres;
                    messageBook.Cells["AG" + (31 + i).ToString()].Value = resultD[i].C65___69_Mujeres;
                    messageBook.Cells["AH" + (31 + i).ToString()].Value = resultD[i].C70___74_Hombres;
                    messageBook.Cells["AI" + (31 + i).ToString()].Value = resultD[i].C70___74_Mujeres;
                    messageBook.Cells["AJ" + (31 + i).ToString()].Value = resultD[i].C75___79_Hombres;
                    messageBook.Cells["AK" + (31 + i).ToString()].Value = resultD[i].C75___79_Mujeres;
                    messageBook.Cells["AL" + (31 + i).ToString()].Value = resultD[i].C80_y_mas_Hombres;
                    messageBook.Cells["AM" + (31 + i).ToString()].Value = resultD[i].C80_y_mas_Mujeres;
                    messageBook.Cells["AN" + (31 + i).ToString()].Value = resultD[i].Beneficiarios;
                    messageBook.Cells["AO" + (31 + i).ToString()].Value = resultD[i].Hospitalización_Domiciliaria;
                }

                for (int x = 3; x <= 41; x++)
                {
                    messageBook.Cells[30, x].Formula = "=SUM(" + messageBook.Cells[31, x].Address + ":" + messageBook.Cells[37, x].Address + ")";
                }

                messageBook.Cells["AO34:AO37"].Value = "";

                //Seccion F
                for (int i = 0; i < resultF.Count(); i++)
                {
                    messageBook.Cells["A" + (42 + i).ToString() + ":B" + (42 + i).ToString()].Value = resultF[i].TIPO_DE_PACIENTES;
                    messageBook.Cells["C" + (42 + i).ToString()].Value = resultF[i].TOTAL_Ambos_Sexos;
                    messageBook.Cells["D" + (42 + i).ToString()].Value = resultF[i].TOTAL_Hombres;
                    messageBook.Cells["E" + (42 + i).ToString()].Value = resultF[i].TOTAL_Mujeres;
                    messageBook.Cells["F" + (42 + i).ToString()].Value = resultF[i].C0___4_Hombres;
                    messageBook.Cells["G" + (42 + i).ToString()].Value = resultF[i].C0___4_Mujeres;
                    messageBook.Cells["H" + (42 + i).ToString()].Value = resultF[i].C10___14_Hombres;
                    messageBook.Cells["I" + (42 + i).ToString()].Value = resultF[i].C10___14_Mujeres;
                    messageBook.Cells["J" + (42 + i).ToString()].Value = resultF[i].C15___19_Hombres;
                    messageBook.Cells["K" + (42 + i).ToString()].Value = resultF[i].C15___19_Mujeres;
                    messageBook.Cells["L" + (42 + i).ToString()].Value = resultF[i].C20___24_Hombres;
                    messageBook.Cells["M" + (42 + i).ToString()].Value = resultF[i].C20___24_Mujeres;
                    messageBook.Cells["N" + (42 + i).ToString()].Value = resultF[i].C25___29_Hombres;
                    messageBook.Cells["O" + (42 + i).ToString()].Value = resultF[i].C25___29_Mujeres;
                    messageBook.Cells["P" + (42 + i).ToString()].Value = resultF[i].C30___34_Hombres;
                    messageBook.Cells["Q" + (42 + i).ToString()].Value = resultF[i].C30___34_Mujeres;
                    messageBook.Cells["R" + (42 + i).ToString()].Value = resultF[i].C35___39_Hombres;
                    messageBook.Cells["S" + (42 + i).ToString()].Value = resultF[i].C35___39_Mujeres;
                    messageBook.Cells["T" + (42 + i).ToString()].Value = resultF[i].C40___44_Hombres;
                    messageBook.Cells["U" + (42 + i).ToString()].Value = resultF[i].C40___44_Mujeres;
                    messageBook.Cells["V" + (42 + i).ToString()].Value = resultF[i].C45___49_Hombres;
                    messageBook.Cells["W" + (42 + i).ToString()].Value = resultF[i].C45___49_Mujeres;
                    messageBook.Cells["X" + (42 + i).ToString()].Value = resultF[i].C50___54_Hombres;
                    messageBook.Cells["Y" + (42 + i).ToString()].Value = resultF[i].C50___54_Mujeres;
                    messageBook.Cells["Z" + (42 + i).ToString()].Value = resultF[i].C55___59_Hombres;
                    messageBook.Cells["AA" + (42 + i).ToString()].Value = resultF[i].C55___59_Mujeres;
                    messageBook.Cells["AB" + (42 + i).ToString()].Value = resultF[i].C5___9_Hombres;
                    messageBook.Cells["AC" + (42 + i).ToString()].Value = resultF[i].C5___9_Mujeres;
                    messageBook.Cells["AD" + (42 + i).ToString()].Value = resultF[i].C60___64_Hombres;
                    messageBook.Cells["AE" + (42 + i).ToString()].Value = resultF[i].C60___64_Mujeres;
                    messageBook.Cells["AF" + (42 + i).ToString()].Value = resultF[i].C65___69_Hombres;
                    messageBook.Cells["AG" + (42 + i).ToString()].Value = resultF[i].C65___69_Mujeres;
                    messageBook.Cells["AH" + (42 + i).ToString()].Value = resultF[i].C70___74_Hombres;
                    messageBook.Cells["AI" + (42 + i).ToString()].Value = resultF[i].C70___74_Mujeres;
                    messageBook.Cells["AJ" + (42 + i).ToString()].Value = resultF[i].C75___79_Hombres;
                    messageBook.Cells["AK" + (42 + i).ToString()].Value = resultF[i].C75___79_Mujeres;
                    messageBook.Cells["AL" + (42 + i).ToString()].Value = resultF[i].C80_y_mas_Hombres;
                    messageBook.Cells["AM" + (42 + i).ToString()].Value = resultF[i].C80_y_mas_Mujeres;
                    messageBook.Cells["AN" + (42 + i).ToString()].Value = resultF[i].Beneficiarios;
                }
                #endregion

                var stream = new MemoryStream(excel.GetAsByteArray());

                return File(stream, "application/excel", "REM_A08_" + (meses[Math.Abs(1 - mes)]).ToString().ToUpper() + '_' + (anio).ToString() + ".xlsx");
            }    
        }

        //Selects Formulario
        public JsonResult getAnioREM()
        {
            var bd = bdBuilder.GetEntiCorporativa();
            var result = bd.REM_GetAños().ToList();

            return Json(result);
        }

        public JsonResult getMesREM(int anio)
        {
            var bd = bdBuilder.GetEntiCorporativa();
            var result = bd.REM_GetMeses(anio).ToList();
            var list = new List<Tuple<string, int>>();
            string[] meses = {"Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"};
            int aux = 0;
            for (int i = 0; i < result.Count(); i++)
            {
                aux = result[i] ?? default(int);
                list.Add(new Tuple<string, int>(meses[Math.Abs(1 - aux)], aux));
            }
            return Json(list);
        }

        //Resize EPPlus
        //Al cambiar el ancho no toma el valor real correspondiente, por lo cual se deja este procedimiento.
        public static double GetTrueColumnWidth(double width)
        {
            double z = 1d;
            if (width >= (1 + 2 / 3))
            {
                z = Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2);
            }
            else
            {
                z = Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);
            }
            double errorAmt = width - z;
            double adj = 0d;
            if (width >= (1 + 2 / 3))
            {
                adj = (Math.Round(7 * errorAmt - 7 / 256, 0)) / 7;
            }
            else
            {
                adj = ((Math.Round(12 * errorAmt - 12 / 256, 0)) / 12) + (2 / 12);
            }
            if (z > 0)
            {
                return width + adj;
            }
            return 0d;
        }
    }
}