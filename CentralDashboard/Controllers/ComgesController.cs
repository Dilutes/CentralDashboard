using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CentralDashboard.Models.EntiCorporativa;
using OfficeOpenXml;
using HtmlAgilityPack;
using OfficeOpenXml.Style;

namespace CentralDashboard.Controllers
{
    public class ComgesController : AppController
    {
        // GET: Comges
        public ActionResult Index()
        {
            return View();
        }

        public void ObtenerReporteComges(int idMes, int anio)
        {
            string mes = ObtenerMes(idMes);
            MemoryStream stream = Comges(idMes, mes, anio);
            string name = "COMGES_Abril_"+ anio + ".xlsx";
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=" + name);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.BinaryWrite(stream.ToArray());
            Response.End();
        }

        private string ObtenerMes(int idMes)
        {
            switch (idMes)
            {
                case 1:
                    return "Enero";
                case 2:
                    return "Febrero";
                case 3:
                    return "Marzo";
                case 4:
                    return "Abril";
                case 5:
                    return "Mayo";
                case 6:
                    return "Junio";
                case 7:
                    return "Julio";
                case 8:
                    return "Agosto";
                case 9:
                    return "Septiembre";
                case 10:
                    return "Octubre";
                case 11:
                    return "Noviembre";
                case 12:
                    return "Diciembre";
                default:
                    return "";
            }
        }

        public void ObtenerReporteSISQ(int idMes, int anio)
        {
            string mes = ObtenerMes(idMes);
            MemoryStream stream = Sisq_ueh(idMes, mes, anio);
            string name = "SISQ_UEH_"+ anio + ".xlsx";
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=" + name);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.BinaryWrite(stream.ToArray());
            Response.End();
        }

        private MemoryStream Sisq_ueh(int idMes, string mes, int año)
        {
            var bdEnti = bdBuilder.GetEntiCorporativa();
            RPT_COMGES_UEH_Result resumen = bdEnti.RPT_COMGES_UEH(idMes, año).FirstOrDefault();
            List<RPT_COMGES_UEH_base_Result> comges = bdEnti.RPT_COMGES_UEH_base(idMes, año).ToList();
            DataTable dt = CabeceraComges(BuildDataTable<RPT_COMGES_UEH_base_Result>(comges));
            dt.TableName = mes;
            return Sisq_uehDataSetToExcelXlsx(resumen, dt, mes);
        }

        private MemoryStream Comges(int idMes, string mes, int año)
        {
            var bdEnti = bdBuilder.GetEntiCorporativa();
            RPT_COMGES_UEH_Result resumen = bdEnti.RPT_COMGES_UEH(idMes, año).FirstOrDefault();
            List<RPT_COMGES_UEH_base_Result> comges = bdEnti.RPT_COMGES_UEH_base(idMes, año).ToList();
            DataTable dt = CabeceraComges(BuildDataTable<RPT_COMGES_UEH_base_Result>(comges));
            dt.TableName = mes;
            return ComgesDataSetToExcelXlsx(resumen, dt, mes);
        }

        private DataTable CabeceraComges(DataTable dataTable)
        {
            foreach (DataColumn row in dataTable.Columns)
            {
                row.ColumnName = row.ColumnName.Replace('_', ' ');
            }

            return dataTable;
        }

        private MemoryStream ComgesDataSetToExcelXlsx(RPT_COMGES_UEH_Result resumen, DataTable dt, string mes)
        {
            System.Drawing.Color azulOscuro = System.Drawing.Color.FromArgb(34, 43, 53);
            System.Drawing.Color azulMedio = System.Drawing.Color.FromArgb(68, 84, 106);
            System.Drawing.Color azulClaro = System.Drawing.Color.FromArgb(217, 225, 242);

            MemoryStream result = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws;

            ws = pack.Workbook.Worksheets.Add("COMGES 11.2");
            using (ExcelRange rango = ws.Cells["B3:E3"])
            {
                rango.Merge = true;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Value = "11.2 Porcentaje de usuarios categorizados C2 atendidos oportunamente en las Unidades de Emergencia Hospitalaria";
                rango.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rango.Style.Fill.BackgroundColor.SetColor(azulOscuro);
                rango.Style.Font.Color.SetColor(Color.White);
                rango.Style.WrapText = true;
                rango.Style.Font.Bold = true;
            }

            double correccion = 0.71;
            ws.Column(1).Width = 10.71 + correccion;
            ws.Column(2).Width = 15 + correccion;
            ws.Column(3).Width = 53.71 + correccion;
            ws.Column(4).Width = 10.71 + correccion;
            ws.Column(5).Width = 10.71 + correccion;

            ws.Row(3).Height = 30.5;
            ws.Row(4).Height = 17;
            ws.Row(5).Height = 17;
            ws.Row(6).Height = 17;
            ws.Row(7).Height = 30;
            ws.Row(8).Height = 17;
            ws.Row(9).Height = 17;
            ws.Row(10).Height = 17;

            using (ExcelRange rango = ws.Cells["B4:C5"])
            {
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(azulOscuro);
            }

            using (ExcelRange rango = ws.Cells["D4:E4"])
            {
                rango.Merge = true;
                rango.Style.WrapText = true;
                rango.Style.Font.Bold = true;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));
                rango.Value = mes;
                rango.Style.Font.Color.SetColor(Color.White);
                rango.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rango.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rango.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                rango.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            }

            ws.Cells["D5"].Value = "N";
            ws.Cells["D5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["D5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["D5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));
            ws.Cells["D5"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["D5"].Style.WrapText = true;
            ws.Cells["D5"].Style.Font.Bold = true;
            ws.Cells["D5"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D5"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D5"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D5"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["E5"].Value = "%";
            ws.Cells["E5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["E5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));
            ws.Cells["E5"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["E5"].Style.WrapText = true;
            ws.Cells["E5"].Style.Font.Bold = true;
            ws.Cells["E5"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E5"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E5"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E5"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["B6:B7"].Merge = true;
            ws.Cells["B6:B7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B6:B7"].Value = "Adulto + Pediatría + G - O";
            ws.Cells["B6:B7"].Style.WrapText = true;
            ws.Cells["B6:B7"].Style.Font.Bold = true;
            ws.Cells["B6:B7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["B6:B7"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["B6:B7"].Style.Font.Color.SetColor(Color.White);

            ws.Cells["C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C6"].Value = "N° total de usuarios C2 atendidos en UEH";
            ws.Cells["C6"].Style.WrapText = true;
            ws.Cells["C6"].Style.Font.Bold = true;
            ws.Cells["C6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["C6"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["C6"].Style.Font.Color.SetColor(Color.White);

            ws.Cells["C7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C7"].Value = "N° total de usuarios C2 con primera atención médica en 30 minutos o menos";
            ws.Cells["C7"].Style.WrapText = true;
            ws.Cells["C7"].Style.Font.Bold = true;
            ws.Cells["C7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["C7"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["C7"].Style.Font.Color.SetColor(Color.White);

            ws.Cells["D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["D6"].Value = resumen.total_C2;
            ws.Cells["D6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["D6"].Style.Fill.BackgroundColor.SetColor(azulClaro);
            ws.Cells["D6"].Style.Font.Color.SetColor(Color.Black);
            ws.Cells["D6"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D6"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D6"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D6"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["D7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["D7"].Value = resumen.menor_a_30;
            ws.Cells["D7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["D7"].Style.Fill.BackgroundColor.SetColor(azulClaro);
            ws.Cells["D7"].Style.Font.Color.SetColor(Color.Black);
            ws.Cells["D7"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D7"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D7"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["D7"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["E6:E7"].Merge = true;
            ws.Cells["E6:E7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E6:E7"].Value = resumen.Porcentaje;
            ws.Cells["E6:E7"].Style.Numberformat.Format = "#0\\.00%";
            ws.Cells["E6:E7"].Style.Font.Bold = true;
            ws.Cells["E6:E7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["E6:E7"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            ws.Cells["E6:E7"].Style.Fill.BackgroundColor.SetColor(azulClaro);
            ws.Cells["E6:E7"].Style.Font.Color.SetColor(Color.Black);
            ws.Cells["E6:E7"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E6:E7"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E6:E7"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["E6:E7"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["B9"].Value = "* Se consideran los pacientes según la fecha y hora del alta médica, ya que debe ser coincidente con el reporte REM A08 (B58)";
            ws.Cells["B10"].Value = "**Se considera pacientes adulto + pediátrico + G - O para que sea consistente con REM A08(B58)";



            ws = pack.Workbook.Worksheets.Add(dt.TableName);
            ws.Cells["A1"].LoadFromDataTable(dt, true);

            ws.Cells[ws.Dimension.Address].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].AutoFitColumns();

            ws.Row(1).Height = 46.25;


            ws.Cells["A1:J1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:J1"].Style.Font.Bold = true;
            ws.Cells["A1:J1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(azulOscuro);
            ws.Cells["A1:J1"].Style.Font.Color.SetColor(Color.White);

            pack.SaveAs(result);
            return result;
        }

        public static DataTable BuildDataTable<T>(IList<T> lst)
        {
            DataTable tbl = CreateTable<T>();
            Type entType = typeof(T);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entType);
            foreach (T item in lst)
            {
                DataRow row = tbl.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item);
                }
                tbl.Rows.Add(row);
            }
            return tbl;
        }

        private static DataTable CreateTable<T>()
        {
            Type entType = typeof(T);
            DataTable tbl = new DataTable(entType.Name);
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(entType);
            foreach (PropertyDescriptor prop in properties)
            {
                Type propType = prop.PropertyType;
                if (propType.Name.Contains("Nullable"))
                {
                    propType = "".GetType();
                }
                tbl.Columns.Add(prop.Name, propType);//prop.PropertyType
            }
            return tbl;
        }
        

        private MemoryStream Sisq_uehDataSetToExcelXlsx(RPT_COMGES_UEH_Result resumen, DataTable dt, string mes)
        {
            System.Drawing.Color azulOscuro = System.Drawing.Color.FromArgb(34, 43, 53);
            System.Drawing.Color azulMedio = System.Drawing.Color.FromArgb(68, 84, 106);
            System.Drawing.Color azulClaro = System.Drawing.Color.FromArgb(217, 225, 242);

            MemoryStream result = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws;

            ws = pack.Workbook.Worksheets.Add("a) B.4_1.2");
            using (ExcelRange rango = ws.Cells["B4:C4"])
            {
                rango.Merge = true;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Value = "Porcentaje de Pacientes Atendidos dentro del estándar en Unidades de Emergencia Hospitalaria (B.4_1.2)";
                rango.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rango.Style.Fill.BackgroundColor.SetColor(azulOscuro);
                rango.Style.Font.Color.SetColor(Color.White);
                rango.Style.WrapText = true;
                rango.Style.Font.Bold = true;
            }

            ws.Cells["B5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["B5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B5"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["B5"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["B5"].Style.WrapText = true;
            ws.Cells["B5"].Style.Font.Bold = true;

            ws.Cells["C5"].Value = "Adulto + Pediátrico (no incluye gineco-obstetra y dental)";
            ws.Cells["C5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["C5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C5"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["C5"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C5"].Style.WrapText = true;
            ws.Cells["C5"].Style.Font.Bold = true;

            ws.Cells["B6"].Value = "Pacientes con estadía <= 6 horas en UEH";
            ws.Cells["B6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B6"].Style.Fill.BackgroundColor.SetColor(azulOscuro);
            ws.Cells["B6"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["B6"].Style.WrapText = true;
            ws.Cells["B6"].Style.Font.Bold = true;

            ws.Cells["B7"].Value = "Total de atenciones";
            ws.Cells["B7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["B7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B7"].Style.Fill.BackgroundColor.SetColor(azulOscuro);
            ws.Cells["B7"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["B7"].Style.WrapText = true;
            ws.Cells["B7"].Style.Font.Bold = true;

            ws.Cells["B8"].Value = "% de cumplimiento";
            ws.Cells["B8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["B8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["B8"].Style.Fill.BackgroundColor.SetColor(azulOscuro);
            ws.Cells["B8"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["B8"].Style.WrapText = true;
            ws.Cells["B8"].Style.Font.Bold = true;


            ws.Cells["C5"].Value = "Adulto + Pediátrico (no incluye gineco-obstetra y dental)";
            ws.Cells["C5"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["C5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C5"].Style.Fill.BackgroundColor.SetColor(azulMedio);
            ws.Cells["C5"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C5"].Style.WrapText = true;
            ws.Cells["C5"].Style.Font.Bold = true;

            ws.Cells["C6"].Value = resumen.total_C2;
            ws.Cells["C6"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C6"].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells["C6"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C6"].Style.WrapText = true;
            ws.Cells["C6"].Style.Font.Bold = true;
            ws.Cells["C6"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C6"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C6"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C6"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["C7"].Value = resumen.total_C2;
            ws.Cells["C7"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["C7"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C7"].Style.Fill.BackgroundColor.SetColor(Color.White);
            ws.Cells["C7"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C7"].Style.WrapText = true;
            ws.Cells["C7"].Style.Font.Bold = true;
            ws.Cells["C7"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C7"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C7"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C7"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            ws.Cells["C8"].Value = resumen.Porcentaje;
            ws.Cells["C8"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["C8"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C8"].Style.Fill.BackgroundColor.SetColor(azulClaro);
            ws.Cells["C8"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C8"].Style.WrapText = true;
            ws.Cells["C8"].Style.Font.Bold = true;
            ws.Cells["C8"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C8"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C8"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells["C8"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

            
            ws = pack.Workbook.Worksheets.Add(dt.TableName);
            ws.Cells["A1"].LoadFromDataTable(dt, true);

            ws.Cells[ws.Dimension.Address].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            ws.Cells[ws.Dimension.Address].AutoFilter = true;

            ws.Row(1).Height = 46.25;


            ws.Cells["A1:G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1:G1"].Style.Font.Bold = true;
            ws.Cells["A1:G1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            ws.Cells["A1:G1"].Style.Fill.BackgroundColor.SetColor(azulOscuro);
            ws.Cells["A1:G1"].Style.Font.Color.SetColor(Color.White);

            pack.SaveAs(result);
            return result;
        }
    }
}