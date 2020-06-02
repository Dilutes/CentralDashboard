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
            MemoryStream stream = Comges();
            string name = "COMGES_Abril_2020.xlsx";
            Response.Clear();
            Response.AddHeader("content-disposition", "attachment;filename=" + name);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.BinaryWrite(stream.ToArray());
            Response.End();
            return View();
        }

        public MemoryStream Comges()
        {
            var bdEnti = bdBuilder.GetEntiCorporativa();
            RPT_COMGES_UEH_Result resumen = bdEnti.RPT_COMGES_UEH(4, 2020).FirstOrDefault();
            List<RPT_COMGES_UEH_base_Result> comges = bdEnti.RPT_COMGES_UEH_base(4, 2020).ToList();
            DataTable dt = CabeceraComges(BuildDataTable<RPT_COMGES_UEH_base_Result>(comges));
            dt.TableName = "Abril";
            return DataSetToExcelXlsx(resumen, dt, "Abril");
        }

        private DataTable CabeceraComges(DataTable dataTable)
        {
            foreach (DataColumn row in dataTable.Columns)
            {
                row.ColumnName = row.ColumnName.Replace('_', ' ');
            }

            return dataTable;
        }

        private MemoryStream DataSetToExcelXlsx(RPT_COMGES_UEH_Result resumen, DataTable dt, string mes)
        {
            System.Drawing.Color azulOscuro = System.Drawing.Color.FromArgb(34, 43, 53);

            MemoryStream result = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws;

            ws = pack.Workbook.Worksheets.Add("COMGES 11.2");
            using (ExcelRange rango = ws.Cells["B3:E3"])
            {
                rango.Merge = true;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Value = "11.2 Porcentaje de usuarios categorizados C2 atendidos oportunamente en las Unidades de Emergencia Hospitalaria";
                rango.Style.WrapText = true;
                rango.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                rango.Style.Fill.BackgroundColor.SetColor(azulOscuro);
                rango.Style.Font.Color.SetColor(Color.White);
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
                rango.Style.Fill.BackgroundColor.SetColor(azulOscuro);
            }

            using (ExcelRange rango = ws.Cells["D4:E4"])
            {
                rango.Style.WrapText = true;
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                rango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));
                rango.Value = mes;
            }

            ws.Cells["D5"].Value = "N";
            ws.Cells["D5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));
            ws.Cells["E5"].Value = "%";
            ws.Cells["E5"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(132, 151, 176));


            ws = pack.Workbook.Worksheets.Add(dt.TableName);
            ws.Cells["A1"].LoadFromDataTable(dt, true);

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
    }
}