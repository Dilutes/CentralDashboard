using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CentralDashboard.Models.EntiCorporativa;
using OfficeOpenXml;

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
            List<RPT_COMGES_UEH_base_Result> comges = bdEnti.RPT_COMGES_UEH_base(4, 2020).ToList();
            DataSet ds = new DataSet();

            DataTable dt = BuildDataTable<RPT_COMGES_UEH_base_Result>(comges);
            ds.Tables.Add(dt);
            return DataSetToExcelXlsx(ds);
        }

        private MemoryStream DataSetToExcelXlsx(DataSet ds)
        {
            MemoryStream result = new MemoryStream();
            ExcelPackage pack = new ExcelPackage();
            ExcelWorksheet ws;

            ws = pack.Workbook.Worksheets.Add("Abril");
            ws.Cells["A1"].LoadFromDataTable(ds.Tables[0], true);

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