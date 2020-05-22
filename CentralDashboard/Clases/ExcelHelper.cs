using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace CentralDashboard.Clases
{
    public class ExcelHelper
    {
        public static void CeldaTitulo(ref OfficeOpenXml.ExcelRange excelRange, string valor)
        {
            excelRange.Worksheet.Column(excelRange.Start.Column).Width = 30;
            excelRange.Worksheet.Row(excelRange.Start.Row).Height = 30;
            excelRange.Style.WrapText = true;
            excelRange.Value = valor;
            excelRange.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            excelRange.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            excelRange.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            excelRange.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Double;
            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#FFFFFF");
            excelRange.Style.Fill.BackgroundColor.SetColor(colFromHex);
        }
    }
}