using ClosedXML.Excel;

using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;

using System;

using System.Collections.Generic;

namespace Wpf_Excel_to_Datagrid
{
    public class Metric
    {
        public int Alpha { get; set; }
        public int Beta { get; set; }
        public int Gamma { get; set; }
        public int Delta { get; set; }
    }
    public class Class1
    {

      
    }
   
}

// пример с сайта https://www.nuget.org/packages/ClosedXML
//    using (var workbook = new XLWorkbook())
//{
//    var worksheet = workbook.Worksheets.Add("Sample Sheet");
//    worksheet.Cell("A1").Value = "Hello World!";
//    worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
//    workbook.SaveAs("HelloWorld.xlsx");
//}
