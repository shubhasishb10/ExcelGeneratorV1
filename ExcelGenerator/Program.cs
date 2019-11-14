using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ExcelGenerator.Model;
using Newtonsoft.Json;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var json = "";
            CSet cs = JsonConvert.DeserializeObject<CSet>(json);
            new Program().CreateExcel(cs);
        }
        private void CreateExcel(CSet dataset)
        {
            Application excelApp = new Application();
            Workbook workbook = workbook = excelApp.Workbooks.Add(Type.Missing); ;
            for (int i = 0; i < 2; i++)
            {
                Worksheet worksheet = null;
                if (i == 0)
                    worksheet = (Worksheet)workbook.ActiveSheet;
                else
                    worksheet = (Worksheet)workbook.Worksheets.Add(After: workbook.ActiveSheet);
                Range headerCell = worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 8]];
                headerCell.Merge();
                Borders border = headerCell.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerCell.Font.Size = 50;
                headerCell.Font.Color = System.Drawing.Color.Red;
                worksheet.Name = "Tab"+(i+1);
                headerCell.Value = "Student Report Card";
            }
            workbook.SaveAs(@"U:\DOT NET Project Work Space\ExcelGenerator\ExcelGenerator\output\gg.xlsx");
            workbook.Close();
            excelApp.Quit();
        }
    }
}
