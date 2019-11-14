using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using ExcelGenerator.Model;
using Newtonsoft.Json;
using System.IO;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = File.OpenRead(@"U:\DOT NET Project Work Space\ExcelGenerator\ExcelGenerator\Data\model.json");
            StreamReader reader = new StreamReader(data);
            string json = reader.ReadToEnd();
            CSet cs = JsonConvert.DeserializeObject<CSet>(json);
            new Program().CreateExcel(cs);
        }
        private void CreateExcel(CSet dataset)
        {
            Application excelApp = new Application();
            Workbook workbook = workbook = excelApp.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dataset.CTables.Count; i++)
            {
                Worksheet worksheet = null;
                if (i == 0)
                    worksheet = (Worksheet)workbook.ActiveSheet;
                else
                    worksheet = (Worksheet)workbook.Worksheets.Add(After: workbook.ActiveSheet);
                worksheet.Name = dataset.CTables[i].TabName;
                Range headerCell = worksheet.Range[worksheet.Cells[2, 6], worksheet.Cells[2, 16]];
                headerCell.Merge();
                //headerCell.AutoFit();
                Borders border = headerCell.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerCell.Font.Size = 50;
                headerCell.Font.Bold = true;
                headerCell.Font.Color = System.Drawing.Color.Red;
                headerCell.Value = dataset.CTables[i].Header;
                int col = 1;
                int row = 4;
                for (int j = 0; j < dataset.CTables[i].CRows.Count; j++)
                {
                    for(int k = 0; k < dataset.CTables[i].CRows[j].Datarow.Count; k++)
                    {
                        //Range r = worksheet.Range[worksheet.Cells[row, col]];
                        worksheet.Cells[4, (k+1)] = dataset.CTables[i].CRows[j].Datarow[k];
                    }
                    Range r = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4+dataset.CTables[i].CRows.Count, dataset.CTables[i].CRows[j].Datarow.Count]];
                    Borders b = r.Borders;
                    r.Font.Size = 30;
                    b.LineStyle = XlLineStyle.xlContinuous;
                }
            }
            workbook.SaveAs(@"U:\DOT NET Project Work Space\ExcelGenerator\ExcelGenerator\output\gg.xlsx");
            workbook.Close();
            excelApp.Quit();
        }
    }
}
