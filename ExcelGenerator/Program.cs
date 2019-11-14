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
            var data = File.OpenRead("..\\..\\data\\model.json");
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
                Range headerCell = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[2, 3]];
                headerCell.Merge();
                //headerCell.AutoFit();
                //Borders border = headerCell.Borders;
                //border.LineStyle = XlLineStyle.xlContinuous;
                headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                headerCell.Font.Size = 50;
                headerCell.Font.Bold = true;
                //border.Weight = 3d;
                headerCell.Font.Color = System.Drawing.Color.BlueViolet;
                headerCell.Value = dataset.CTables[i].Header;
                for (int j = 0; j < dataset.CTables[i].CRows.Count; j++)
                {
                    for(int k = 0; k < dataset.CTables[i].CRows[j].Datarow.Count; k++)
                    {
                        //Range r = worksheet.Range[worksheet.Cells[row, col]];
                        worksheet.Cells[(4+j), (k+1)] = dataset.CTables[i].CRows[j].Datarow[k];
                        Range rRange = worksheet.Range[worksheet.Cells[4 + j, k + 1], worksheet.Cells[4 + j, dataset.CTables[i].CRows[j].Datarow.Count]];
                        //rRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        rRange.Font.Size = 13;
                    }
                    Range r = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[4, dataset.CTables[i].CRows[j].Datarow.Count]];
                    //Borders b = r.Borders;
                    r.Font.Size = 15;
                    r.Font.Bold = true;
                    //b.LineStyle = XlLineStyle.xlContinuous;
                }
                Range fullRange = worksheet.Range[worksheet.Cells[4, 1], worksheet.Cells[dataset.CTables[i].CRows.Count+4, dataset.CTables[i].CRows[1].Datarow.Count]];
                fullRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                fullRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                fullRange.Columns.ColumnWidth = 30;
            }
            workbook.SaveAs(@"C:\DNprojects\ExcelGeneratorV1\ExcelGenerator\output\new.xlsx");
            workbook.Close();
            excelApp.Quit();
        }
    }
}
