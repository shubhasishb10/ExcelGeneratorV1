using ExcelGenerator.CellCalculator;
using ExcelGenerator.Enum;
using ExcelGenerator.Model;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.Utils
{
    class XUtil
    {
        public readonly static string sourceFileLocation = "..\\..\\data\\model.json";
        public readonly static string destinationFileLocation = @"U:\DOT NET Project Work Space\ExcelGenerator\ExcelGenerator\output\today.xlsx";

        public void CreateExcel(CSet dataset)
        {
            Application excelApp = new Application();
            Workbook workbook = workbook = excelApp.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dataset.CTables.Count; i++)
            {
                CelCalc celCalc = new CelCalc(dataset.CTables[i]);
                Worksheet worksheet = null;
                if (i == 0)
                    worksheet = (Worksheet)workbook.ActiveSheet;
                else
                    worksheet = (Worksheet)workbook.Worksheets.Add(After: workbook.ActiveSheet);
                // Get the Starting Position of the record table
                Coordinates co = celCalc.GetTableStartingPosition();
                worksheet.Name = dataset.CTables[i].TabName;
                // Header of the Title of the Sheet
                if (dataset.CTables[i].Header != null && dataset.CTables[i].Header != "")
                {
                    Coordinates hCo = celCalc.GetSheetHeaderStartingPosition();
                    Range headerSpan = worksheet.Range[worksheet.Cells[hCo.Row, hCo.Column], worksheet.Cells[hCo.Row, hCo.Column + CellPosition.HeaderSpan]];
                    headerSpan.Merge();
                    headerSpan.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    headerSpan.Font.Size = CellPosition.SheetHeaderFontSize;
                    headerSpan.Font.Bold = true;
                    headerSpan.Font.Color = SheetHeaderColor.sheetHeaderColor;
                    headerSpan.Value = dataset.CTables[i].Header;
                    headerSpan.Borders.ThemeColor = System.Drawing.Color.Aqua;
                }
                for (int j = 0; j < dataset.CTables[i].CRows.Count; j++)
                {
                    for (int k = 0; k < dataset.CTables[i].CRows[j].Datarow.Count; k++)
                    {
                        worksheet.Cells[(co.Row + j), (k + 1)] = dataset.CTables[i].CRows[j].Datarow[k];
                        Range rRange = worksheet.Range[worksheet.Cells[co.Row + j, k + co.Column], worksheet.Cells[co.Row + j, dataset.CTables[i].CRows[j].Datarow.Count]];
                        rRange.Font.Size = CellPosition.TableRecordFontSize;
                    }
                    // Header Row
                    Range r = worksheet.Range[worksheet.Cells[co.Row, co.Column], worksheet.Cells[co.Row, dataset.CTables[i].CRows[j].Datarow.Count]];
                    r.Font.Size = CellPosition.TableHeaderFontSize;
                    r.Font.Bold = true;
                }
                // Full Table Border & Alignment set
                Range fullRange = worksheet.Range[worksheet.Cells[co.Row, co.Column], worksheet.Cells[dataset.CTables[i].CRows.Count + co.Row - 1, dataset.CTables[i].CRows[1].Datarow.Count]];
                fullRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                fullRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                fullRange.Columns.ColumnWidth = CellPosition.ColumnWidth;
            }
            excelApp.DisplayAlerts = false;
            workbook.SaveAs(destinationFileLocation,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            workbook.Close();
            excelApp.Quit();
        }
    }
}
