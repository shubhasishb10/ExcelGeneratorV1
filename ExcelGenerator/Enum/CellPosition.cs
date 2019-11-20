using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.Enum
{
    internal enum CellPosition
    {
        ColumnWidth = 30,
        TableRecordFontSize = 13,
        TableHeaderFontSize = 15,
        SheetHeaderFontSize = 40,
        HeaderSpan = 2
    }
    internal class SheetHeaderColor
    {
        public readonly static System.Drawing.Color sheetHeaderColor = System.Drawing.Color.DarkGoldenrod;
    }
}
