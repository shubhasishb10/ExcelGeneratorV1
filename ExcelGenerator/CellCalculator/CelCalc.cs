using ExcelGenerator.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.CellCalculator
{
    class CelCalc
    {
        private CTable cTable;
        public CelCalc(CTable cTable)
        {
            this.cTable = cTable;
        }
        public Coordinates GetTableStartingPosition()
        {
            if (this.cTable.Header != null && this.cTable.Header != "")
                return new Coordinates() { Row = 4, Column = 1 };
            else
                return new Coordinates() { Row = 2, Column = 1 };
        }
        public Coordinates GetSheetHeaderStartingPosition()
        {
            return new Coordinates() { Row = 2, Column = 1};
        }
    }
}
