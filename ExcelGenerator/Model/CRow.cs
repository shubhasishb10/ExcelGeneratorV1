using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelGenerator.Model
{
    class CRow
    {
        public List<string> Datarow { set; get; }
        public Properties Props { set; get; }
    }
    class Properties
    {
        public string Color { set; get; }
        public string FSize { set; get; }
    }
    class CTable
    {
        public List<CRow> Rows { set; get; }
        public string Header { set; get; }
    }
    class CSet
    {
       public ICollection<CTable> CTables { set; get; }
    }
}
