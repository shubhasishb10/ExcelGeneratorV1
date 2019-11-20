using System;
using Microsoft.Office.Interop.Excel;
using ExcelGenerator.Model;
using Newtonsoft.Json;
using System.IO;
using ExcelGenerator.CellCalculator;
using ExcelGenerator.Enum;
using ExcelGenerator.Utils;

namespace ExcelGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = File.OpenRead(XUtil.sourceFileLocation);
            StreamReader reader = new StreamReader(data);
            string json = reader.ReadToEnd();
            CSet cs = JsonConvert.DeserializeObject<CSet>(json);
            XUtil xUtil = new XUtil();
            xUtil.CreateExcel(cs);
        }
    }
}
