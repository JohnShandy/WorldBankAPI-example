using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldBankTest
{
    class WBDataItem
    {
        public string Indicator { get; set; }
        public string Country { get; set; }
        public string Date { get; set; }
        public int Value { get; set; }
    }
    
    class WorldBankTest
    {
        static void Main(string[] args)
        {
            string uri = "http://api.worldbank.org/countries/USA/indicators/AG.AGR.TRAC.NO?per_page=10&date=2000:2010";
            XDocument document = XDocument.Load(uri);
            XNamespace ns = "http://www.worldbank.org";

            var wbData = document.Element(ns + "data");

            List<WBDataItem> WBDataItems = (
                from datapoint in wbData.Elements(ns + "data")
                select new WBDataItem
                {
                    Indicator = (string)datapoint.Element(ns + "indicator"),
                    Country = (string)datapoint.Element(ns + "country"),
                    Date = (string)datapoint.Element(ns + "date"),
                    Value = string.IsNullOrEmpty((string)datapoint.Element(ns + "value")) ? 0 : int.Parse((string)datapoint.Element(ns + "value"))
                }).ToList<WBDataItem>();

            Console.WriteLine(WBDataItems[0].Country + ": " + WBDataItems[0].Indicator);

            foreach (WBDataItem i in WBDataItems)
            {
                Console.WriteLine(i.Date + ": " + i.Value.ToString());
            }

            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Add(misValue);

            xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            
            xlWorksheet.Cells[1, 1] = WBDataItems[0].Country + ": " + WBDataItems[0].Indicator;
            for (int i = 0; i < WBDataItems.Count; i++)
            {
                xlWorksheet.Cells[i + 2, 1] = WBDataItems[i].Date;
                xlWorksheet.Cells[i + 2, 2] = WBDataItems[i].Value;
            }

            xlWorkbook.SaveAs(
                @"D:\output.xls",
                Excel.XlFileFormat.xlWorkbookNormal,
                misValue, misValue, misValue, misValue,
                Excel.XlSaveAsAccessMode.xlExclusive,
                misValue, misValue, misValue, misValue, misValue);

            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorksheet);
            releaseObject(xlWorkbook);
            releaseObject(xlApp);

            Console.WriteLine("Output written to Excel workbook.");
            Console.ReadLine();
        }

        private static void releaseObject(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
                o = null;
            }
            catch (Exception e)
            {
                o = null;
                Console.WriteLine(e.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
