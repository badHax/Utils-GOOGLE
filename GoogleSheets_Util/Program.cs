using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheets_Util
{
    class Program
    {
        static void Main(string[] args)
        {
            // print a sheet
            SheetsUtil.ConsoleLogSheet(
                SheetsUtil.GetSheet("15qN7pMrXn8Au_abM4LD7k5mIEx3Z_JZoTsl3DlJz2y4"));

            //print a row
            SheetsUtil.ConsoleLogRow(SheetsUtil.GetRow("15qN7pMrXn8Au_abM4LD7k5mIEx3Z_JZoTsl3DlJz2y4", 2));

            // print cell
            SheetsUtil.ConsoleLogRow(SheetsUtil.GetCell("15qN7pMrXn8Au_abM4LD7k5mIEx3Z_JZoTsl3DlJz2y4", 4, "c"));

            // print put cell
            Console.WriteLine(SheetsUtil.PutCell("15qN7pMrXn8Au_abM4LD7k5mIEx3Z_JZoTsl3DlJz2y4", 11, "g", "test"));

            Console.Read();
        }
    }
}
