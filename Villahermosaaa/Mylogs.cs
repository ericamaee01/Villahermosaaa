using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Villahermosaaa
{
    class Mylogs
    {
        Workbook book = new Workbook();

        public void insertLogs(string user, string message)
        {
            //logs
            Workbook book = new Workbook();
            book.LoadFromFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
            Worksheet sh = book.Worksheets[1];
            int r = sh.LastRow + 1;

            sh.Range[r, 1].Value = user;
            sh.Range[r, 2].Value = message;
            sh.Range[r, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[r, 4].Value = DateTime.Now.ToString("hh:mm:ss tt");

            book.SaveToFile("C:\\Users\\Erica Mae\\source\\repos\\Villahermosaaa\\book\\book1.xlsx");
        }
    }
    
}
