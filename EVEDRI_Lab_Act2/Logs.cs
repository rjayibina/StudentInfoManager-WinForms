using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EVEDRI_Lab_Act2
{
    class Logs
    {
        Workbook book = new Workbook();

        public void insertLogs(string user, string message)
        {
            string fileLocation = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;

            string folder = "assets";
            string file = "Book1.xlsx";
            string path = Path.Combine(fileLocation, folder, file);

            Workbook book = new Workbook();

            book.LoadFromFile(path);

            Worksheet sh = book.Worksheets[1];

            int r = sh.Rows.Length + 1;

            sh.Range[r, 1].Value = user;
            sh.Range[r, 2].Value = message;
            sh.Range[r, 3].Value = DateTime.Now.ToString("MM/dd/yyyy");
            sh.Range[r, 4].Value = DateTime.Now.ToString("hh:mm:ss tt");

            book.SaveToFile(path, ExcelVersion.Version2016);
        }
    }
}
