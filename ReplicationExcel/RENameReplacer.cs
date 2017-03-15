using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReplicationExcel
{

     

    class RENameReplacer
    {
        string ClasseurPath = "";
        Excel.Workbook WorkBook = null;
        Excel.Application ExcelApplication = null;
        Excel.Worksheet Sheet = null;


        // Ouvre le classeur à l'emplacement spécifié par path
        public void Initialize(string path, Excel.Worksheet recapitulatif)
        {
            ClasseurPath = path;

            ExcelApplication = new Excel.Application();
            ExcelApplication.Visible = false;

            WorkBook = ExcelApplication.Workbooks.Open(ClasseurPath);
            Sheet = recapitulatif; 
        }

    }
}
