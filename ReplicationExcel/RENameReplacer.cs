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
        // fonction sert a remplacer les pokemons par la liste des noms d'eleves 
        public void raplacespokemons(List<string> names)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;


            Excel.Range places = Sheet.get_Range("B7","B14");

            currentFind = places.Find("pokemon", "B8", Excel.XlFindLookIn.xlValues);


            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

                currentFind = places.FindNext(currentFind);
            }


        }

    }
}
