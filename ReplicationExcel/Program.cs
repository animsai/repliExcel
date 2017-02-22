/* Version : 1.0
 * Auteurs : Bacaicoa Thomas, Geier Phillip, Argelli Angelo, Ehlers Thomas, Schupbach Loïc
 * Date    : 12 mars 2015
 * Classe  : IFA-P3B
 */

/* Version : 2.0
 * Auteurs : Jessica Sulzbach, Dilan Marques, Gabor Tagliabue, Sean Metry
 * Date    : 25 janvier 2017 - [ajouter - fin]
 * Classe  : I.IN-P4B
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReplicationExcel
{
    static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
