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
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ReplicationExcel
{
    public class ConvertListToFile
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="URL"></param>
        /// <param name="encoding"></param>
        /// <returns></returns>
        public List<string> ExtractFromFile(string URL, Encoding encoding)
        {
            List<string> lines = new List<string>();
            string line = "";
            // Read the file and display it line by line.
            StreamReader file = new StreamReader(URL, encoding);
            while ((line = file.ReadLine()) != null)
            {
                lines.Add(line);
            }
            file.Close();
            return lines;
        }
        /// <summary>
        /// Cette fonction permet d'enregistrer un fichier
        /// </summary>
        /// <param name="URL">Url basique, Peut demander le FileSaveName de OpenFile</param>
        /// <param name="items"></param>
        /// <param name="encoding"> 
        ///     Objet qui contient le type d'encodage du fichier 
        ///         Exemple : Encoding encode = Encoding.UTF8;  || Encoding encode = Encoding.GetEncoding("latin1");
        /// </param>
        public void ExportToFile(string URL, List<string> items, Encoding encoding)
        {
            List<string> lines = items;
            // Read the file and display it line by line.
            StreamWriter file = new StreamWriter(@URL, false, encoding);
            for (int i = 0; i < lines.Count(); i++)
            {
                file.WriteLine(lines[i]);
            }
            file.Close();
        }
    }
}
