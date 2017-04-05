using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace ReplicationExcel
{
    [DefaultSheet("NameDefaultSheet")] // rendre dynamique
    public class Student
    {
        private List<string> _Name;

        [FromRange("D3","D15")]

        public List<string> Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        private List<string> _Family;
        [FromRange("C3","C15")]
        public List<string> Family
        {
            get { return _Family; }
            set { _Family = value; }
        }

        private List<string> _Numbers;
        [FromRange("A3", "A13")]

        public List<string> Numbers
        {
            get { return _Numbers; }
            set { _Numbers = value; }
        }
    }
}
