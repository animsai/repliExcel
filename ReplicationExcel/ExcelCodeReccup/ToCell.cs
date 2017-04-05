using System;
using System.Collections.Generic;
using System.Text;

namespace ReplicationExcel
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ToCell : Attribute
    {
        private string _CellAddress;
        public string CellAddress
        {
            get { return _CellAddress; }
        }

        public ToCell(string cellAddress)
        {
            _CellAddress = cellAddress;
        }

    }
}
