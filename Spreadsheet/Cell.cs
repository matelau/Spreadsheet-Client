// Written by Asaeli Matelau for CS3500 Assignment PS4
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SS
{
    /// <summary>
    /// Represents a Single Cell of Spreadsheet
    /// Each Cell consists of a name and contents
    /// The contents maybe a double, a formula error, a string, or if null is implied to be the empty string ""
    /// </summary>
    class Cell
    {

        private object contents;
        private bool isFormula;
        private object value; 

        public bool IsFormula
        {
            get { return isFormula; }
            set { isFormula = value; }
        }



        public Cell( object contents1, bool formula)
        {
            contents = contents1;
            isFormula = formula;

        }

        public object Contents
        {
            get { return contents; }
            set { contents = value; }
        }

        public object Value
        {
            get { return value; }
            set { this.value = value; }
        }
    }
}
