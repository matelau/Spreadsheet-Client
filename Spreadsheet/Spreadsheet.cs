// Written by Asaeli Matelau for CS3500 Assignment PS6
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;
using System.Xml;

namespace SS
{
    /// <summary>
    /// Represents the State of a Simple Spreadsheet
    /// </summary>
    public class Spreadsheet : AbstractSpreadsheet
    {
        // Abstraction function: 
        // A Spreadsheet is a collection of named cells which contain a valid formula, a double, a string, or an empty string

        // Representation invariant:
        // cells must have valid names and must not contain a circular dependency, representation only contains NonEmptyCells != Empty String

        // need to change representation from HashSet to Dictionary <String Name, Cell>
        private Dictionary<String, Cell> representation;
        //private HashSet<Cell> representation;
        private DependencyGraph DG;

        /// <summary>
        /// True if the spreadsheet Has been Changed
        /// </summary>
        public override bool Changed
        {
            get;
            protected set;
        }

        // Zero Arg Constructor, isValid is always True, Strings Remain unmanipulated
        public Spreadsheet()
            : base(S => true, S => S, "default")
        {
            // Each cell must have a unique name
            representation = new Dictionary<String, Cell>();
            // DG used to determine dependecies
            DG = new DependencyGraph();   
        }

        // New 3 parameter Constructor
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version)
            : base(isValid, normalize, version)
        {
            // Each cell must have a unique name
            representation = new Dictionary<String, Cell>();
            // DG used to determine dependecies
            DG = new DependencyGraph();
        }


        public Spreadsheet(String Path, Func<string, bool> isValid, Func<string, string> normalize, string version)
            : base(isValid, normalize, version)
        {
            // create spreadsheet from file at "Path" 
            // DG used to determine dependecies
            DG = new DependencyGraph();

            GetSavedVersion(Path);

        }


        public override string GetSavedVersion(string filename)
        {
            try
            {

                using (XmlReader reader = XmlReader.Create(filename))
                {
                    representation = new Dictionary<String, Cell>();
                    string returnString = "";
                    String name = "";

                    while (reader.Read())
                    {

                        if (reader.IsStartElement())
                        {
                            string test = reader.Name;
                            switch (test)
                            {
                                case "spreadsheet":
                                    // check if version information is included 
                                    if (!(reader.HasAttributes))
                                    {
                                        throw new SpreadsheetReadWriteException("Version information was not included");
                                    }
                                    // save version info
                                    else
                                    {
                                        returnString = reader.GetAttribute(0);
                                        if (Version.Equals("default"))
                                        {
                                            Version = returnString;
                                        }
                                        if (!(returnString.Equals(Version)))
                                        { throw new SpreadsheetReadWriteException("Version information mismatch"); }
                                    }
                                    break;

                                case "cell":
                                    break;

                                //save name
                                case "name":
                                    name = reader.ReadElementContentAsString();
                                    
                                    //get the content
                                    String Content= "";
                                    reader.MoveToContent();
                                    Content = reader.ReadElementContentAsString();
                                    if (Content.StartsWith("="))
                                    {
                                        representation.Add(name, new Cell(Content, true));
                                    }
                                    else
                                    { representation.Add(name, new Cell(Content, false)); }

                                    break;
                                //add to representation
                                case "contents":
                                    //need to check if contents is a formula and get rid of the preappended =
                                    //string contents = reader.ReadContentAsString();
                                    //SetContentsOfCell(name, contents);

                                    break;
                            }
                        }
                    }
                    if (!(Version.Equals(returnString)))
                    {
                        throw new SpreadsheetReadWriteException("Version information did not match");
                    }
                    return returnString;
                }
            }
            catch (System.IO.DirectoryNotFoundException)
            {
                throw new SpreadsheetReadWriteException("The Path was Invalid");
            }
            catch (System.Xml.XmlException)
            {
                throw new SpreadsheetReadWriteException("The XML File is Invalid");
            }
            
        }

        public override void Save(string filename)
        {
            try
            {
                using (XmlWriter writer = XmlWriter.Create(filename))
                {
                    Changed = false;

                    writer.WriteStartDocument();
                    writer.WriteStartElement("spreadsheet");
                    writer.WriteAttributeString("version", Version);
                    
                    // Get all the NonemptyCells
                    foreach (String s in GetNamesOfAllNonemptyCells())
                    {
                        
                        writer.WriteStartElement("cell");
                        writer.WriteElementString("name", s);
                        object contents = GetCellContents(s);
                        double d;
                        //check if contents is a formula
                        if (representation[s].IsFormula)
                        {
                            String result = "=";
                            writer.WriteElementString("contents", result + contents.ToString());
                        }
                        else if (double.TryParse(representation[s].Contents.ToString(), out d))
                        {
                            writer.WriteElementString("contents", d.ToString());
                        }
                        else
                        { writer.WriteElementString("contents", (string)contents); }

                        //close cell
                        writer.WriteEndElement();
                        
                    }

                    
                    //close cell
                    writer.WriteEndElement();
                    // close spreadsheet
                    writer.WriteEndDocument();
                }
            }

            catch (System.IO.DirectoryNotFoundException)
            {
                throw new SpreadsheetReadWriteException("The Path was Invalid");
            }
        }

        public override object GetCellValue(string name)
        {

            String NewName = Normalize(name);
            checkValid(name);
            // Determine if cell is empty, if it contains a double, text, or a formula

            // if representation does not contain name 
            if (!(representation.ContainsKey(NewName)))
            {
                return "";
            }

            object contents = GetCellContents(NewName);
            // if the value is a double
            double value;
            if (double.TryParse(contents.ToString(), out value))
            {
                return value;
            }

            //alternative formula created with an xml writer and not using my representation
            else if (contents.ToString().StartsWith("="))
            {
                String NewContents = contents.ToString().Substring(1, contents.ToString().Length - 1);
                Formula f = new Formula(NewContents, IsValid, Normalize);
                object evaluated;
                try
                {
                    evaluated = f.Evaluate(lookup);

                }
                catch (ArgumentException)
                {
                    return new FormulaError("The Value of this Cell is Undefined");
                }

                return evaluated;
            }

             // value is a formula
            else if (representation[NewName].IsFormula  )
            {
                
                Formula f = (Formula)representation[NewName].Contents;
                object evaluated;
                try
                {
                    evaluated = f.Evaluate(lookup);
                    
                }
                catch (ArgumentException)
                {
                    return new FormulaError("The Value of this Cell is Undefined");
                }

                return evaluated;
            }
          
            
            // value must contain text
            else
                return contents.ToString();
        }


        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            List<string> returnList = new List<string>();
            foreach (String s in representation.Keys)
            {
                returnList.Add(s);
            }

            return returnList;
        }

        private Double lookup(String variable)
        {     
             // check if contents is a double or a string
            if (representation.ContainsKey(variable))
            {
                object contents = representation[variable].Contents;
                // if the variable contains a formula recursively call lookup on those variables
                if (representation[variable].IsFormula)
                {
                    Formula f = (Formula)representation[variable].Contents;
                    object x = f.Evaluate(lookup);
                    return (double)x;
                }


                double tryToParse;
                if (double.TryParse(contents.ToString(), out tryToParse))
                {
                    return tryToParse;
                }
            }

               throw new ArgumentException();

           
        }


        public override object GetCellContents(string name)
        {
            //Determines wether of not the cells name is valid 
            String newName = Normalize(name); 
            if (checkValid(newName))
            {
                //Check if cell is nonempty (already set in representation)
                if (representation.ContainsKey(newName))
                { return representation[newName].Contents; }
            }

            
            // else return empty string
            return "";
        }


        public override ISet<string> SetContentsOfCell(string name, string content)
        {
            checkValid(name);
            String NewName = Normalize(name);
            if (content == null)
            {
                throw new ArgumentNullException();
            }

            //this.Changed = true;
            // content is a double
            double value;
            if (double.TryParse(content, out value))
            {
                return SetCellContents(NewName, value);
            }


            // content is a formula
            else if (content.StartsWith("="))
            {
                String formula = content.Substring(1, content.Length-1);
                formula = Normalize(formula);
                Formula x = new Formula(formula, IsValid, Normalize);
                return SetCellContents(NewName, x);
            }

            // content is text
            else
                return SetCellContents(NewName, content);
        }

        protected override ISet<string> SetCellContents(string name, double number)
        {
            //Check Valid

            checkValid(name);

            Changed = true;
            //Check if cell is nonempty
            if (representation.ContainsKey(name))
            {
                //set contents
                representation[name].Contents = number;
                representation[name].IsFormula = false; // reset isFormula incase it has been changed from a formula

            }

            // else add cell to representation and set its contents
            else
            {
                representation.Add(name, (new Cell( number, false)));
            }

            IEnumerable<String> returnS = GetCellsToRecalculate(name);

      
            //prepare ReturnSet later to be used with Recalculate
            HashSet<String> returnSet = new HashSet<string>();
            foreach (String s in returnS)
            {
                returnSet.Add(s);
               // representation[s].Value = GetCellValue(s);
            }
            returnSet.Add(name);
            return returnSet;
            
        }

        protected override ISet<string> SetCellContents(string name, string text)
        {
            //Check Valid
            checkValid(name);
            Changed = true;
            if (text != "")
            {
                //check if cell name is already represented
                if (representation.ContainsKey(name))
                {
                    representation[name].Contents = text;
                    representation[name].IsFormula = false; // reset isFormula incase it has been changed from a formula
                }



               // else add cell to representation and set its contents
                else { representation.Add(name, (new Cell(text, false))); }
            }

            IEnumerable<String> returnS = GetCellsToRecalculate(name);
            HashSet<String> returnSet = new HashSet<string>();
            foreach (String s in returnS)
            {
               //representation[s].Value = GetCellValue(s);
               returnSet.Add(s);

           }
            returnSet.Add(name);
            return returnSet;
        
        }

        protected override ISet<string> SetCellContents(string name, Formula formula)
        {
            //Check Valid
            checkValid(name);
            
            //determine if change would cause a circular dependency
            ISet<String> recalc = new HashSet<String>();
            recalc.Add(name);
            

            foreach (String s in formula.GetVariables())
            {
                DG.AddDependency(s, name);
            }


            IEnumerable<String> returnS;
            // Check for circular dependency
            try
            {
                returnS = GetCellsToRecalculate(recalc);
            }
            // if Circular Exception is found remove Dependency and throw again 
            catch (CircularException)
            {
                foreach (String s in formula.GetVariables())
                {
                    DG.RemoveDependency(s, name);
                }

                throw new CircularException(); 

            }


            //Check if cell is nonempty
            if (representation.ContainsKey(name))
            {
                // locate proper cell
                representation[name].Contents = formula;
                representation[name].IsFormula = true; // set is formula

            }
            // else add cell to representation and set its contents
            else { representation.Add(name, (new Cell( formula, true))); }

            Changed = true;
            // setup returnSet
            HashSet<String> returnSet = new HashSet<string>();
            foreach (String s in returnS)
            {
                returnSet.Add(s);
                //representation[s].Value = GetCellValue(s);
            }
            returnSet.Add(name);
            return returnSet;
        }

        protected override IEnumerable<string> GetDirectDependents(string name)
        {
            //Check Valid
            if (name == null)
            { throw new ArgumentNullException(); }
            string NewName = Normalize(name);
            checkValid(NewName);

            List<String> returnSet = new List<string>();
            foreach (String s in DG.GetDependents(NewName))
            {
                yield return s;
            }
        
        }

        /// <summary>
        /// Helper method that determines if a cells name is valid
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private bool checkValid(String name)
        {
            // should modify string
            String NewName = Normalize(name);
            String varPattern = "^[a-zA-Z]+[0-9][0-9]*$";
            if (NewName == null)
            { throw new InvalidNameException(); }
            if (IsValid(NewName) && Regex.IsMatch(name,varPattern ))
            { return true; }
            else
                throw new InvalidNameException(); 
        }

    }


}
