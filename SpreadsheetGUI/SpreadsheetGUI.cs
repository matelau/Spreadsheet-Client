using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;
using System.IO;

namespace SS
{
    public partial class SpreadsheetGUI : Form
    {
        private Spreadsheet backend;
        private char[] columns = {'A','B','C', 'D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'};
        int row, col;
        bool editing;
        String filename; 
        
        public Spreadsheet Backend
        {
            get { return backend; }
            set { backend = value; }
        } 

        public SpreadsheetGUI()
        {
            InitializeComponent();
            filename = null;
            editing = true;
            row = 0;
            col = 0; 
            backend = new Spreadsheet(s =>true, s => s.ToUpper(), "ps6");   
        }


        /// <summary>
        /// Updates Cell Selection, Value, and Content text box to correspond with the current selection.
        /// Also stores the cell name.
        /// </summary>
        /// <param name="ss"></param>
        private void spreadsheetPanel1_SelectionChanged(SpreadsheetPanel ss)
        {
            String cellSelection = "";

            ss.GetSelection(out col, out row);
            //clear contents initially
            content.Text = "";
            // update row for display use
            row++;
            // display the selected cell in the cell selected window store the selected cell in a var
            cellSelection = ""+columns[col]+row;
            cellSelected.Text = cellSelection;
            row--;
            // display the value 
            value.Text = backend.GetCellValue(cellSelection).ToString();
            //display the contents
            content.Text = backend.GetCellContents(cellSelection).ToString();

        }


     


        /// <summary>
        /// Method used to interact with the gui and backend upon keypress events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void spreadsheetPanel1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // get row and col 

            bool editing = true;

            // get current value
            String value1;
            spreadsheetPanel1.GetValue(col, row, out value1);
            //update value with last key press

            // if backspace is pressed update value & gui
            if (e.KeyChar.Equals('\b') && value1.Length >= 1)
            {
                //expensive concatenating strings
                value1 = value1.Substring(0, value1.Length - 1);
                spreadsheetPanel1.SetValue(col, row, value1);
            }

            //update due to other input 
            else
            {
                value1 = value1 + e.KeyChar;
                spreadsheetPanel1.SetValue(col, row, value1);
            }


            //pressing enter signals the editing process is complete allowing the program to proceed
            if (e.KeyChar.Equals('\r'))
            {
                editing = false; 
            }

            String cell = cellName(row, col);
            //insure at minimal cell is evaluated

            //if no longer editing set the cells content
            if (!editing)
            {
                // try to Set the cells contents and only set the value once there is no errors
                //update backend
                setContent(value1, cell);
            }
               
        }


        /// <summary>
        /// Helper method, used to convert SpreadsheetPanel grid coordinates to a cell name
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private String cellName(int row, int col)
        {
            row++;
            String cellName = "" + columns[col] + row;
            return cellName;
        }


        /// <summary>
        /// sets the Row and Column attributes according to the cell name
        /// </summary>
        /// <param name="cellName"></param>
        private void getRowCol(String cellName)
        {
            char[] var = cellName.ToCharArray();

            col = Array.IndexOf(columns, var[0]); 
            String rowParse;
            if (var.Length > 2)
            {
                rowParse = "" + var[1] + var[2];
            }
            else
            { rowParse = "" + var[1]; }
            
            Int32.TryParse(rowParse, out row); 
            //adjust row to agree with gui
            row = row - 1;


        }

        private void content_TextChanged(object sender, EventArgs e)
        {
            //check if the user is entering a formula via the textbox, otherwise ignore input
            
            bool containsVar = false;
            editing = true;
            //&& !(content.Text.Equals("SpreadsheetUtilities.FormulaError"))
            if (!editing )
            {
                if (content.Text.Contains("+") || content.Text.Contains("*") || content.Text.Contains("/"))
                {
                    containsVar = true;
                    //MessageBox.Show("I'm Working");
                }
                if (content.Text.StartsWith("="))
                {
                    setContent(content.Text, cellName(row, col));
                }
                else if (!containsVar)
                {
                    setContent(content.Text, cellName(row, col));
                }
            }
           
        }


        /// <summary>
        /// Helper Method used to update the content
        /// </summary>
        private void setContent(String Content, String cellA)
        {
            String cell = cellA;

            String value1 = Content;

            bool flag = false;

            //insure at minimal cell is evaluated for value
            HashSet<String> eval = new HashSet<string>();
            //eval.Add(cell);

            try
            {
                //spreadsheetPanel1.SetValue(col, row, value1);
                eval = (HashSet<String>)backend.SetContentsOfCell(cell, value1);
                flag = true;
            }
            catch (SpreadsheetUtilities.FormulaFormatException)
            {
                flag = false;
            }
            catch (CircularException)
            {
                flag = false;
                MessageBox.Show("The formula entered caused a Circular Dependency");
            }
            // if input causes a Formula Format Exception or Circular Exception it will be "ignored" 
            if (!flag)
            {
                spreadsheetPanel1.SetValue(col, row, "");
                backend.SetContentsOfCell(cell, "");
            }


            else
            {
                //update the values of the call and all impacted cells
                Object value2 = backend.GetCellValue(cell);
                foreach (String cellName1 in eval)
                {
                    getRowCol(cellName1);
                    value2 = backend.GetCellValue(cellName1);
                    String ValueToBeSet = value2.ToString();
                    if (ValueToBeSet.Contains("SpreadsheetUtilities."))
                    {
                        spreadsheetPanel1.SetValue(col, row, "Formula Error");
                    }
                    else
                    {
                        spreadsheetPanel1.SetValue(col, row, ValueToBeSet);
                    }
  
                }
                //spreadsheetPanel1.SetValue(col, row, value2.ToString());
                //reset row and column
                getRowCol(cell);

            }

            
        }

        private void content_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyValue.Equals(13) )
            {
                editing = false;
            }
            bool containsVar = false;
            if (!editing)
            {
                if (content.Text.StartsWith("="))
                {
                    setContent(content.Text, cellName(row, col));
                }
                else if (content.Text.Contains("+") || content.Text.Contains("*") || content.Text.Contains("/"))
                {
                    containsVar = true;
                    MessageBox.Show("Only valid formulas are allowed to contain operators consider adding a = symbol to the beginning of your statement");
                }
                
                else if (!containsVar)
                {
                    setContent(content.Text, cellName(row, col));
                }
            }
        }

        // Deals with the New menu
        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Tell the application context to run the form on the same
            // thread as the other forms.
            SpreadsheetApplicationContext.getAppContext().RunForm(new SpreadsheetGUI());
        }

        // Deals with the Close menu
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (backend.Changed)
            {
                DialogResult result2 = MessageBox.Show("Would you like to Save your previous Work?", "Important Query", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    if (filename == null)
                    {
                        saveFileDialog1.ShowDialog();
                    }
                    else
                    {
                        backend.Save(filename);
                        MessageBox.Show("Saved");
                    }
                   
                    openFileDialog1.ShowDialog();
                }
                else if (result2 == DialogResult.No)
                {
                    openFileDialog1.ShowDialog();
                }
            }
            else
            {
                openFileDialog1.ShowDialog();
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (filename == null)
            {
                saveFileDialog1.ShowDialog();
            }
            else
            {
                backend.Save(filename);
                MessageBox.Show("Saved");
            }
            
           
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            // Get file name.
            filename = saveFileDialog1.FileName;
            // Write to the file name selected.
            backend.Save(filename);
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            filename = openFileDialog1.FileName;
            //string directoryPath = Path.GetDirectoryName(name);

            try
            {
                backend = new Spreadsheet(filename, s => true, s => s.ToUpper(), "ps6");
                foreach (String s in backend.GetNamesOfAllNonemptyCells())
                {
                    Object Content1 = backend.GetCellContents(s);
                    setContent(Content1.ToString(), s);
                }  

            }
            catch (ArgumentException)
            {
                MessageBox.Show("Error");
            }
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("To use this Spreadsheet you will need to be aware of a few simple rules." +"\n"+"\n"+"To enter contents into a cell you can enter information directly by clicking a cell and typing or you can use the provided contents panel, but you may not be allowed to use both if you try to alternate you will be stuck with the contents panel."+ "\n"+"\n" +"Also if you want to enter text directly by clicking on a cell you are not allowed to use the Left, Right, Up, and Down keys." + "\n"+"\n" +"Finally, Only Valid Input will be accepted. You are not allowed to use operators such as +, -, *, or / unless they are part of a valid formula ");
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
        }

        private void SpreadsheetGUI_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (backend.Changed)
            {
                DialogResult result2 = MessageBox.Show("Would you like to Save your previous Work?", "Important Query", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result2 == DialogResult.Yes)
                {
                    saveFileDialog1.ShowDialog();
                }
            }

            

        }
  

        
    }
}
