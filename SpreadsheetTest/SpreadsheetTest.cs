using SS;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using SpreadsheetUtilities;
using System.Text.RegularExpressions;

namespace SpreadsheetTest
{
    
    
    /// <summary>
    ///This is a test class for SpreadsheetTest and is intended
    ///to contain all SpreadsheetTest Unit Tests
    ///</summary>
    [TestClass()]
    public class SpreadsheetTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion

        /// <summary>
        /// Helper Method used for Testing 
        /// </summary>
        /// <param name="submitted"></param>
        /// <returns></returns>
        private bool isValidHelper(string submitted)
        {   
            String varPattern = "^[a-zA-Z]+[1-9][0-9]*$";
            if(Regex.IsMatch(submitted,varPattern))
            {return true;}
            else return false;
        }

        /// <summary>
        /// Helper Method used for Testing 
        /// </summary>
        /// <param name="submitted"></param>
        /// <returns></returns>
        private bool isValidHelper1(string submitted)
        {
            String varPattern = "^[A-Z]+[1-9][0-9]*$";
            if (Regex.IsMatch(submitted, varPattern))
            { return true; }
            else return false;
        }

        /// <summary>
        /// Helper Method used to Testing 
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string normalizeToUpper(String input)
        {
            return input.ToUpper();
        }

        /// <summary>
        /// Helper Method Used for Testing
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private string normalizeToLower(String input)
        {
            return input.ToLower();
        }

        /// <summary>
        ///A test for Spreadsheet Constructor
        ///</summary>
        [TestMethod()]
        public void SpreadsheetConstructorTest()
        {
            // test version information is returned correctly
            Spreadsheet target = new Spreadsheet(isValidHelper, normalizeToUpper, "One");
            target.Save("test1.xml");
            Assert.AreEqual("One", target.GetSavedVersion("test1.xml"));

        }

        /// <summary>
        ///A test for Spreadsheet Constructor
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(SpreadsheetReadWriteException))]
        public void SpreadsheetConstructorTest1()
        {
            //try to construct a spreadsheet from an invalid file
            Spreadsheet target = new Spreadsheet("C:\\Users\\Matelau\\Documents\\Visual Studio 2010\\Projects\\PS5\\states1.xml", isValidHelper, normalizeToUpper, "One"); 
        }

        /// <summary>
        ///A test for Spreadsheet Constructor
        ///</summary>
        [TestMethod()]
        public void SpreadsheetConstructorTest2()
        {
            //test isValidHelper1 and normalizeToUpper is working
            Spreadsheet target = new Spreadsheet(isValidHelper1, normalizeToUpper, "One");
            target.SetContentsOfCell("a1", "I Should be Capitol");
            String actual =target.GetCellValue("A1").ToString();
            String expected = "I Should be Capitol";
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for GetCellContents
        ///</summary>
        [TestMethod()]
        public void GetCellContentsTest()
        {
            Spreadsheet target = new Spreadsheet(isValidHelper1, normalizeToUpper,"three"); 
            string name ="b2";
            object expected = ""; 
            object actual = target.GetCellContents(name);
            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for GetCellValue
        ///</summary>
        [TestMethod()]
        public void GetCellValueTest()
        {
            Spreadsheet target = new Spreadsheet(isValidHelper1,normalizeToUpper, "two");
            string name = "a1";
            target.SetContentsOfCell(name, "2.0");

            object expected = 2.0;
            object actual = target.GetCellValue("A1");
            actual = target.GetCellValue(name);
            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void GetCellValue1()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "text");

            object actual = target.GetCellValue("a1");

            Assert.AreEqual("text", actual);
        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void GetCellValue2()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            object actual = target.GetCellValue("a1");
            Assert.AreEqual("", actual);
        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellValue3()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            object actual = target.GetCellValue("az");
        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        [ExpectedException(typeof(InvalidNameException))]
        public void GetCellValue4()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            object actual = target.GetCellValue(null);
        }


        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        [ExpectedException(typeof(FormulaFormatException))]
        public void GetCellValue5()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "=++2");
        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void GetCellValue6()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("A11", "=b2+b2*b3");
            object actual = target.GetCellValue("A11");
            Assert.AreEqual(new FormulaError("The provided Variable Was not valid"), actual);
        }

        /// <summary>
        ///A test for getCellValue
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void GetCellValue7()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("A11", "=b2+b2*b3");
            target.SetContentsOfCell("b2", "=b3+b3");
            target.SetContentsOfCell("b3", "=b4*b4");
            target.SetContentsOfCell("b4", "3.0");
            object actual = target.GetCellValue("b3");
            Assert.AreEqual(9.0, actual);
        }



        /// <summary>
        ///A test for GetSavedVersion
        ///</summary>
        [TestMethod()]
        public void GetSavedVersionTest()
        {
            Spreadsheet target = new Spreadsheet(); 
            string filename = "GSVtest.xml";
            target.Save(filename);
            string expected = "default"; 
            string actual;
            actual = target.GetSavedVersion(filename);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for GetSavedVersion
        ///</summary>
        [TestMethod()]
        public void GetSavedVersionTest1()
        {
            Spreadsheet target = new Spreadsheet();
            double x = 25;
            for (double i = 1; i < x; i++)
            {
                double d = i + 1;
                String add ="a"+i;
                target.SetContentsOfCell(add, "="+"a"+d+"+"+"a"+d);   
            }
            string filename = "GSVtest.xml";
            target.Save(filename);
            string expected = "default";
            string actual;
            actual = target.GetSavedVersion(filename);
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for Save
        ///</summary>
        [TestMethod()]
        public void SaveTest()
        {
            Spreadsheet target = new Spreadsheet(); // TODO: Initialize to an appropriate value
            string filename = "testMeNow.xml";
            //target.SetContentsOfCell("a1", "Hello");
            //GetSavedVersion(filename);
            target.Save(filename);
            //Spreadsheet testTargert = new Spreadsheet(
        }

   

        /// <summary>
        ///A test for SetContentsOfCell
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SetContentsOfCellTest()
        {
            Spreadsheet target = new Spreadsheet();
            string name = "b2";
            string content = null; 
            target.SetContentsOfCell(name, content);

        }

        /// <summary>
        ///A test for SetContentsOfCell
        ///</summary>
        [TestMethod()]
        public void SetContentsOfCellTest1()
        {
            Spreadsheet target = new Spreadsheet();
            string name = "b2";
            string content = "3.0";
            target.SetContentsOfCell(name, content);
            target.SetContentsOfCell(name, "4.0");
            object actual = target.GetCellValue("b2");
            Assert.AreEqual(4.0, actual);
        }

        /// <summary>
        ///A test for SetContentsOfCell
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SetContentsOfCellTest2()
        {
            Spreadsheet target = new Spreadsheet();
            string name = "b2";
            string content = null;
            target.SetContentsOfCell(name, content);

        }

        /// <summary>
        ///A test for SetContentsOfCell
        ///</summary>
        [TestMethod()]
        public void SetContentsOfCellTest3()
        {
            Spreadsheet target = new Spreadsheet();
            string name = "b2";
            string content = "text";
            target.SetContentsOfCell(name, content);
            target.SetContentsOfCell(name, "text2");
            object actual = target.GetCellValue("b2");
            Assert.AreEqual("text2", actual);
        }

        /// <summary>
        ///A test for SetContentsOfCell
        ///</summary>
        [TestMethod()]
        [ExpectedException(typeof(CircularException))]
        public void SetContentsOfCellTest4()
        {
            Spreadsheet target = new Spreadsheet();
            string name = "b2";
            string content = "=b3+b3";
            target.SetContentsOfCell(name, content);
            target.SetContentsOfCell(name, "=b3+b4");
            target.SetContentsOfCell("b4", "=b2+b2");
            //object actual = target.GetCellValue("b2");
            //Assert.AreEqual("text2", actual);
        }


        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void lookupTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "3.0");

            object expected = 3.0; 
            object actual = target.GetCellValue("a1")   ;
            
            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void lookupTest1()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "=b1+b1");

            object expected = new FormulaError("The provided Variable Was not valid");
            object actual = target.GetCellValue("a1");

            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void lookupTest2()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "=b1+b1");
            target.SetContentsOfCell("b1", "3.0");


            object expected = 6.0;
            object actual = target.GetCellValue("a1");

            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void lookupTest3()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "=a2+a2");//512
            target.SetContentsOfCell("a2", "=a3+a3");//256
            target.SetContentsOfCell("a3", "=a4+a4");//128
            target.SetContentsOfCell("a4", "=a5+a5");//64
            target.SetContentsOfCell("a5", "=a6+a6");//32
            target.SetContentsOfCell("a6", "=a7+a7");//16
            target.SetContentsOfCell("a7", "=a8+a8");//8
            target.SetContentsOfCell("a8", "=a9+a9");//4
            target.SetContentsOfCell("a9", "=a10+a10");//2
            target.SetContentsOfCell("a10", "1.0");

            target.Save("test2.txt");


            double expected = 512;
            object actual = target.GetCellValue("a1");

            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void lookupTest4()
        {
            //stress test
             //have x items add 1+1 together
            // result should equal 2^x-1 
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            double x = 24;
            ISet<String> CheckSet = new HashSet<String>();
            for (double i = 1; i < x; i++)
            {
                double d = i + 1;
                String add ="a"+i;
               target.SetContentsOfCell(add, "="+"a"+d+"+"+"a"+d);
               CheckSet.Add(add);                
            }
            CheckSet.Add("a" + x);

            // set final piece and get return set of dependents
            ISet<String> checkMatch = target.SetContentsOfCell("a" + x, "1.0");

            // compare sets
           Assert.AreEqual(true, CheckSet.SetEquals(checkMatch));

            // check that the solution equals 2^x-1 
            double expected = Math.Pow(2, x-1);
            object actual = target.GetCellValue("a1");

            Assert.AreEqual(expected, actual);
        }


        /// <summary>
        ///A test for lookup
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        //[ExpectedException(typeof(ArgumentException))]
        public void lookupTest5()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor();
            target.SetContentsOfCell("a1", "=b1+b1");
            target.SetContentsOfCell("b1", "text");

            FormulaError actual = (FormulaError)target.GetCellValue("a1");
            FormulaError expected = new FormulaError("The provided Variable Was not valid");
            Assert.AreEqual(expected, actual);

        } 
        /// <summary>
        ///A test for Changed
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void ChangedTest()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor(isValidHelper, normalizeToLower,"1");
            bool expected = false;
            bool actual;

            actual = target.Changed;
            Assert.AreEqual(expected, actual);

        }

        /// <summary>
        ///A test for Changed
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void ChangedTest1()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor(isValidHelper, normalizeToLower, "1");
            target.SetContentsOfCell("b1", "test");
            bool expected = true;
            bool actual;

            actual = target.Changed;
            Assert.AreEqual(expected, actual);
        }

        /// <summary>
        ///A test for Changed
        ///</summary>
        [TestMethod()]
        [DeploymentItem("PS5.dll")]
        public void ChangedTest2()
        {
            Spreadsheet_Accessor target = new Spreadsheet_Accessor(isValidHelper, normalizeToLower, "1");
            target.SetContentsOfCell("b1", "test");
            bool expected = false;
            bool actual;
            target.Save("test2.xml");
            actual = target.Changed;
            Assert.AreEqual(expected, actual);

        }
    }
}
