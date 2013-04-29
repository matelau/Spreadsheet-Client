using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;


namespace SpreadSheetGUITest
{
    /// <summary>
    /// Summary description for SpreadsheetGUITEst
    /// </summary>
    [CodedUITest]
    public class SpreadsheetGUITEst
    {
        ApplicationUnderTest app;

        public SpreadsheetGUITEst()
        {
        }

        [TestMethod]
        public void CodedUITestMethod1()
        {
            // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
            // For more information on generated code, see http://go.microsoft.com/fwlink/?LinkId=179463
        }

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        public void mytestinitialize()
        {

            this.UIMap.AssertMethod5();
            this.UIMap.RecordedMethod6();
            this.UIMap.AssertMethod6();
            this.UIMap.RecordedMethod7();
            this.UIMap.AssertMethod7();
            this.UIMap.RecordedMethod8();
            this.UIMap.AssertMethod8();
            this.UIMap.RecordedMethod9();
            this.UIMap.AssertMethod9();
            this.UIMap.RecordedMethod5();
            app = ApplicationUnderTest.Launch(@"..\..\..\SpreadsheetGUI\bin\Debug\SpreadsheetGUI.exe");
            this.UIMap.RecordedMethod3();
            this.UIMap.AssertMethod3();
            this.UIMap.RecordedMethod4();
            this.UIMap.AssertMethod4();

        }

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        public void MyTestCleanup()
        {
            app.Close();
            this.UIMap.AssertMethod1();
            this.UIMap.RecordedMethod2();
            this.UIMap.AssertMethod2();
        }

        #endregion

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
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
