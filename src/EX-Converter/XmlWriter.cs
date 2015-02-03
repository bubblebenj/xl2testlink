using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Xml;

namespace EX_Converter
{
    internal class XmlWriter
    {
        #region Fields and const values.
        private readonly string newFilePath;
        private readonly bool writingTestSuite;

        private XmlDocument document;
        private XmlDeclaration declaration;
        private XmlElement root;

        private const string NODENAME_TEST_SUITE = "testsuite";
        private const string NODENAME_DETAILS = "details";
        private const string NODENAME_TEST_CASES = "testcases";
        private const string NODENAME_TEST_CASE = "testcase";
        private const string NODENAME_NAME = "name";
        private const string NODENAME_INTERNALID = "internalid";

        private const string NODENAME_NODE_ORDER = "node_order";
        private const string NODENAME_EXTERNALID = "externalid";
        private const string NODENAME_VERSION = "version";
        private const string NODENAME_SUMMARY = "summary";
        private const string NODENAME_PRECONDITIONS = "preconditions";
        private const string NODENAME_EXECUTION_TYPE = "execution_type";
        private const string NODENAME_IMPORTANCE = "importance";

        private const string NODENAME_STEPS = "steps";
        private const string NODENAME_STEP = "step";
        private const string NODENAME_STEP_NUMBER = "step_number";
        private const string NODENAME_ACTIONS = "actions";
        private const string NODENAME_EXPECTED_RESULTS = "expectedresults";   
        #endregion

        #region Constructor and utility methods
        public XmlWriter(string newPath, bool isWritingSuite)
        {
            this.newFilePath = newPath;
            this.writingTestSuite = isWritingSuite;
            this.CreateXmlFrame();
        }
        private void CreateXmlFrame()
        {
            this.document = new XmlDocument();

            //XML declaration.
            this.declaration = this.document.CreateXmlDeclaration("1.0", "UTF-8", null);
            this.document.AppendChild(this.declaration);

            //Set root node.
            if (this.writingTestSuite)
            {
                this.root = this.document.CreateElement(NODENAME_TEST_SUITE);
                this.root.SetAttribute(NODENAME_NAME, String.Empty);
                this.AppendCDataElement(this.root, NODENAME_NODE_ORDER, String.Empty);
                this.AppendCDataElement(this.root, NODENAME_DETAILS, String.Empty);
            }
            else
            {
                this.root = this.document.CreateElement(NODENAME_TEST_CASES);
            }
            this.document.AppendChild(this.root);
        }
        #endregion

        #region The Writing method and its utility methods.
        //Main method for writting to XML file.
        public void Write(TestSuite suite)
        {
            if (this.writingTestSuite)
            {
                this.WriteToTestSuite(suite);
            }
            else
            {
                this.WriteToTestCases(suite);
            }
            //Finally write to file.
            this.document.Save(newFilePath);
        }
        private void WriteToTestCases(TestSuite suite)
        {
            foreach (ITlElement elem in suite.ChildrenElements)
            {
                if (elem is TestCase)
                {
                    this.AppendTestCase(this.root, elem as TestCase);
                }
                else
                {
                    throw new ArgumentException("Error! XML element:\n"
                        + "Name: " + elem.AttrName + '\n'
                        + "(Node Order: " + elem.NodeOrder + ")\n"
                        + "is not a \"testcase\"");
                }
            }
        }
        private void WriteToTestSuite(TestSuite suite)
        {
            this.AppendTestSuite(this.root, suite);
        }

        #region Methods for appending elements.
        private void AppendCDataElement(XmlNode parentNode, string name, string text)
        {
            XmlElement elem = this.document.CreateElement(name);
            parentNode.AppendChild(elem);

            XmlNode cData = this.document.CreateNode(XmlNodeType.CDATA, String.Empty, String.Empty);
            elem.AppendChild(cData);

            cData.InnerText = text;
        }
        private void AppendSingleStep(XmlNode parentSteps, TestStep step)
        {
            XmlElement newStep = this.document.CreateElement(NODENAME_STEP);
            parentSteps.AppendChild(newStep);

            this.AppendCDataElement(newStep, NODENAME_STEP_NUMBER, Convert.ToString(step.StepNumber));
            this.AppendCDataElement(newStep, NODENAME_ACTIONS, step.Actions);
            this.AppendCDataElement(newStep, NODENAME_EXPECTED_RESULTS, step.ExpectedResults);
            this.AppendCDataElement(newStep, NODENAME_EXECUTION_TYPE, Convert.ToString(step.ExecutionType));
        }
        private void AppendSteps(XmlNode parentTestCase, List<TestStep> steps)
        {
            XmlElement newSteps = this.document.CreateElement(NODENAME_STEPS);
            parentTestCase.AppendChild(newSteps);

            foreach (TestStep step in steps)
            {
                this.AppendSingleStep(newSteps, step);
            }
        }
        private void AppendCaseInfo(XmlNode parentTestCase, TestCase testCase)
        {
            this.AppendCDataElement(parentTestCase, NODENAME_NODE_ORDER, testCase.NodeOrder);
            this.AppendCDataElement(parentTestCase, NODENAME_EXTERNALID, testCase.ExternalId);
            this.AppendCDataElement(parentTestCase, NODENAME_VERSION, testCase.Version);
            this.AppendCDataElement(parentTestCase, NODENAME_SUMMARY, testCase.Summary);
            this.AppendCDataElement(parentTestCase, NODENAME_PRECONDITIONS, testCase.Preconditions);
            this.AppendCDataElement(parentTestCase, NODENAME_EXECUTION_TYPE, 
                            Convert.ToString(testCase.ExecutionType));
            this.AppendCDataElement(parentTestCase, NODENAME_IMPORTANCE, 
                            Convert.ToString(testCase.Importance));
            //All steps.
            this.AppendSteps(parentTestCase, testCase.Steps);
        }
        private void AppendTestCase(XmlNode parentNode, TestCase testCase)
        {
            XmlElement newTestCase = this.document.CreateElement(NODENAME_TEST_CASE);
            newTestCase.SetAttribute(NODENAME_INTERNALID, testCase.AttrInternalId);
            newTestCase.SetAttribute(NODENAME_NAME, testCase.AttrName);
            this.AppendCaseInfo(newTestCase, testCase);

            parentNode.AppendChild(newTestCase);
        }
        private void AppendSuiteInfo(XmlNode parentTestSuite, TestSuite testSuite)
        {
            this.AppendCDataElement(parentTestSuite, NODENAME_NODE_ORDER, String.Empty);
            this.AppendCDataElement(parentTestSuite, NODENAME_DETAILS, testSuite.Details);
        }
        private void AppendTestSuite(XmlNode parentNode, TestSuite testSuite)
        {
            XmlElement newTestSuite = this.document.CreateElement(NODENAME_TEST_SUITE);
            newTestSuite.SetAttribute(NODENAME_NAME, testSuite.AttrName);
            this.AppendSuiteInfo(newTestSuite, testSuite);

            foreach (ITlElement elem in testSuite.ChildrenElements)
            {
                if (elem is TestSuite)
                {
                    //Recursively append all children of this test suite.
                    this.AppendTestSuite(newTestSuite, elem as TestSuite);
                }
                else
                {
                    this.AppendTestCase(newTestSuite, elem as TestCase);
                }
            }

            //Determine if parent node is root. If is root, move newTestSuite nodes to root.
            if (parentNode != this.root)
            {
                parentNode.AppendChild(newTestSuite);
            }
            else
            {
                int origChildCount = newTestSuite.ChildNodes.Count;
                for (int nodeCounter = 0; nodeCounter < origChildCount; nodeCounter++)
                {
                    if (nodeCounter < 2)
                    {
                        newTestSuite.RemoveChild(newTestSuite.FirstChild);
                    }
                    else
                    {
                        this.root.AppendChild(newTestSuite.FirstChild);
                    }
                }
            }
        }
        #endregion
        #endregion
    }
}
