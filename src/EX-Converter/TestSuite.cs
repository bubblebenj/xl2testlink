using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EX_Converter
{
    internal class TestSuite : ITlElement
    {
        #region Fields.
        //Implement ITlElement.
        public string AttrName { get; private set; }
        //Implement ITlElement.
        public string NodeOrder { get; private set; }
        public string Details { get; private set; }
        public List<ITlElement> ChildrenElements { get; private set; }
        #endregion

        public TestSuite(string suiteName)
        {
            this.AttrName = suiteName;
            this.Details = String.Empty;
            this.ChildrenElements = new List<ITlElement>();
        }

        public void AddChild(ITlElement tlElement)
        {
            this.ChildrenElements.Add(tlElement);
        }

        public bool NameEquals(ITlElement other)
        {
            if ((other is TestSuite) && (this.AttrName == other.AttrName))
                return true;
            else
                return false;
        }

        public TestSuite FindSuite(string suiteName)
        {
            foreach (ITlElement elem in this.ChildrenElements)
            {
                if ((elem is TestSuite) && ((elem as TestSuite).AttrName == suiteName))
                {
                    return elem as TestSuite;
                }
            }
            return null;
        }


    }
}
