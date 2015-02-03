using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EX_Converter
{
    internal class TestCase : ITlElement
    {
        #region Fields.
        public string AttrInternalId { get; private set; }
        //Implement ITlElement.
        public string AttrName { get; private set; }
        //Implement ITlElement.
        public string NodeOrder { get; private set; }
        public string ExternalId { get; private set; }
        public string Version { get; private set; }
        public string Summary { get; private set; }
        public string Preconditions { get; private set; }
        public int ExecutionType { get; private set; }
        public int Importance { get; private set; }

        public List<TestStep> Steps { get; private set; }

        //Should refer to TL.
        private const int EXECUTION_TYPE_DEFAULT = 1;
        //Should refer to TL.
        private const int IMPORTANCE_DEFAULT = 2;
        #endregion

        public TestCase(string name, string summary, string preconditions, int importance)
        {
            this.AttrInternalId = String.Empty;
            this.AttrName = name;

            this.NodeOrder = String.Empty;
            this.ExternalId = String.Empty;
            this.Version = String.Empty;
            this.Summary = summary;
            this.Preconditions = preconditions;
            this.ExecutionType = EXECUTION_TYPE_DEFAULT;
            this.Importance = importance;

            this.Steps = new List<TestStep>();
        }

        public void AddTestStep(TestStep newStep)
        {
            this.Steps.Add(newStep);
        }

        public bool NameEquals(ITlElement other)
        {
            if ((other is TestCase) && (this.AttrName == other.AttrName))
                return true;
            else
                return false;
        }

    }
}
