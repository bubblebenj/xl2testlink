using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EX_Converter
{
    internal class TestStep
    {
        public int StepNumber { get; private set; }
        public string Actions { get; private set; }
        public string ExpectedResults { get; private set; }
        public int ExecutionType { get; private set; }


        public TestStep(int stepNumber, string stepActions, string stepExpected, int stepExeType = 1)
        {
            this.StepNumber = stepNumber;
            this.Actions = stepActions;
            this.ExpectedResults = stepExpected;
            this.ExecutionType = stepExeType;
        }
    }
}
