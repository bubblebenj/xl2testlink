using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;


namespace EX_Converter
{
    internal class ExcelReader
    {
        #region Customized eventArgs, event handler and events.
        public class ReaderEventArgs : EventArgs
        {
            public LogType Type { get; private set; }
            public string Log { get; private set; }
            public int Row { get; private set; }

            public ReaderEventArgs(LogType type, int row, string log)
            {
                this.Type = type;
                this.Row = row;
                this.Log = log;
            }
        }
        public delegate void ReaderEventHandler(ExcelReader source, ReaderEventArgs e);
        public event ReaderEventHandler ReaderEvent;
        #endregion

        #region Fields and const values.
        private readonly string FilePath;
        private readonly bool ReadCaseOnly;
        private readonly bool L2_Enabled;
        private readonly bool allowDuplicateSuite;

        private int ActiveSheetIndex { get; set; }
        private int StartRowIndex { get; set; }
        private int EndRowIndex { get; set; }

        private int L1_FolderIndex { get; set; }
        private int L2_FolderIndex { get; set; }

        private int NameColumnIndex { get; set; }
        private int SummaryColumnIndex { get; set; }
        private int PreconditionsColumnIndex { get; set; }
        private int ImportanceColumnIndex { get; set; }

        private int ActionsColumnIndex { get; set; }
        private int ExpectedResultsColumnIndex { get; set; }

        private const string SUITE_DEFAULT_NAME = "#TEST_SUITE_DEFAULT_NAME#";
        private TestCase currentCase = null;
        private TestSuite L1_CurrentSuite = null;
        private TestSuite L2_CurrentSuite = null;

        //Short name.
        private object missing = Missing.Value;
        private int totalCaseCount = 0;
        #endregion

        #region Constructor.
        public ExcelReader(string excelPath, bool readCase, bool l2Enabled, bool allowDupSuite,
                            int activeSheet, int startRow, int endRow,
                            int folderL1, int folderL2,
                            int caseName, int summary, int preconditions, int importance,
                            int actions, int expected)
        {
            this.FilePath = excelPath;
            this.ReadCaseOnly = readCase;
            this.L2_Enabled = l2Enabled;
            this.allowDuplicateSuite = allowDupSuite;

            this.ActiveSheetIndex = activeSheet;
            this.StartRowIndex = startRow;
            this.EndRowIndex = endRow;
            this.L1_FolderIndex = folderL1;
            this.L2_FolderIndex = folderL2;

            this.NameColumnIndex = caseName;
            this.SummaryColumnIndex = summary;
            this.PreconditionsColumnIndex = preconditions;
            this.ImportanceColumnIndex = importance;
            this.ActionsColumnIndex = actions;
            this.ExpectedResultsColumnIndex = expected;
        }
        #endregion

        #region Utility methods for adding new case/step and check-ups.
        private void PushBackNewCase(TestSuite destSuite, string name, string summary, string preconditions, int importance,
                                        string actions, string expected)
        {
            TestCase newCase = new TestCase(name, summary, preconditions, importance);
            this.currentCase = newCase;
            TestStep newStep = new TestStep(1, actions, expected);
            newCase.AddTestStep(newStep);
            destSuite.AddChild(newCase as ITlElement);

            this.totalCaseCount++;
        }
        private void PushBackNewStep(string actions, string expected)
        {
            TestStep newStep = new TestStep(this.currentCase.Steps.Count + 1, actions, expected);
            this.currentCase.AddTestStep(newStep);
        }
        private string GetShortName(ITlElement elem)
        {
            return elem.AttrName.Substring
                (0,elem.AttrName.Length >= 50 ? 50 : elem.AttrName.Length);
        }
        private KeyValuePair<int, bool> ValidateImportance(string importanceString)
        {
            //Set test case importance to "Medium" if column not defined by user.
            if (this.ImportanceColumnIndex == 0)
                return new KeyValuePair<int, bool>(2, true);

            if (importanceString == null)
            {
                return new KeyValuePair<int, bool>(2, false);
            }
            //If importance defined with strings.
            string trimmedStr = importanceString.Trim();
            string lowerStr = trimmedStr.ToLower();

            switch (lowerStr)
            {
                case "high":
                    return new KeyValuePair<int, bool>(3, true);
                case "medium":
                    return new KeyValuePair<int, bool>(2, true);
                case "low":
                    return new KeyValuePair<int, bool>(1, true);
            }
            //If importance defined with numbers.
            int retInt = 0;
            try
            {
                retInt = Convert.ToInt32(importanceString);
                if ((retInt >= 1) && (retInt <= 3))
                {
                    return new KeyValuePair<int, bool>(retInt, true);
                }
                else
                {
                    return new KeyValuePair<int, bool>(2, false);
                }
            }
            catch (Exception err)
            {
                return new KeyValuePair<int, bool>(2, false);
            }
        }
        #endregion

        #region Log info const string fields.
        private const string LOG_EmptyRow = "Empty row.";
        private const string LOG_ContainNoCase = "current row contains no test case.";
        private const string LOG_AddNewStep = "Added a new step to current test case.";
        private const string LOG_AddNewCase = "Added a new test case: \"";
        private const string LOG_AddNewCaseL1 = "Added a new case (to current Level 1 test suite): \"";
        private const string LOG_AddNewCaseL2 = "Added a new case (to current Level 2 test suite): \"";
        private const string LOG_Postfix = "\".";
        private const string LOG_NewL1 = "Added a new Level 1 test suite: \"";
        private const string LOG_DupL1 = "Duplicate Level 1 suite name found. Appended to existing Level 1 test suite: \"";
        private const string LOG_NewL2 = "Added a new Level 2 test suite: \"";
        private const string LOG_DupL2 = "Duplicate Level 2 suite name found. Appended to existing Level 2 test suite: \"";

        private const string WARNING_InvalidImportanceValue =
            "Invalid test importance value found. Set case to default importance value: 2 (Medium).";

        private const string ERROR_CaseNameMissing =
            "Invalid row: Test case name missing. Parsing skipped.";
        private const string ERROR_NoCurrentCase =
            "Invalid row: Destination test case not available. Cannot add new step. Parsing skipped.";
        private const string ERROR_L1Missing =
            "Invalid row: Destination Level_1 test suite not available. Cannot add new case or new step. Parsing skipped.";
        #endregion

        #region Parsing logics.
        //Parse logic for cases only.
        private void Parse_Case(int rowIndex, TestSuite baseSuite, 
                                string name, string summary, string preconditions, string importance,
                                string actions, string expected)
        {
            if ((actions == null) && (expected == null)
                && (name == null) && (summary == null) && (preconditions == null) && (importance == null))
            {
                //Indicate empty row.
                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_EmptyRow));
                return;
            }

            if (name == null)
            {
                if ((actions == null) && (expected == null))
                {
                    //empty...
                    //...invalid row...
                    ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_CaseNameMissing));
                }
                else
                {
                    //add step to current case
                    if (this.currentCase == null)
                    {
                        //err...no current case...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                    ERROR_NoCurrentCase));
                    }
                    else
                    {
                        this.PushBackNewStep(actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewStep));
                    }
                }
            }
            else
            {
                //Validate test importance value
                KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                if (!pairImportance.Value)
                {
                    ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                    WARNING_InvalidImportanceValue));
                }
                //add new case to base suite
                this.PushBackNewCase(baseSuite, name, summary, preconditions, pairImportance.Key, 
                                        actions, expected);
                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCase + this.GetShortName(this.currentCase) + LOG_Postfix));
            }
        }
        //Parse logic for level 1 suite only.
        private void Parse_L1(int rowIndex, TestSuite baseSuite, string L1Val, 
                                string name, string summary, string preconditions, string importance,
                                string actions, string expected)
        {
            if ((actions == null) && (expected == null)
                && (name == null) && (summary == null) && (preconditions == null) && (importance == null)
                && (L1Val == null))
            {
                //Indicate empty row.
                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_EmptyRow));
                return;
            }

            if (L1Val == null)
            {
                if (this.L1_CurrentSuite == null)
                {
                    //no current L1...
                    ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                           ERROR_L1Missing));
                }
                else if ((name == null) && (actions == null) && (expected == null))
                {
                    if ((summary != null) || (preconditions != null) || (importance != null))
                    {
                        //not empty but no case...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                           ERROR_CaseNameMissing));
                    }
                    else
                    {
                        //this row contains no case.
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                        return;
                    }
                }
                else if (name == null)
                {
                    //add step to L1 last case
                    if (this.currentCase == null)
                    {
                        //err...no current case...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                    }
                    else
                    {
                        this.PushBackNewStep(actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewStep));
                    }
                }
                else
                {
                    //Validate test importance value
                    KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                    if (!pairImportance.Value)
                    {
                        ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                        WARNING_InvalidImportanceValue));
                    }
                    //add new case to L1
                    this.PushBackNewCase(this.L1_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                            actions, expected);
                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL1 + this.GetShortName(this.currentCase) + LOG_Postfix));
                }
            }
            else
            {
                //set current L1 to base(search first); reset current case
                TestSuite existingL1Suite;
                if (this.allowDuplicateSuite)
                {
                    existingL1Suite = null;
                }
                else
                {
                    existingL1Suite = baseSuite.FindSuite(L1Val);
                }

                if (existingL1Suite == null)
                {
                    TestSuite newL1Suite = new TestSuite(L1Val);
                    baseSuite.AddChild(newL1Suite as ITlElement);
                    this.L1_CurrentSuite = newL1Suite;

                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                        LOG_NewL1 + this.GetShortName(this.L1_CurrentSuite) + LOG_Postfix));
                }
                else
                {
                    this.L1_CurrentSuite = existingL1Suite;

                    ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                        LOG_DupL1 + this.GetShortName(this.L1_CurrentSuite) + LOG_Postfix));
                }
                this.currentCase = null;



                if ((name == null) && (actions == null) && (expected == null))
                {
                    if ((summary != null) || (preconditions != null) || (importance != null))
                    {
                        //not empty but no case...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                           ERROR_CaseNameMissing));
                    }
                    else
                    {
                        //this row contains no case.
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                        return;
                    }
                }
                else if (name == null)
                {
                    //no new case name...
                    ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                }
                else
                {
                    //Validate test importance value
                    KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                    if (!pairImportance.Value)
                    {
                        ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                        WARNING_InvalidImportanceValue));
                    }
                    //add new case to L1
                    this.PushBackNewCase(this.L1_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                            actions, expected);
                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL1 + this.GetShortName(this.currentCase) + LOG_Postfix));
                }
            }
        }
        //Parse logic for Level 2 suite enabled.
        private void Parse_L1L2(int rowIndex, TestSuite baseSuite, string L1Val, string L2Val, 
                                string name, string summary, string preconditions, string importance, 
                                string actions, string expected)
        {
            if ((actions == null) && (expected == null)
                && (name == null) && (summary == null) && (preconditions == null) && (importance == null)
                && (L1Val == null) && (L2Val == null))
            {
                //Indicate empty row.
                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_EmptyRow));
                return;
            }

            #region L1Val == null
            if (L1Val == null)
            {
                if (this.L1_CurrentSuite == null)
                {
                    //no current L1...
                    ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                           ERROR_L1Missing));
                }
                else if (L2Val == null)
                {
                    if (this.L2_CurrentSuite == null)
                    {
                        if ((name == null) && (actions == null) && (expected == null))
                        {
                            if ((summary != null) || (preconditions != null) || (importance != null))
                            {
                                //not empty but no case...
                                ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                        ERROR_CaseNameMissing));
                            }
                            else
                            {
                                //this row contains no case.
                                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                                return;
                            }
                        }
                        else if (name == null)
                        {
                            //<no new case name... //*** should add new step to current case>
                            //check current case sinc logic not clear...
                            if (this.currentCase == null)
                            {
                                //err...no current case...
                                ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                            ERROR_NoCurrentCase));
                            }
                            else
                            {
                                this.PushBackNewStep(actions, expected);
                                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                        LOG_AddNewStep));
                            }
                        }
                        else
                        {
                            //Validate test importance value
                            KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                            if (!pairImportance.Value)
                            {
                                ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                                WARNING_InvalidImportanceValue));
                            }
                            //add new case to L1
                            this.PushBackNewCase(this.L1_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                                    actions, expected);
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL1 + this.GetShortName(this.currentCase) + LOG_Postfix));
                        }
                    }
                    else if ((name == null) && (actions == null) && (expected == null))
                    {
                        if ((summary != null) || (preconditions != null) || (importance != null))
                        {
                            //not empty but no case...
                            ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_CaseNameMissing));
                        }
                        else
                        {
                            //this row contains no case.
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                            return;
                        }
                    }
                    else if (name == null)
                    {
                        //add step to L2 last case
                        if (this.currentCase == null)
                        {
                            //err...no current case...
                            ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                        }
                        else
                        {
                            this.PushBackNewStep(actions, expected);
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewStep));
                        }
                    }
                    else
                    {
                        //Validate test importance value
                        KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                        if (!pairImportance.Value)
                        {
                            ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                            WARNING_InvalidImportanceValue));
                        }
                        //add new case to L2
                        this.PushBackNewCase(this.L2_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                                actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL2 + this.GetShortName(this.currentCase) + LOG_Postfix));
                    }
                }
                else
                {
                    //set current L2 to L1 (search first); reset current case
                    TestSuite existingL2Suite;
                    if (this.allowDuplicateSuite)
                    {
                        existingL2Suite = null;
                    }
                    else
                    {
                        existingL2Suite = this.L1_CurrentSuite.FindSuite(L2Val);
                    }

                    if (existingL2Suite == null)
                    {
                        TestSuite newL2Suite = new TestSuite(L2Val);
                        this.L1_CurrentSuite.AddChild(newL2Suite as ITlElement);
                        this.L2_CurrentSuite = newL2Suite;

                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_NewL2 + this.GetShortName(this.L2_CurrentSuite) + LOG_Postfix));
                    }
                    else
                    {
                        this.L2_CurrentSuite = existingL2Suite;
                        ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                LOG_DupL2 + this.GetShortName(this.L2_CurrentSuite) + LOG_Postfix));
                    }
                    this.currentCase = null;

                    if ((name == null) && (actions == null) && (expected == null))
                    {
                        if ((summary != null) || (preconditions != null) || (importance != null))
                        {
                            //not empty but no case...
                            ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_CaseNameMissing));
                        }
                        else
                        {
                            //this row contains no case.
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                            return;
                        }
                    }
                    else if (name == null)
                    {
                        //no new case name...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                    }
                    else
                    {
                        //Validate test importance value
                        KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                        if (!pairImportance.Value)
                        {
                            ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                            WARNING_InvalidImportanceValue));
                        }
                        //add new case to L2
                        this.PushBackNewCase(this.L2_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                                actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL2 + this.GetShortName(this.currentCase) + LOG_Postfix));
                    }
                }
            }
            #endregion
            #region L1Val != null
            else
            {
                //set current L1 to base (search first?) and reset L2 to null; and reset current case
                TestSuite existingL1Suite;
                if (this.allowDuplicateSuite)
                {
                    existingL1Suite = null;
                }
                else
                {
                    existingL1Suite = baseSuite.FindSuite(L1Val);
                }

                if (existingL1Suite == null)
                {
                    TestSuite newL1Suite = new TestSuite(L1Val);
                    baseSuite.AddChild(newL1Suite as ITlElement);
                    this.L1_CurrentSuite = newL1Suite;

                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                        LOG_NewL1 + this.GetShortName(this.L1_CurrentSuite) + LOG_Postfix));
                }
                else
                {
                    this.L1_CurrentSuite = existingL1Suite;

                    ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                        LOG_DupL1 + this.GetShortName(this.L1_CurrentSuite) + LOG_Postfix));
                }
                //must reset L2 and current case
                this.L2_CurrentSuite = null;
                this.currentCase = null;

                if (L2Val == null)
                {
                    if ((name == null) && (actions == null) && (expected == null))
                    {
                        if ((summary != null) || (preconditions != null) || (importance != null))
                        {
                            //not empty but no case...
                            ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_CaseNameMissing));
                        }
                        else
                        {
                            //this row contains no case.
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                            return;
                        }
                    }
                    else if (name == null)
                    {
                        //no new case name...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                    }
                    else
                    {
                        //Validate test importance value
                        KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                        if (!pairImportance.Value)
                        {
                            ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                            WARNING_InvalidImportanceValue));
                        }
                        //add new case to L1
                        this.PushBackNewCase(this.L1_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                                actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL1 + this.GetShortName(this.currentCase) + LOG_Postfix));
                    }
                }
                else
                {
                    //set current L2 to L1 (search first); reset current case
                    TestSuite existingL2Suite;
                    if (this.allowDuplicateSuite)
                    {
                        existingL2Suite = null;
                    }
                    else
                    {
                        existingL2Suite = this.L1_CurrentSuite.FindSuite(L2Val);
                    }
                    if (existingL2Suite == null)
                    {
                        TestSuite newL2Suite = new TestSuite(L2Val);
                        this.L1_CurrentSuite.AddChild(newL2Suite as ITlElement);
                        this.L2_CurrentSuite = newL2Suite;

                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_NewL2 + this.GetShortName(this.L2_CurrentSuite) + LOG_Postfix));
                    }
                    else
                    {
                        this.L2_CurrentSuite = existingL2Suite;
                        ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                LOG_DupL2 + this.GetShortName(this.L2_CurrentSuite) + LOG_Postfix));
                    }
                    this.currentCase = null;

                    if ((name == null) && (actions == null) && (expected == null))
                    {
                        if ((summary != null) || (preconditions != null) || (importance != null))
                        {
                            //not empty but no case...
                            ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_CaseNameMissing));
                        }
                        else
                        {
                            //this row contains no case.
                            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex, LOG_ContainNoCase));
                            return;
                        }
                    }
                    else if (name == null)
                    {
                        //no new case name...
                        ReaderEvent(this, new ReaderEventArgs(LogType.Error, rowIndex,
                                ERROR_NoCurrentCase));
                    }
                    else
                    {
                        //Validate test importance value
                        KeyValuePair<int, bool> pairImportance = this.ValidateImportance(importance);
                        if (!pairImportance.Value)
                        {
                            ReaderEvent(this, new ReaderEventArgs(LogType.Warning, rowIndex,
                                            WARNING_InvalidImportanceValue));
                        }
                        //add new case to L2
                        this.PushBackNewCase(this.L2_CurrentSuite, name, summary, preconditions, pairImportance.Key, 
                                                actions, expected);
                        ReaderEvent(this, new ReaderEventArgs(LogType.Normal, rowIndex,
                                LOG_AddNewCaseL2 + this.GetShortName(this.currentCase) + LOG_Postfix));
                    }
                }
            }
            #endregion
        }
        #endregion

        #region Read file method.
        public TestSuite Read()
        {
            if (this.StartRowIndex > this.EndRowIndex)
                throw new ArgumentOutOfRangeException("End row number should not be less than start row number.");

            Application excelApplication = new Application();
            if (excelApplication == null)
            {
                throw new NullReferenceException("Fail to initiate Excel application!");
            }

            ReaderEvent(this, new ReaderEventArgs(LogType.Normal, 0,
                                "Excel application initiated successfully."));

            try
            {
                excelApplication.Visible = false;
                Workbook destWorkbook = excelApplication.Application.Workbooks.Open(this.FilePath, missing, true, missing,
                    missing, missing, missing, missing, missing, true, missing, missing, missing,
                    missing, missing);

                ReaderEvent(this, new ReaderEventArgs(LogType.Normal, 0,
                                "Destination Excel file opened successfully."));

                //Result test suite to return.
                TestSuite resultTestSuite = new TestSuite(SUITE_DEFAULT_NAME);

                try
                {
                    Worksheet destSheet = (Worksheet)destWorkbook.Worksheets[this.ActiveSheetIndex];
                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, 0,
                                "Destination worksheet opened successfully."));
                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, 0, "Start parsing rows. (" 
                        + this.StartRowIndex.ToString() + " - " + this.EndRowIndex.ToString() + ")"));

                    for (int row = this.StartRowIndex; row <= this.EndRowIndex; row++)
                    {
                        #region Get strings from mapped cells.
                        string L1Value = null;
                        if ((!this.ReadCaseOnly) && (this.L1_FolderIndex != 0))
                        {
                            L1Value = Convert.ToString(((Range)destSheet.Cells[row, this.L1_FolderIndex]).get_Value(missing));
                        }

                        string L2Value = null;
                        if ((this.L2_Enabled) && (this.L2_FolderIndex != 0))
                        {
                            L2Value = Convert.ToString(((Range)destSheet.Cells[row, this.L2_FolderIndex]).get_Value(missing));
                        }

                        string nameValue = null;
                        if (this.NameColumnIndex != 0)
                        {
                            nameValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.NameColumnIndex]).get_Value(missing));
                        }

                        string summaryValue = null;
                        if (this.SummaryColumnIndex != 0)
                        {
                            summaryValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.SummaryColumnIndex]).get_Value(missing));
                        }

                        string preconditionsValue = null;
                        if (this.PreconditionsColumnIndex != 0)
                        {
                            preconditionsValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.PreconditionsColumnIndex]).get_Value(missing));
                        }

                        string importanceValue = null;
                        if (this.ImportanceColumnIndex != 0)
                        {
                            importanceValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.ImportanceColumnIndex]).get_Value(missing));
                        }

                        string actionsValue = null;
                        if (this.ActionsColumnIndex != 0)
                        {
                            actionsValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.ActionsColumnIndex]).get_Value(missing));
                        }

                        string expectedValue = null;
                        if (this.ExpectedResultsColumnIndex != 0)
                        {
                            expectedValue =
                                Convert.ToString(((Range)destSheet.Cells[row, this.ExpectedResultsColumnIndex]).get_Value(missing));
                        }
                        #endregion

                        if (this.ReadCaseOnly)
                        {
                            //Call parse method "Case".
                            this.Parse_Case(row, resultTestSuite, 
                                            nameValue, summaryValue, preconditionsValue, importanceValue, 
                                            actionsValue, expectedValue);
                        }
                        else if (this.L2_Enabled)
                        {
                            //Call parse method "L1L2".
                            this.Parse_L1L2(row, resultTestSuite, L1Value, L2Value,
                                            nameValue, summaryValue, preconditionsValue, importanceValue, 
                                            actionsValue, expectedValue);
                        }
                        else
                        {
                            //Call parse method "L1".
                            this.Parse_L1(row, resultTestSuite, L1Value, 
                                            nameValue, summaryValue, preconditionsValue, importanceValue, 
                                            actionsValue, expectedValue);
                        }

                        //Collect garbage once in each 10 rows.
                        if (row % 10 == 0)
                        {
                            GC.Collect();
                        }
                    }

                    ReaderEvent(this, new ReaderEventArgs(LogType.Normal, 0, 
                                    "Reading Excel data finished. " + this.totalCaseCount + " test cases generated in total."));
                    return resultTestSuite;
                }
                catch (Exception err)
                {
                    throw err;
                }
                finally
                {
                    destWorkbook.Close(false, missing, missing);
                    destWorkbook = null;
                }
            }
            catch (Exception err)
            {
                ReaderEvent(this, new ReaderEventArgs(LogType.Error, 0,
                                "Severe Error! Parsing failed! Message: " + err.Message));
                return null;
            }
            finally
            {
                excelApplication.Quit();
                excelApplication = null;
                GC.Collect();
            }
        }
        #endregion

    }
}
