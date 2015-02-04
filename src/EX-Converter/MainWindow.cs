using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Threading;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace EX_Converter
{
    public partial class MainWindow : Form
    {
        #region Fields
        private string excelPath = String.Empty;
        private string xmlPath = String.Empty;

        private Thread convertWorker;
        private bool writeCasesOnly;
        private bool L2Enabled;
        private bool allowDuplicateSuiteName;
        private bool readerErrDetected;

        private int numActiveSheet;
        private int numStartRow;
        private int numEndRow;
        private int numL1Name;
        private int numL1Details;
        private int numL2Name;
        private int numL2Details;
        private int numCaseName;
        private int numSummary;
        private int numPreconditions;
        private int numImportance;
        private int numActions;
        private int numExpected;

        private const int MAX_SHEET = 256;
        private const int MAX_ROW = 1048576;
        private const int MAX_COLUMN = 16384;
        #endregion

        #region Constructors and initial utility methods.
        public MainWindow()
        {
            InitializeComponent();
            this.SetInitialValues();
        }
        private void SetInitialValues()
        {
            //File selections.
            this.excelPath = String.Empty;
            this.xmlPath = String.Empty;
            this.textBoxExcelPath.Text = String.Empty;
            this.textBoxXmlPath.Text = String.Empty;

            //Template.
            this.RadioButtonCases.Select();
            this.writeCasesOnly = true;

            this.textBoxL1Name.Enabled = false;
            this.textBoxL1Details.Enabled = false;

            this.L2Enabled = false;
            this.CheckBoxEnableL2.Checked = false;
            this.textBoxL2Name.Enabled = false;
            this.textBoxL2Details.Enabled = false;
            
            this.CheckBoxAllowDupSuite.Checked = false;
            this.allowDuplicateSuiteName = false;
            this.readerErrDetected = false;

            //Log window.
            this.logWindow.HideSelection = true;
        }
        #endregion

        #region Input validation methods.
        private bool CheckNonEmptyMappings()
        {
            if ((this.textBoxActiveSheet.Text == String.Empty)
                || (this.textBoxStartRow.Text == String.Empty)
                || (this.textBoxEndRow.Text == String.Empty)
                || (this.textBoxCaseName.Text == String.Empty)
                || ((!this.writeCasesOnly) && (this.textBoxL1Name.Text == String.Empty))
                || ((this.L2Enabled) && (this.textBoxL2Name.Text == String.Empty))
                )
            {
                MessageBox.Show("Excel Mapping Error: Mandatory mappings (in bold texts) cannot be empty!");
                return false;
            }
            else
                return true;
        }
        private bool ValidateMappings()
        {
            if (!this.ValidateSheet(this.textBoxActiveSheet.Text, ref numActiveSheet))
                return false;
            if (!this.ValidateRow(this.textBoxStartRow.Text, ref numStartRow))
                return false;
            if (!this.ValidateRow(this.textBoxEndRow.Text, ref numEndRow))
                return false;

            if (!this.ValidateColumn(this.textBoxL1Name.Text, ref numL1Name))
                return false;
            if (!this.ValidateColumn(this.textBoxL1Details.Text, ref numL1Details))
                return false;

            if (!this.ValidateColumn(this.textBoxL2Name.Text, ref numL2Name))
                return false;
            if (!this.ValidateColumn(this.textBoxL2Details.Text, ref numL2Details))
                return false;

            if (!this.ValidateColumn(this.textBoxCaseName.Text, ref numCaseName))
                return false;
            if (!this.ValidateColumn(this.textBoxSummary.Text, ref numSummary))
                return false;
            if (!this.ValidateColumn(this.textBoxPreconditions.Text, ref numPreconditions))
                return false;
            if (!this.ValidateColumn(this.textBoxImportance.Text, ref numImportance))
                return false;
            if (!this.ValidateColumn(this.textBoxActions.Text, ref numActions))
                return false;
            if (!this.ValidateColumn(this.textBoxExpectedResult.Text, ref numExpected))
                return false;

            else
                return true;
        }
        private bool ValidateSheet(string text, ref int destNum)
        {
            try
            {
                int sheetNum = Convert.ToInt32(text);

                if ((sheetNum > 0) && (sheetNum < MAX_SHEET))
                {
                    destNum = sheetNum;
                    return true;
                }
                else
                    throw new ArgumentException("Sheet number: " + sheetNum + " is out of range.(1 - "
                        + MAX_SHEET + ")");
            }
            catch (Exception err)
            {
                MessageBox.Show("Sheet Mapping Error: " + err.Message);
                return false;
            }
        }
        private bool ValidateColumn(string text, ref int destNum)
        {
            //To accept signal alphabetic char (a - z)
            if (text.Length == 1)
            {
                string newText = text.ToLower();
                int numText = Convert.ToInt32(newText[0]);
                if ((numText >= 97) && (numText <= 122))
                {
                    destNum = numText - 96;
                    return true;
                }
            }

            if (text == String.Empty)
            {
                destNum = 0;
                return true;
            }
            else
            {
                try
                {
                    int colNum = Convert.ToInt32(text);

                    if ((colNum > 0) && (colNum < MAX_COLUMN))
                    {
                        destNum = colNum;
                        return true;
                    }
                    else
                        throw new ArgumentException("Column number: " + colNum + " is out of range. (1 - "
                            + MAX_COLUMN + ")");
                }
                catch (Exception err)
                {
                    MessageBox.Show("Column Mapping Error: " + err.Message);
                    return false;
                }
            }
        }
        private bool ValidateRow(string text, ref int destNum)
        {
            try
            {
                int rowNum = Convert.ToInt32(text);

                if ((rowNum > 0) && (rowNum < MAX_COLUMN))
                {
                    destNum = rowNum;
                    return true;
                }
                else
                    throw new ArgumentException("Row number: " + rowNum + " is out of range. (1 - "
                        + MAX_ROW + ")");
            }
            catch (Exception err)
            {
                MessageBox.Show("Row Mapping Error: " + err.Message);
                return false;
            }
        }
        #endregion

        #region Cross-thread method invokes
        private void StartConvert(object threadParam)
        {
            try
            {
                this.PrintLog(LogType.Normal, 0, "Start reading data from Excel file: \""+ this.excelPath +"\"");

                ExcelReader reader = new ExcelReader(
                    this.excelPath,
                    this.writeCasesOnly,
                    this.L2Enabled,
                    this.allowDuplicateSuiteName,
                    this.numActiveSheet,
                    this.numStartRow,
                    this.numEndRow,
                    this.numL1Name,
                    this.numL1Details,
                    this.numL2Name,
                    this.numL2Details,
                    this.numCaseName,
                    this.numSummary,
                    this.numPreconditions,
                    this.numImportance,
                    this.numActions,
                    this.numExpected
                    );
                //Reading process.
                reader.ReaderEvent += new ExcelReader.ReaderEventHandler(this.HandleReaderLog);
                TestSuite result = reader.Read();
                reader.ReaderEvent -= new ExcelReader.ReaderEventHandler(this.HandleReaderLog);

                if (result != null)
                {
                    //this.PrintLog(LogType.Normal, 0, "Reading Excel data finished.");
                    this.PrintLog(LogType.Normal, 0, "Start writing data to XML file: \""
                        + this.xmlPath + "\"");
                    //Writing process.
                    XmlWriter writer = new XmlWriter(this.xmlPath, (!this.writeCasesOnly));
                    writer.Write(result);
                    this.PrintLog(LogType.Normal, 0, "Writing XML data finished.");

                    ConvertDoneDelegate done = threadParam as ConvertDoneDelegate;
                    if (!this.readerErrDetected)
                    {
                        done("Converting successfully done!");
                    }
                    else
                    {
                        done("Converting done but with errors or warnings! Please check the log information.");
                    }
                }
                else
                {
                    ConvertDoneDelegate notDone = threadParam as ConvertDoneDelegate;
                    notDone("Converting aborted due to Excel reading failure!");
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Error: " + err.Message);
                MessageBox.Show("Converting aborted!");
            }
            finally
            {
                this.EnableConvertButton();
            }
        }
        private delegate void ConvertDoneDelegate(string message);
        private void ShowConvertDoneMessage(string message)
        {
            MessageBox.Show(message);
        }
        //Reader event handler.
        private void HandleReaderLog(ExcelReader sourceReader, ExcelReader.ReaderEventArgs e)
        {
            if (e.Row != 0)
            {
                this.UpdateProgress(e.Row);
            }
            if (e.Type != LogType.Normal)
            {
                this.readerErrDetected = true;
            }
            this.PrintLog(e.Type, e.Row, e.Log);
        }
        //Method for pdating progress bar.
        private delegate void UpdateProgressDele(int row);
        private void UpdateProgress(int row)
        {
            if (this.ProgressBar.InvokeRequired)
            {
                UpdateProgressDele upd = new UpdateProgressDele(this.UpdateProgress);
                object[] updateParams = { row };
                this.Invoke(upd, updateParams);
            }
            else
            {
                this.ProgressBar.Minimum = this.numStartRow;
                this.ProgressBar.Maximum = this.numEndRow;
                this.ProgressBar.Value = row;
            }
        }
        //Method for printing log.
        private delegate void PrintLogDele(LogType type, int row, string message);
        private void PrintLog(LogType type, int row, string message)
        {
            if (this.logWindow.InvokeRequired)
            {
                PrintLogDele pd = new PrintLogDele(this.PrintLog);
                object[] printParams = { type, row, message };
                this.Invoke(pd, printParams);
            }
            else
            {                
                this.logWindow.Focus();
                this.logWindow.SelectionColor = Color.Black;

                string rowFormatString;
                rowFormatString = (row != 0) ? "<Row " + row.ToString() + "> " : String.Empty;

                if (type == LogType.Error)
                {
                    this.logWindow.SelectionColor = Color.Red;
                    this.logWindow.AppendText("[ERROR]:: " + rowFormatString + message + '\n');
                }
                else if (type == LogType.Warning)
                {
                    this.logWindow.SelectionColor = Color.DarkOrange;
                    this.logWindow.AppendText("[WARNING]:: " + rowFormatString + message + '\n');
                }
                else
                {
                    this.logWindow.AppendText("[INFO]:: " + rowFormatString + message + '\n');
                }
            }
        }
        //Method for button manipulating.
        private delegate void EnableConvertButtonDele();
        private void EnableConvertButton()
        {
            if (this.buttonConvert.InvokeRequired)
            {
                EnableConvertButtonDele ecbd = new EnableConvertButtonDele(this.EnableConvertButton);
                this.Invoke(ecbd);
            }
            else
            {
                this.buttonConvert.Enabled = true;
            }
        }
        #endregion

        #region MainWindow event handling methods.
        //Buttons.
        private void buttonConvert_Click(object sender, EventArgs e)
        {
            if (this.excelPath == string.Empty)
            {
                MessageBox.Show("Select the Excel file that you want to convert.");
            }
            else if (this.xmlPath == string.Empty)
            {
                MessageBox.Show("Select or create your destination XML file.");
            }
            else if (this.CheckNonEmptyMappings() && this.ValidateMappings())
            {
                this.buttonConvert.Enabled = false;
                this.readerErrDetected = false;

                try
                {
                    ConvertDoneDelegate convertDone = this.ShowConvertDoneMessage;

                    ParameterizedThreadStart paramStart = new ParameterizedThreadStart(this.StartConvert);
                    this.convertWorker = new Thread(paramStart);
                    this.convertWorker.IsBackground = true;
                    this.convertWorker.Start(convertDone);
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }
            GC.Collect();
        }
        private void buttonSelectExcel_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxExcelPath.Text = this.openFileDialog.FileName;
                this.excelPath = this.openFileDialog.FileName;
            }
        }
        private void buttonSelectXml_Click(object sender, EventArgs e)
        {
            if (this.saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBoxXmlPath.Text = this.saveFileDialog.FileName;
                this.xmlPath = this.saveFileDialog.FileName;
            }
        }
        //Template options.
        private void RadioButtonCases_CheckedChanged(object sender, EventArgs e)
        {
            if (this.RadioButtonCases.Checked)
            {
                this.writeCasesOnly = true;

                this.textBoxL1Name.Enabled = false;
                this.textBoxL1Name.Text = null;
                this.textBoxL1Details.Enabled = false;
                this.textBoxL1Details.Text = null;

                this.CheckBoxEnableL2.Checked = false;
                this.CheckBoxEnableL2.Enabled = false;

                this.CheckBoxAllowDupSuite.Checked = false;
                this.CheckBoxAllowDupSuite.Enabled = false;
            }
        }
        private void RadioButtonSuite_CheckedChanged(object sender, EventArgs e)
        {
            if (this.RadioButtonSuite.Checked)
            {
                this.writeCasesOnly = false;

                this.textBoxL1Name.Enabled = true;
                this.textBoxL1Details.Enabled = true;

                this.CheckBoxEnableL2.Checked = false;
                this.CheckBoxEnableL2.Enabled = true;

                this.CheckBoxAllowDupSuite.Checked = false;
                this.CheckBoxAllowDupSuite.Enabled = true;
            }
        }
        private void CheckBoxEnableL2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.CheckBoxEnableL2.Checked)
            {
                this.L2Enabled = true;
                this.textBoxL2Name.Enabled = true;
                this.textBoxL2Details.Enabled = true;
            }
            else
            {
                this.L2Enabled = false;
                this.textBoxL2Name.Enabled = false;
                this.textBoxL2Name.Text = null;
                this.textBoxL2Details.Enabled = false;
                this.textBoxL2Details.Text = null;
            }
        }
        private void checkBoxAllowDupSuite_CheckedChanged(object sender, EventArgs e)
        {
            if (this.CheckBoxAllowDupSuite.Checked)
            {
                this.allowDuplicateSuiteName = true;
            }
            else
            {
                this.allowDuplicateSuiteName = false;
            }
        }
        //Other buttons.
        private void ButtonClear_Click(object sender, EventArgs e)
        {
            this.logWindow.Text = string.Empty;

            this.ProgressBar.Minimum = 0;
            this.ProgressBar.Maximum = 100;
            this.ProgressBar.Value = 0;
        }
        private void ButtonReset_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("All defined values will be reset. Continue?", "Reset Values", MessageBoxButtons.YesNo)
                == System.Windows.Forms.DialogResult.Yes)
            {
                this.SetInitialValues();

                this.textBoxActiveSheet.Text = string.Empty;
                this.textBoxStartRow.Text = string.Empty;
                this.textBoxEndRow.Text = string.Empty;

                this.textBoxL1Name.Text = string.Empty;
                this.textBoxL1Details.Text = string.Empty;

                this.textBoxL2Name.Text = string.Empty;
                this.textBoxL2Details.Text = string.Empty;

                this.textBoxCaseName.Text = string.Empty;
                this.textBoxSummary.Text = string.Empty;
                this.textBoxPreconditions.Text = string.Empty;
                this.textBoxImportance.Text = string.Empty;
                this.textBoxActions.Text = string.Empty;
                this.textBoxExpectedResult.Text = string.Empty;

                this.logWindow.Text = string.Empty;
                this.ProgressBar.Minimum = 0;
                this.ProgressBar.Maximum = 100;
                this.ProgressBar.Value = 0;
            }
        }

        private void label12_Click(object sender, EventArgs e)
        {
            MessageBox.Show("\n\nVersion: EX-Converter v1.2.1"
                + "\nAuthor: Jack Zhang <sf.jackzhang@gmail.com>"
                + "\n\nRespects to TestLink - a great tool for software test management!",  
                "About EX-Converter");
        }
        #endregion

        private void label_L2_Click(object sender, EventArgs e)
        {

        }

        private void label_L1_Click(object sender, EventArgs e)
        {

        }

        private void labelCaseName_Click(object sender, EventArgs e)
        {

        }

        private void textBoxCaseName_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBoxImportance_TextChanged(object sender, EventArgs e)
        {

        }

        private void labelImportance_Click(object sender, EventArgs e)
        {

        }

        private void textBoxSummary_TextChanged(object sender, EventArgs e)
        {

        }

        private void labelSummary_Click(object sender, EventArgs e)
        {

        }

        private void textBoxPreconditions_TextChanged(object sender, EventArgs e)
        {

        }

        private void labelPreconditions_Click(object sender, EventArgs e)
        {

        }

        private void labelActions_Click(object sender, EventArgs e)
        {

        }

        private void textBoxActions_TextChanged(object sender, EventArgs e)
        {

        }

        private void labelExpectedResults_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
