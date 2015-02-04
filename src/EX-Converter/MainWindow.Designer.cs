namespace EX_Converter
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
            this.TemplateTypeGroupBox = new System.Windows.Forms.GroupBox();
            this.CheckBoxAllowDupSuite = new System.Windows.Forms.CheckBox();
            this.CheckBoxEnableL2 = new System.Windows.Forms.CheckBox();
            this.RadioButtonSuite = new System.Windows.Forms.RadioButton();
            this.RadioButtonCases = new System.Windows.Forms.RadioButton();
            this.groupBoxSelectPath = new System.Windows.Forms.GroupBox();
            this.buttonSelectXml = new System.Windows.Forms.Button();
            this.buttonSelectExcel = new System.Windows.Forms.Button();
            this.labelXmlPath = new System.Windows.Forms.Label();
            this.labelExcelPath = new System.Windows.Forms.Label();
            this.textBoxXmlPath = new System.Windows.Forms.TextBox();
            this.textBoxExcelPath = new System.Windows.Forms.TextBox();
            this.groupboxOperations = new System.Windows.Forms.GroupBox();
            this.labelImportance = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.labelL2Name = new System.Windows.Forms.Label();
            this.labelL1Name = new System.Windows.Forms.Label();
            this.labelEndRow = new System.Windows.Forms.Label();
            this.labelStartRow = new System.Windows.Forms.Label();
            this.labelActions = new System.Windows.Forms.Label();
            this.labelExpectedResults = new System.Windows.Forms.Label();
            this.labelPreconditions = new System.Windows.Forms.Label();
            this.labelSummary = new System.Windows.Forms.Label();
            this.labelCaseName = new System.Windows.Forms.Label();
            this.textBoxSummary = new System.Windows.Forms.TextBox();
            this.textBoxExpectedResult = new System.Windows.Forms.TextBox();
            this.textBoxEndRow = new System.Windows.Forms.TextBox();
            this.textBoxImportance = new System.Windows.Forms.TextBox();
            this.textBoxActions = new System.Windows.Forms.TextBox();
            this.textBoxCaseName = new System.Windows.Forms.TextBox();
            this.textBoxStartRow = new System.Windows.Forms.TextBox();
            this.textBoxPreconditions = new System.Windows.Forms.TextBox();
            this.textBoxL2Name = new System.Windows.Forms.TextBox();
            this.textBoxL1Name = new System.Windows.Forms.TextBox();
            this.textBoxActiveSheet = new System.Windows.Forms.TextBox();
            this.labelActiveSheet = new System.Windows.Forms.Label();
            this.buttonConvert = new System.Windows.Forms.Button();
            this.logWindow = new System.Windows.Forms.RichTextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.ProgressBar = new System.Windows.Forms.ProgressBar();
            this.ButtonClear = new System.Windows.Forms.Button();
            this.ButtonReset = new System.Windows.Forms.Button();
            this.toolTip_12s = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip_5s = new System.Windows.Forms.ToolTip(this.components);
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.label12 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelL1Details = new System.Windows.Forms.Label();
            this.textBoxL1Details = new System.Windows.Forms.TextBox();
            this.labelL2Details = new System.Windows.Forms.Label();
            this.textBoxL2Details = new System.Windows.Forms.TextBox();
            this.TemplateTypeGroupBox.SuspendLayout();
            this.groupBoxSelectPath.SuspendLayout();
            this.groupboxOperations.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TemplateTypeGroupBox
            // 
            this.TemplateTypeGroupBox.Controls.Add(this.CheckBoxAllowDupSuite);
            this.TemplateTypeGroupBox.Controls.Add(this.CheckBoxEnableL2);
            this.TemplateTypeGroupBox.Controls.Add(this.RadioButtonSuite);
            this.TemplateTypeGroupBox.Controls.Add(this.RadioButtonCases);
            this.TemplateTypeGroupBox.Location = new System.Drawing.Point(6, 5);
            this.TemplateTypeGroupBox.Name = "TemplateTypeGroupBox";
            this.TemplateTypeGroupBox.Size = new System.Drawing.Size(210, 122);
            this.TemplateTypeGroupBox.TabIndex = 0;
            this.TemplateTypeGroupBox.TabStop = false;
            this.TemplateTypeGroupBox.Text = "Template Type";
            // 
            // CheckBoxAllowDupSuite
            // 
            this.CheckBoxAllowDupSuite.AutoSize = true;
            this.CheckBoxAllowDupSuite.Location = new System.Drawing.Point(22, 85);
            this.CheckBoxAllowDupSuite.Name = "CheckBoxAllowDupSuite";
            this.CheckBoxAllowDupSuite.Size = new System.Drawing.Size(157, 17);
            this.CheckBoxAllowDupSuite.TabIndex = 0;
            this.CheckBoxAllowDupSuite.Text = "Allow Duplicate Suite Name";
            this.toolTip_12s.SetToolTip(this.CheckBoxAllowDupSuite, resources.GetString("CheckBoxAllowDupSuite.ToolTip"));
            this.CheckBoxAllowDupSuite.UseVisualStyleBackColor = true;
            this.CheckBoxAllowDupSuite.CheckedChanged += new System.EventHandler(this.checkBoxAllowDupSuite_CheckedChanged);
            // 
            // CheckBoxEnableL2
            // 
            this.CheckBoxEnableL2.AutoSize = true;
            this.CheckBoxEnableL2.Location = new System.Drawing.Point(22, 64);
            this.CheckBoxEnableL2.Name = "CheckBoxEnableL2";
            this.CheckBoxEnableL2.Size = new System.Drawing.Size(124, 17);
            this.CheckBoxEnableL2.TabIndex = 0;
            this.CheckBoxEnableL2.Text = "Enable Level 2 Suite";
            this.toolTip_12s.SetToolTip(this.CheckBoxEnableL2, "Check to enable level 2 folders when generating a test specification.\r\nYou need to map an Excel sheet column to level 2 folders.");
            this.CheckBoxEnableL2.UseVisualStyleBackColor = true;
            this.CheckBoxEnableL2.CheckedChanged += new System.EventHandler(this.CheckBoxEnableL2_CheckedChanged);
            // 
            // RadioButtonSuite
            // 
            this.RadioButtonSuite.AutoSize = true;
            this.RadioButtonSuite.Location = new System.Drawing.Point(9, 42);
            this.RadioButtonSuite.Name = "RadioButtonSuite";
            this.RadioButtonSuite.Size = new System.Drawing.Size(73, 17);
            this.RadioButtonSuite.TabIndex = 0;
            this.RadioButtonSuite.TabStop = true;
            this.RadioButtonSuite.Text = "Test Suite";
            this.toolTip_12s.SetToolTip(this.RadioButtonSuite, "Generate an XML test specification with type of \"Test Suite\".\r\n(Used by \"Import Test Suite\" in TestLink.)");
            this.RadioButtonSuite.UseVisualStyleBackColor = true;
            this.RadioButtonSuite.CheckedChanged += new System.EventHandler(this.RadioButtonSuite_CheckedChanged);
            // 
            // RadioButtonCases
            // 
            this.RadioButtonCases.AutoSize = true;
            this.RadioButtonCases.Location = new System.Drawing.Point(9, 21);
            this.RadioButtonCases.Name = "RadioButtonCases";
            this.RadioButtonCases.Size = new System.Drawing.Size(78, 17);
            this.RadioButtonCases.TabIndex = 0;
            this.RadioButtonCases.TabStop = true;
            this.RadioButtonCases.Text = "Test Cases";
            this.toolTip_12s.SetToolTip(this.RadioButtonCases, "Generate an XML test specification with the type of \"Test Cases\".\r\n(Used by \"Import Test Cases\" in TestLink.)");
            this.RadioButtonCases.UseVisualStyleBackColor = true;
            this.RadioButtonCases.CheckedChanged += new System.EventHandler(this.RadioButtonCases_CheckedChanged);
            // 
            // groupBoxSelectPath
            // 
            this.groupBoxSelectPath.Controls.Add(this.buttonSelectXml);
            this.groupBoxSelectPath.Controls.Add(this.buttonSelectExcel);
            this.groupBoxSelectPath.Controls.Add(this.labelXmlPath);
            this.groupBoxSelectPath.Controls.Add(this.labelExcelPath);
            this.groupBoxSelectPath.Controls.Add(this.textBoxXmlPath);
            this.groupBoxSelectPath.Controls.Add(this.textBoxExcelPath);
            this.groupBoxSelectPath.Location = new System.Drawing.Point(225, 5);
            this.groupBoxSelectPath.Name = "groupBoxSelectPath";
            this.groupBoxSelectPath.Size = new System.Drawing.Size(475, 121);
            this.groupBoxSelectPath.TabIndex = 0;
            this.groupBoxSelectPath.TabStop = false;
            this.groupBoxSelectPath.Text = "Select Files:";
            // 
            // buttonSelectXml
            // 
            this.buttonSelectXml.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectXml.Location = new System.Drawing.Point(436, 86);
            this.buttonSelectXml.Name = "buttonSelectXml";
            this.buttonSelectXml.Size = new System.Drawing.Size(31, 25);
            this.buttonSelectXml.TabIndex = 2;
            this.buttonSelectXml.Text = "...";
            this.buttonSelectXml.UseVisualStyleBackColor = true;
            this.buttonSelectXml.Click += new System.EventHandler(this.buttonSelectXml_Click);
            // 
            // buttonSelectExcel
            // 
            this.buttonSelectExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectExcel.Location = new System.Drawing.Point(436, 34);
            this.buttonSelectExcel.Name = "buttonSelectExcel";
            this.buttonSelectExcel.Size = new System.Drawing.Size(31, 25);
            this.buttonSelectExcel.TabIndex = 1;
            this.buttonSelectExcel.Text = "...";
            this.buttonSelectExcel.UseVisualStyleBackColor = true;
            this.buttonSelectExcel.Click += new System.EventHandler(this.buttonSelectExcel_Click);
            // 
            // labelXmlPath
            // 
            this.labelXmlPath.AutoSize = true;
            this.labelXmlPath.Location = new System.Drawing.Point(6, 72);
            this.labelXmlPath.Name = "labelXmlPath";
            this.labelXmlPath.Size = new System.Drawing.Size(107, 13);
            this.labelXmlPath.TabIndex = 0;
            this.labelXmlPath.Text = "Destination XML File:";
            // 
            // labelExcelPath
            // 
            this.labelExcelPath.AutoSize = true;
            this.labelExcelPath.Location = new System.Drawing.Point(6, 20);
            this.labelExcelPath.Name = "labelExcelPath";
            this.labelExcelPath.Size = new System.Drawing.Size(92, 13);
            this.labelExcelPath.TabIndex = 0;
            this.labelExcelPath.Text = "Source Excel File:";
            // 
            // textBoxXmlPath
            // 
            this.textBoxXmlPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxXmlPath.BackColor = System.Drawing.Color.LightGray;
            this.textBoxXmlPath.Location = new System.Drawing.Point(6, 88);
            this.textBoxXmlPath.Name = "textBoxXmlPath";
            this.textBoxXmlPath.ReadOnly = true;
            this.textBoxXmlPath.Size = new System.Drawing.Size(424, 20);
            this.textBoxXmlPath.TabIndex = 0;
            // 
            // textBoxExcelPath
            // 
            this.textBoxExcelPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxExcelPath.BackColor = System.Drawing.Color.LightGray;
            this.textBoxExcelPath.Location = new System.Drawing.Point(6, 36);
            this.textBoxExcelPath.Name = "textBoxExcelPath";
            this.textBoxExcelPath.ReadOnly = true;
            this.textBoxExcelPath.Size = new System.Drawing.Size(424, 20);
            this.textBoxExcelPath.TabIndex = 0;
            // 
            // groupboxOperations
            // 
            this.groupboxOperations.Controls.Add(this.labelL2Details);
            this.groupboxOperations.Controls.Add(this.textBoxL2Details);
            this.groupboxOperations.Controls.Add(this.labelL1Details);
            this.groupboxOperations.Controls.Add(this.textBoxL1Details);
            this.groupboxOperations.Controls.Add(this.labelImportance);
            this.groupboxOperations.Controls.Add(this.pictureBox1);
            this.groupboxOperations.Controls.Add(this.labelL2Name);
            this.groupboxOperations.Controls.Add(this.labelL1Name);
            this.groupboxOperations.Controls.Add(this.labelEndRow);
            this.groupboxOperations.Controls.Add(this.labelStartRow);
            this.groupboxOperations.Controls.Add(this.labelActions);
            this.groupboxOperations.Controls.Add(this.labelExpectedResults);
            this.groupboxOperations.Controls.Add(this.labelPreconditions);
            this.groupboxOperations.Controls.Add(this.labelSummary);
            this.groupboxOperations.Controls.Add(this.labelCaseName);
            this.groupboxOperations.Controls.Add(this.textBoxSummary);
            this.groupboxOperations.Controls.Add(this.textBoxExpectedResult);
            this.groupboxOperations.Controls.Add(this.textBoxEndRow);
            this.groupboxOperations.Controls.Add(this.textBoxImportance);
            this.groupboxOperations.Controls.Add(this.textBoxActions);
            this.groupboxOperations.Controls.Add(this.textBoxCaseName);
            this.groupboxOperations.Controls.Add(this.textBoxStartRow);
            this.groupboxOperations.Controls.Add(this.textBoxPreconditions);
            this.groupboxOperations.Controls.Add(this.textBoxL2Name);
            this.groupboxOperations.Controls.Add(this.textBoxL1Name);
            this.groupboxOperations.Controls.Add(this.textBoxActiveSheet);
            this.groupboxOperations.Controls.Add(this.labelActiveSheet);
            this.groupboxOperations.Location = new System.Drawing.Point(6, 130);
            this.groupboxOperations.Name = "groupboxOperations";
            this.groupboxOperations.Size = new System.Drawing.Size(694, 227);
            this.groupboxOperations.TabIndex = 2;
            this.groupboxOperations.TabStop = false;
            this.groupboxOperations.Text = "Excel Mappings:";
            // 
            // labelImportance
            // 
            this.labelImportance.AutoSize = true;
            this.labelImportance.Location = new System.Drawing.Point(121, 178);
            this.labelImportance.Name = "labelImportance";
            this.labelImportance.Size = new System.Drawing.Size(60, 13);
            this.labelImportance.TabIndex = 13;
            this.labelImportance.Text = "Importance";
            this.labelImportance.Click += new System.EventHandler(this.labelImportance_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::EX_Converter.Properties.Resources.TestLink_logo;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(596, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(70, 38);
            this.pictureBox1.TabIndex = 16;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // labelL2Name
            // 
            this.labelL2Name.AutoSize = true;
            this.labelL2Name.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelL2Name.Location = new System.Drawing.Point(8, 124);
            this.labelL2Name.Name = "labelL2Name";
            this.labelL2Name.Size = new System.Drawing.Size(96, 12);
            this.labelL2Name.TabIndex = 11;
            this.labelL2Name.Text = "L2 Suite Name";
            this.labelL2Name.Click += new System.EventHandler(this.label_L2_Click);
            // 
            // labelL1Name
            // 
            this.labelL1Name.AutoSize = true;
            this.labelL1Name.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelL1Name.Location = new System.Drawing.Point(8, 71);
            this.labelL1Name.Name = "labelL1Name";
            this.labelL1Name.Size = new System.Drawing.Size(96, 12);
            this.labelL1Name.TabIndex = 10;
            this.labelL1Name.Text = "L1 Suite Name";
            this.labelL1Name.Click += new System.EventHandler(this.label_L1_Click);
            // 
            // labelEndRow
            // 
            this.labelEndRow.AutoSize = true;
            this.labelEndRow.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelEndRow.Location = new System.Drawing.Point(233, 20);
            this.labelEndRow.Name = "labelEndRow";
            this.labelEndRow.Size = new System.Drawing.Size(54, 12);
            this.labelEndRow.TabIndex = 9;
            this.labelEndRow.Text = "End Row";
            // 
            // labelStartRow
            // 
            this.labelStartRow.AutoSize = true;
            this.labelStartRow.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelStartRow.Location = new System.Drawing.Point(119, 20);
            this.labelStartRow.Name = "labelStartRow";
            this.labelStartRow.Size = new System.Drawing.Size(68, 12);
            this.labelStartRow.TabIndex = 8;
            this.labelStartRow.Text = "Start Row";
            // 
            // labelActions
            // 
            this.labelActions.AutoSize = true;
            this.labelActions.Location = new System.Drawing.Point(464, 178);
            this.labelActions.Name = "labelActions";
            this.labelActions.Size = new System.Drawing.Size(42, 13);
            this.labelActions.TabIndex = 7;
            this.labelActions.Text = "Actions";
            this.labelActions.Click += new System.EventHandler(this.labelActions_Click);
            // 
            // labelExpectedResults
            // 
            this.labelExpectedResults.AutoSize = true;
            this.labelExpectedResults.Location = new System.Drawing.Point(576, 178);
            this.labelExpectedResults.Name = "labelExpectedResults";
            this.labelExpectedResults.Size = new System.Drawing.Size(90, 13);
            this.labelExpectedResults.TabIndex = 6;
            this.labelExpectedResults.Text = "Expected Results";
            this.labelExpectedResults.Click += new System.EventHandler(this.labelExpectedResults_Click);
            // 
            // labelPreconditions
            // 
            this.labelPreconditions.AutoSize = true;
            this.labelPreconditions.Location = new System.Drawing.Point(350, 178);
            this.labelPreconditions.Name = "labelPreconditions";
            this.labelPreconditions.Size = new System.Drawing.Size(71, 13);
            this.labelPreconditions.TabIndex = 5;
            this.labelPreconditions.Text = "Preconditions";
            this.labelPreconditions.Click += new System.EventHandler(this.labelPreconditions_Click);
            // 
            // labelSummary
            // 
            this.labelSummary.AutoSize = true;
            this.labelSummary.Location = new System.Drawing.Point(235, 178);
            this.labelSummary.Name = "labelSummary";
            this.labelSummary.Size = new System.Drawing.Size(50, 13);
            this.labelSummary.TabIndex = 4;
            this.labelSummary.Text = "Summary";
            this.labelSummary.Click += new System.EventHandler(this.labelSummary_Click);
            // 
            // labelCaseName
            // 
            this.labelCaseName.AutoSize = true;
            this.labelCaseName.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelCaseName.Location = new System.Drawing.Point(8, 178);
            this.labelCaseName.Name = "labelCaseName";
            this.labelCaseName.Size = new System.Drawing.Size(68, 12);
            this.labelCaseName.TabIndex = 3;
            this.labelCaseName.Text = "Case Name";
            this.labelCaseName.Click += new System.EventHandler(this.labelCaseName_Click);
            // 
            // textBoxSummary
            // 
            this.textBoxSummary.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxSummary.Location = new System.Drawing.Point(237, 194);
            this.textBoxSummary.Name = "textBoxSummary";
            this.textBoxSummary.Size = new System.Drawing.Size(100, 20);
            this.textBoxSummary.TabIndex = 10;
            this.toolTip_5s.SetToolTip(this.textBoxSummary, "Numbers or single alphabetical character.");
            this.textBoxSummary.TextChanged += new System.EventHandler(this.textBoxSummary_TextChanged);
            // 
            // textBoxExpectedResult
            // 
            this.textBoxExpectedResult.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxExpectedResult.Location = new System.Drawing.Point(577, 194);
            this.textBoxExpectedResult.Name = "textBoxExpectedResult";
            this.textBoxExpectedResult.Size = new System.Drawing.Size(100, 20);
            this.textBoxExpectedResult.TabIndex = 13;
            this.toolTip_5s.SetToolTip(this.textBoxExpectedResult, "Numbers or single alphabetical character.");
            // 
            // textBoxEndRow
            // 
            this.textBoxEndRow.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxEndRow.Location = new System.Drawing.Point(235, 36);
            this.textBoxEndRow.Name = "textBoxEndRow";
            this.textBoxEndRow.Size = new System.Drawing.Size(100, 20);
            this.textBoxEndRow.TabIndex = 5;
            this.toolTip_5s.SetToolTip(this.textBoxEndRow, "Numbers only.");
            // 
            // textBoxImportance
            // 
            this.textBoxImportance.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxImportance.Location = new System.Drawing.Point(121, 194);
            this.textBoxImportance.Name = "textBoxImportance";
            this.textBoxImportance.Size = new System.Drawing.Size(100, 20);
            this.textBoxImportance.TabIndex = 9;
            this.toolTip_5s.SetToolTip(this.textBoxImportance, "Numbers or single alphabetical character.");
            this.textBoxImportance.TextChanged += new System.EventHandler(this.textBoxImportance_TextChanged);
            // 
            // textBoxActions
            // 
            this.textBoxActions.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxActions.Location = new System.Drawing.Point(464, 194);
            this.textBoxActions.Name = "textBoxActions";
            this.textBoxActions.Size = new System.Drawing.Size(100, 20);
            this.textBoxActions.TabIndex = 12;
            this.toolTip_5s.SetToolTip(this.textBoxActions, "Numbers or single alphabetical character.");
            this.textBoxActions.TextChanged += new System.EventHandler(this.textBoxActions_TextChanged);
            // 
            // textBoxCaseName
            // 
            this.textBoxCaseName.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxCaseName.Location = new System.Drawing.Point(10, 194);
            this.textBoxCaseName.Name = "textBoxCaseName";
            this.textBoxCaseName.Size = new System.Drawing.Size(100, 20);
            this.textBoxCaseName.TabIndex = 8;
            this.toolTip_5s.SetToolTip(this.textBoxCaseName, "Numbers or single alphabetical character.");
            this.textBoxCaseName.TextChanged += new System.EventHandler(this.textBoxCaseName_TextChanged);
            // 
            // textBoxStartRow
            // 
            this.textBoxStartRow.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxStartRow.Location = new System.Drawing.Point(121, 36);
            this.textBoxStartRow.Name = "textBoxStartRow";
            this.textBoxStartRow.Size = new System.Drawing.Size(100, 20);
            this.textBoxStartRow.TabIndex = 4;
            this.toolTip_5s.SetToolTip(this.textBoxStartRow, "Numbers only.");
            // 
            // textBoxPreconditions
            // 
            this.textBoxPreconditions.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxPreconditions.Location = new System.Drawing.Point(350, 194);
            this.textBoxPreconditions.Name = "textBoxPreconditions";
            this.textBoxPreconditions.Size = new System.Drawing.Size(100, 20);
            this.textBoxPreconditions.TabIndex = 11;
            this.toolTip_5s.SetToolTip(this.textBoxPreconditions, "Numbers or single alphabetical character.");
            this.textBoxPreconditions.TextChanged += new System.EventHandler(this.textBoxPreconditions_TextChanged);
            // 
            // textBoxL2Name
            // 
            this.textBoxL2Name.Location = new System.Drawing.Point(10, 139);
            this.textBoxL2Name.Name = "textBoxL2Name";
            this.textBoxL2Name.Size = new System.Drawing.Size(100, 20);
            this.textBoxL2Name.TabIndex = 7;
            this.toolTip_5s.SetToolTip(this.textBoxL2Name, "Numbers or single alphabetical character.");
            // 
            // textBoxL1Name
            // 
            this.textBoxL1Name.Location = new System.Drawing.Point(10, 86);
            this.textBoxL1Name.Name = "textBoxL1Name";
            this.textBoxL1Name.Size = new System.Drawing.Size(100, 20);
            this.textBoxL1Name.TabIndex = 6;
            this.toolTip_5s.SetToolTip(this.textBoxL1Name, "Numbers or single alphabetical character.");
            // 
            // textBoxActiveSheet
            // 
            this.textBoxActiveSheet.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxActiveSheet.Location = new System.Drawing.Point(10, 36);
            this.textBoxActiveSheet.Name = "textBoxActiveSheet";
            this.textBoxActiveSheet.Size = new System.Drawing.Size(100, 20);
            this.textBoxActiveSheet.TabIndex = 3;
            this.toolTip_5s.SetToolTip(this.textBoxActiveSheet, "Numbers only.");
            // 
            // labelActiveSheet
            // 
            this.labelActiveSheet.AutoSize = true;
            this.labelActiveSheet.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelActiveSheet.Location = new System.Drawing.Point(7, 20);
            this.labelActiveSheet.Name = "labelActiveSheet";
            this.labelActiveSheet.Size = new System.Drawing.Size(89, 12);
            this.labelActiveSheet.TabIndex = 0;
            this.labelActiveSheet.Text = "Active Sheet";
            // 
            // buttonConvert
            // 
            this.buttonConvert.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonConvert.Location = new System.Drawing.Point(322, 369);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(85, 30);
            this.buttonConvert.TabIndex = 13;
            this.buttonConvert.Text = "Convert!";
            this.buttonConvert.UseVisualStyleBackColor = true;
            this.buttonConvert.Click += new System.EventHandler(this.buttonConvert_Click);
            // 
            // logWindow
            // 
            this.logWindow.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logWindow.BackColor = System.Drawing.Color.White;
            this.logWindow.Location = new System.Drawing.Point(6, 422);
            this.logWindow.Name = "logWindow";
            this.logWindow.ReadOnly = true;
            this.logWindow.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
            this.logWindow.Size = new System.Drawing.Size(694, 178);
            this.logWindow.TabIndex = 3;
            this.logWindow.Text = "";
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // ProgressBar
            // 
            this.ProgressBar.Location = new System.Drawing.Point(6, 403);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(695, 16);
            this.ProgressBar.TabIndex = 4;
            // 
            // ButtonClear
            // 
            this.ButtonClear.Location = new System.Drawing.Point(614, 372);
            this.ButtonClear.Name = "ButtonClear";
            this.ButtonClear.Size = new System.Drawing.Size(85, 25);
            this.ButtonClear.TabIndex = 15;
            this.ButtonClear.Text = "Clear Logs";
            this.ButtonClear.UseVisualStyleBackColor = true;
            this.ButtonClear.Click += new System.EventHandler(this.ButtonClear_Click);
            // 
            // ButtonReset
            // 
            this.ButtonReset.Location = new System.Drawing.Point(509, 372);
            this.ButtonReset.Name = "ButtonReset";
            this.ButtonReset.Size = new System.Drawing.Size(85, 25);
            this.ButtonReset.TabIndex = 14;
            this.ButtonReset.Text = "Reset Values";
            this.ButtonReset.UseVisualStyleBackColor = true;
            this.ButtonReset.Click += new System.EventHandler(this.ButtonReset_Click);
            // 
            // toolTip_12s
            // 
            this.toolTip_12s.AutoPopDelay = 12000;
            this.toolTip_12s.InitialDelay = 500;
            this.toolTip_12s.ReshowDelay = 0;
            this.toolTip_12s.ShowAlways = true;
            // 
            // toolTip_5s
            // 
            this.toolTip_5s.AutoPopDelay = 5000;
            this.toolTip_5s.InitialDelay = 500;
            this.toolTip_5s.ReshowDelay = 0;
            this.toolTip_5s.ShowAlways = true;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel1.Location = new System.Drawing.Point(2, 3);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(95, 13);
            this.linkLabel1.TabIndex = 17;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "TestLink Web Site";
            this.toolTip_5s.SetToolTip(this.linkLabel1, "http://www.teamst.org/");
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.linkLabel2.Location = new System.Drawing.Point(3, 21);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(86, 13);
            this.linkLabel2.TabIndex = 18;
            this.linkLabel2.TabStop = true;
            this.linkLabel2.Text = "SourceForge.net";
            this.toolTip_5s.SetToolTip(this.linkLabel2, "http://sourceforge.net/");
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(105, 10);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(101, 13);
            this.label12.TabIndex = 19;
            this.label12.Text = "About EX-Converter";
            this.label12.Click += new System.EventHandler(this.label12_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.Gold;
            this.panel1.Controls.Add(this.linkLabel2);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.linkLabel1);
            this.panel1.Location = new System.Drawing.Point(6, 363);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(212, 38);
            this.panel1.TabIndex = 20;
            // 
            // labelL1Details
            // 
            this.labelL1Details.AutoSize = true;
            this.labelL1Details.Location = new System.Drawing.Point(121, 70);
            this.labelL1Details.Name = "labelL1Details";
            this.labelL1Details.Size = new System.Drawing.Size(81, 13);
            this.labelL1Details.TabIndex = 18;
            this.labelL1Details.Text = "L1 Suite Details";
            this.labelL1Details.Click += new System.EventHandler(this.label1_Click);
            // 
            // textBoxL1Details
            // 
            this.textBoxL1Details.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxL1Details.Location = new System.Drawing.Point(121, 86);
            this.textBoxL1Details.Name = "textBoxL1Details";
            this.textBoxL1Details.Size = new System.Drawing.Size(100, 20);
            this.textBoxL1Details.TabIndex = 17;
            this.toolTip_5s.SetToolTip(this.textBoxL1Details, "Numbers or single alphabetical character.");
            this.textBoxL1Details.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // labelL2Details
            // 
            this.labelL2Details.AutoSize = true;
            this.labelL2Details.Location = new System.Drawing.Point(121, 123);
            this.labelL2Details.Name = "labelL2Details";
            this.labelL2Details.Size = new System.Drawing.Size(81, 13);
            this.labelL2Details.TabIndex = 20;
            this.labelL2Details.Text = "L2 Suite Details";
            // 
            // textBoxL2Details
            // 
            this.textBoxL2Details.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxL2Details.Location = new System.Drawing.Point(121, 139);
            this.textBoxL2Details.Name = "textBoxL2Details";
            this.textBoxL2Details.Size = new System.Drawing.Size(100, 20);
            this.textBoxL2Details.TabIndex = 19;
            this.toolTip_5s.SetToolTip(this.textBoxL2Details, "Numbers or single alphabetical character.");
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 609);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.ButtonReset);
            this.Controls.Add(this.ButtonClear);
            this.Controls.Add(this.TemplateTypeGroupBox);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.logWindow);
            this.Controls.Add(this.groupboxOperations);
            this.Controls.Add(this.groupBoxSelectPath);
            this.MaximumSize = new System.Drawing.Size(720, 647);
            this.MinimumSize = new System.Drawing.Size(720, 647);
            this.Name = "MainWindow";
            this.Text = "EX-Converter v1.2.1";
            this.TemplateTypeGroupBox.ResumeLayout(false);
            this.TemplateTypeGroupBox.PerformLayout();
            this.groupBoxSelectPath.ResumeLayout(false);
            this.groupBoxSelectPath.PerformLayout();
            this.groupboxOperations.ResumeLayout(false);
            this.groupboxOperations.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxSelectPath;
        private System.Windows.Forms.Button buttonSelectXml;
        private System.Windows.Forms.Button buttonSelectExcel;
        private System.Windows.Forms.Label labelXmlPath;
        private System.Windows.Forms.Label labelExcelPath;
        private System.Windows.Forms.TextBox textBoxXmlPath;
        private System.Windows.Forms.TextBox textBoxExcelPath;
        private System.Windows.Forms.GroupBox groupboxOperations;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.RichTextBox logWindow;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Label labelActiveSheet;
        private System.Windows.Forms.TextBox textBoxActiveSheet;
        private System.Windows.Forms.Label labelCaseName;
        private System.Windows.Forms.TextBox textBoxCaseName;
        private System.Windows.Forms.Label labelExpectedResults;
        private System.Windows.Forms.Label labelPreconditions;
        private System.Windows.Forms.Label labelSummary;
        private System.Windows.Forms.TextBox textBoxSummary;
        private System.Windows.Forms.TextBox textBoxExpectedResult;
        private System.Windows.Forms.TextBox textBoxPreconditions;
        private System.Windows.Forms.Label labelActions;
        private System.Windows.Forms.TextBox textBoxActions;
        private System.Windows.Forms.ProgressBar ProgressBar;
        private System.Windows.Forms.GroupBox TemplateTypeGroupBox;
        private System.Windows.Forms.RadioButton RadioButtonSuite;
        private System.Windows.Forms.RadioButton RadioButtonCases;
        private System.Windows.Forms.Label labelEndRow;
        private System.Windows.Forms.Label labelStartRow;
        private System.Windows.Forms.TextBox textBoxEndRow;
        private System.Windows.Forms.TextBox textBoxStartRow;
        private System.Windows.Forms.Label labelL1Name;
        private System.Windows.Forms.Label labelL2Name;
        private System.Windows.Forms.TextBox textBoxL1Name;
        private System.Windows.Forms.TextBox textBoxL2Name;
        private System.Windows.Forms.CheckBox CheckBoxEnableL2;
        private System.Windows.Forms.CheckBox CheckBoxAllowDupSuite;
        private System.Windows.Forms.Button ButtonClear;
        private System.Windows.Forms.Button ButtonReset;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelImportance;
        private System.Windows.Forms.TextBox textBoxImportance;
        private System.Windows.Forms.ToolTip toolTip_12s;
        private System.Windows.Forms.ToolTip toolTip_5s;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label labelL1Details;
        private System.Windows.Forms.TextBox textBoxL1Details;
        private System.Windows.Forms.Label labelL2Details;
        private System.Windows.Forms.TextBox textBoxL2Details;
    }
}

