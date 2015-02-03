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
            this.checkBoxAllowDupSuite = new System.Windows.Forms.CheckBox();
            this.CheckBoxEnableL2 = new System.Windows.Forms.CheckBox();
            this.RadioButtonSuite = new System.Windows.Forms.RadioButton();
            this.RadioButtonCases = new System.Windows.Forms.RadioButton();
            this.groupBoxSelectPath = new System.Windows.Forms.GroupBox();
            this.buttonSelectXml = new System.Windows.Forms.Button();
            this.buttonSelectExcel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxXmlPath = new System.Windows.Forms.TextBox();
            this.textBoxExcelPath = new System.Windows.Forms.TextBox();
            this.groupboxOperations = new System.Windows.Forms.GroupBox();
            this.label11 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label_L2 = new System.Windows.Forms.Label();
            this.label_L1 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxSummary = new System.Windows.Forms.TextBox();
            this.textBoxExpectedResult = new System.Windows.Forms.TextBox();
            this.textBoxEndRow = new System.Windows.Forms.TextBox();
            this.textBoxImportance = new System.Windows.Forms.TextBox();
            this.textBoxActions = new System.Windows.Forms.TextBox();
            this.textBoxName = new System.Windows.Forms.TextBox();
            this.textBoxStartRow = new System.Windows.Forms.TextBox();
            this.textBoxPreconditions = new System.Windows.Forms.TextBox();
            this.textBoxLevel_2 = new System.Windows.Forms.TextBox();
            this.textBoxlevel_1 = new System.Windows.Forms.TextBox();
            this.textBoxActiveSheet = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
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
            this.TemplateTypeGroupBox.SuspendLayout();
            this.groupBoxSelectPath.SuspendLayout();
            this.groupboxOperations.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TemplateTypeGroupBox
            // 
            this.TemplateTypeGroupBox.Controls.Add(this.checkBoxAllowDupSuite);
            this.TemplateTypeGroupBox.Controls.Add(this.CheckBoxEnableL2);
            this.TemplateTypeGroupBox.Controls.Add(this.RadioButtonSuite);
            this.TemplateTypeGroupBox.Controls.Add(this.RadioButtonCases);
            this.TemplateTypeGroupBox.Location = new System.Drawing.Point(6, 5);
            this.TemplateTypeGroupBox.Name = "TemplateTypeGroupBox";
            this.TemplateTypeGroupBox.Size = new System.Drawing.Size(210, 113);
            this.TemplateTypeGroupBox.TabIndex = 0;
            this.TemplateTypeGroupBox.TabStop = false;
            this.TemplateTypeGroupBox.Text = "Template Type";
            // 
            // checkBoxAllowDupSuite
            // 
            this.checkBoxAllowDupSuite.AutoSize = true;
            this.checkBoxAllowDupSuite.Location = new System.Drawing.Point(22, 78);
            this.checkBoxAllowDupSuite.Name = "checkBoxAllowDupSuite";
            this.checkBoxAllowDupSuite.Size = new System.Drawing.Size(180, 16);
            this.checkBoxAllowDupSuite.TabIndex = 0;
            this.checkBoxAllowDupSuite.Text = "Allow Duplicate Suite Name";
            this.toolTip_12s.SetToolTip(this.checkBoxAllowDupSuite, resources.GetString("checkBoxAllowDupSuite.ToolTip"));
            this.checkBoxAllowDupSuite.UseVisualStyleBackColor = true;
            this.checkBoxAllowDupSuite.CheckedChanged += new System.EventHandler(this.checkBoxAllowDupSuite_CheckedChanged);
            // 
            // CheckBoxEnableL2
            // 
            this.CheckBoxEnableL2.AutoSize = true;
            this.CheckBoxEnableL2.Location = new System.Drawing.Point(22, 59);
            this.CheckBoxEnableL2.Name = "CheckBoxEnableL2";
            this.CheckBoxEnableL2.Size = new System.Drawing.Size(150, 16);
            this.CheckBoxEnableL2.TabIndex = 0;
            this.CheckBoxEnableL2.Text = "Enable Level 2 Folder";
            this.toolTip_12s.SetToolTip(this.CheckBoxEnableL2, "Check to enable level 2 folders when generating a test specification.\r\nYou need t" +
                    "o map an Excel sheet column to level 2 folders.");
            this.CheckBoxEnableL2.UseVisualStyleBackColor = true;
            this.CheckBoxEnableL2.CheckedChanged += new System.EventHandler(this.CheckBoxEnableL2_CheckedChanged);
            // 
            // RadioButtonSuite
            // 
            this.RadioButtonSuite.AutoSize = true;
            this.RadioButtonSuite.Location = new System.Drawing.Point(9, 39);
            this.RadioButtonSuite.Name = "RadioButtonSuite";
            this.RadioButtonSuite.Size = new System.Drawing.Size(83, 16);
            this.RadioButtonSuite.TabIndex = 0;
            this.RadioButtonSuite.TabStop = true;
            this.RadioButtonSuite.Text = "Test Suite";
            this.toolTip_12s.SetToolTip(this.RadioButtonSuite, "Generate an XML test specification with type of \"Test Suite\".\r\n(Used by \"Improt T" +
                    "est Suite\" in TestLink.)");
            this.RadioButtonSuite.UseVisualStyleBackColor = true;
            this.RadioButtonSuite.CheckedChanged += new System.EventHandler(this.RadioButtonSuite_CheckedChanged);
            // 
            // RadioButtonCases
            // 
            this.RadioButtonCases.AutoSize = true;
            this.RadioButtonCases.Location = new System.Drawing.Point(9, 19);
            this.RadioButtonCases.Name = "RadioButtonCases";
            this.RadioButtonCases.Size = new System.Drawing.Size(83, 16);
            this.RadioButtonCases.TabIndex = 0;
            this.RadioButtonCases.TabStop = true;
            this.RadioButtonCases.Text = "Test Cases";
            this.toolTip_12s.SetToolTip(this.RadioButtonCases, "Generate an XML test specification with the type of \"Test Cases\".\r\n(Used by \"Impr" +
                    "ot Test Cases\" in TestLink.)");
            this.RadioButtonCases.UseVisualStyleBackColor = true;
            this.RadioButtonCases.CheckedChanged += new System.EventHandler(this.RadioButtonCases_CheckedChanged);
            // 
            // groupBoxSelectPath
            // 
            this.groupBoxSelectPath.Controls.Add(this.buttonSelectXml);
            this.groupBoxSelectPath.Controls.Add(this.buttonSelectExcel);
            this.groupBoxSelectPath.Controls.Add(this.label2);
            this.groupBoxSelectPath.Controls.Add(this.label1);
            this.groupBoxSelectPath.Controls.Add(this.textBoxXmlPath);
            this.groupBoxSelectPath.Controls.Add(this.textBoxExcelPath);
            this.groupBoxSelectPath.Location = new System.Drawing.Point(225, 5);
            this.groupBoxSelectPath.Name = "groupBoxSelectPath";
            this.groupBoxSelectPath.Size = new System.Drawing.Size(475, 112);
            this.groupBoxSelectPath.TabIndex = 0;
            this.groupBoxSelectPath.TabStop = false;
            this.groupBoxSelectPath.Text = "Select Files:";
            // 
            // buttonSelectXml
            // 
            this.buttonSelectXml.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectXml.Location = new System.Drawing.Point(436, 79);
            this.buttonSelectXml.Name = "buttonSelectXml";
            this.buttonSelectXml.Size = new System.Drawing.Size(31, 23);
            this.buttonSelectXml.TabIndex = 2;
            this.buttonSelectXml.Text = "...";
            this.buttonSelectXml.UseVisualStyleBackColor = true;
            this.buttonSelectXml.Click += new System.EventHandler(this.buttonSelectXml_Click);
            // 
            // buttonSelectExcel
            // 
            this.buttonSelectExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonSelectExcel.Location = new System.Drawing.Point(436, 31);
            this.buttonSelectExcel.Name = "buttonSelectExcel";
            this.buttonSelectExcel.Size = new System.Drawing.Size(31, 23);
            this.buttonSelectExcel.TabIndex = 1;
            this.buttonSelectExcel.Text = "...";
            this.buttonSelectExcel.UseVisualStyleBackColor = true;
            this.buttonSelectExcel.Click += new System.EventHandler(this.buttonSelectExcel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "Destination XML File:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(113, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Source Excel File:";
            // 
            // textBoxXmlPath
            // 
            this.textBoxXmlPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxXmlPath.BackColor = System.Drawing.Color.LightGray;
            this.textBoxXmlPath.Location = new System.Drawing.Point(6, 81);
            this.textBoxXmlPath.Name = "textBoxXmlPath";
            this.textBoxXmlPath.ReadOnly = true;
            this.textBoxXmlPath.Size = new System.Drawing.Size(424, 21);
            this.textBoxXmlPath.TabIndex = 0;
            // 
            // textBoxExcelPath
            // 
            this.textBoxExcelPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxExcelPath.BackColor = System.Drawing.Color.LightGray;
            this.textBoxExcelPath.Location = new System.Drawing.Point(6, 33);
            this.textBoxExcelPath.Name = "textBoxExcelPath";
            this.textBoxExcelPath.ReadOnly = true;
            this.textBoxExcelPath.Size = new System.Drawing.Size(424, 21);
            this.textBoxExcelPath.TabIndex = 0;
            // 
            // groupboxOperations
            // 
            this.groupboxOperations.Controls.Add(this.label11);
            this.groupboxOperations.Controls.Add(this.pictureBox1);
            this.groupboxOperations.Controls.Add(this.label_L2);
            this.groupboxOperations.Controls.Add(this.label_L1);
            this.groupboxOperations.Controls.Add(this.label10);
            this.groupboxOperations.Controls.Add(this.label9);
            this.groupboxOperations.Controls.Add(this.label8);
            this.groupboxOperations.Controls.Add(this.label7);
            this.groupboxOperations.Controls.Add(this.label6);
            this.groupboxOperations.Controls.Add(this.label5);
            this.groupboxOperations.Controls.Add(this.label4);
            this.groupboxOperations.Controls.Add(this.textBoxSummary);
            this.groupboxOperations.Controls.Add(this.textBoxExpectedResult);
            this.groupboxOperations.Controls.Add(this.textBoxEndRow);
            this.groupboxOperations.Controls.Add(this.textBoxImportance);
            this.groupboxOperations.Controls.Add(this.textBoxActions);
            this.groupboxOperations.Controls.Add(this.textBoxName);
            this.groupboxOperations.Controls.Add(this.textBoxStartRow);
            this.groupboxOperations.Controls.Add(this.textBoxPreconditions);
            this.groupboxOperations.Controls.Add(this.textBoxLevel_2);
            this.groupboxOperations.Controls.Add(this.textBoxlevel_1);
            this.groupboxOperations.Controls.Add(this.textBoxActiveSheet);
            this.groupboxOperations.Controls.Add(this.label3);
            this.groupboxOperations.Location = new System.Drawing.Point(6, 120);
            this.groupboxOperations.Name = "groupboxOperations";
            this.groupboxOperations.Size = new System.Drawing.Size(694, 108);
            this.groupboxOperations.TabIndex = 2;
            this.groupboxOperations.TabStop = false;
            this.groupboxOperations.Text = "Excel Mappings:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(121, 61);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(65, 12);
            this.label11.TabIndex = 13;
            this.label11.Text = "Importance";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::EX_Converter.Properties.Resources.TestLink_logo;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(596, 18);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(70, 35);
            this.pictureBox1.TabIndex = 16;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Visible = false;
            // 
            // label_L2
            // 
            this.label_L2.AutoSize = true;
            this.label_L2.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label_L2.Location = new System.Drawing.Point(465, 17);
            this.label_L2.Name = "label_L2";
            this.label_L2.Size = new System.Drawing.Size(103, 12);
            this.label_L2.TabIndex = 11;
            this.label_L2.Text = "Level 2 Folder";
            // 
            // label_L1
            // 
            this.label_L1.AutoSize = true;
            this.label_L1.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label_L1.Location = new System.Drawing.Point(347, 18);
            this.label_L1.Name = "label_L1";
            this.label_L1.Size = new System.Drawing.Size(103, 12);
            this.label_L1.TabIndex = 10;
            this.label_L1.Text = "Level 1 Folder";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(234, 18);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(54, 12);
            this.label10.TabIndex = 9;
            this.label10.Text = "End Row";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(120, 18);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(68, 12);
            this.label9.TabIndex = 8;
            this.label9.Text = "Start Row";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(464, 61);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(47, 12);
            this.label8.TabIndex = 7;
            this.label8.Text = "Actions";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(576, 61);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(101, 12);
            this.label7.TabIndex = 6;
            this.label7.Text = "Expected Results";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(350, 61);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(83, 12);
            this.label6.TabIndex = 5;
            this.label6.Text = "Preconditions";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(235, 61);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 12);
            this.label5.TabIndex = 4;
            this.label5.Text = "Summary";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(8, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(33, 12);
            this.label4.TabIndex = 3;
            this.label4.Text = "Name";
            // 
            // textBoxSummary
            // 
            this.textBoxSummary.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxSummary.Location = new System.Drawing.Point(237, 76);
            this.textBoxSummary.Name = "textBoxSummary";
            this.textBoxSummary.Size = new System.Drawing.Size(100, 21);
            this.textBoxSummary.TabIndex = 10;
            this.toolTip_5s.SetToolTip(this.textBoxSummary, "Numbers or single alphabetical character.");
            // 
            // textBoxExpectedResult
            // 
            this.textBoxExpectedResult.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxExpectedResult.Location = new System.Drawing.Point(577, 76);
            this.textBoxExpectedResult.Name = "textBoxExpectedResult";
            this.textBoxExpectedResult.Size = new System.Drawing.Size(100, 21);
            this.textBoxExpectedResult.TabIndex = 13;
            this.toolTip_5s.SetToolTip(this.textBoxExpectedResult, "Numbers or single alphabetical character.");
            // 
            // textBoxEndRow
            // 
            this.textBoxEndRow.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxEndRow.Location = new System.Drawing.Point(235, 33);
            this.textBoxEndRow.Name = "textBoxEndRow";
            this.textBoxEndRow.Size = new System.Drawing.Size(100, 21);
            this.textBoxEndRow.TabIndex = 5;
            this.toolTip_5s.SetToolTip(this.textBoxEndRow, "Numbers only.");
            // 
            // textBoxImportance
            // 
            this.textBoxImportance.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxImportance.Location = new System.Drawing.Point(121, 76);
            this.textBoxImportance.Name = "textBoxImportance";
            this.textBoxImportance.Size = new System.Drawing.Size(100, 21);
            this.textBoxImportance.TabIndex = 9;
            this.toolTip_5s.SetToolTip(this.textBoxImportance, "Numbers or single alphabetical character.");
            // 
            // textBoxActions
            // 
            this.textBoxActions.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxActions.Location = new System.Drawing.Point(464, 76);
            this.textBoxActions.Name = "textBoxActions";
            this.textBoxActions.Size = new System.Drawing.Size(100, 21);
            this.textBoxActions.TabIndex = 12;
            this.toolTip_5s.SetToolTip(this.textBoxActions, "Numbers or single alphabetical character.");
            // 
            // textBoxName
            // 
            this.textBoxName.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxName.Location = new System.Drawing.Point(10, 76);
            this.textBoxName.Name = "textBoxName";
            this.textBoxName.Size = new System.Drawing.Size(100, 21);
            this.textBoxName.TabIndex = 8;
            this.toolTip_5s.SetToolTip(this.textBoxName, "Numbers or single alphabetical character.");
            // 
            // textBoxStartRow
            // 
            this.textBoxStartRow.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxStartRow.Location = new System.Drawing.Point(121, 33);
            this.textBoxStartRow.Name = "textBoxStartRow";
            this.textBoxStartRow.Size = new System.Drawing.Size(100, 21);
            this.textBoxStartRow.TabIndex = 4;
            this.toolTip_5s.SetToolTip(this.textBoxStartRow, "Numbers only.");
            // 
            // textBoxPreconditions
            // 
            this.textBoxPreconditions.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxPreconditions.Location = new System.Drawing.Point(350, 76);
            this.textBoxPreconditions.Name = "textBoxPreconditions";
            this.textBoxPreconditions.Size = new System.Drawing.Size(100, 21);
            this.textBoxPreconditions.TabIndex = 11;
            this.toolTip_5s.SetToolTip(this.textBoxPreconditions, "Numbers or single alphabetical character.");
            // 
            // textBoxLevel_2
            // 
            this.textBoxLevel_2.Location = new System.Drawing.Point(464, 33);
            this.textBoxLevel_2.Name = "textBoxLevel_2";
            this.textBoxLevel_2.Size = new System.Drawing.Size(100, 21);
            this.textBoxLevel_2.TabIndex = 7;
            this.toolTip_5s.SetToolTip(this.textBoxLevel_2, "Numbers or single alphabetical character.");
            // 
            // textBoxlevel_1
            // 
            this.textBoxlevel_1.Location = new System.Drawing.Point(349, 33);
            this.textBoxlevel_1.Name = "textBoxlevel_1";
            this.textBoxlevel_1.Size = new System.Drawing.Size(100, 21);
            this.textBoxlevel_1.TabIndex = 6;
            this.toolTip_5s.SetToolTip(this.textBoxlevel_1, "Numbers or single alphabetical character.");
            // 
            // textBoxActiveSheet
            // 
            this.textBoxActiveSheet.BackColor = System.Drawing.SystemColors.Window;
            this.textBoxActiveSheet.Location = new System.Drawing.Point(10, 33);
            this.textBoxActiveSheet.Name = "textBoxActiveSheet";
            this.textBoxActiveSheet.Size = new System.Drawing.Size(100, 21);
            this.textBoxActiveSheet.TabIndex = 3;
            this.toolTip_5s.SetToolTip(this.textBoxActiveSheet, "Numbers only.");
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(8, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "Active Sheet";
            // 
            // buttonConvert
            // 
            this.buttonConvert.Font = new System.Drawing.Font("SimSun", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonConvert.Location = new System.Drawing.Point(322, 234);
            this.buttonConvert.Name = "buttonConvert";
            this.buttonConvert.Size = new System.Drawing.Size(85, 28);
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
            this.logWindow.Location = new System.Drawing.Point(6, 287);
            this.logWindow.Name = "logWindow";
            this.logWindow.ReadOnly = true;
            this.logWindow.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
            this.logWindow.Size = new System.Drawing.Size(694, 267);
            this.logWindow.TabIndex = 3;
            this.logWindow.Text = "";
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // ProgressBar
            // 
            this.ProgressBar.Location = new System.Drawing.Point(6, 266);
            this.ProgressBar.Name = "ProgressBar";
            this.ProgressBar.Size = new System.Drawing.Size(695, 15);
            this.ProgressBar.TabIndex = 4;
            // 
            // ButtonClear
            // 
            this.ButtonClear.Location = new System.Drawing.Point(614, 237);
            this.ButtonClear.Name = "ButtonClear";
            this.ButtonClear.Size = new System.Drawing.Size(85, 23);
            this.ButtonClear.TabIndex = 15;
            this.ButtonClear.Text = "Clear Logs";
            this.ButtonClear.UseVisualStyleBackColor = true;
            this.ButtonClear.Click += new System.EventHandler(this.ButtonClear_Click);
            // 
            // ButtonReset
            // 
            this.ButtonReset.Location = new System.Drawing.Point(509, 237);
            this.ButtonReset.Name = "ButtonReset";
            this.ButtonReset.Size = new System.Drawing.Size(85, 23);
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
            this.linkLabel2.Location = new System.Drawing.Point(3, 19);
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
            this.label12.Location = new System.Drawing.Point(105, 9);
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
            this.panel1.Location = new System.Drawing.Point(6, 229);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(212, 35);
            this.panel1.TabIndex = 20;
            // 
            // MainWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(712, 566);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.ButtonReset);
            this.Controls.Add(this.ButtonClear);
            this.Controls.Add(this.TemplateTypeGroupBox);
            this.Controls.Add(this.ProgressBar);
            this.Controls.Add(this.buttonConvert);
            this.Controls.Add(this.logWindow);
            this.Controls.Add(this.groupboxOperations);
            this.Controls.Add(this.groupBoxSelectPath);
            this.MaximumSize = new System.Drawing.Size(720, 600);
            this.MinimumSize = new System.Drawing.Size(720, 600);
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
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxXmlPath;
        private System.Windows.Forms.TextBox textBoxExcelPath;
        private System.Windows.Forms.GroupBox groupboxOperations;
        private System.Windows.Forms.Button buttonConvert;
        private System.Windows.Forms.RichTextBox logWindow;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxActiveSheet;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxSummary;
        private System.Windows.Forms.TextBox textBoxExpectedResult;
        private System.Windows.Forms.TextBox textBoxPreconditions;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBoxActions;
        private System.Windows.Forms.ProgressBar ProgressBar;
        private System.Windows.Forms.GroupBox TemplateTypeGroupBox;
        private System.Windows.Forms.RadioButton RadioButtonSuite;
        private System.Windows.Forms.RadioButton RadioButtonCases;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textBoxEndRow;
        private System.Windows.Forms.TextBox textBoxStartRow;
        private System.Windows.Forms.Label label_L1;
        private System.Windows.Forms.TextBox textBoxLevel_2;
        private System.Windows.Forms.TextBox textBoxlevel_1;
        private System.Windows.Forms.Label label_L2;
        private System.Windows.Forms.CheckBox CheckBoxEnableL2;
        private System.Windows.Forms.CheckBox checkBoxAllowDupSuite;
        private System.Windows.Forms.Button ButtonClear;
        private System.Windows.Forms.Button ButtonReset;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBoxImportance;
        private System.Windows.Forms.ToolTip toolTip_12s;
        private System.Windows.Forms.ToolTip toolTip_5s;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Panel panel1;
    }
}

