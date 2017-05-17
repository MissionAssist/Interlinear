namespace Interlinear
{
    partial class Interlinear
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
            this.openLegacyFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveLegacyFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.WordsPerLine = new System.Windows.Forms.NumericUpDown();
            this.saveExcelFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.saveUnicodeFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.Wordcount = new System.Windows.Forms.ToolTip(this.components);
            this.chkCloseOnCompletion = new System.Windows.Forms.CheckBox();
            this.boxExtension = new System.Windows.Forms.ComboBox();
            this.UpdownFontSize = new System.Windows.Forms.NumericUpDown();
            this.updownThreshold = new System.Windows.Forms.NumericUpDown();
            this.UpdownInterval = new System.Windows.Forms.NumericUpDown();
            this.DebugCheckBox = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.Setup = new System.Windows.Forms.TabPage();
            this.BtnBothToExcel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.BtnSegmentBoth = new System.Windows.Forms.Button();
            this.grpUnicode = new System.Windows.Forms.GroupBox();
            this.chkUnicodeAddSpace = new System.Windows.Forms.CheckBox();
            this.BtnUnicodeToExcel = new System.Windows.Forms.Button();
            this.chkUnicodeToExcel = new System.Windows.Forms.CheckBox();
            this.txtUnicodeInput = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnGetUnicodeInput = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.txtUnicodeOutput = new System.Windows.Forms.TextBox();
            this.GetUnicodeOutput = new System.Windows.Forms.Button();
            this.txtUnicodeWordCount = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.BtnSegmentUnicode = new System.Windows.Forms.Button();
            this.grpLegacy = new System.Windows.Forms.GroupBox();
            this.chkLegacyAddSpace = new System.Windows.Forms.CheckBox();
            this.BtnLegacyToExcel = new System.Windows.Forms.Button();
            this.chkLegacyToExcel = new System.Windows.Forms.CheckBox();
            this.txtLegacyInput = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGetLegacyInputFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLegacyOutput = new System.Windows.Forms.TextBox();
            this.btnGetLegacyOutputFile = new System.Windows.Forms.Button();
            this.txtLegacyWordCount = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.BtnSegmentLegacy = new System.Windows.Forms.Button();
            this.btnGetExcelOutput = new System.Windows.Forms.Button();
            this.txtExcelOutput = new System.Windows.Forms.TextBox();
            this.Progress = new System.Windows.Forms.TabPage();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.boxProgress = new System.Windows.Forms.ListBox();
            this.openUnicodeFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.BtnClose = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.documentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.licenseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.BtnPauseResume = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UpdownFontSize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.updownThreshold)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UpdownInterval)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.Setup.SuspendLayout();
            this.grpUnicode.SuspendLayout();
            this.grpLegacy.SuspendLayout();
            this.Progress.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openLegacyFileDialog
            // 
            this.openLegacyFileDialog.DefaultExt = "docx";
            this.openLegacyFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx|RTF files|*.rtf|Text files|*.txt|" +
    "OpenDocument Text|*.odt";
            this.openLegacyFileDialog.FilterIndex = 2;
            // 
            // saveLegacyFileDialog
            // 
            this.saveLegacyFileDialog.DefaultExt = "docx";
            this.saveLegacyFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            this.saveLegacyFileDialog.FilterIndex = 2;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 66);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Words per line";
            // 
            // WordsPerLine
            // 
            this.WordsPerLine.Increment = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.WordsPerLine.Location = new System.Drawing.Point(96, 63);
            this.WordsPerLine.Maximum = new decimal(new int[] {
            20,
            0,
            0,
            0});
            this.WordsPerLine.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.WordsPerLine.Name = "WordsPerLine";
            this.WordsPerLine.Size = new System.Drawing.Size(37, 20);
            this.WordsPerLine.TabIndex = 14;
            this.WordsPerLine.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.Wordcount.SetToolTip(this.WordsPerLine, "The number of words (i.e. text separated by spaces) on each line.");
            this.WordsPerLine.Value = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.WordsPerLine.ValueChanged += new System.EventHandler(this.WordsPerLine_ValueChanged);
            // 
            // saveExcelFileDialog
            // 
            this.saveExcelFileDialog.DefaultExt = "xlsx";
            this.saveExcelFileDialog.Filter = "Excel WorkBook | *.xlsx";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(6, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(698, 37);
            this.label3.TabIndex = 23;
            this.label3.Text = "Interlinear Comparison";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // saveUnicodeFileDialog
            // 
            this.saveUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx|RTF files|*.rtf";
            this.saveUnicodeFileDialog.FilterIndex = 2;
            // 
            // chkCloseOnCompletion
            // 
            this.chkCloseOnCompletion.AutoSize = true;
            this.chkCloseOnCompletion.Location = new System.Drawing.Point(581, 330);
            this.chkCloseOnCompletion.Name = "chkCloseOnCompletion";
            this.chkCloseOnCompletion.Size = new System.Drawing.Size(121, 17);
            this.chkCloseOnCompletion.TabIndex = 37;
            this.chkCloseOnCompletion.Text = "Close on completion";
            this.chkCloseOnCompletion.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.Wordcount.SetToolTip(this.chkCloseOnCompletion, "This closes the application when it has finished doing what you have asked it to " +
        "do.");
            this.chkCloseOnCompletion.UseVisualStyleBackColor = true;
            // 
            // boxExtension
            // 
            this.boxExtension.DisplayMember = "Text";
            this.boxExtension.FormattingEnabled = true;
            this.boxExtension.Items.AddRange(new object[] {
            ".doc",
            ".docx",
            ".rtf",
            ".txt",
            ".odt"});
            this.boxExtension.Location = new System.Drawing.Point(277, 63);
            this.boxExtension.Name = "boxExtension";
            this.boxExtension.Size = new System.Drawing.Size(53, 21);
            this.boxExtension.TabIndex = 31;
            this.boxExtension.Text = ".docx";
            this.Wordcount.SetToolTip(this.boxExtension, "The default input file extension you want to use.");
            this.boxExtension.ValueMember = "Text";
            this.boxExtension.SelectedIndexChanged += new System.EventHandler(this.BoxExtension_SelectedIndexChanged);
            // 
            // UpdownFontSize
            // 
            this.UpdownFontSize.Location = new System.Drawing.Point(390, 63);
            this.UpdownFontSize.Maximum = new decimal(new int[] {
            32,
            0,
            0,
            0});
            this.UpdownFontSize.Minimum = new decimal(new int[] {
            8,
            0,
            0,
            0});
            this.UpdownFontSize.Name = "UpdownFontSize";
            this.UpdownFontSize.Size = new System.Drawing.Size(45, 20);
            this.UpdownFontSize.TabIndex = 33;
            this.Wordcount.SetToolTip(this.UpdownFontSize, "The size of the font in the Excel workbook.  A large font makes it easier to chec" +
        "k on accents.");
            this.UpdownFontSize.Value = new decimal(new int[] {
            16,
            0,
            0,
            0});
            this.UpdownFontSize.ValueChanged += new System.EventHandler(this.UpdownZoom_ValueChanged);
            // 
            // updownThreshold
            // 
            this.updownThreshold.Increment = new decimal(new int[] {
            5,
            0,
            0,
            0});
            this.updownThreshold.Location = new System.Drawing.Point(548, 61);
            this.updownThreshold.Name = "updownThreshold";
            this.updownThreshold.Size = new System.Drawing.Size(47, 20);
            this.updownThreshold.TabIndex = 36;
            this.Wordcount.SetToolTip(this.updownThreshold, "If the character copy rate drops below this, we pause to let Word catch up.  Zero" +
        " means we never stop to let Word catch up.");
            this.updownThreshold.ValueChanged += new System.EventHandler(this.UpdownThreshold_ValueChanged);
            // 
            // UpdownInterval
            // 
            this.UpdownInterval.Increment = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.UpdownInterval.Location = new System.Drawing.Point(684, 56);
            this.UpdownInterval.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.UpdownInterval.Name = "UpdownInterval";
            this.UpdownInterval.Size = new System.Drawing.Size(42, 20);
            this.UpdownInterval.TabIndex = 38;
            this.Wordcount.SetToolTip(this.UpdownInterval, "The interval between saving the output file in seconds.\r\nZero means don\'t save.");
            this.UpdownInterval.ValueChanged += new System.EventHandler(this.UpdownInterval_ValueChanged);
            // 
            // DebugCheckBox
            // 
            this.DebugCheckBox.AutoSize = true;
            this.DebugCheckBox.Location = new System.Drawing.Point(23, 89);
            this.DebugCheckBox.Name = "DebugCheckBox";
            this.DebugCheckBox.Size = new System.Drawing.Size(64, 17);
            this.DebugCheckBox.TabIndex = 39;
            this.DebugCheckBox.Text = "Debug?";
            this.Wordcount.SetToolTip(this.DebugCheckBox, "Check this if you want to display progress as the copying process proceeds.");
            this.DebugCheckBox.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.Setup);
            this.tabControl1.Controls.Add(this.Progress);
            this.tabControl1.Location = new System.Drawing.Point(10, 123);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(732, 412);
            this.tabControl1.TabIndex = 26;
            // 
            // Setup
            // 
            this.Setup.Controls.Add(this.chkCloseOnCompletion);
            this.Setup.Controls.Add(this.BtnBothToExcel);
            this.Setup.Controls.Add(this.label4);
            this.Setup.Controls.Add(this.BtnSegmentBoth);
            this.Setup.Controls.Add(this.grpUnicode);
            this.Setup.Controls.Add(this.grpLegacy);
            this.Setup.Controls.Add(this.btnGetExcelOutput);
            this.Setup.Controls.Add(this.txtExcelOutput);
            this.Setup.Location = new System.Drawing.Point(4, 22);
            this.Setup.Name = "Setup";
            this.Setup.Padding = new System.Windows.Forms.Padding(3);
            this.Setup.Size = new System.Drawing.Size(724, 386);
            this.Setup.TabIndex = 0;
            this.Setup.Text = "Setup";
            this.Setup.UseVisualStyleBackColor = true;
            // 
            // btnBothToExcel
            // 
            this.BtnBothToExcel.Enabled = false;
            this.BtnBothToExcel.Location = new System.Drawing.Point(396, 316);
            this.BtnBothToExcel.Name = "btnBothToExcel";
            this.BtnBothToExcel.Size = new System.Drawing.Size(143, 42);
            this.BtnBothToExcel.TabIndex = 35;
            this.BtnBothToExcel.Text = "Already segmented,  just send to Excel";
            this.BtnBothToExcel.UseVisualStyleBackColor = true;
            this.BtnBothToExcel.Click += new System.EventHandler(this.BothToExcel_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 276);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 13);
            this.label4.TabIndex = 34;
            this.label4.Text = "Excel Output File";
            // 
            // btnSegmentBoth
            // 
            this.BtnSegmentBoth.Enabled = false;
            this.BtnSegmentBoth.Location = new System.Drawing.Point(164, 316);
            this.BtnSegmentBoth.Name = "btnSegmentBoth";
            this.BtnSegmentBoth.Size = new System.Drawing.Size(107, 43);
            this.BtnSegmentBoth.TabIndex = 33;
            this.BtnSegmentBoth.Text = "Segment Both Files";
            this.BtnSegmentBoth.UseVisualStyleBackColor = true;
            this.BtnSegmentBoth.Click += new System.EventHandler(this.BtnSegmentBoth_Click);
            // 
            // grpUnicode
            // 
            this.grpUnicode.Controls.Add(this.chkUnicodeAddSpace);
            this.grpUnicode.Controls.Add(this.BtnUnicodeToExcel);
            this.grpUnicode.Controls.Add(this.chkUnicodeToExcel);
            this.grpUnicode.Controls.Add(this.txtUnicodeInput);
            this.grpUnicode.Controls.Add(this.label8);
            this.grpUnicode.Controls.Add(this.btnGetUnicodeInput);
            this.grpUnicode.Controls.Add(this.label9);
            this.grpUnicode.Controls.Add(this.txtUnicodeOutput);
            this.grpUnicode.Controls.Add(this.GetUnicodeOutput);
            this.grpUnicode.Controls.Add(this.txtUnicodeWordCount);
            this.grpUnicode.Controls.Add(this.label10);
            this.grpUnicode.Controls.Add(this.BtnSegmentUnicode);
            this.grpUnicode.Location = new System.Drawing.Point(0, 139);
            this.grpUnicode.Name = "grpUnicode";
            this.grpUnicode.Size = new System.Drawing.Size(724, 130);
            this.grpUnicode.TabIndex = 32;
            this.grpUnicode.TabStop = false;
            this.grpUnicode.Text = "Unicode";
            // 
            // chkUnicodeAddSpace
            // 
            this.chkUnicodeAddSpace.AutoSize = true;
            this.chkUnicodeAddSpace.Location = new System.Drawing.Point(571, 96);
            this.chkUnicodeAddSpace.Name = "chkUnicodeAddSpace";
            this.chkUnicodeAddSpace.Size = new System.Drawing.Size(131, 17);
            this.chkUnicodeAddSpace.TabIndex = 29;
            this.chkUnicodeAddSpace.Text = "Add space after range";
            this.chkUnicodeAddSpace.UseVisualStyleBackColor = true;
            // 
            // btnUnicodeToExcel
            // 
            this.BtnUnicodeToExcel.Enabled = false;
            this.BtnUnicodeToExcel.Location = new System.Drawing.Point(344, 79);
            this.BtnUnicodeToExcel.Name = "btnUnicodeToExcel";
            this.BtnUnicodeToExcel.Size = new System.Drawing.Size(123, 42);
            this.BtnUnicodeToExcel.TabIndex = 28;
            this.BtnUnicodeToExcel.Text = "Already segmented,  just send to Excel";
            this.BtnUnicodeToExcel.UseVisualStyleBackColor = true;
            this.BtnUnicodeToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
            // 
            // chkUnicodeToExcel
            // 
            this.chkUnicodeToExcel.AutoSize = true;
            this.chkUnicodeToExcel.Enabled = false;
            this.chkUnicodeToExcel.Location = new System.Drawing.Point(473, 97);
            this.chkUnicodeToExcel.Name = "chkUnicodeToExcel";
            this.chkUnicodeToExcel.Size = new System.Drawing.Size(92, 17);
            this.chkUnicodeToExcel.TabIndex = 27;
            this.chkUnicodeToExcel.Text = "Send to Excel";
            this.chkUnicodeToExcel.UseVisualStyleBackColor = true;
            this.chkUnicodeToExcel.CheckStateChanged += new System.EventHandler(this.ChkSendtoExcel_Change);
            // 
            // txtUnicodeInput
            // 
            this.txtUnicodeInput.Location = new System.Drawing.Point(111, 19);
            this.txtUnicodeInput.Name = "txtUnicodeInput";
            this.txtUnicodeInput.Size = new System.Drawing.Size(500, 20);
            this.txtUnicodeInput.TabIndex = 1;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(-2, 19);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(76, 13);
            this.label8.TabIndex = 0;
            this.label8.Text = "Input file name";
            this.label8.UseMnemonic = false;
            // 
            // btnGetUnicodeInput
            // 
            this.btnGetUnicodeInput.Location = new System.Drawing.Point(617, 19);
            this.btnGetUnicodeInput.Name = "btnGetUnicodeInput";
            this.btnGetUnicodeInput.Size = new System.Drawing.Size(75, 23);
            this.btnGetUnicodeInput.TabIndex = 2;
            this.btnGetUnicodeInput.Text = "Browse";
            this.btnGetUnicodeInput.UseVisualStyleBackColor = true;
            this.btnGetUnicodeInput.Click += new System.EventHandler(this.BtnGetInputFile_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(-2, 56);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(111, 13);
            this.label9.TabIndex = 5;
            this.label9.Text = "Segmented File Name";
            // 
            // txtUnicodeOutput
            // 
            this.txtUnicodeOutput.Location = new System.Drawing.Point(111, 53);
            this.txtUnicodeOutput.Name = "txtUnicodeOutput";
            this.txtUnicodeOutput.Size = new System.Drawing.Size(500, 20);
            this.txtUnicodeOutput.TabIndex = 6;
            // 
            // GetUnicodeOutput
            // 
            this.GetUnicodeOutput.Location = new System.Drawing.Point(617, 53);
            this.GetUnicodeOutput.Name = "GetUnicodeOutput";
            this.GetUnicodeOutput.Size = new System.Drawing.Size(75, 20);
            this.GetUnicodeOutput.TabIndex = 7;
            this.GetUnicodeOutput.Text = "Browse";
            this.GetUnicodeOutput.UseVisualStyleBackColor = true;
            this.GetUnicodeOutput.Click += new System.EventHandler(this.BtnGetOutputFile_Click);
            // 
            // txtUnicodeWordCount
            // 
            this.txtUnicodeWordCount.Enabled = false;
            this.txtUnicodeWordCount.Location = new System.Drawing.Point(194, 94);
            this.txtUnicodeWordCount.Name = "txtUnicodeWordCount";
            this.txtUnicodeWordCount.ReadOnly = true;
            this.txtUnicodeWordCount.Size = new System.Drawing.Size(108, 20);
            this.txtUnicodeWordCount.TabIndex = 16;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(124, 97);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(64, 13);
            this.label10.TabIndex = 15;
            this.label10.Text = "Word Count";
            // 
            // btnSegmentUnicode
            // 
            this.BtnSegmentUnicode.Enabled = false;
            this.BtnSegmentUnicode.Location = new System.Drawing.Point(10, 90);
            this.BtnSegmentUnicode.Name = "btnSegmentUnicode";
            this.BtnSegmentUnicode.Size = new System.Drawing.Size(112, 27);
            this.BtnSegmentUnicode.TabIndex = 3;
            this.BtnSegmentUnicode.Text = "Segment  File";
            this.BtnSegmentUnicode.UseVisualStyleBackColor = true;
            this.BtnSegmentUnicode.Click += new System.EventHandler(this.BtnSegmentInput_Click);
            // 
            // grpLegacy
            // 
            this.grpLegacy.BackColor = System.Drawing.Color.Transparent;
            this.grpLegacy.Controls.Add(this.chkLegacyAddSpace);
            this.grpLegacy.Controls.Add(this.BtnLegacyToExcel);
            this.grpLegacy.Controls.Add(this.chkLegacyToExcel);
            this.grpLegacy.Controls.Add(this.txtLegacyInput);
            this.grpLegacy.Controls.Add(this.label1);
            this.grpLegacy.Controls.Add(this.btnGetLegacyInputFile);
            this.grpLegacy.Controls.Add(this.label2);
            this.grpLegacy.Controls.Add(this.txtLegacyOutput);
            this.grpLegacy.Controls.Add(this.btnGetLegacyOutputFile);
            this.grpLegacy.Controls.Add(this.txtLegacyWordCount);
            this.grpLegacy.Controls.Add(this.label6);
            this.grpLegacy.Controls.Add(this.BtnSegmentLegacy);
            this.grpLegacy.Location = new System.Drawing.Point(4, 0);
            this.grpLegacy.Name = "grpLegacy";
            this.grpLegacy.Size = new System.Drawing.Size(717, 133);
            this.grpLegacy.TabIndex = 31;
            this.grpLegacy.TabStop = false;
            this.grpLegacy.Text = "Legacy";
            // 
            // chkLegacyAddSpace
            // 
            this.chkLegacyAddSpace.AutoSize = true;
            this.chkLegacyAddSpace.Location = new System.Drawing.Point(567, 93);
            this.chkLegacyAddSpace.Name = "chkLegacyAddSpace";
            this.chkLegacyAddSpace.Size = new System.Drawing.Size(131, 17);
            this.chkLegacyAddSpace.TabIndex = 19;
            this.chkLegacyAddSpace.Text = "Add space after range";
            this.chkLegacyAddSpace.UseVisualStyleBackColor = true;
            // 
            // btnLegacyToExcel
            // 
            this.BtnLegacyToExcel.Enabled = false;
            this.BtnLegacyToExcel.Location = new System.Drawing.Point(340, 79);
            this.BtnLegacyToExcel.Name = "btnLegacyToExcel";
            this.BtnLegacyToExcel.Size = new System.Drawing.Size(123, 42);
            this.BtnLegacyToExcel.TabIndex = 18;
            this.BtnLegacyToExcel.Text = "Already segmented,  just send to Excel";
            this.BtnLegacyToExcel.UseVisualStyleBackColor = true;
            this.BtnLegacyToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
            // 
            // chkLegacyToExcel
            // 
            this.chkLegacyToExcel.AutoSize = true;
            this.chkLegacyToExcel.Enabled = false;
            this.chkLegacyToExcel.Location = new System.Drawing.Point(469, 93);
            this.chkLegacyToExcel.Name = "chkLegacyToExcel";
            this.chkLegacyToExcel.Size = new System.Drawing.Size(92, 17);
            this.chkLegacyToExcel.TabIndex = 17;
            this.chkLegacyToExcel.Text = "Send to Excel";
            this.chkLegacyToExcel.UseVisualStyleBackColor = true;
            this.chkLegacyToExcel.CheckStateChanged += new System.EventHandler(this.ChkSendtoExcel_Change);
            // 
            // txtLegacyInput
            // 
            this.txtLegacyInput.Location = new System.Drawing.Point(107, 19);
            this.txtLegacyInput.Name = "txtLegacyInput";
            this.txtLegacyInput.Size = new System.Drawing.Size(504, 20);
            this.txtLegacyInput.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Input file name";
            this.label1.UseMnemonic = false;
            // 
            // btnGetLegacyInputFile
            // 
            this.btnGetLegacyInputFile.Location = new System.Drawing.Point(617, 19);
            this.btnGetLegacyInputFile.Name = "btnGetLegacyInputFile";
            this.btnGetLegacyInputFile.Size = new System.Drawing.Size(75, 23);
            this.btnGetLegacyInputFile.TabIndex = 2;
            this.btnGetLegacyInputFile.Text = "Browse";
            this.btnGetLegacyInputFile.UseVisualStyleBackColor = true;
            this.btnGetLegacyInputFile.Click += new System.EventHandler(this.BtnGetInputFile_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(-3, 57);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(111, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Segmented File Name";
            // 
            // txtLegacyOutput
            // 
            this.txtLegacyOutput.Location = new System.Drawing.Point(107, 53);
            this.txtLegacyOutput.Name = "txtLegacyOutput";
            this.txtLegacyOutput.Size = new System.Drawing.Size(504, 20);
            this.txtLegacyOutput.TabIndex = 6;
            // 
            // btnGetLegacyOutputFile
            // 
            this.btnGetLegacyOutputFile.Location = new System.Drawing.Point(617, 53);
            this.btnGetLegacyOutputFile.Name = "btnGetLegacyOutputFile";
            this.btnGetLegacyOutputFile.Size = new System.Drawing.Size(75, 20);
            this.btnGetLegacyOutputFile.TabIndex = 7;
            this.btnGetLegacyOutputFile.Text = "Browse";
            this.btnGetLegacyOutputFile.UseVisualStyleBackColor = true;
            this.btnGetLegacyOutputFile.Click += new System.EventHandler(this.BtnGetOutputFile_Click);
            // 
            // txtLegacyWordCount
            // 
            this.txtLegacyWordCount.Enabled = false;
            this.txtLegacyWordCount.Location = new System.Drawing.Point(194, 94);
            this.txtLegacyWordCount.Name = "txtLegacyWordCount";
            this.txtLegacyWordCount.ReadOnly = true;
            this.txtLegacyWordCount.Size = new System.Drawing.Size(108, 20);
            this.txtLegacyWordCount.TabIndex = 16;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(124, 97);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 13);
            this.label6.TabIndex = 15;
            this.label6.Text = "Word Count";
            // 
            // btnSegmentLegacy
            // 
            this.BtnSegmentLegacy.Enabled = false;
            this.BtnSegmentLegacy.Location = new System.Drawing.Point(10, 90);
            this.BtnSegmentLegacy.Name = "btnSegmentLegacy";
            this.BtnSegmentLegacy.Size = new System.Drawing.Size(112, 27);
            this.BtnSegmentLegacy.TabIndex = 3;
            this.BtnSegmentLegacy.Text = "Segment  File";
            this.BtnSegmentLegacy.UseVisualStyleBackColor = true;
            this.BtnSegmentLegacy.Click += new System.EventHandler(this.BtnSegmentInput_Click);
            // 
            // btnGetExcelOutput
            // 
            this.btnGetExcelOutput.Location = new System.Drawing.Point(621, 272);
            this.btnGetExcelOutput.Name = "btnGetExcelOutput";
            this.btnGetExcelOutput.Size = new System.Drawing.Size(75, 23);
            this.btnGetExcelOutput.TabIndex = 30;
            this.btnGetExcelOutput.Text = "Browse";
            this.btnGetExcelOutput.UseVisualStyleBackColor = true;
            this.btnGetExcelOutput.Click += new System.EventHandler(this.BtnGetExcelOutput_Click);
            // 
            // txtExcelOutput
            // 
            this.txtExcelOutput.Location = new System.Drawing.Point(131, 272);
            this.txtExcelOutput.Name = "txtExcelOutput";
            this.txtExcelOutput.Size = new System.Drawing.Size(490, 20);
            this.txtExcelOutput.TabIndex = 29;
            // 
            // Progress
            // 
            this.Progress.Controls.Add(this.statusStrip1);
            this.Progress.Controls.Add(this.boxProgress);
            this.Progress.Location = new System.Drawing.Point(4, 22);
            this.Progress.Name = "Progress";
            this.Progress.Padding = new System.Windows.Forms.Padding(3);
            this.Progress.Size = new System.Drawing.Size(724, 386);
            this.Progress.TabIndex = 1;
            this.Progress.Text = "Progress";
            this.Progress.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.progressBar1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(3, 361);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(718, 22);
            this.statusStrip1.TabIndex = 29;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Overflow = System.Windows.Forms.ToolStripItemOverflow.Always;
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // progressBar1
            // 
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(400, 16);
            // 
            // boxProgress
            // 
            this.boxProgress.HorizontalScrollbar = true;
            this.boxProgress.Location = new System.Drawing.Point(9, 6);
            this.boxProgress.Name = "boxProgress";
            this.boxProgress.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.boxProgress.Size = new System.Drawing.Size(709, 355);
            this.boxProgress.TabIndex = 28;
            // 
            // openUnicodeFileDialog
            // 
            this.openUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx|RTF files|*.rtf";
            this.openUnicodeFileDialog.FilterIndex = 2;
            this.openUnicodeFileDialog.Multiselect = true;
            // 
            // btnClose
            // 
            this.BtnClose.Location = new System.Drawing.Point(460, 545);
            this.BtnClose.Name = "btnClose";
            this.BtnClose.Size = new System.Drawing.Size(112, 44);
            this.BtnClose.TabIndex = 27;
            this.BtnClose.Text = "Close";
            this.BtnClose.UseVisualStyleBackColor = true;
            this.BtnClose.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(754, 24);
            this.menuStrip1.TabIndex = 29;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2,
            this.toolStripMenuItem1});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.AutoToolTip = true;
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(166, 22);
            this.toolStripMenuItem2.Text = "New Comparison";
            this.toolStripMenuItem2.ToolTipText = "Clear the text boxes so you can start a new comparison.";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.ToolStripMenuItem2_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(166, 22);
            this.toolStripMenuItem1.Text = "Exit";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.BtnClose_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.documentationToolStripMenuItem,
            this.licenseToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // documentationToolStripMenuItem
            // 
            this.documentationToolStripMenuItem.Name = "documentationToolStripMenuItem";
            this.documentationToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.documentationToolStripMenuItem.Text = "Documentation";
            this.documentationToolStripMenuItem.Click += new System.EventHandler(this.DocumentationToolStripMenuItem_Click);
            // 
            // licenseToolStripMenuItem
            // 
            this.licenseToolStripMenuItem.Name = "licenseToolStripMenuItem";
            this.licenseToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.licenseToolStripMenuItem.Text = "License";
            this.licenseToolStripMenuItem.Click += new System.EventHandler(this.LicenseToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.AboutToolStripMenuItem_Click);
            // 
            // btnPauseResume
            // 
            this.BtnPauseResume.Enabled = false;
            this.BtnPauseResume.Location = new System.Drawing.Point(219, 545);
            this.BtnPauseResume.Name = "btnPauseResume";
            this.BtnPauseResume.Size = new System.Drawing.Size(112, 44);
            this.BtnPauseResume.TabIndex = 30;
            this.BtnPauseResume.Text = "Pause";
            this.BtnPauseResume.UseVisualStyleBackColor = true;
            this.BtnPauseResume.Click += new System.EventHandler(this.BtnPauseResume_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(141, 66);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(131, 13);
            this.label7.TabIndex = 32;
            this.label7.Text = "Default input file extension";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(338, 65);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(49, 13);
            this.label11.TabIndex = 34;
            this.label11.Text = "Font size";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(441, 56);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(100, 26);
            this.label12.TabIndex = 35;
            this.label12.Text = "Character copy rate\r\nthreshold.";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(601, 53);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(67, 26);
            this.label13.TabIndex = 37;
            this.label13.Text = "Output file\r\nsave interval";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // Interlinear
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(754, 611);
            this.Controls.Add(this.DebugCheckBox);
            this.Controls.Add(this.UpdownInterval);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.updownThreshold);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.UpdownFontSize);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.boxExtension);
            this.Controls.Add(this.BtnPauseResume);
            this.Controls.Add(this.BtnClose);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.WordsPerLine);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Interlinear";
            this.Text = "Interlinear comparison";
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UpdownFontSize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.updownThreshold)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UpdownInterval)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.Setup.ResumeLayout(false);
            this.Setup.PerformLayout();
            this.grpUnicode.ResumeLayout(false);
            this.grpUnicode.PerformLayout();
            this.grpLegacy.ResumeLayout(false);
            this.grpLegacy.PerformLayout();
            this.Progress.ResumeLayout(false);
            this.Progress.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openLegacyFileDialog;
        private System.Windows.Forms.SaveFileDialog saveLegacyFileDialog;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown WordsPerLine;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.SaveFileDialog saveUnicodeFileDialog;
        private System.Windows.Forms.ToolTip Wordcount;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage Setup;
        private System.Windows.Forms.GroupBox grpUnicode;
        private System.Windows.Forms.TextBox txtUnicodeInput;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnGetUnicodeInput;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtUnicodeOutput;
        private System.Windows.Forms.Button GetUnicodeOutput;
        private System.Windows.Forms.TextBox txtUnicodeWordCount;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Button BtnSegmentUnicode;
        private System.Windows.Forms.GroupBox grpLegacy;
        private System.Windows.Forms.TextBox txtLegacyInput;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGetLegacyInputFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLegacyOutput;
        private System.Windows.Forms.Button btnGetLegacyOutputFile;
        private System.Windows.Forms.TextBox txtLegacyWordCount;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button BtnSegmentLegacy;
        private System.Windows.Forms.Button btnGetExcelOutput;
        private System.Windows.Forms.TextBox txtExcelOutput;
        private System.Windows.Forms.TabPage Progress;
        private System.Windows.Forms.ListBox boxProgress;
        private System.Windows.Forms.OpenFileDialog openUnicodeFileDialog;
        private System.Windows.Forms.Button BtnSegmentBoth;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox chkUnicodeToExcel;
        private System.Windows.Forms.CheckBox chkLegacyToExcel;
        private System.Windows.Forms.Button BtnClose;
        public System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
        private System.Windows.Forms.Button BtnBothToExcel;
        private System.Windows.Forms.Button BtnUnicodeToExcel;
        private System.Windows.Forms.Button BtnLegacyToExcel;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem documentationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem licenseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripProgressBar progressBar1;
        private System.Windows.Forms.Button BtnPauseResume;
        private System.Windows.Forms.CheckBox chkUnicodeAddSpace;
        private System.Windows.Forms.CheckBox chkLegacyAddSpace;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.CheckBox chkCloseOnCompletion;
        private System.Windows.Forms.ComboBox boxExtension;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.NumericUpDown UpdownFontSize;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.NumericUpDown updownThreshold;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.NumericUpDown UpdownInterval;
        private System.Windows.Forms.CheckBox DebugCheckBox;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}

