namespace Interlinear
{
    partial class Form1
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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.Setup = new System.Windows.Forms.TabPage();
            this.btnInterlinear = new System.Windows.Forms.Button();
            this.btnBothToExcel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnSegmentBoth = new System.Windows.Forms.Button();
            this.grpUnicode = new System.Windows.Forms.GroupBox();
            this.chkUnicodeAddSpace = new System.Windows.Forms.CheckBox();
            this.btnUnicodeToExcel = new System.Windows.Forms.Button();
            this.chkUnicodeToExcel = new System.Windows.Forms.CheckBox();
            this.txtUnicodeInput = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnGetUnicodeInput = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.txtUnicodeOutput = new System.Windows.Forms.TextBox();
            this.GetUnicodeOutput = new System.Windows.Forms.Button();
            this.txtUnicodeWordCount = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btnSegmentUnicode = new System.Windows.Forms.Button();
            this.grpLegacy = new System.Windows.Forms.GroupBox();
            this.chkLegacyAddSpace = new System.Windows.Forms.CheckBox();
            this.btnLegacyToExcel = new System.Windows.Forms.Button();
            this.chkLegacyToExcel = new System.Windows.Forms.CheckBox();
            this.txtLegacyInput = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGetLegacyInputFile = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtLegacyOutput = new System.Windows.Forms.TextBox();
            this.btnGetLegacyOutputFile = new System.Windows.Forms.Button();
            this.txtLegacyWordCount = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnSegmentLegacy = new System.Windows.Forms.Button();
            this.btnGetExcelOutput = new System.Windows.Forms.Button();
            this.txtExcelOutput = new System.Windows.Forms.TextBox();
            this.Progress = new System.Windows.Forms.TabPage();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.progressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.boxProgress = new System.Windows.Forms.ListBox();
            this.openUnicodeFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnClose = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.documentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.licenseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPauseResume = new System.Windows.Forms.Button();
            this.chkDebug = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).BeginInit();
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
            this.openLegacyFileDialog.DefaultExt = "doc";
            this.openLegacyFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            this.openLegacyFileDialog.FilterIndex = 2;
            // 
            // saveLegacyFileDialog
            // 
            this.saveLegacyFileDialog.DefaultExt = "doc";
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
            this.WordsPerLine.Location = new System.Drawing.Point(111, 63);
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
            this.label3.Size = new System.Drawing.Size(698, 42);
            this.label3.TabIndex = 23;
            this.label3.Text = "Interlinear Comparison";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // saveUnicodeFileDialog
            // 
            this.saveUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            this.saveUnicodeFileDialog.FilterIndex = 2;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.Setup);
            this.tabControl1.Controls.Add(this.Progress);
            this.tabControl1.Location = new System.Drawing.Point(10, 85);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(732, 452);
            this.tabControl1.TabIndex = 26;
            // 
            // Setup
            // 
            this.Setup.Controls.Add(this.btnInterlinear);
            this.Setup.Controls.Add(this.btnBothToExcel);
            this.Setup.Controls.Add(this.label4);
            this.Setup.Controls.Add(this.btnSegmentBoth);
            this.Setup.Controls.Add(this.grpUnicode);
            this.Setup.Controls.Add(this.grpLegacy);
            this.Setup.Controls.Add(this.btnGetExcelOutput);
            this.Setup.Controls.Add(this.txtExcelOutput);
            this.Setup.Location = new System.Drawing.Point(4, 22);
            this.Setup.Name = "Setup";
            this.Setup.Padding = new System.Windows.Forms.Padding(3);
            this.Setup.Size = new System.Drawing.Size(724, 426);
            this.Setup.TabIndex = 0;
            this.Setup.Text = "Setup";
            this.Setup.UseVisualStyleBackColor = true;
            // 
            // btnInterlinear
            // 
            this.btnInterlinear.Enabled = false;
            this.btnInterlinear.Location = new System.Drawing.Point(322, 317);
            this.btnInterlinear.Name = "btnInterlinear";
            this.btnInterlinear.Size = new System.Drawing.Size(106, 41);
            this.btnInterlinear.TabIndex = 36;
            this.btnInterlinear.Text = "Build Interlinear Worksheet";
            this.btnInterlinear.UseVisualStyleBackColor = true;
            this.btnInterlinear.Click += new System.EventHandler(this.btnInterlinear_Click);
            // 
            // btnBothToExcel
            // 
            this.btnBothToExcel.Enabled = false;
            this.btnBothToExcel.Location = new System.Drawing.Point(159, 315);
            this.btnBothToExcel.Name = "btnBothToExcel";
            this.btnBothToExcel.Size = new System.Drawing.Size(143, 42);
            this.btnBothToExcel.TabIndex = 35;
            this.btnBothToExcel.Text = "Already segmented,  just send to Excel";
            this.btnBothToExcel.UseVisualStyleBackColor = true;
            this.btnBothToExcel.Click += new System.EventHandler(this.BothToExcel_Click);
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
            this.btnSegmentBoth.Enabled = false;
            this.btnSegmentBoth.Location = new System.Drawing.Point(9, 315);
            this.btnSegmentBoth.Name = "btnSegmentBoth";
            this.btnSegmentBoth.Size = new System.Drawing.Size(107, 43);
            this.btnSegmentBoth.TabIndex = 33;
            this.btnSegmentBoth.Text = "Segment Both Files";
            this.btnSegmentBoth.UseVisualStyleBackColor = true;
            this.btnSegmentBoth.Click += new System.EventHandler(this.btnSegmentBoth_Click);
            // 
            // grpUnicode
            // 
            this.grpUnicode.Controls.Add(this.chkUnicodeAddSpace);
            this.grpUnicode.Controls.Add(this.btnUnicodeToExcel);
            this.grpUnicode.Controls.Add(this.chkUnicodeToExcel);
            this.grpUnicode.Controls.Add(this.txtUnicodeInput);
            this.grpUnicode.Controls.Add(this.label8);
            this.grpUnicode.Controls.Add(this.btnGetUnicodeInput);
            this.grpUnicode.Controls.Add(this.label9);
            this.grpUnicode.Controls.Add(this.txtUnicodeOutput);
            this.grpUnicode.Controls.Add(this.GetUnicodeOutput);
            this.grpUnicode.Controls.Add(this.txtUnicodeWordCount);
            this.grpUnicode.Controls.Add(this.label10);
            this.grpUnicode.Controls.Add(this.btnSegmentUnicode);
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
            this.btnUnicodeToExcel.Enabled = false;
            this.btnUnicodeToExcel.Location = new System.Drawing.Point(344, 79);
            this.btnUnicodeToExcel.Name = "btnUnicodeToExcel";
            this.btnUnicodeToExcel.Size = new System.Drawing.Size(123, 42);
            this.btnUnicodeToExcel.TabIndex = 28;
            this.btnUnicodeToExcel.Text = "Already segmented,  just send to Excel";
            this.btnUnicodeToExcel.UseVisualStyleBackColor = true;
            this.btnUnicodeToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
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
            this.chkUnicodeToExcel.CheckStateChanged += new System.EventHandler(this.chkSendtoExcel_Change);
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
            this.btnGetUnicodeInput.Click += new System.EventHandler(this.btnGetInputFile_Click);
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
            this.GetUnicodeOutput.Click += new System.EventHandler(this.btnGetOutputFile_Click);
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
            this.btnSegmentUnicode.Enabled = false;
            this.btnSegmentUnicode.Location = new System.Drawing.Point(10, 90);
            this.btnSegmentUnicode.Name = "btnSegmentUnicode";
            this.btnSegmentUnicode.Size = new System.Drawing.Size(112, 27);
            this.btnSegmentUnicode.TabIndex = 3;
            this.btnSegmentUnicode.Text = "Segment  File";
            this.btnSegmentUnicode.UseVisualStyleBackColor = true;
            this.btnSegmentUnicode.Click += new System.EventHandler(this.btnSegmentInput_Click);
            // 
            // grpLegacy
            // 
            this.grpLegacy.BackColor = System.Drawing.SystemColors.Control;
            this.grpLegacy.Controls.Add(this.chkLegacyAddSpace);
            this.grpLegacy.Controls.Add(this.btnLegacyToExcel);
            this.grpLegacy.Controls.Add(this.chkLegacyToExcel);
            this.grpLegacy.Controls.Add(this.txtLegacyInput);
            this.grpLegacy.Controls.Add(this.label1);
            this.grpLegacy.Controls.Add(this.btnGetLegacyInputFile);
            this.grpLegacy.Controls.Add(this.label2);
            this.grpLegacy.Controls.Add(this.txtLegacyOutput);
            this.grpLegacy.Controls.Add(this.btnGetLegacyOutputFile);
            this.grpLegacy.Controls.Add(this.txtLegacyWordCount);
            this.grpLegacy.Controls.Add(this.label6);
            this.grpLegacy.Controls.Add(this.btnSegmentLegacy);
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
            this.btnLegacyToExcel.Enabled = false;
            this.btnLegacyToExcel.Location = new System.Drawing.Point(340, 79);
            this.btnLegacyToExcel.Name = "btnLegacyToExcel";
            this.btnLegacyToExcel.Size = new System.Drawing.Size(123, 42);
            this.btnLegacyToExcel.TabIndex = 18;
            this.btnLegacyToExcel.Text = "Already segmented,  just send to Excel";
            this.btnLegacyToExcel.UseVisualStyleBackColor = true;
            this.btnLegacyToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
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
            this.chkLegacyToExcel.CheckStateChanged += new System.EventHandler(this.chkSendtoExcel_Change);
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
            this.btnGetLegacyInputFile.Click += new System.EventHandler(this.btnGetInputFile_Click);
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
            this.btnGetLegacyOutputFile.Click += new System.EventHandler(this.btnGetOutputFile_Click);
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
            this.btnSegmentLegacy.Enabled = false;
            this.btnSegmentLegacy.Location = new System.Drawing.Point(10, 90);
            this.btnSegmentLegacy.Name = "btnSegmentLegacy";
            this.btnSegmentLegacy.Size = new System.Drawing.Size(112, 27);
            this.btnSegmentLegacy.TabIndex = 3;
            this.btnSegmentLegacy.Text = "Segment  File";
            this.btnSegmentLegacy.UseVisualStyleBackColor = true;
            this.btnSegmentLegacy.Click += new System.EventHandler(this.btnSegmentInput_Click);
            // 
            // btnGetExcelOutput
            // 
            this.btnGetExcelOutput.Location = new System.Drawing.Point(621, 272);
            this.btnGetExcelOutput.Name = "btnGetExcelOutput";
            this.btnGetExcelOutput.Size = new System.Drawing.Size(75, 23);
            this.btnGetExcelOutput.TabIndex = 30;
            this.btnGetExcelOutput.Text = "Browse";
            this.btnGetExcelOutput.UseVisualStyleBackColor = true;
            this.btnGetExcelOutput.Click += new System.EventHandler(this.btnGetExcelOutput_Click);
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
            this.Progress.Size = new System.Drawing.Size(724, 426);
            this.Progress.TabIndex = 1;
            this.Progress.Text = "Progress";
            this.Progress.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.progressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(3, 401);
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
            this.boxProgress.Size = new System.Drawing.Size(709, 394);
            this.boxProgress.TabIndex = 28;
            // 
            // openUnicodeFileDialog
            // 
            this.openUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            this.openUnicodeFileDialog.FilterIndex = 2;
            this.openUnicodeFileDialog.Multiselect = true;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(460, 545);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(112, 44);
            this.btnClose.TabIndex = 27;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
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
            this.toolStripMenuItem1});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(92, 22);
            this.toolStripMenuItem1.Text = "Exit";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.btnClose_Click);
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
            this.documentationToolStripMenuItem.Click += new System.EventHandler(this.documentationToolStripMenuItem_Click);
            // 
            // licenseToolStripMenuItem
            // 
            this.licenseToolStripMenuItem.Name = "licenseToolStripMenuItem";
            this.licenseToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.licenseToolStripMenuItem.Text = "License";
            this.licenseToolStripMenuItem.Click += new System.EventHandler(this.licenseToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // btnPauseResume
            // 
            this.btnPauseResume.Enabled = false;
            this.btnPauseResume.Location = new System.Drawing.Point(219, 545);
            this.btnPauseResume.Name = "btnPauseResume";
            this.btnPauseResume.Size = new System.Drawing.Size(112, 44);
            this.btnPauseResume.TabIndex = 30;
            this.btnPauseResume.Text = "Pause";
            this.btnPauseResume.UseVisualStyleBackColor = true;
            this.btnPauseResume.Click += new System.EventHandler(this.btnPauseResume_Click);
            // 
            // chkDebug
            // 
            this.chkDebug.AutoSize = true;
            this.chkDebug.Location = new System.Drawing.Point(208, 66);
            this.chkDebug.Name = "chkDebug";
            this.chkDebug.Size = new System.Drawing.Size(126, 17);
            this.chkDebug.TabIndex = 31;
            this.chkDebug.Text = "Debug messages on.";
            this.chkDebug.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(754, 601);
            this.Controls.Add(this.chkDebug);
            this.Controls.Add(this.btnPauseResume);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.WordsPerLine);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Interlinear comparison";
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).EndInit();
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
        private System.Windows.Forms.Button btnSegmentUnicode;
        private System.Windows.Forms.GroupBox grpLegacy;
        private System.Windows.Forms.TextBox txtLegacyInput;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGetLegacyInputFile;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtLegacyOutput;
        private System.Windows.Forms.Button btnGetLegacyOutputFile;
        private System.Windows.Forms.TextBox txtLegacyWordCount;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnSegmentLegacy;
        private System.Windows.Forms.Button btnGetExcelOutput;
        private System.Windows.Forms.TextBox txtExcelOutput;
        private System.Windows.Forms.TabPage Progress;
        private System.Windows.Forms.ListBox boxProgress;
        private System.Windows.Forms.OpenFileDialog openUnicodeFileDialog;
        private System.Windows.Forms.Button btnSegmentBoth;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox chkUnicodeToExcel;
        private System.Windows.Forms.CheckBox chkLegacyToExcel;
        private System.Windows.Forms.Button btnClose;
        public System.Windows.Forms.SaveFileDialog saveExcelFileDialog;
        private System.Windows.Forms.Button btnBothToExcel;
        private System.Windows.Forms.Button btnUnicodeToExcel;
        private System.Windows.Forms.Button btnLegacyToExcel;
        private System.Windows.Forms.Button btnInterlinear;
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
        private System.Windows.Forms.Button btnPauseResume;
        private System.Windows.Forms.CheckBox chkUnicodeAddSpace;
        private System.Windows.Forms.CheckBox chkLegacyAddSpace;
        private System.Windows.Forms.CheckBox chkDebug;
    }
}

