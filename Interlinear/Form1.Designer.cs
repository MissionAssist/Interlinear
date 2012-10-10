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
            this.btnBothToExcel = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnSegmentBoth = new System.Windows.Forms.Button();
            this.grpUnicode = new System.Windows.Forms.GroupBox();
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
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.boxProgress = new System.Windows.Forms.ListBox();
            this.openUnicodeFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnClose = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.Setup.SuspendLayout();
            this.grpUnicode.SuspendLayout();
            this.grpLegacy.SuspendLayout();
            this.Progress.SuspendLayout();
            this.SuspendLayout();
            // 
            // openLegacyFileDialog
            // 
            this.openLegacyFileDialog.DefaultExt = "doc";
            this.openLegacyFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            // 
            // saveLegacyFileDialog
            // 
            this.saveLegacyFileDialog.DefaultExt = "doc";
            this.saveLegacyFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(20, 55);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Words per line";
            // 
            // WordsPerLine
            // 
            this.WordsPerLine.Location = new System.Drawing.Point(111, 53);
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
            this.label3.Location = new System.Drawing.Point(6, 9);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(698, 41);
            this.label3.TabIndex = 23;
            this.label3.Text = "Interlinear Comparison";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // saveUnicodeFileDialog
            // 
            this.saveUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.Setup);
            this.tabControl1.Controls.Add(this.Progress);
            this.tabControl1.Location = new System.Drawing.Point(10, 79);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(732, 416);
            this.tabControl1.TabIndex = 26;
            // 
            // Setup
            // 
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
            this.Setup.Size = new System.Drawing.Size(724, 390);
            this.Setup.TabIndex = 0;
            this.Setup.Text = "Setup";
            this.Setup.UseVisualStyleBackColor = true;
            // 
            // btnBothToExcel
            // 
            this.btnBothToExcel.Enabled = false;
            this.btnBothToExcel.Location = new System.Drawing.Point(359, 316);
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
            this.grpUnicode.Size = new System.Drawing.Size(751, 130);
            this.grpUnicode.TabIndex = 32;
            this.grpUnicode.TabStop = false;
            this.grpUnicode.Text = "Unicode";
            // 
            // btnUnicodeToExcel
            // 
            this.btnUnicodeToExcel.Enabled = false;
            this.btnUnicodeToExcel.Location = new System.Drawing.Point(359, 79);
            this.btnUnicodeToExcel.Name = "btnUnicodeToExcel";
            this.btnUnicodeToExcel.Size = new System.Drawing.Size(143, 42);
            this.btnUnicodeToExcel.TabIndex = 28;
            this.btnUnicodeToExcel.Text = "Already segmented,  just send to Excel";
            this.btnUnicodeToExcel.UseVisualStyleBackColor = true;
            this.btnUnicodeToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
            // 
            // chkUnicodeToExcel
            // 
            this.chkUnicodeToExcel.AutoSize = true;
            this.chkUnicodeToExcel.Enabled = false;
            this.chkUnicodeToExcel.Location = new System.Drawing.Point(523, 100);
            this.chkUnicodeToExcel.Name = "chkUnicodeToExcel";
            this.chkUnicodeToExcel.Size = new System.Drawing.Size(92, 17);
            this.chkUnicodeToExcel.TabIndex = 27;
            this.chkUnicodeToExcel.Text = "Send to Excel";
            this.chkUnicodeToExcel.UseVisualStyleBackColor = true;
            // 
            // txtUnicodeInput
            // 
            this.txtUnicodeInput.Location = new System.Drawing.Point(107, 19);
            this.txtUnicodeInput.Name = "txtUnicodeInput";
            this.txtUnicodeInput.Size = new System.Drawing.Size(504, 20);
            this.txtUnicodeInput.TabIndex = 1;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(15, 22);
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
            this.label9.Location = new System.Drawing.Point(12, 56);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(89, 13);
            this.label9.TabIndex = 5;
            this.label9.Text = "Output File Name";
            // 
            // txtUnicodeOutput
            // 
            this.txtUnicodeOutput.Location = new System.Drawing.Point(107, 53);
            this.txtUnicodeOutput.Name = "txtUnicodeOutput";
            this.txtUnicodeOutput.Size = new System.Drawing.Size(504, 20);
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
            this.grpLegacy.Location = new System.Drawing.Point(4, 16);
            this.grpLegacy.Name = "grpLegacy";
            this.grpLegacy.Size = new System.Drawing.Size(751, 132);
            this.grpLegacy.TabIndex = 31;
            this.grpLegacy.TabStop = false;
            this.grpLegacy.Text = "Legacy";
            // 
            // btnLegacyToExcel
            // 
            this.btnLegacyToExcel.Enabled = false;
            this.btnLegacyToExcel.Location = new System.Drawing.Point(355, 82);
            this.btnLegacyToExcel.Name = "btnLegacyToExcel";
            this.btnLegacyToExcel.Size = new System.Drawing.Size(143, 42);
            this.btnLegacyToExcel.TabIndex = 18;
            this.btnLegacyToExcel.Text = "Already segmented,  just send to Excel";
            this.btnLegacyToExcel.UseVisualStyleBackColor = true;
            this.btnLegacyToExcel.Click += new System.EventHandler(this.SendToExcel_Click);
            // 
            // chkLegacyToExcel
            // 
            this.chkLegacyToExcel.AutoSize = true;
            this.chkLegacyToExcel.Enabled = false;
            this.chkLegacyToExcel.Location = new System.Drawing.Point(519, 96);
            this.chkLegacyToExcel.Name = "chkLegacyToExcel";
            this.chkLegacyToExcel.Size = new System.Drawing.Size(92, 17);
            this.chkLegacyToExcel.TabIndex = 17;
            this.chkLegacyToExcel.Text = "Send to Excel";
            this.chkLegacyToExcel.UseVisualStyleBackColor = true;
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
            this.label1.Location = new System.Drawing.Point(15, 22);
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
            this.label2.Location = new System.Drawing.Point(12, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Output File Name";
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
            this.Progress.Controls.Add(this.progressBar1);
            this.Progress.Controls.Add(this.boxProgress);
            this.Progress.Location = new System.Drawing.Point(4, 22);
            this.Progress.Name = "Progress";
            this.Progress.Padding = new System.Windows.Forms.Padding(3);
            this.Progress.Size = new System.Drawing.Size(724, 390);
            this.Progress.TabIndex = 1;
            this.Progress.Text = "Progress";
            this.Progress.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 462);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(697, 18);
            this.progressBar1.TabIndex = 29;
            // 
            // boxProgress
            // 
            this.boxProgress.HorizontalScrollbar = true;
            this.boxProgress.Location = new System.Drawing.Point(9, 6);
            this.boxProgress.Name = "boxProgress";
            this.boxProgress.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.boxProgress.Size = new System.Drawing.Size(709, 381);
            this.boxProgress.TabIndex = 28;
            // 
            // openUnicodeFileDialog
            // 
            this.openUnicodeFileDialog.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(304, 501);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(112, 44);
            this.btnClose.TabIndex = 27;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(754, 598);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.WordsPerLine);
            this.Controls.Add(this.label5);
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
        private System.Windows.Forms.ProgressBar progressBar1;
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
    }
}

