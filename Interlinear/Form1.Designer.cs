namespace WindowsFormsApplication1
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.txtInput = new System.Windows.Forms.TextBox();
            this.btnGetInputFile = new System.Windows.Forms.Button();
            this.btnSegmentInput = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnBrowseOutput = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.boxProgress = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.WordsPerLine = new System.Windows.Forms.NumericUpDown();
            this.label6 = new System.Windows.Forms.Label();
            this.txtWordCount = new System.Windows.Forms.TextBox();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtExcelOutput = new System.Windows.Forms.TextBox();
            this.btnGetExcelOutput = new System.Windows.Forms.Button();
            this.saveFileDialog2 = new System.Windows.Forms.SaveFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "doc";
            this.openFileDialog1.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Input file name";
            this.label1.UseMnemonic = false;
            // 
            // txtInput
            // 
            this.txtInput.Location = new System.Drawing.Point(107, 53);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(504, 20);
            this.txtInput.TabIndex = 1;
            // 
            // btnGetInputFile
            // 
            this.btnGetInputFile.Location = new System.Drawing.Point(617, 53);
            this.btnGetInputFile.Name = "btnGetInputFile";
            this.btnGetInputFile.Size = new System.Drawing.Size(75, 23);
            this.btnGetInputFile.TabIndex = 2;
            this.btnGetInputFile.Text = "Browse";
            this.btnGetInputFile.UseVisualStyleBackColor = true;
            this.btnGetInputFile.Click += new System.EventHandler(this.btnGetInputFile_Click);
            // 
            // btnSegmentInput
            // 
            this.btnSegmentInput.Enabled = false;
            this.btnSegmentInput.Location = new System.Drawing.Point(-4, 144);
            this.btnSegmentInput.Name = "btnSegmentInput";
            this.btnSegmentInput.Size = new System.Drawing.Size(112, 52);
            this.btnSegmentInput.TabIndex = 3;
            this.btnSegmentInput.Text = "Segment  File";
            this.btnSegmentInput.UseVisualStyleBackColor = true;
            this.btnSegmentInput.Click += new System.EventHandler(this.btnSegmentInput_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(286, 373);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(112, 44);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Output File Name";
            // 
            // txtOutput
            // 
            this.txtOutput.Location = new System.Drawing.Point(107, 87);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(504, 20);
            this.txtOutput.TabIndex = 6;
            // 
            // btnBrowseOutput
            // 
            this.btnBrowseOutput.Location = new System.Drawing.Point(617, 90);
            this.btnBrowseOutput.Name = "btnBrowseOutput";
            this.btnBrowseOutput.Size = new System.Drawing.Size(75, 23);
            this.btnBrowseOutput.TabIndex = 7;
            this.btnBrowseOutput.Text = "Browse";
            this.btnBrowseOutput.UseVisualStyleBackColor = true;
            this.btnBrowseOutput.Click += new System.EventHandler(this.btnBrowseOutput_Click);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "doc";
            this.saveFileDialog1.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(5, 349);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(697, 18);
            this.progressBar1.TabIndex = 10;
            // 
            // boxProgress
            // 
            this.boxProgress.FormattingEnabled = true;
            this.boxProgress.Location = new System.Drawing.Point(301, 144);
            this.boxProgress.Name = "boxProgress";
            this.boxProgress.Size = new System.Drawing.Size(392, 147);
            this.boxProgress.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(478, 128);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 13);
            this.label4.TabIndex = 12;
            this.label4.Text = "Progress";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 120);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(75, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Words per line";
            // 
            // WordsPerLine
            // 
            this.WordsPerLine.Location = new System.Drawing.Point(107, 118);
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
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(117, 147);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 13);
            this.label6.TabIndex = 15;
            this.label6.Text = "Word Count";
            // 
            // txtWordCount
            // 
            this.txtWordCount.Enabled = false;
            this.txtWordCount.Location = new System.Drawing.Point(187, 144);
            this.txtWordCount.Name = "txtWordCount";
            this.txtWordCount.ReadOnly = true;
            this.txtWordCount.Size = new System.Drawing.Size(108, 20);
            this.txtWordCount.TabIndex = 16;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Location = new System.Drawing.Point(10, 19);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(84, 17);
            this.radioButton1.TabIndex = 17;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Legacy Font";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(10, 39);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(89, 17);
            this.radioButton2.TabIndex = 18;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Unicode Font";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton2);
            this.groupBox1.Controls.Add(this.radioButton1);
            this.groupBox1.Location = new System.Drawing.Point(5, 212);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(112, 71);
            this.groupBox1.TabIndex = 19;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Font in file";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(14, 306);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(87, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "Excel Output File";
            // 
            // txtExcelOutput
            // 
            this.txtExcelOutput.Location = new System.Drawing.Point(107, 299);
            this.txtExcelOutput.Name = "txtExcelOutput";
            this.txtExcelOutput.Size = new System.Drawing.Size(490, 20);
            this.txtExcelOutput.TabIndex = 21;
            // 
            // btnGetExcelOutput
            // 
            this.btnGetExcelOutput.Location = new System.Drawing.Point(604, 298);
            this.btnGetExcelOutput.Name = "btnGetExcelOutput";
            this.btnGetExcelOutput.Size = new System.Drawing.Size(75, 23);
            this.btnGetExcelOutput.TabIndex = 22;
            this.btnGetExcelOutput.Text = "Browse";
            this.btnGetExcelOutput.UseVisualStyleBackColor = true;
            // 
            // saveFileDialog2
            // 
            this.saveFileDialog2.Filter = "Excel WorkBook | .docx; .docx";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 414);
            this.Controls.Add(this.btnGetExcelOutput);
            this.Controls.Add(this.txtExcelOutput);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtWordCount);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.WordsPerLine);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.boxProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnBrowseOutput);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSegmentInput);
            this.Controls.Add(this.btnGetInputFile);
            this.Controls.Add(this.txtInput);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Interlinear comparison";
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtInput;
        private System.Windows.Forms.Button btnGetInputFile;
        private System.Windows.Forms.Button btnSegmentInput;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnBrowseOutput;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox boxProgress;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown WordsPerLine;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtWordCount;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtExcelOutput;
        private System.Windows.Forms.Button btnGetExcelOutput;
        private System.Windows.Forms.SaveFileDialog saveFileDialog2;
    }
}

