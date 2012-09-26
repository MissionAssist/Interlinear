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
            this.txtLegacy = new System.Windows.Forms.TextBox();
            this.btnGetLegacyFile = new System.Windows.Forms.Button();
            this.btnSegmentLegacy = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnBrowseOutput = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.txtLineCount = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.boxProgress = new System.Windows.Forms.ListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.WordsPerLine = new System.Windows.Forms.NumericUpDown();
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).BeginInit();
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
            // txtLegacy
            // 
            this.txtLegacy.Location = new System.Drawing.Point(107, 53);
            this.txtLegacy.Name = "txtLegacy";
            this.txtLegacy.Size = new System.Drawing.Size(504, 20);
            this.txtLegacy.TabIndex = 1;
            // 
            // btnGetLegacyFile
            // 
            this.btnGetLegacyFile.Location = new System.Drawing.Point(617, 53);
            this.btnGetLegacyFile.Name = "btnGetLegacyFile";
            this.btnGetLegacyFile.Size = new System.Drawing.Size(75, 23);
            this.btnGetLegacyFile.TabIndex = 2;
            this.btnGetLegacyFile.Text = "Browse";
            this.btnGetLegacyFile.UseVisualStyleBackColor = true;
            this.btnGetLegacyFile.Click += new System.EventHandler(this.btnGetLegacyFile_Click);
            // 
            // btnSegmentLegacy
            // 
            this.btnSegmentLegacy.Enabled = false;
            this.btnSegmentLegacy.Location = new System.Drawing.Point(-4, 144);
            this.btnSegmentLegacy.Name = "btnSegmentLegacy";
            this.btnSegmentLegacy.Size = new System.Drawing.Size(112, 52);
            this.btnSegmentLegacy.TabIndex = 3;
            this.btnSegmentLegacy.Text = "Segment  File";
            this.btnSegmentLegacy.UseVisualStyleBackColor = true;
            this.btnSegmentLegacy.Click += new System.EventHandler(this.btnSegmentLegacy_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(-4, 344);
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
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(130, 183);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Lines segmented";
            // 
            // txtLineCount
            // 
            this.txtLineCount.Location = new System.Drawing.Point(223, 176);
            this.txtLineCount.Name = "txtLineCount";
            this.txtLineCount.ReadOnly = true;
            this.txtLineCount.Size = new System.Drawing.Size(108, 20);
            this.txtLineCount.TabIndex = 9;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(-5, 394);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(697, 18);
            this.progressBar1.TabIndex = 10;
            // 
            // boxProgress
            // 
            this.boxProgress.FormattingEnabled = true;
            this.boxProgress.Location = new System.Drawing.Point(496, 176);
            this.boxProgress.Name = "boxProgress";
            this.boxProgress.Size = new System.Drawing.Size(196, 212);
            this.boxProgress.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(568, 160);
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
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 416);
            this.Controls.Add(this.WordsPerLine);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.boxProgress);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtLineCount);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnBrowseOutput);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnSegmentLegacy);
            this.Controls.Add(this.btnGetLegacyFile);
            this.Controls.Add(this.txtLegacy);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Interlinear comparison";
            ((System.ComponentModel.ISupportInitialize)(this.WordsPerLine)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtLegacy;
        private System.Windows.Forms.Button btnGetLegacyFile;
        private System.Windows.Forms.Button btnSegmentLegacy;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnBrowseOutput;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtLineCount;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ListBox boxProgress;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown WordsPerLine;
    }
}

