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
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.DefaultExt = "doc";
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(33, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Legacy file name";
            // 
            // txtLegacy
            // 
            this.txtLegacy.Location = new System.Drawing.Point(138, 53);
            this.txtLegacy.Name = "txtLegacy";
            this.txtLegacy.Size = new System.Drawing.Size(178, 20);
            this.txtLegacy.TabIndex = 1;
            // 
            // btnGetLegacyFile
            // 
            this.btnGetLegacyFile.Location = new System.Drawing.Point(336, 49);
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
            this.btnSegmentLegacy.Location = new System.Drawing.Point(140, 93);
            this.btnSegmentLegacy.Name = "btnSegmentLegacy";
            this.btnSegmentLegacy.Size = new System.Drawing.Size(93, 52);
            this.btnSegmentLegacy.TabIndex = 3;
            this.btnSegmentLegacy.Text = "Segment Legacy File";
            this.btnSegmentLegacy.UseVisualStyleBackColor = true;
            this.btnSegmentLegacy.Click += new System.EventHandler(this.btnSegmentLegacy_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 298);
            this.Controls.Add(this.btnSegmentLegacy);
            this.Controls.Add(this.btnGetLegacyFile);
            this.Controls.Add(this.txtLegacy);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Interlinear comparison";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtLegacy;
        private System.Windows.Forms.Button btnGetLegacyFile;
        private System.Windows.Forms.Button btnSegmentLegacy;
    }
}

