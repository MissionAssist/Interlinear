namespace CharacterCounter
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
            this.label1 = new System.Windows.Forms.Label();
            this.openInputDialogue = new System.Windows.Forms.OpenFileDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.OutputFileBox = new System.Windows.Forms.TextBox();
            this.btnOutputFile = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.saveExcelDialogue = new System.Windows.Forms.SaveFileDialog();
            this.btnAnalyse = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripContainer1 = new System.Windows.Forms.ToolStripContainer();
            this.btnAggregateFile = new System.Windows.Forms.Button();
            this.AggregateStatsBox = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabFonts = new System.Windows.Forms.TabPage();
            this.btnSaveFontList = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.FontList = new System.Windows.Forms.ListBox();
            this.btnListFonts = new System.Windows.Forms.Button();
            this.tabStyles = new System.Windows.Forms.TabPage();
            this.btnSaveStyles = new System.Windows.Forms.Button();
            this.listStyles = new System.Windows.Forms.DataGridView();
            this.Style = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.theDefaultFont = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label5 = new System.Windows.Forms.Label();
            this.btnGetStyles = new System.Windows.Forms.Button();
            this.ErrorTab = new System.Windows.Forms.TabPage();
            this.label11 = new System.Windows.Forms.Label();
            this.btnSaveErrorList = new System.Windows.Forms.Button();
            this.listNormalisedErrors = new System.Windows.Forms.DataGridView();
            this.MappedCharacter = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PossibleCharacter = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnGetEncoding = new System.Windows.Forms.Button();
            this.EncodingTextBox = new System.Windows.Forms.TextBox();
            this.btnGetFont = new System.Windows.Forms.Button();
            this.FontBox = new System.Windows.Forms.TextBox();
            this.FontLabel = new System.Windows.Forms.Label();
            this.btnDecompGlyph = new System.Windows.Forms.Button();
            this.DecompGlyphBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.WriteIndividualFile = new System.Windows.Forms.CheckBox();
            this.btnSaveAggregateStats = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.FileCounter = new System.Windows.Forms.Label();
            this.btnPause = new System.Windows.Forms.Button();
            this.AggregateStats = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.CombDecomposedChars = new System.Windows.Forms.CheckBox();
            this.AnalyseByFont = new System.Windows.Forms.CheckBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.documentationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.CombiningCharacters = new System.Windows.Forms.ToolStripMenuItem();
            this.LicenseMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.IndivOrBulk = new System.Windows.Forms.TabControl();
            this.IndividualFile = new System.Windows.Forms.TabPage();
            this.label2 = new System.Windows.Forms.Label();
            this.btnGetInput = new System.Windows.Forms.Button();
            this.InputFileBox = new System.Windows.Forms.TextBox();
            this.btnSaveXML = new System.Windows.Forms.Button();
            this.XMLFileBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.FontListFileBox = new System.Windows.Forms.TextBox();
            this.StyleListFileBox = new System.Windows.Forms.TextBox();
            this.btnErrorList = new System.Windows.Forms.Button();
            this.ErrorListBox = new System.Windows.Forms.TextBox();
            this.btnFontListFile = new System.Windows.Forms.Button();
            this.btnStyleListFile = new System.Windows.Forms.Button();
            this.btnXMLFile = new System.Windows.Forms.Button();
            this.Bulk = new System.Windows.Forms.TabPage();
            this.btnSelectFiles = new System.Windows.Forms.Button();
            this.OutputFileSuffixBox = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.btnInputFolder = new System.Windows.Forms.Button();
            this.InputFolderBox = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.btnCharStatFolder = new System.Windows.Forms.Button();
            this.OutputFolderBox = new System.Windows.Forms.TextBox();
            this.BulkFontListFileBox = new System.Windows.Forms.TextBox();
            this.BulkStyleListBox = new System.Windows.Forms.TextBox();
            this.btnBulkErrorList = new System.Windows.Forms.Button();
            this.BulkErrorListbox = new System.Windows.Forms.TextBox();
            this.btnBulkFontListFile = new System.Windows.Forms.Button();
            this.btnBulkStyleListFile = new System.Windows.Forms.Button();
            this.toolStripContainer2 = new System.Windows.Forms.ToolStripContainer();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.saveXMLDialogue = new System.Windows.Forms.SaveFileDialog();
            this.toolTipCombine = new System.Windows.Forms.ToolTip(this.components);
            this.OpenGlyphFileDialogue = new System.Windows.Forms.OpenFileDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.FolderDialogue = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1.SuspendLayout();
            this.toolStripContainer1.ContentPanel.SuspendLayout();
            this.toolStripContainer1.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.tabFonts.SuspendLayout();
            this.tabStyles.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listStyles)).BeginInit();
            this.ErrorTab.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listNormalisedErrors)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.IndivOrBulk.SuspendLayout();
            this.IndividualFile.SuspendLayout();
            this.Bulk.SuspendLayout();
            this.toolStripContainer2.ContentPanel.SuspendLayout();
            this.toolStripContainer2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(96, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(453, 37);
            this.label1.TabIndex = 0;
            this.label1.Text = "Count Glyphs in a Document";
            // 
            // openInputDialogue
            // 
            this.openInputDialogue.DefaultExt = "docx";
            this.openInputDialogue.Filter = "Word 2000 files |*.doc|Word 2007+ files |*.docx|Rich Text Format|*.rtf|All files|" +
    " *.*";
            this.openInputDialogue.FilterIndex = 2;
            this.openInputDialogue.Title = "Input File";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(94, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Character stats file";
            // 
            // OutputFileBox
            // 
            this.OutputFileBox.Location = new System.Drawing.Point(122, 54);
            this.OutputFileBox.Name = "OutputFileBox";
            this.OutputFileBox.Size = new System.Drawing.Size(408, 20);
            this.OutputFileBox.TabIndex = 5;
            // 
            // btnOutputFile
            // 
            this.btnOutputFile.Location = new System.Drawing.Point(537, 52);
            this.btnOutputFile.Name = "btnOutputFile";
            this.btnOutputFile.Size = new System.Drawing.Size(75, 23);
            this.btnOutputFile.TabIndex = 6;
            this.btnOutputFile.Text = "Browse";
            this.btnOutputFile.UseVisualStyleBackColor = true;
            this.btnOutputFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(631, 618);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(101, 30);
            this.btnClose.TabIndex = 7;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // saveExcelDialogue
            // 
            this.saveExcelDialogue.Filter = "Excel WorkBook | *.xlsx";
            // 
            // btnAnalyse
            // 
            this.btnAnalyse.Enabled = false;
            this.btnAnalyse.Location = new System.Drawing.Point(631, 360);
            this.btnAnalyse.Name = "btnAnalyse";
            this.btnAnalyse.Size = new System.Drawing.Size(101, 33);
            this.btnAnalyse.TabIndex = 8;
            this.btnAnalyse.Text = "Analyse";
            this.btnAnalyse.UseVisualStyleBackColor = true;
            this.btnAnalyse.Click += new System.EventHandler(this.btnAnalyse_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(4, 640);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(219, 22);
            this.statusStrip1.TabIndex = 9;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(200, 16);
            // 
            // toolStripContainer1
            // 
            // 
            // toolStripContainer1.ContentPanel
            // 
            this.toolStripContainer1.ContentPanel.AutoScroll = true;
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnAggregateFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AggregateStatsBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label13);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.tabControl1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnGetEncoding);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.EncodingTextBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnAnalyse);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnGetFont);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FontBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FontLabel);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnDecompGlyph);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.DecompGlyphBox);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label9);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.WriteIndividualFile);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnSaveAggregateStats);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label14);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.FileCounter);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnPause);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.btnClose);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AggregateStats);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label12);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.CombDecomposedChars);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.AnalyseByFont);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.textBox1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.statusStrip1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.menuStrip1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.label1);
            this.toolStripContainer1.ContentPanel.Controls.Add(this.IndivOrBulk);
            this.toolStripContainer1.ContentPanel.Size = new System.Drawing.Size(767, 662);
            this.toolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer1.LeftToolStripPanelVisible = false;
            this.toolStripContainer1.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer1.Name = "toolStripContainer1";
            this.toolStripContainer1.RightToolStripPanelVisible = false;
            this.toolStripContainer1.Size = new System.Drawing.Size(767, 662);
            this.toolStripContainer1.TabIndex = 10;
            this.toolStripContainer1.Text = "toolStripContainer1";
            this.toolStripContainer1.TopToolStripPanelVisible = false;
            // 
            // btnAggregateFile
            // 
            this.btnAggregateFile.Location = new System.Drawing.Point(554, 567);
            this.btnAggregateFile.Name = "btnAggregateFile";
            this.btnAggregateFile.Size = new System.Drawing.Size(75, 23);
            this.btnAggregateFile.TabIndex = 64;
            this.btnAggregateFile.Text = "Browse";
            this.btnAggregateFile.UseVisualStyleBackColor = true;
            this.btnAggregateFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // AggregateStatsBox
            // 
            this.AggregateStatsBox.Location = new System.Drawing.Point(138, 568);
            this.AggregateStatsBox.Name = "AggregateStatsBox";
            this.AggregateStatsBox.Size = new System.Drawing.Size(408, 20);
            this.AggregateStatsBox.TabIndex = 63;
            this.toolTip1.SetToolTip(this.AggregateStatsBox, "File that aggregates statistics from aggregate input files.\r\n");
            this.AggregateStatsBox.TextChanged += new System.EventHandler(this.AggregateStatsBox_TextChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(11, 572);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(97, 13);
            this.label13.TabIndex = 62;
            this.label13.Text = "Aggregate stats file";
            this.label13.UseMnemonic = false;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabFonts);
            this.tabControl1.Controls.Add(this.tabStyles);
            this.tabControl1.Controls.Add(this.ErrorTab);
            this.tabControl1.Location = new System.Drawing.Point(13, 321);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(611, 183);
            this.tabControl1.TabIndex = 61;
            // 
            // tabFonts
            // 
            this.tabFonts.Controls.Add(this.btnSaveFontList);
            this.tabFonts.Controls.Add(this.label4);
            this.tabFonts.Controls.Add(this.FontList);
            this.tabFonts.Controls.Add(this.btnListFonts);
            this.tabFonts.Location = new System.Drawing.Point(4, 22);
            this.tabFonts.Name = "tabFonts";
            this.tabFonts.Padding = new System.Windows.Forms.Padding(3);
            this.tabFonts.Size = new System.Drawing.Size(603, 157);
            this.tabFonts.TabIndex = 0;
            this.tabFonts.Text = "Get fonts";
            this.tabFonts.UseVisualStyleBackColor = true;
            // 
            // btnSaveFontList
            // 
            this.btnSaveFontList.Enabled = false;
            this.btnSaveFontList.Location = new System.Drawing.Point(10, 98);
            this.btnSaveFontList.Name = "btnSaveFontList";
            this.btnSaveFontList.Size = new System.Drawing.Size(119, 45);
            this.btnSaveFontList.TabIndex = 21;
            this.btnSaveFontList.Text = "Save font list";
            this.btnSaveFontList.UseVisualStyleBackColor = true;
            this.btnSaveFontList.Click += new System.EventHandler(this.btnListFonts_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(-2, 5);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(131, 13);
            this.label4.TabIndex = 20;
            this.label4.Text = "The fonts in the document";
            // 
            // FontList
            // 
            this.FontList.FormattingEnabled = true;
            this.FontList.Location = new System.Drawing.Point(135, 5);
            this.FontList.Name = "FontList";
            this.FontList.Size = new System.Drawing.Size(455, 147);
            this.FontList.TabIndex = 19;
            // 
            // btnListFonts
            // 
            this.btnListFonts.Enabled = false;
            this.btnListFonts.Location = new System.Drawing.Point(10, 47);
            this.btnListFonts.Name = "btnListFonts";
            this.btnListFonts.Size = new System.Drawing.Size(119, 45);
            this.btnListFonts.TabIndex = 18;
            this.btnListFonts.Text = "List the fonts";
            this.btnListFonts.UseVisualStyleBackColor = true;
            this.btnListFonts.Click += new System.EventHandler(this.btnListFonts_Click);
            // 
            // tabStyles
            // 
            this.tabStyles.Controls.Add(this.btnSaveStyles);
            this.tabStyles.Controls.Add(this.listStyles);
            this.tabStyles.Controls.Add(this.label5);
            this.tabStyles.Controls.Add(this.btnGetStyles);
            this.tabStyles.Location = new System.Drawing.Point(4, 22);
            this.tabStyles.Name = "tabStyles";
            this.tabStyles.Padding = new System.Windows.Forms.Padding(3);
            this.tabStyles.Size = new System.Drawing.Size(603, 157);
            this.tabStyles.TabIndex = 1;
            this.tabStyles.Text = "Get Styles";
            this.tabStyles.UseVisualStyleBackColor = true;
            // 
            // btnSaveStyles
            // 
            this.btnSaveStyles.Enabled = false;
            this.btnSaveStyles.Location = new System.Drawing.Point(3, 125);
            this.btnSaveStyles.Name = "btnSaveStyles";
            this.btnSaveStyles.Size = new System.Drawing.Size(102, 24);
            this.btnSaveStyles.TabIndex = 3;
            this.btnSaveStyles.Text = "Save styles";
            this.btnSaveStyles.UseVisualStyleBackColor = true;
            this.btnSaveStyles.Click += new System.EventHandler(this.btnGetStyles_Click);
            // 
            // listStyles
            // 
            this.listStyles.AllowUserToAddRows = false;
            this.listStyles.AllowUserToDeleteRows = false;
            this.listStyles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.listStyles.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Style,
            this.theDefaultFont});
            this.listStyles.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnF2;
            this.listStyles.Location = new System.Drawing.Point(118, 4);
            this.listStyles.MultiSelect = false;
            this.listStyles.Name = "listStyles";
            this.listStyles.ReadOnly = true;
            this.listStyles.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.listStyles.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;
            this.listStyles.RowTemplate.ReadOnly = true;
            this.listStyles.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.listStyles.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.listStyles.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.listStyles.ShowCellToolTips = false;
            this.listStyles.ShowEditingIcon = false;
            this.listStyles.Size = new System.Drawing.Size(484, 150);
            this.listStyles.TabIndex = 2;
            this.listStyles.TabStop = false;
            // 
            // Style
            // 
            this.Style.HeaderText = "Style";
            this.Style.Name = "Style";
            this.Style.ReadOnly = true;
            this.Style.Width = 230;
            // 
            // theDefaultFont
            // 
            this.theDefaultFont.HeaderText = "Default Font";
            this.theDefaultFont.Name = "theDefaultFont";
            this.theDefaultFont.ReadOnly = true;
            this.theDefaultFont.Width = 230;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(113, 92);
            this.label5.TabIndex = 2;
            this.label5.Text = "This lets you inspect\r\nthe styles and their\r\ndefault fonts. It is\r\nthere so you c" +
    "an\r\ncheck if the\r\napplication is\r\nworking properly\r\n";
            this.label5.UseCompatibleTextRendering = true;
            // 
            // btnGetStyles
            // 
            this.btnGetStyles.Enabled = false;
            this.btnGetStyles.Location = new System.Drawing.Point(3, 95);
            this.btnGetStyles.Name = "btnGetStyles";
            this.btnGetStyles.Size = new System.Drawing.Size(102, 24);
            this.btnGetStyles.TabIndex = 0;
            this.btnGetStyles.Text = "Get styles";
            this.btnGetStyles.UseVisualStyleBackColor = true;
            this.btnGetStyles.Click += new System.EventHandler(this.btnGetStyles_Click);
            // 
            // ErrorTab
            // 
            this.ErrorTab.Controls.Add(this.label11);
            this.ErrorTab.Controls.Add(this.btnSaveErrorList);
            this.ErrorTab.Controls.Add(this.listNormalisedErrors);
            this.ErrorTab.Location = new System.Drawing.Point(4, 22);
            this.ErrorTab.Name = "ErrorTab";
            this.ErrorTab.Padding = new System.Windows.Forms.Padding(3);
            this.ErrorTab.Size = new System.Drawing.Size(603, 157);
            this.ErrorTab.TabIndex = 2;
            this.ErrorTab.Text = "Normalisation suggestions";
            this.ErrorTab.UseVisualStyleBackColor = true;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(14, 16);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(71, 65);
            this.label11.TabIndex = 2;
            this.label11.Text = "This lists the\r\nsuggested\r\ncharacters to\r\nuse after\r\nnormalisation.";
            // 
            // btnSaveErrorList
            // 
            this.btnSaveErrorList.Enabled = false;
            this.btnSaveErrorList.Location = new System.Drawing.Point(7, 98);
            this.btnSaveErrorList.Name = "btnSaveErrorList";
            this.btnSaveErrorList.Size = new System.Drawing.Size(87, 53);
            this.btnSaveErrorList.TabIndex = 1;
            this.btnSaveErrorList.Text = "Save data";
            this.btnSaveErrorList.UseVisualStyleBackColor = true;
            this.btnSaveErrorList.Click += new System.EventHandler(this.btnSaveErrorList_Click);
            // 
            // listNormalisedErrors
            // 
            this.listNormalisedErrors.AllowUserToAddRows = false;
            this.listNormalisedErrors.AllowUserToDeleteRows = false;
            this.listNormalisedErrors.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.listNormalisedErrors.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.MappedCharacter,
            this.PossibleCharacter});
            this.listNormalisedErrors.Location = new System.Drawing.Point(107, 3);
            this.listNormalisedErrors.Name = "listNormalisedErrors";
            this.listNormalisedErrors.ReadOnly = true;
            this.listNormalisedErrors.Size = new System.Drawing.Size(490, 150);
            this.listNormalisedErrors.TabIndex = 0;
            // 
            // MappedCharacter
            // 
            this.MappedCharacter.HeaderText = "Mapped Character";
            this.MappedCharacter.Name = "MappedCharacter";
            this.MappedCharacter.ReadOnly = true;
            this.MappedCharacter.Width = 250;
            // 
            // PossibleCharacter
            // 
            this.PossibleCharacter.HeaderText = "Possible Character";
            this.PossibleCharacter.Name = "PossibleCharacter";
            this.PossibleCharacter.ReadOnly = true;
            this.PossibleCharacter.Width = 200;
            // 
            // btnGetEncoding
            // 
            this.btnGetEncoding.Enabled = false;
            this.btnGetEncoding.Location = new System.Drawing.Point(16, 537);
            this.btnGetEncoding.Name = "btnGetEncoding";
            this.btnGetEncoding.Size = new System.Drawing.Size(75, 23);
            this.btnGetEncoding.TabIndex = 59;
            this.btnGetEncoding.Text = "Encoding";
            this.btnGetEncoding.UseVisualStyleBackColor = true;
            this.btnGetEncoding.Click += new System.EventHandler(this.btnGetEncoding_Click);
            // 
            // EncodingTextBox
            // 
            this.EncodingTextBox.Enabled = false;
            this.EncodingTextBox.Location = new System.Drawing.Point(138, 538);
            this.EncodingTextBox.Name = "EncodingTextBox";
            this.EncodingTextBox.Size = new System.Drawing.Size(180, 20);
            this.EncodingTextBox.TabIndex = 60;
            // 
            // btnGetFont
            // 
            this.btnGetFont.Enabled = false;
            this.btnGetFont.Location = new System.Drawing.Point(556, 537);
            this.btnGetFont.Name = "btnGetFont";
            this.btnGetFont.Size = new System.Drawing.Size(75, 23);
            this.btnGetFont.TabIndex = 58;
            this.btnGetFont.Text = "Browse";
            this.btnGetFont.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.toolTip1.SetToolTip(this.btnGetFont, "Get the font you want to use.");
            this.btnGetFont.UseVisualStyleBackColor = true;
            this.btnGetFont.Click += new System.EventHandler(this.btnGetFont_Click);
            // 
            // FontBox
            // 
            this.FontBox.Enabled = false;
            this.FontBox.Location = new System.Drawing.Point(361, 538);
            this.FontBox.Name = "FontBox";
            this.FontBox.Size = new System.Drawing.Size(184, 20);
            this.FontBox.TabIndex = 57;
            this.FontBox.Text = "Calibri";
            // 
            // FontLabel
            // 
            this.FontLabel.AutoSize = true;
            this.FontLabel.Enabled = false;
            this.FontLabel.Location = new System.Drawing.Point(328, 542);
            this.FontLabel.Name = "FontLabel";
            this.FontLabel.Size = new System.Drawing.Size(28, 13);
            this.FontLabel.TabIndex = 56;
            this.FontLabel.Text = "Font";
            // 
            // btnDecompGlyph
            // 
            this.btnDecompGlyph.Location = new System.Drawing.Point(557, 508);
            this.btnDecompGlyph.Name = "btnDecompGlyph";
            this.btnDecompGlyph.Size = new System.Drawing.Size(75, 23);
            this.btnDecompGlyph.TabIndex = 55;
            this.btnDecompGlyph.Text = "Browse";
            this.toolTip1.SetToolTip(this.btnDecompGlyph, "Browse for the font for text files");
            this.btnDecompGlyph.UseVisualStyleBackColor = true;
            this.btnDecompGlyph.Click += new System.EventHandler(this.btnGetInput_Click);
            // 
            // DecompGlyphBox
            // 
            this.DecompGlyphBox.Location = new System.Drawing.Point(137, 509);
            this.DecompGlyphBox.Name = "DecompGlyphBox";
            this.DecompGlyphBox.Size = new System.Drawing.Size(408, 20);
            this.DecompGlyphBox.TabIndex = 54;
            this.toolTip1.SetToolTip(this.DecompGlyphBox, "An Excel files with the decomposed characters and their fonts in cells A2 downwar" +
        "ds.  It is only useful if you are analysing by font.");
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(17, 513);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(114, 13);
            this.label9.TabIndex = 53;
            this.label9.Text = "Decomposed glyph file";
            // 
            // WriteIndividualFile
            // 
            this.WriteIndividualFile.AutoSize = true;
            this.WriteIndividualFile.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.WriteIndividualFile.Checked = true;
            this.WriteIndividualFile.CheckState = System.Windows.Forms.CheckState.Checked;
            this.WriteIndividualFile.Location = new System.Drawing.Point(127, 596);
            this.WriteIndividualFile.Name = "WriteIndividualFile";
            this.WriteIndividualFile.Size = new System.Drawing.Size(114, 17);
            this.WriteIndividualFile.TabIndex = 51;
            this.WriteIndividualFile.Text = "Write individual file";
            this.WriteIndividualFile.UseVisualStyleBackColor = true;
            // 
            // btnSaveAggregateStats
            // 
            this.btnSaveAggregateStats.Enabled = false;
            this.btnSaveAggregateStats.Location = new System.Drawing.Point(631, 564);
            this.btnSaveAggregateStats.Name = "btnSaveAggregateStats";
            this.btnSaveAggregateStats.Size = new System.Drawing.Size(101, 26);
            this.btnSaveAggregateStats.TabIndex = 50;
            this.btnSaveAggregateStats.Text = "Save aggregate";
            this.toolTip1.SetToolTip(this.btnSaveAggregateStats, "Save aggregate statistics");
            this.btnSaveAggregateStats.UseVisualStyleBackColor = true;
            this.btnSaveAggregateStats.Click += new System.EventHandler(this.btnSaveAggregateStats_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(259, 619);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(70, 13);
            this.label14.TabIndex = 49;
            this.label14.Text = "Files counted";
            // 
            // FileCounter
            // 
            this.FileCounter.AutoSize = true;
            this.FileCounter.Location = new System.Drawing.Point(335, 619);
            this.FileCounter.Name = "FileCounter";
            this.FileCounter.Size = new System.Drawing.Size(13, 13);
            this.FileCounter.TabIndex = 48;
            this.FileCounter.Text = "0";
            // 
            // btnPause
            // 
            this.btnPause.Enabled = false;
            this.btnPause.Location = new System.Drawing.Point(631, 405);
            this.btnPause.Name = "btnPause";
            this.btnPause.Size = new System.Drawing.Size(101, 30);
            this.btnPause.TabIndex = 11;
            this.btnPause.Text = "Pause";
            this.btnPause.UseVisualStyleBackColor = true;
            this.btnPause.Click += new System.EventHandler(this.btnPause_Click);
            // 
            // AggregateStats
            // 
            this.AggregateStats.AutoSize = true;
            this.AggregateStats.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.AggregateStats.Enabled = false;
            this.AggregateStats.Location = new System.Drawing.Point(127, 617);
            this.AggregateStats.Name = "AggregateStats";
            this.AggregateStats.Size = new System.Drawing.Size(121, 17);
            this.AggregateStats.TabIndex = 44;
            this.AggregateStats.Text = "Aggregate File Stats";
            this.AggregateStats.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(61, 61);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(525, 13);
            this.label12.TabIndex = 36;
            this.label12.Text = "This program runs hidden copies of Word and Excel. Don\'t close instances of them " +
    "before closing this program";
            // 
            // CombDecomposedChars
            // 
            this.CombDecomposedChars.AutoSize = true;
            this.CombDecomposedChars.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CombDecomposedChars.Checked = true;
            this.CombDecomposedChars.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CombDecomposedChars.Cursor = System.Windows.Forms.Cursors.Default;
            this.CombDecomposedChars.Location = new System.Drawing.Point(402, 596);
            this.CombDecomposedChars.Name = "CombDecomposedChars";
            this.CombDecomposedChars.Size = new System.Drawing.Size(184, 17);
            this.CombDecomposedChars.TabIndex = 29;
            this.CombDecomposedChars.Text = "Combine decomposed characters";
            this.toolTipCombine.SetToolTip(this.CombDecomposedChars, "Some characters have to be mapped to two Unicode characters that are displayed to" +
        "gether as a single glyph. \r\nChecking this box causes the program to try to count" +
        " them as a single character.");
            this.CombDecomposedChars.UseVisualStyleBackColor = true;
            this.CombDecomposedChars.CheckStateChanged += new System.EventHandler(this.CombDecomposedChars_CheckStateChanged);
            // 
            // AnalyseByFont
            // 
            this.AnalyseByFont.AutoSize = true;
            this.AnalyseByFont.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.AnalyseByFont.Checked = true;
            this.AnalyseByFont.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AnalyseByFont.Location = new System.Drawing.Point(262, 595);
            this.AnalyseByFont.Name = "AnalyseByFont";
            this.AnalyseByFont.Size = new System.Drawing.Size(101, 17);
            this.AnalyseByFont.TabIndex = 14;
            this.AnalyseByFont.Text = "Analyse by Font";
            this.toolTip1.SetToolTip(this.AnalyseByFont, "Analyse by Font is slightly slower.");
            this.AnalyseByFont.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Location = new System.Drawing.Point(64, 64);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(452, 13);
            this.textBox1.TabIndex = 13;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Dock = System.Windows.Forms.DockStyle.None;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(4, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(89, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(92, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.documentationToolStripMenuItem,
            this.CombiningCharacters,
            this.LicenseMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // documentationToolStripMenuItem
            // 
            this.documentationToolStripMenuItem.Name = "documentationToolStripMenuItem";
            this.documentationToolStripMenuItem.Size = new System.Drawing.Size(291, 22);
            this.documentationToolStripMenuItem.Text = "Documentation";
            this.documentationToolStripMenuItem.Click += new System.EventHandler(this.documentationToolStripMenuItem_Click);
            // 
            // CombiningCharacters
            // 
            this.CombiningCharacters.Name = "CombiningCharacters";
            this.CombiningCharacters.Size = new System.Drawing.Size(291, 22);
            this.CombiningCharacters.Text = "Legacy Combining Characters Workbook";
            this.CombiningCharacters.Click += new System.EventHandler(this.CombiningCharacters_Click);
            // 
            // LicenseMenuItem
            // 
            this.LicenseMenuItem.Name = "LicenseMenuItem";
            this.LicenseMenuItem.Size = new System.Drawing.Size(291, 22);
            this.LicenseMenuItem.Text = "License";
            this.LicenseMenuItem.Click += new System.EventHandler(this.LicenseMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(291, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // IndivOrBulk
            // 
            this.IndivOrBulk.Controls.Add(this.IndividualFile);
            this.IndivOrBulk.Controls.Add(this.Bulk);
            this.IndivOrBulk.Location = new System.Drawing.Point(10, 81);
            this.IndivOrBulk.Name = "IndivOrBulk";
            this.IndivOrBulk.SelectedIndex = 0;
            this.IndivOrBulk.Size = new System.Drawing.Size(732, 234);
            this.IndivOrBulk.TabIndex = 52;
            // 
            // IndividualFile
            // 
            this.IndividualFile.Controls.Add(this.label2);
            this.IndividualFile.Controls.Add(this.btnGetInput);
            this.IndividualFile.Controls.Add(this.InputFileBox);
            this.IndividualFile.Controls.Add(this.btnSaveXML);
            this.IndividualFile.Controls.Add(this.XMLFileBox);
            this.IndividualFile.Controls.Add(this.label3);
            this.IndividualFile.Controls.Add(this.label6);
            this.IndividualFile.Controls.Add(this.label7);
            this.IndividualFile.Controls.Add(this.label10);
            this.IndividualFile.Controls.Add(this.label8);
            this.IndividualFile.Controls.Add(this.btnOutputFile);
            this.IndividualFile.Controls.Add(this.OutputFileBox);
            this.IndividualFile.Controls.Add(this.FontListFileBox);
            this.IndividualFile.Controls.Add(this.StyleListFileBox);
            this.IndividualFile.Controls.Add(this.btnErrorList);
            this.IndividualFile.Controls.Add(this.ErrorListBox);
            this.IndividualFile.Controls.Add(this.btnFontListFile);
            this.IndividualFile.Controls.Add(this.btnStyleListFile);
            this.IndividualFile.Controls.Add(this.btnXMLFile);
            this.IndividualFile.Location = new System.Drawing.Point(4, 22);
            this.IndividualFile.Name = "IndividualFile";
            this.IndividualFile.Padding = new System.Windows.Forms.Padding(3);
            this.IndividualFile.Size = new System.Drawing.Size(724, 208);
            this.IndividualFile.TabIndex = 0;
            this.IndividualFile.Text = "Individual";
            this.IndividualFile.UseVisualStyleBackColor = true;
            this.IndividualFile.Click += new System.EventHandler(this.Bulk_Entered);
            this.IndividualFile.Enter += new System.EventHandler(this.Individual_Entered);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 30);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 29;
            this.label2.Text = "Input Document";
            // 
            // btnGetInput
            // 
            this.btnGetInput.Location = new System.Drawing.Point(537, 27);
            this.btnGetInput.Name = "btnGetInput";
            this.btnGetInput.Size = new System.Drawing.Size(75, 23);
            this.btnGetInput.TabIndex = 45;
            this.btnGetInput.Text = "Browse";
            this.btnGetInput.UseVisualStyleBackColor = true;
            this.btnGetInput.Click += new System.EventHandler(this.btnGetInput_Click);
            // 
            // InputFileBox
            // 
            this.InputFileBox.Location = new System.Drawing.Point(123, 27);
            this.InputFileBox.Name = "InputFileBox";
            this.InputFileBox.Size = new System.Drawing.Size(408, 20);
            this.InputFileBox.TabIndex = 44;
            this.InputFileBox.TextChanged += new System.EventHandler(this.InputFileBox_TextChanged);
            // 
            // btnSaveXML
            // 
            this.btnSaveXML.Enabled = false;
            this.btnSaveXML.Location = new System.Drawing.Point(626, 158);
            this.btnSaveXML.Name = "btnSaveXML";
            this.btnSaveXML.Size = new System.Drawing.Size(92, 30);
            this.btnSaveXML.TabIndex = 28;
            this.btnSaveXML.Text = "Save XML";
            this.btnSaveXML.UseVisualStyleBackColor = true;
            this.btnSaveXML.Click += new System.EventHandler(this.btnSaveXML_Click);
            // 
            // XMLFileBox
            // 
            this.XMLFileBox.Location = new System.Drawing.Point(122, 162);
            this.XMLFileBox.Name = "XMLFileBox";
            this.XMLFileBox.Size = new System.Drawing.Size(408, 20);
            this.XMLFileBox.TabIndex = 26;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(11, 80);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(59, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Font list file";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 115);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(61, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Style list file";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(11, 141);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(60, 13);
            this.label10.TabIndex = 33;
            this.label10.Text = "Error list file";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(11, 167);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 13);
            this.label8.TabIndex = 25;
            this.label8.Text = "XML file";
            // 
            // FontListFileBox
            // 
            this.FontListFileBox.Location = new System.Drawing.Point(123, 81);
            this.FontListFileBox.Name = "FontListFileBox";
            this.FontListFileBox.Size = new System.Drawing.Size(408, 20);
            this.FontListFileBox.TabIndex = 20;
            this.FontListFileBox.TextChanged += new System.EventHandler(this.FontListFileBox_TextChanged);
            // 
            // StyleListFileBox
            // 
            this.StyleListFileBox.Location = new System.Drawing.Point(122, 108);
            this.StyleListFileBox.Name = "StyleListFileBox";
            this.StyleListFileBox.Size = new System.Drawing.Size(408, 20);
            this.StyleListFileBox.TabIndex = 23;
            this.StyleListFileBox.TextChanged += new System.EventHandler(this.StyleListFileBox_TextChanged);
            // 
            // btnErrorList
            // 
            this.btnErrorList.Location = new System.Drawing.Point(537, 136);
            this.btnErrorList.Name = "btnErrorList";
            this.btnErrorList.Size = new System.Drawing.Size(75, 23);
            this.btnErrorList.TabIndex = 35;
            this.btnErrorList.Text = "Browse";
            this.btnErrorList.UseVisualStyleBackColor = true;
            this.btnErrorList.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // ErrorListBox
            // 
            this.ErrorListBox.Location = new System.Drawing.Point(122, 135);
            this.ErrorListBox.Name = "ErrorListBox";
            this.ErrorListBox.Size = new System.Drawing.Size(408, 20);
            this.ErrorListBox.TabIndex = 34;
            // 
            // btnFontListFile
            // 
            this.btnFontListFile.Location = new System.Drawing.Point(537, 80);
            this.btnFontListFile.Name = "btnFontListFile";
            this.btnFontListFile.Size = new System.Drawing.Size(75, 23);
            this.btnFontListFile.TabIndex = 21;
            this.btnFontListFile.Text = "Browse";
            this.btnFontListFile.UseVisualStyleBackColor = true;
            this.btnFontListFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // btnStyleListFile
            // 
            this.btnStyleListFile.Location = new System.Drawing.Point(537, 107);
            this.btnStyleListFile.Name = "btnStyleListFile";
            this.btnStyleListFile.Size = new System.Drawing.Size(75, 23);
            this.btnStyleListFile.TabIndex = 24;
            this.btnStyleListFile.Text = "Browse";
            this.btnStyleListFile.UseVisualStyleBackColor = true;
            this.btnStyleListFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // btnXMLFile
            // 
            this.btnXMLFile.Location = new System.Drawing.Point(537, 163);
            this.btnXMLFile.Name = "btnXMLFile";
            this.btnXMLFile.Size = new System.Drawing.Size(75, 23);
            this.btnXMLFile.TabIndex = 27;
            this.btnXMLFile.Text = "Browse";
            this.btnXMLFile.UseVisualStyleBackColor = true;
            this.btnXMLFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // Bulk
            // 
            this.Bulk.Controls.Add(this.btnSelectFiles);
            this.Bulk.Controls.Add(this.OutputFileSuffixBox);
            this.Bulk.Controls.Add(this.label21);
            this.Bulk.Controls.Add(this.label15);
            this.Bulk.Controls.Add(this.btnInputFolder);
            this.Bulk.Controls.Add(this.InputFolderBox);
            this.Bulk.Controls.Add(this.label16);
            this.Bulk.Controls.Add(this.label17);
            this.Bulk.Controls.Add(this.label18);
            this.Bulk.Controls.Add(this.label19);
            this.Bulk.Controls.Add(this.btnCharStatFolder);
            this.Bulk.Controls.Add(this.OutputFolderBox);
            this.Bulk.Controls.Add(this.BulkFontListFileBox);
            this.Bulk.Controls.Add(this.BulkStyleListBox);
            this.Bulk.Controls.Add(this.btnBulkErrorList);
            this.Bulk.Controls.Add(this.BulkErrorListbox);
            this.Bulk.Controls.Add(this.btnBulkFontListFile);
            this.Bulk.Controls.Add(this.btnBulkStyleListFile);
            this.Bulk.Location = new System.Drawing.Point(4, 22);
            this.Bulk.Name = "Bulk";
            this.Bulk.Padding = new System.Windows.Forms.Padding(3);
            this.Bulk.Size = new System.Drawing.Size(724, 208);
            this.Bulk.TabIndex = 1;
            this.Bulk.Text = "Bulk";
            this.Bulk.UseVisualStyleBackColor = true;
            this.Bulk.Enter += new System.EventHandler(this.Bulk_Entered);
            // 
            // btnSelectFiles
            // 
            this.btnSelectFiles.Enabled = false;
            this.btnSelectFiles.Location = new System.Drawing.Point(630, 5);
            this.btnSelectFiles.Name = "btnSelectFiles";
            this.btnSelectFiles.Size = new System.Drawing.Size(87, 35);
            this.btnSelectFiles.TabIndex = 68;
            this.btnSelectFiles.Text = "Select files";
            this.btnSelectFiles.UseVisualStyleBackColor = true;
            this.btnSelectFiles.Click += new System.EventHandler(this.btnSelectFiles_Click);
            // 
            // OutputFileSuffixBox
            // 
            this.OutputFileSuffixBox.Location = new System.Drawing.Point(502, 41);
            this.OutputFileSuffixBox.Name = "OutputFileSuffixBox";
            this.OutputFileSuffixBox.Size = new System.Drawing.Size(122, 20);
            this.OutputFileSuffixBox.TabIndex = 65;
            this.OutputFileSuffixBox.TextChanged += new System.EventHandler(this.OutputFileSuffixBox_TextChanged);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(401, 45);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(79, 13);
            this.label21.TabIndex = 64;
            this.label21.Text = "File name suffix";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(23, 16);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(60, 13);
            this.label15.TabIndex = 58;
            this.label15.Text = "Input folder";
            // 
            // btnInputFolder
            // 
            this.btnInputFolder.Location = new System.Drawing.Point(549, 13);
            this.btnInputFolder.Name = "btnInputFolder";
            this.btnInputFolder.Size = new System.Drawing.Size(75, 23);
            this.btnInputFolder.TabIndex = 63;
            this.btnInputFolder.Text = "Browse";
            this.btnInputFolder.UseVisualStyleBackColor = true;
            this.btnInputFolder.Click += new System.EventHandler(this.btnInputFolder_Click);
            // 
            // InputFolderBox
            // 
            this.InputFolderBox.Location = new System.Drawing.Point(135, 13);
            this.InputFolderBox.Name = "InputFolderBox";
            this.InputFolderBox.Size = new System.Drawing.Size(407, 20);
            this.InputFolderBox.TabIndex = 62;
            this.InputFolderBox.TextChanged += new System.EventHandler(this.InputFolderBox_TextChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(23, 45);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(107, 13);
            this.label16.TabIndex = 46;
            this.label16.Text = "Character stats folder";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(23, 66);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(59, 13);
            this.label17.TabIndex = 49;
            this.label17.Text = "Font list file";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(23, 101);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(61, 13);
            this.label18.TabIndex = 52;
            this.label18.Text = "Style list file";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(23, 127);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(60, 13);
            this.label19.TabIndex = 59;
            this.label19.Text = "Error list file";
            // 
            // btnCharStatFolder
            // 
            this.btnCharStatFolder.Location = new System.Drawing.Point(310, 40);
            this.btnCharStatFolder.Name = "btnCharStatFolder";
            this.btnCharStatFolder.Size = new System.Drawing.Size(75, 23);
            this.btnCharStatFolder.TabIndex = 48;
            this.btnCharStatFolder.Text = "Browse";
            this.btnCharStatFolder.UseVisualStyleBackColor = true;
            this.btnCharStatFolder.Click += new System.EventHandler(this.btnCharStatFolder_Click);
            // 
            // OutputFolderBox
            // 
            this.OutputFolderBox.Location = new System.Drawing.Point(134, 41);
            this.OutputFolderBox.Name = "OutputFolderBox";
            this.OutputFolderBox.Size = new System.Drawing.Size(170, 20);
            this.OutputFolderBox.TabIndex = 47;
            // 
            // BulkFontListFileBox
            // 
            this.BulkFontListFileBox.Location = new System.Drawing.Point(135, 67);
            this.BulkFontListFileBox.Name = "BulkFontListFileBox";
            this.BulkFontListFileBox.Size = new System.Drawing.Size(408, 20);
            this.BulkFontListFileBox.TabIndex = 50;
            // 
            // BulkStyleListBox
            // 
            this.BulkStyleListBox.Location = new System.Drawing.Point(134, 94);
            this.BulkStyleListBox.Name = "BulkStyleListBox";
            this.BulkStyleListBox.Size = new System.Drawing.Size(408, 20);
            this.BulkStyleListBox.TabIndex = 53;
            // 
            // btnBulkErrorList
            // 
            this.btnBulkErrorList.Location = new System.Drawing.Point(549, 122);
            this.btnBulkErrorList.Name = "btnBulkErrorList";
            this.btnBulkErrorList.Size = new System.Drawing.Size(75, 23);
            this.btnBulkErrorList.TabIndex = 61;
            this.btnBulkErrorList.Text = "Browse";
            this.btnBulkErrorList.UseVisualStyleBackColor = true;
            this.btnBulkErrorList.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // BulkErrorListbox
            // 
            this.BulkErrorListbox.Location = new System.Drawing.Point(134, 121);
            this.BulkErrorListbox.Name = "BulkErrorListbox";
            this.BulkErrorListbox.Size = new System.Drawing.Size(408, 20);
            this.BulkErrorListbox.TabIndex = 60;
            this.BulkErrorListbox.TextChanged += new System.EventHandler(this.BulkErrorListbox_TextChanged);
            // 
            // btnBulkFontListFile
            // 
            this.btnBulkFontListFile.Location = new System.Drawing.Point(549, 66);
            this.btnBulkFontListFile.Name = "btnBulkFontListFile";
            this.btnBulkFontListFile.Size = new System.Drawing.Size(75, 23);
            this.btnBulkFontListFile.TabIndex = 51;
            this.btnBulkFontListFile.Text = "Browse";
            this.btnBulkFontListFile.UseVisualStyleBackColor = true;
            this.btnBulkFontListFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // btnBulkStyleListFile
            // 
            this.btnBulkStyleListFile.Location = new System.Drawing.Point(549, 93);
            this.btnBulkStyleListFile.Name = "btnBulkStyleListFile";
            this.btnBulkStyleListFile.Size = new System.Drawing.Size(75, 23);
            this.btnBulkStyleListFile.TabIndex = 54;
            this.btnBulkStyleListFile.Text = "Browse";
            this.btnBulkStyleListFile.UseVisualStyleBackColor = true;
            this.btnBulkStyleListFile.Click += new System.EventHandler(this.btnGetOutput_Click);
            // 
            // toolStripContainer2
            // 
            // 
            // toolStripContainer2.ContentPanel
            // 
            this.toolStripContainer2.ContentPanel.AutoScroll = true;
            this.toolStripContainer2.ContentPanel.Controls.Add(this.toolStripContainer1);
            this.toolStripContainer2.ContentPanel.Size = new System.Drawing.Size(767, 662);
            this.toolStripContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.toolStripContainer2.LeftToolStripPanelVisible = false;
            this.toolStripContainer2.Location = new System.Drawing.Point(0, 0);
            this.toolStripContainer2.Name = "toolStripContainer2";
            this.toolStripContainer2.RightToolStripPanelVisible = false;
            this.toolStripContainer2.Size = new System.Drawing.Size(767, 662);
            this.toolStripContainer2.TabIndex = 11;
            this.toolStripContainer2.Text = "toolStripContainer2";
            this.toolStripContainer2.TopToolStripPanelVisible = false;
            // 
            // toolTip1
            // 
            this.toolTip1.ToolTipTitle = "Analyse by Font";
            // 
            // saveXMLDialogue
            // 
            this.saveXMLDialogue.DefaultExt = "xml";
            this.saveXMLDialogue.Filter = "XML File | *.xml";
            // 
            // toolTipCombine
            // 
            this.toolTipCombine.BackColor = System.Drawing.Color.Khaki;
            this.toolTipCombine.ToolTipTitle = "Combine decomposed characters";
            // 
            // OpenGlyphFileDialogue
            // 
            this.OpenGlyphFileDialogue.DefaultExt = "xlsm";
            this.OpenGlyphFileDialogue.Filter = "Excel Files | *.xlsx|Excel Macro Enabled Files |*.xlsm";
            this.OpenGlyphFileDialogue.FilterIndex = 2;
            this.OpenGlyphFileDialogue.Title = "Decomposed Glyph File";
            // 
            // fontDialog1
            // 
            this.fontDialog1.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            // 
            // FolderDialogue
            // 
            this.FolderDialogue.Description = "Select the input folder you want to analyse.";
            this.FolderDialogue.RootFolder = System.Environment.SpecialFolder.MyComputer;
            this.FolderDialogue.ShowNewFolderButton = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 662);
            this.Controls.Add(this.toolStripContainer2);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Count Glyphs";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolStripContainer1.ContentPanel.ResumeLayout(false);
            this.toolStripContainer1.ContentPanel.PerformLayout();
            this.toolStripContainer1.ResumeLayout(false);
            this.toolStripContainer1.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            this.tabFonts.ResumeLayout(false);
            this.tabFonts.PerformLayout();
            this.tabStyles.ResumeLayout(false);
            this.tabStyles.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listStyles)).EndInit();
            this.ErrorTab.ResumeLayout(false);
            this.ErrorTab.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.listNormalisedErrors)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.IndivOrBulk.ResumeLayout(false);
            this.IndividualFile.ResumeLayout(false);
            this.IndividualFile.PerformLayout();
            this.Bulk.ResumeLayout(false);
            this.Bulk.PerformLayout();
            this.toolStripContainer2.ContentPanel.ResumeLayout(false);
            this.toolStripContainer2.ResumeLayout(false);
            this.toolStripContainer2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openInputDialogue;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox OutputFileBox;
        private System.Windows.Forms.Button btnOutputFile;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.SaveFileDialog saveExcelDialogue;
        private System.Windows.Forms.Button btnAnalyse;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripContainer toolStripContainer1;
        private System.Windows.Forms.ToolStripContainer toolStripContainer2;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem documentationToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem LicenseMenuItem;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnPause;
        private System.Windows.Forms.CheckBox AnalyseByFont;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btnFontListFile;
        private System.Windows.Forms.TextBox FontListFileBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnXMLFile;
        private System.Windows.Forms.TextBox XMLFileBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnStyleListFile;
        private System.Windows.Forms.TextBox StyleListFileBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.SaveFileDialog saveXMLDialogue;
        private System.Windows.Forms.Button btnSaveXML;
        private System.Windows.Forms.CheckBox CombDecomposedChars;
        private System.Windows.Forms.ToolTip toolTipCombine;
        private System.Windows.Forms.OpenFileDialog OpenGlyphFileDialogue;
        private System.Windows.Forms.Button btnErrorList;
        private System.Windows.Forms.TextBox ErrorListBox;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label FileCounter;
        private System.Windows.Forms.CheckBox AggregateStats;
        private System.Windows.Forms.Button btnSaveAggregateStats;
        private System.Windows.Forms.CheckBox WriteIndividualFile;
        private System.Windows.Forms.ToolStripMenuItem CombiningCharacters;
        private System.Windows.Forms.TabControl IndivOrBulk;
        private System.Windows.Forms.TabPage IndividualFile;
        private System.Windows.Forms.Button btnGetInput;
        private System.Windows.Forms.TextBox InputFileBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TabPage Bulk;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabFonts;
        private System.Windows.Forms.Button btnSaveFontList;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ListBox FontList;
        private System.Windows.Forms.Button btnListFonts;
        private System.Windows.Forms.TabPage tabStyles;
        private System.Windows.Forms.Button btnSaveStyles;
        internal System.Windows.Forms.DataGridView listStyles;
        private System.Windows.Forms.DataGridViewTextBoxColumn Style;
        private System.Windows.Forms.DataGridViewTextBoxColumn theDefaultFont;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnGetStyles;
        private System.Windows.Forms.TabPage ErrorTab;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button btnSaveErrorList;
        private System.Windows.Forms.DataGridView listNormalisedErrors;
        private System.Windows.Forms.DataGridViewTextBoxColumn MappedCharacter;
        private System.Windows.Forms.DataGridViewTextBoxColumn PossibleCharacter;
        private System.Windows.Forms.Button btnGetEncoding;
        private System.Windows.Forms.TextBox EncodingTextBox;
        private System.Windows.Forms.Button btnGetFont;
        private System.Windows.Forms.TextBox FontBox;
        private System.Windows.Forms.Label FontLabel;
        private System.Windows.Forms.Button btnDecompGlyph;
        private System.Windows.Forms.TextBox DecompGlyphBox;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox OutputFileSuffixBox;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button btnInputFolder;
        private System.Windows.Forms.TextBox InputFolderBox;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.Button btnCharStatFolder;
        private System.Windows.Forms.TextBox OutputFolderBox;
        private System.Windows.Forms.TextBox BulkFontListFileBox;
        private System.Windows.Forms.TextBox BulkStyleListBox;
        private System.Windows.Forms.Button btnBulkErrorList;
        private System.Windows.Forms.TextBox BulkErrorListbox;
        private System.Windows.Forms.Button btnBulkFontListFile;
        private System.Windows.Forms.Button btnBulkStyleListFile;
        private System.Windows.Forms.Button btnAggregateFile;
        private System.Windows.Forms.TextBox AggregateStatsBox;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnSelectFiles;
        private System.Windows.Forms.FolderBrowserDialog FolderDialogue;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

