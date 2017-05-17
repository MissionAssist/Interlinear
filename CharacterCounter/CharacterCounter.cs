/*
 *  This program counts characters in MS Word documents and writes the font, character code, glyph and count to an Excel workbook.
 *  It counts the characters in headers, footers, text boxes and other such containers as well as the main text.
 *  By storing the character information in a dictionary, we count the characters associated with each font in the document.
 *  It queries the XML representation of a Word document as this is much faster than using the Word object model and allows us to
 *  count the occurrences of characters entered using Insert Symbol.
 *  
 *  The copyright is owned by MissionAssist as the work was carried out on their behalf.
 * 
 *  Written by Stephen Palmstrom, last modified 11 December 2015
 *  
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using System.Xml;
using System.Xml.XPath;
using System.Runtime.InteropServices;

using Microsoft.Win32;
using WordApp = Microsoft.Office.Interop.Word._Application;
using WordRoot = Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word._Document;
using ExcelApp = Microsoft.Office.Interop.Excel._Application;
using ExcelRoot = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel.Application;
using WorkBook = Microsoft.Office.Interop.Excel._Workbook;


namespace CharacterCounter
{
    public partial class CharacterCounter : Form
    {
        // Some global variables
        private WordApp wrdApp;
        private ExcelApp excelApp;
        object missing = Type.Missing;
        private Document theDocument = null;
        private string InputDir = "";
        private string OutputDir = "";
        private string StyleDir = "";
        private string FontDir = "";
        private string XMLDir = "";
        private string GlyphDir = "";
        private string ErrorDir = "";
        private string AggregateDir = "";

        private string theFirstFont = "";

        private bool AggregateSaved = false;
        private bool Individual = true;
        private bool WordWasRunning = false;
        private bool ExcelWasRunning = false;

        private int FileType = WordDoc;
        //private XmlDocument theXMLDocument = new XmlDocument();  // make it global so we can save it.
        //private string theTextDocument = "";  // This is to load a taxt document
        // needed for XML lookup
        const string wordmlNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
        const string wordmlxNamespace = "http://schemas.microsoft.com/office/word/2003/auxHint";
        const string userRoot = "HKEY_CURRENT_USER";
        const string subkey = "Software\\MissionAssist\\CountCharacters";
        const string keyName = userRoot + "\\" + subkey;

        private Dictionary<string, string> theStyleDictionary = new Dictionary<string, string>(10); // to hold all defined styles
        private Dictionary<string, string> theDefaultStyleDictionary = new Dictionary<string, string>(5); // to hold all default styles
        private Dictionary<string, string> theBreakDictionary = new Dictionary<string, string>(3); // to hold the characters corresponding to breaks
        private Dictionary<string, string> theGlyphDictionary = new Dictionary<string, string>(5); // to hold the regular expressions for decomposed characters by font.
        private Dictionary<string, Encoding> theEncodingDictionary = new Dictionary<string, Encoding>(30); // to hold the encoding dictionary
        private Dictionary<CharacterDescriptor, int> theAggregateDictionary =
        new Dictionary<CharacterDescriptor, int>(500, new CharacterEqualityComparer());
        private string AggregateFileList = "";


        private Dictionary<string, XmlDocument> theXMLDictionary = new Dictionary<string, XmlDocument>(10); // To hold the XML documents
        private Dictionary<string, string> theTextDictionary = new Dictionary<string, string>(10);  // to hold the text documents

        private string InputFileName = null;
        private bool GlyphsLoaded = false;
        private Encoding theEncoding = null;  // A place to store the encoding
        private XmlNamespaceManager nsManager;

        //
        //  Special characters
        //
        string[] SpecialCharacterKeys = { "w:endnoteRef", "w:footnoteRef", "w:tab", "w:noBreakHyphen", 
                                            "w:softHyphen", "w:separator", "w:continuationSeparator"};
        string[] SpecialCharacterValues =
            { 
                Convert.ToString("\x0002"),
                Convert.ToString("\x0002"),
                "\t",
                Convert.ToString("\x001E"),
                Convert.ToString("\x001F"),
                Convert.ToString("\x0003"),
                Convert.ToString("\x0004")
           };
        /*
         * Variables for handling Pause and Resume
         */
        private bool Paused = false;
        private bool CloseApp = false;

        /*
         * Character codes can be displayed as decimal, hexadecimal or USV
         */
        const int dec = 0;
        const int hex = 1;
        const int USV = 2;
        /*
         * File types
         */
        const int TextDoc = 0;
        const int WordDoc = 1;

        public CharacterCounter()
        {
            InitializeComponent();
            System.Windows.Forms.Application.ApplicationExit += new EventHandler(this.CloseApps);
            // Start Word and Excel
            try
            {
                wrdApp = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "Word.Application") as Word;
                WordWasRunning = true; // Remember we were running Word
            }
            catch
            {
                /*
                 * Word isn't running, so we run it.
                 */
                wrdApp = new Word();
                WordWasRunning = false;
            }
            wrdApp.Visible = false;
            try
            {
                excelApp = System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "Excel.Application") as Excel;
                ExcelWasRunning = true; // Remember we were running Excel
            }
            catch
            {
                /*
                 * Excel isn't running, so we run it.
                 */
                excelApp = new Excel();
                ExcelWasRunning = false;
            }
            excelApp.Visible = false;
            /*
             * If the registry subkey doesn't exist, create it
             */
            if (Registry.CurrentUser.OpenSubKey(subkey, true) == null)
            {
                Registry.CurrentUser.CreateSubKey(subkey);
            }
            Registry.CurrentUser.Close(); // Close it
            //
            // Some registry settings
            //
            try
            {
                InputDir = GetDirectory("InputDir");
                InputFolderBox.Text = InputDir;
                OutputDir = GetDirectory("OutputDir", InputDir);
                OutputFolderBox.Text = OutputDir;
                StyleDir = GetDirectory("StyleDir", OutputDir);
                FontDir = GetDirectory("FontDir", OutputDir);
                XMLDir = GetDirectory("XMLDir", OutputDir);
                ErrorDir = GetDirectory("ErrorDir", OutputDir);
                GlyphDir = GetDirectory("GlyphDir", InputDir);
                AggregateDir = GetDirectory("AggregateDir", OutputDir);
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Failed to get directories", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseApps();
            }
            //
            //  Types of break
            //

            theBreakDictionary.Add("page", "\f");
            theBreakDictionary.Add("column", Convert.ToString("\x000E"));
            theBreakDictionary.Add("text-wrapping", "\v");



        }
        private string GetDirectory(string ValueName, string DefaultPath = "")
        {
            string theDirectory = "";
            if (DefaultPath == "")
            {
                DefaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            }
            try
            {
                if (Registry.GetValue(keyName, ValueName, DefaultPath) != null)
                {
                    theDirectory = Registry.GetValue(keyName, ValueName, DefaultPath).ToString();
                }
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace + "\rkeyName " + keyName + "\rValueName " + ValueName +
                "\rDefaultPath " + DefaultPath, "Can't read registry", MessageBoxButtons.OK);
                CloseApps();
            }

            return theDirectory;
        }






        private void btnGetInput_Click(object sender, EventArgs e)
        {
            Control theControl = (Control)sender;
            switch (theControl.Name)
            {
                case "btnGetInput":
                    openInputDialogue.Multiselect = false;
                    if (InputFileBox.Text != "")
                    {
                        openInputDialogue.InitialDirectory = InputFileBox.Text;
                    }
                    else
                    {
                        openInputDialogue.InitialDirectory = InputDir;
                    }
                    if (openInputDialogue.ShowDialog() == DialogResult.OK)
                    {
                        InputFileBox.Text = openInputDialogue.FileName;
                    }
                    break;
                case "btnDecompGlyph":
                    if (DecompGlyphBox.Text != "")
                    {
                        OpenGlyphFileDialogue.InitialDirectory = DecompGlyphBox.Text;
                    }
                    else
                    {
                        OpenGlyphFileDialogue.InitialDirectory = GlyphDir;
                    }
                    if (OpenGlyphFileDialogue.ShowDialog() == DialogResult.OK)
                    {
                        DecompGlyphBox.Text = OpenGlyphFileDialogue.FileName;
                        GlyphDir = Path.GetDirectoryName(OpenGlyphFileDialogue.FileName);
                        Registry.SetValue(keyName, "GlyphDir", GlyphDir);
                    }
                    break;
            }
        }
        private int GetFileType(string FileName, bool JustType = true)
        {
            btnPause.Enabled = false;
            switch (Path.GetExtension(FileName).ToLower())
            {
                // There may be times we just want the file type and nothing else.
                case ".doc":
                case ".docx":
                case ".rtf":
                    if (!JustType)
                    {
                        btnListFonts.Enabled = true;
                        btnGetStyles.Enabled = true;
                        btnSaveFontList.Enabled = true && (Individual || FontListFileBox.Text != "");
                        btnSaveStyles.Enabled = true && (Individual || StyleListFileBox.Text != "");
                        btnSaveXML.Enabled = true;
                        btnGetFont.Enabled = false && !Individual;
                        btnGetEncoding.Enabled = false && !Individual;
                        if (Individual)
                        {
                            AnalyseByFont.Enabled = true;  // Don't turn on if we are doing a bulk analysis
                        }
                        FontLabel.Enabled = false;
                        FontBox.Text = "";
                    }
                    return WordDoc;
                default:
                    if (!JustType)
                    {
                        btnGetFont.Enabled = true;
                        btnGetEncoding.Enabled = true;
                        if (Individual)
                        {
                            AnalyseByFont.Checked = false;  // Don't turn off if we are doing a bulk analysis.
                        }
                        AnalyseByFont.Enabled = false;
                        btnListFonts.Enabled = false;
                        btnSaveXML.Enabled = false;
                        btnSaveFontList.Enabled = false;
                        btnGetStyles.Enabled = false;
                        btnSaveStyles.Enabled = false;
                        FontLabel.Enabled = true;
                        if (FontBox.Text == "")
                        {
                            FontBox.Text = Registry.GetValue(keyName, "Font", "Calibri").ToString();
                        }
                        if (EncodingTextBox.Text == "")
                        {
                            EncodingTextBox.Text = Registry.GetValue(keyName, "Encoding", "Western European (Windows)").ToString();
                        }
                    }
                    return TextDoc;

            }


        }

        private void btnGetOutput_Click(object sender, EventArgs e)
        {
            //
            // Handle the output files
            //
            Control theControl = (Control)sender;  // cast the sender as a control.
            System.Windows.Forms.SaveFileDialog theDialogue = null;
            string theDirectory = "";
            string ValueName = "";
            TextBox theTextBox = null;
            Button theButton = null;
            UpdateFolders(); // update the output folders if relevant box checked.
            switch (theControl.Name)
            {
                case "btnOutputFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = OutputFileBox;
                    theDialogue.InitialDirectory = OutputDir;
                    theButton = btnAnalyse;
                    ValueName = "OutputDir";
                    break;
                case "btnStyleListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = StyleListFileBox;
                    theDialogue.InitialDirectory = StyleDir;
                    theButton = btnSaveStyles;
                    ValueName = "StyleDir";
                    break;
                case "btnBulkStyleListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkStyleListBox;
                    theDialogue.InitialDirectory = StyleDir;
                    theButton = btnSaveStyles;
                    ValueName = "StyleDir";
                    break;
                case "btnBulkFontListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkFontListFileBox;
                    theDialogue.InitialDirectory = FontDir;
                    theButton = btnSaveFontList;
                    ValueName = "FontDir";
                    break;
                case "btnFontListFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = FontListFileBox;
                    theDialogue.InitialDirectory = FontDir;
                    theButton = btnSaveFontList;
                    ValueName = "FontDir";
                    break;
                case "btnXMLFile":
                    theDialogue = saveXMLDialogue;
                    theTextBox = XMLFileBox;
                    theDialogue.InitialDirectory = XMLDir;
                    theButton = btnSaveXML;
                    ValueName = "XMLDir";
                    break;
                case "btnErrorList":
                    theDialogue = saveExcelDialogue;
                    theTextBox = ErrorListBox;
                    theDialogue.InitialDirectory = ErrorDir;
                    theButton = btnSaveErrorList;
                    ValueName = "ErrorDir";
                    break;
                case "btnBulkErrorList":
                    theDialogue = saveExcelDialogue;
                    theTextBox = BulkErrorListBox;
                    theDialogue.InitialDirectory = ErrorDir;
                    theButton = btnSaveErrorList;
                    ValueName = "ErrorDir";
                    break;

                case "btnAggregateFile":
                    theDialogue = saveExcelDialogue;
                    theTextBox = AggregateStatsBox;
                    theDialogue.InitialDirectory = AggregateDir;
                    theButton = btnAggregateFile;
                    ValueName = "AggregateDir";
                    break;

            }

            theDialogue.FileName = theTextBox.Text;
            if (theDialogue.ShowDialog() == DialogResult.OK)
            {
                theTextBox.Text = theDialogue.FileName;
                theDirectory = Path.GetDirectoryName(theDialogue.FileName);
                Registry.SetValue(keyName, ValueName, theDirectory);
                if (theButton != null)
                {
                    theButton.Enabled = true;
                }
            }
        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            CloseApp = true;
            System.Windows.Forms.Application.DoEvents();
            this.Close();
            System.Windows.Forms.Application.Exit();
        }
        private void CloseApps(object sender = null, EventArgs e = null)
        {
            toolStripStatusLabel1.Text = "Shutting down...";
            System.Windows.Forms.Application.DoEvents();
            // Close Excel and Word, but don't flag an error if they are already closed.
            try {
                     if (WordWasRunning)
                    {
                        wrdApp.Visible = true;  // See Word again
                    }
                    else
                    {
                        wrdApp.Quit(ref missing, ref missing, ref missing);
                    }
                    
                }
                catch (Exception ex)
                {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                try
                {
                    // Shut down Excel
                    if (excelApp.ActiveWorkbook != null)
                    {
                        excelApp.ActiveWorkbook.Close(true);
                    }
                    if (ExcelWasRunning)
                    {
                        excelApp.Visible = true;
                    }
                    else
                    {
                        excelApp.Quit();
                        NAR(excelApp);  // release any objects like workbooks because Excel doesn't always quit.
                        System.Threading.Thread.Sleep(5000); // and sleep five seconds
                        excelApp.Quit(); // try again
                        excelApp = null;
                    }
                    
                }
                catch
                {
                }


            this.Close();

        }
        private void NAR(object o)
        {
            try
            {
                while (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0) ;
            }
            catch { }
            finally
            {
                o = null;  // clear the object
            }
        }


        private void btnAnalyse_Click(object sender, EventArgs e)
        {
            /*
             * Here is where we start the analysis
             * 
             * The dictionary holds the character counts.  Its key is the character, including details of the associated font.
             * This means that we can output the counts for a given code for each associated font and therefore glyph.
             * 
             * The CharacterDescriptor class holds both the font and text information.
             */
            Button theButton = (Button)sender;
            try
            {
                Stopwatch theStopwatch = new Stopwatch();

                foreach (string theFile in openInputDialogue.FileNames)
                {
                    Dictionary<CharacterDescriptor, int> theDictionary =
                            new Dictionary<CharacterDescriptor, int>(500, new CharacterEqualityComparer());
                    toolStripProgressBar1.Value = 0;
                    // This is so we can remember the textboxes as we find them.
                    TimeSpan TimeToFinish = TimeSpan.Zero;
                    toolStripStatusLabel1.Text = "Analysing " + theFile + "...";
                    System.Windows.Forms.Application.DoEvents();
                    /*
                     * Open the Word file
                     */
                    EnableButtons(false, FileType); // Disable a whole lot of buttons
                    //lngLink = theDocument.Sections[1].Headers[theIndex].Range.StoryType;
                    int CharacterCount = 0;  // Counter to measure progress
                    // Count the total characters
                    toolStripStatusLabel1.Text = theFile + " opened... ";
                    btnPause.Enabled = true;
                    btnPause.Text = "Pause";
                    Paused = false;
                    theFirstFont = "";
                    toolStripProgressBar1.Value = 0;
                    System.Windows.Forms.Application.DoEvents();
                    /*
                     * We now load the document and analyse it
                     */
                    string theFileName = Path.GetFileName(theFile);
                    switch (GetFileType(theFile))
                    {
                        case WordDoc:

                            if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                            {
                                CharacterCount = AnalyseDocument(theDictionary, theXMLDictionary[theFileName], ref theFirstFont, nsManager, CharacterCount,
                                    theStopwatch);
                            }
                            break;
                        default:
                            DocumentLoader(theFile, theTextDictionary);
                            CharacterCount = AnalyseText(theDictionary, theTextDictionary[theFileName], FontBox.Text, CharacterCount, theStopwatch);
                            theFirstFont = FontBox.Text;
                            break;
                    }
                    /*
                     * Load the aggregate file statistics dictionary
                     */
                    if (AggregateStats.Checked)
                    {
                        AggregateFileList = LoadAggregateStats(theAggregateDictionary, theDictionary, AggregateFileList, theFileName);
                        AggregateSaved = false;  // the list has changed.
                        btnSaveAggregateStats.Enabled = !AggregateSaved;
                    }

                    /*
                      * Create the Excel worksheet
                      */
                    toolStripProgressBar1.Value = toolStripProgressBar1.Maximum;
                    if (WriteIndividualFile.Checked)
                    {
                        string OutputFile = "";
                        if (Individual)
                        {
                            OutputFile = OutputFileBox.Text;
                        }
                        else
                        {
                            OutputFile = Path.Combine(OutputDir, Path.GetFileNameWithoutExtension(theFile) +
                            OutputFileSuffixBox.Text + ".xlsx");
                        }
                        WriteOutput(theDictionary, theFirstFont, OutputDir, OutputFile, theFile, false);
                    }
                    EnableButtons(true, FileType);
                    btnSaveErrorList.Enabled = (listNormalisedErrors.Rows.Count > 0);
                }
                if (!Individual && AggregateStats.Checked)
                {
                    btnSaveAggregateStats_Click(sender, e);  // Pretend we clicked the Save Aggregate Stats button
                }
                theStopwatch.Stop();
                toolStripStatusLabel1.Text = "Finished in " + theStopwatch.Elapsed.ToString(@"hh\:mm\:ss");
                AnalyseByFont.Enabled = true;
                CombDecomposedChars.Enabled = true;
                toolStripProgressBar1.Value = 0;
                Registry.SetValue(keyName, "OutputDir", OutputDir);
                theButton.Enabled = false;  // disable so you can't analyse the same file twice by mistake.
                System.Media.SystemSounds.Beep.Play();  // and beep
            }
            catch (Exception theException)
            {
                // Catch any unexpected errors
                System.Windows.Forms.MessageBox.Show(theException.Message + theException.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CloseApps(this);
            }


        }
        private void EnableButtons(bool Enable, int theFileType)
        {
            // Disable or enable a number of buttons
            btnClose.Enabled = Enable;  // Disable the close - we are analysing
            btnAnalyse.Enabled = Enable;
            btnErrorList.Enabled = Enable;
            btnSaveErrorList.Enabled = Enable;
            CombDecomposedChars.Enabled = Enable;
            AnalyseByFont.Enabled = Enable;
            btnDecompGlyph.Enabled = Enable && CombDecomposedChars.Checked;
            if (theFileType == WordDoc)
            {
                // These only apply to Word documents
                btnOutputFile.Enabled = Enable;
                btnGetInput.Enabled = Enable;
                btnDecompGlyph.Enabled = Enable;
                btnFontListFile.Enabled = Enable;
                btnSaveFontList.Enabled = Enable;
                btnStyleListFile.Enabled = Enable;
                btnSaveStyles.Enabled = Enable;
                btnXMLFile.Enabled = Enable;
                btnListFonts.Enabled = Enable;
                btnGetStyles.Enabled = Enable;
            }
        }

        private bool WriteOutput(Dictionary<CharacterDescriptor, int> theDictionary, string theFont, string OutputDir, string OutputFile,
            string InputFile, bool Aggregate)
        {
            Stopwatch theStopwatch = new Stopwatch();
            theStopwatch.Start();
            // We write the results to Excel.
            btnPause.Enabled = false;  // We can no longer pause
            btnClose.Enabled = false;  // nor can we click on close
            System.Windows.Forms.Application.DoEvents();
            if (!Path.IsPathRooted(OutputFile))
            {
                OutputFile = Path.Combine(OutputDir, OutputFile);  // Make it a complete directory
            }
            toolStripStatusLabel1.Text = "Writing to Excel workbook " + Path.GetFileName(OutputFile) + "...";
            System.Windows.Forms.Application.DoEvents();
            if (!DeleteFile(OutputFile))
            {
                return false; // Do no more if not successfully deleted.
            }
            bool Retry = true;
            ExcelRoot.Workbook theWorkbook = null;
            while (Retry)
            {
                try
                {
                    theWorkbook = excelApp.Workbooks.Add();  // Create it
                    excelApp.Visible = false;  // make sure it is invisible.
                    Retry = false;
                }
                catch (COMException ComEx)
                {
                    if (ComEx.ErrorCode == -2147023174)  // RPC Exception - we've lost Excel
                    {
                        excelApp = new Excel();  // Recreate it
                        Retry = true;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show(ComEx.Message + "\r" + ComEx.StackTrace, "Failed to open Excel", MessageBoxButtons.OK);
                        CloseApps();
                        return false;
                    }
                }
            }


            ExcelRoot.Worksheet theSheet = theWorkbook.ActiveSheet;
            /*
                * Column headers
                */
            int Column = 1;
            string FontColumnLetter = "A";
            string ValueColumnLetter = "B";
            string ColumnLetter = "E";
            string GlyphLetter = "D";
            int FontNumber = (int)Convert.ToChar(FontColumnLetter);
            int ValueNumber = (int)Convert.ToChar(ValueColumnLetter);
            int ColumnNumber = (int)Convert.ToChar(ColumnLetter);
            int GlyphNumber = (int)Convert.ToChar(GlyphLetter);
            toolStripProgressBar1.Value = 0; // reset
            toolStripProgressBar1.Maximum = theDictionary.Count;
            if (Aggregate)
            {
                theSheet.Cells[1, Column++].Value = "Filename";
                // Increment the column numbers for glyph and the final column
                ValueNumber++;
                FontNumber++;
                ColumnNumber++;
                GlyphNumber++;
            }
            if (AnalyseByFont.Checked)
            {
                theSheet.Cells[1, Column++].Value = "Font";
                ValueNumber++;
                ColumnNumber++;
                GlyphNumber++;
            }
            ValueColumnLetter = Convert.ToChar(ValueNumber).ToString();
            ColumnLetter = Convert.ToChar(ColumnNumber).ToString();
            GlyphLetter = Convert.ToChar(GlyphNumber).ToString();
            FontColumnLetter = Convert.ToChar(FontNumber).ToString();
            string RangeString = "A1:" + ColumnLetter + "1";

            theSheet.Cells[1, Column++].Value = "Dec";
            theSheet.Cells[1, Column++].value = "MS Hex";
            theSheet.Cells[1, Column++].Value = "USV";
            theSheet.Cells[1, Column++].Value = "Glyph";
            theSheet.Cells[1, Column].Value = "Count";
            theSheet.Range[RangeString].Font.Bold = true;  // Make the headings bold.
            theSheet.Range[RangeString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;  // Centre

            int theRow = 2;
            foreach (KeyValuePair<CharacterDescriptor, int> kvp in theDictionary)
            {
                /*
                    * Go through the dictionary writing to the worksheet
                    */
                Column = 1;
                if (Aggregate)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.FileName;
                }
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column++].Value = kvp.Key.Font;
                }
                theSheet.Cells[theRow, Column].Value = GetCodes(kvp.Key.Text, dec);  // Decimal
                theSheet.Cells[theRow, Column++].NumberFormat = "@";  // Text format
                theSheet.Cells[theRow, Column].NumberFormat = "@";
                theSheet.Cells[theRow, Column++].Value = GetCodes(kvp.Key.Text, hex);  // Hexadecimal
                theSheet.Cells[theRow, Column++].Value = GetCodes(kvp.Key.Text, USV);  // USV
                if (AnalyseByFont.Checked)
                {
                    theSheet.Cells[theRow, Column].Font.Name = kvp.Key.Font;  // and set the font.
                }
                else
                {
                    theSheet.Cells[theRow, Column].Font.Name = theFont;  // Set the cell to the first font we found
                }

                theSheet.Cells[theRow, Column++].Value = kvp.Key.Text;  // Write the glyph
                theSheet.Cells[theRow, Column].Value = kvp.Value;  // the count
                theRow++;
                toolStripProgressBar1.Value = theRow - 2;
                System.Windows.Forms.Application.DoEvents();
            }
            theWorkbook.SaveAs(OutputFile);  // Save before sorting
            string theRowString = (theRow - 1).ToString();
            /*
                * Now sort by USV value and, if analysing by font, font name
                */

            ExcelRoot.Range CharStats = excelApp.get_Range("A1", "F" + theRowString);
            // Now format all cells
            if (AnalyseByFont.Checked)
            {
                CharStats.Sort(CharStats.Columns[4], ExcelRoot.XlSortOrder.xlAscending,
                    CharStats.Columns[1], missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);
                string FontRange = FontColumnLetter + ":" + FontColumnLetter;
                theSheet.Range[FontRange].EntireColumn.ColumnWidth = 30;  // allow for 30 character font names
                theSheet.Range[FontRange].EntireColumn.HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            }
            else
            {
                CharStats.Sort(CharStats.Columns[3], ExcelRoot.XlSortOrder.xlAscending,
                    missing, missing, ExcelRoot.XlSortOrder.xlAscending,
                    missing, ExcelRoot.XlSortOrder.xlAscending, ExcelRoot.XlYesNoGuess.xlYes);

            }
            theSheet.Range[ValueColumnLetter + "1:" + ColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            theSheet.Range["A1:" + ColumnLetter + theRowString].VerticalAlignment = ExcelRoot.XlVAlign.xlVAlignBottom;
            // and the counts
            theSheet.Range[GlyphLetter + "2:" + GlyphLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;
            theSheet.Range[ColumnLetter + "2:" + ColumnLetter + theRowString].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignRight;
            theSheet.Range[ColumnLetter + "2:" + ColumnLetter + theRowString].NumberFormat = "#,##0";
            // and return to A1
            theSheet.Range["A1"].Select();
            // and freeze the top row
            excelApp.ActiveWindow.SplitColumn = 0;
            excelApp.ActiveWindow.SplitRow = 1;
            excelApp.ActiveWindow.FreezePanes = true;
            theSheet.Name = "Statistics";
            //
            //Create a new worksheet and write to it
            //
            theSheet = theWorkbook.Sheets.Add(missing, theSheet, 1, ExcelRoot.XlSheetType.xlWorksheet);
            theSheet.Name = "MetaData";
            theSheet.Range["A1"].Value = "Character Counter Version";
            theSheet.Range["B1"].Value = String.Format("{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
            theSheet.Range["A2"].Value = "Filename(s)";
            theSheet.Range["B2"].Value = InputFile;
            if (FileType == TextDoc)
            {
                theSheet.Range["A3"].Value = "Encoding";
                theSheet.Range["B3"].Value = EncodingTextBox.Text;
            }
            theSheet.Columns["A"].ColumnWidth = 25;
            theWorkbook.Sheets["Statistics"].Activate();  // go to the statistics sheet

            // and save it
            theWorkbook.Save();
            theWorkbook.Close();
            theStopwatch.Stop();
            //System.Media.SystemSounds.Beep.Play();  // and beep
            toolStripStatusLabel1.Text = "Finished writing to Excel workbook " + Path.GetFileName(OutputFile) + " in " +
                ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            toolStripProgressBar1.Value = 0; // reset
            btnClose.Enabled = true;
            return true;

        }
        private bool DeleteFile(string theFileName)
        {
            DialogResult Success = DialogResult.Retry;
            while (Success == DialogResult.Retry)
            {
                try
                {
                    File.Delete(theFileName); // Delete the existing file
                    Success = DialogResult.OK;  // we succeeded
                }
                catch (Exception Ex)
                {
                    Success = System.Windows.Forms.MessageBox.Show(Ex.Message, "Failed to delete Excel file", MessageBoxButtons.RetryCancel);
                    if (Success == DialogResult.Cancel)
                    {
                        return false; // Don't try to save to Excel
                    }
                }
            }
            return true;
        }
        private void WriteFontList(string theFileName, string theHeader, ListBox theListBox)
        {
            btnClose.Enabled = false;
            int RowCounter = 2;
            // Write the contents of a list box to Excel
            if (!DeleteFile(theFileName))
            {
                return;
            }
            toolStripStatusLabel1.Text = "Writing to " + Path.GetFileName(theFileName);
            ExcelRoot.Workbook theWorkBook = excelApp.Workbooks.Add();
            excelApp.Visible = false;  // make sure we hide it.
            ExcelRoot.Worksheet theSheet = theWorkBook.ActiveSheet;
            theSheet.Range["A1"].Value = theHeader;
            theSheet.Range["A1"].Font.Bold = true;
            foreach (string theItem in theListBox.Items)
            {
                theSheet.Range["A" + RowCounter.ToString()].Value = theItem;
                RowCounter++;
            }
            theWorkBook.SaveAs(theFileName);
            theWorkBook.Close();
            toolStripStatusLabel1.Text = "Finished writing to " + Path.GetFileName(theFileName);
            btnClose.Enabled = true;
            System.Media.SystemSounds.Beep.Play();  // and beep
            return;
        }
        private void WriteDataGridView(string theFileName, DataGridView theDataGridView)
        {
            btnClose.Enabled = false;
            int RowCounter = 2;
            // Write the contents of a list box to Excel
            if (DeleteFile(theFileName))
            {
                toolStripStatusLabel1.Text = "Writing to " + Path.GetFileName(theFileName);
                ExcelRoot.Workbook theWorkBook = excelApp.Workbooks.Add();
                excelApp.Visible = false;
                ExcelRoot.Worksheet theSheet = theWorkBook.ActiveSheet;
                theSheet.Range["A1"].Value = theDataGridView.Columns[0].HeaderText;
                theSheet.Range["B1"].Value = theDataGridView.Columns[1].HeaderText;
                theSheet.Range["A1:B1"].Font.Bold = true;
                theSheet.Range["A1:B1"].HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignCenter;
                theSheet.Range["A:B"].EntireColumn.ColumnWidth = 100;

                foreach (DataGridViewRow theRow in theDataGridView.Rows)
                {

                    theSheet.Range["A" + RowCounter.ToString()].Value = theRow.Cells[0].Value;
                    theSheet.Range["B" + RowCounter.ToString()].Value = theRow.Cells[1].Value;
                    RowCounter++;
                }
                theWorkBook.SaveAs(theFileName);
                theWorkBook.Close();
                toolStripStatusLabel1.Text = "Finished writing to " + Path.GetFileName(theFileName);
                System.Media.SystemSounds.Beep.Play();  // and beep
            }
            btnClose.Enabled = true;
            return;
        }
        private int AnalyseText(Dictionary<CharacterDescriptor, int> theDictionary,
            string theText, string theFont, int CharacterCount, Stopwatch theStopwatch)
        {
            // Analyse a text document
            int CharactersInText = theText.Length;
            toolStripStatusLabel1.Text = "Counting characters";
            CharacterCount += AnalyseString(theDictionary, theFont, theText, CharactersInText, CharacterCount, theStopwatch);
            listNormalisedErrors.Sort(listNormalisedErrors.Columns[0], ListSortDirection.Ascending);  // Sort first
            btnPause.Enabled = false;
            return CharacterCount;


        }
        private int AnalyseDocument(Dictionary<CharacterDescriptor, int> theDictionary,
            XmlDocument theXMLDocument, ref string theFirstFont, XmlNamespaceManager nsManager, int CharacterCount, Stopwatch theStopwatch)
        {
            // Analyse the contents of a character string
            // We repeat ourselves to avoid having to do a logic test through each iteration of the loop.
            string TextString = "";
            string FontName = "";
            int RangeCharacterCount = 0;
            int TextCount = 0;  // Count separately for troubleshooting purposes
            int OtherCount = 0;
            theDictionary.Clear();  // Clear the dictionary
            //
            //  Count the characters
            //
            toolStripStatusLabel1.Text = "Counting characters";
            XmlNode theRoot = theXMLDocument.DocumentElement;
            XmlNodeList theNodeList = theRoot.SelectNodes(@"(//w:body//w:r/w:t | //w:body//w:r/w:sym | //w:body//w:r/w:tab | //w:body//w:r/w:noBreakHyphen | //w:body//w:r/w:softHyphen | //w:body//w:r/w:br)", nsManager);

            foreach (XmlNode theData in theNodeList)
            {
                // we look the range structures
                switch (theData.Name)
                {
                    case "w:t":
                        // we have text
                        TextCount += theData.InnerText.Length;
                        break;
                    case "w:sym":
                        // we have a symbol
                        TextCount++;
                        break;

                    default:
                        // Anything else we simply increment the counter
                        OtherCount++;
                        break;

                }
            }
            // now count paragraphs and section and breaks
            theNodeList = theRoot.SelectNodes(@"(//w:body//w:p | //w:body//w:sectPr)", nsManager);
            if (theNodeList != null)
            {
                OtherCount += theNodeList.Count;
            }
            RangeCharacterCount = TextCount + OtherCount;

            toolStripStatusLabel1.Text = "Counted " + RangeCharacterCount + " characters in "
                + ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";

            toolStripProgressBar1.Maximum = RangeCharacterCount;  // To show progress, but the count isn't accurate.
            btnPause.Enabled = true;
            System.Windows.Forms.Application.DoEvents();
            /*
                * Look for text or symbols in the document
            */
            try
            {
                // Get the styles in use
                GetStylesInUse(theRoot, nsManager, theStyleDictionary);

                // Load decomposed glyphs if we have specified a file
                if (!GlyphsLoaded && CombDecomposedChars.Checked && DecompGlyphBox.Text != "")
                {
                    GlyphsLoaded = LoadDecomposedGlyphs(theGlyphDictionary, excelApp);
                }

                if (AnalyseByFont.Checked)
                {
                    try
                    {
                        string theParagraphFont = "";
                        theNodeList = theRoot.SelectNodes(@"//w:body//w:p", nsManager);  // Find the paragraphs
                        foreach (XmlNode theParagraphData in theNodeList)
                        {
                            // Check if the paragraph has a font - we ignore this, it gave misleading results.

                            //FontName = XmlLookup(theParagraphData, "w:pPr/w:rPr/wx:font", nsManager, "wx:val", "");
                            //FontName = "";
                            //if (FontName == "")
                            //{
                            // Determine the paragraph's style
                            string theParagraphStyleID = XmlLookup(theParagraphData, "w:pPr/w:pStyle", nsManager, "w:val", "DefaultParagraphFont");
                            if (theStyleDictionary.Keys.Contains(theParagraphStyleID))
                            {
                                FontName = theStyleDictionary[theParagraphStyleID];
                            }
                            else
                            {
                                FontName = GetDefaultFont(theStyleDictionary, theParagraphData);
                            }
                            //}
                            theParagraphFont = FontName;  // Remember the paragraph font for the end of line.
                            XmlNodeList theRanges = theParagraphData.SelectNodes("w:r", nsManager);
                            TextString = "";
                            /*
                             * We go through the document a range at a time.  If we find a symbol whose font is the same as that of an existing range
                             * we concatenate the symbol to that range.
                             */
                            string OldFontName = "";
                            foreach (XmlNode theRangeData in theRanges)
                            {
                                XmlNode theSymbol = theRangeData.SelectSingleNode("w:sym", nsManager);
                                if (theSymbol != null)
                                {
                                    // we have a symbol
                                    FontName = theSymbol.Attributes["w:font"].Value;
                                    string theSymbolValue = theSymbol.Attributes["w:char"].Value;
                                    char theChar = Convert.ToChar(Convert.ToUInt16(theSymbolValue, 16));  // get the character number
                                    if (FontName == OldFontName)
                                    {
                                        // Concatenate the text string
                                        TextString += Convert.ToString(theChar); // make it a string concatenating it with previous symbols.
                                    }
                                    else
                                    {
                                        // Analyse the text string, then remember the old font and start a new text string
                                        CharacterCount = AnalyseString(theDictionary, OldFontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                                        OldFontName = FontName;
                                        TextString = Convert.ToString(theChar); // make it a string concatenating it with previous symbols. 
                                    }

                                }
                                else
                                {

                                    // See if there is a font defined in the range and use that
                                    FontName = XmlLookup(theRangeData, "w:rPr/wx:font", nsManager, "wx:val", "");
                                    if (FontName == "")
                                    {
                                        string theStyleID = XmlLookup(theRangeData, "w:rPr/w:rStyle", nsManager, "w:val", "");
                                        if (theStyleID != "" && theStyleDictionary.Keys.Contains(theStyleID))
                                        {
                                            // If we have no style nor do we have a font for the style, we do nothing
                                            // Otherwise we get the font name for the style.
                                            FontName = theStyleDictionary[theStyleID];
                                        }
                                        else
                                        {
                                            FontName = theParagraphFont; // we pick up the paragraph font
                                        }
                                    }

                                    // Look for text
                                    XmlNode theText = theRangeData.SelectSingleNode("w:t", nsManager);
                                    if (theText != null)
                                    {
                                        if (FontName == OldFontName)
                                        {
                                            TextString += theText.InnerText;
                                        }
                                        else
                                        {
                                            CharacterCount = AnalyseString(theDictionary, OldFontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                                            OldFontName = FontName;
                                            TextString = theText.InnerText;
                                        }
                                    }
                                    for (int Counter = 0; Counter < SpecialCharacterKeys.Count(); Counter++)
                                    {
                                        XmlNode theSpecialChar = theRangeData.SelectSingleNode(SpecialCharacterKeys[Counter], nsManager);
                                        if (theSpecialChar != null)
                                        {
                                            // We've found a special character
                                            TextString += SpecialCharacterValues[Counter];
                                        }
                                    }
                                    //
                                    // Look for break characters
                                    //
                                    XmlNode theBreak = theRangeData.SelectSingleNode("w:br", nsManager);
                                    if (theBreak != null)
                                    {
                                        if (theBreak.Attributes.Count > 0)
                                        {
                                            try
                                            {
                                                TextString += theBreakDictionary[theBreak.Attributes["w:type"].Value];
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        else
                                        {
                                            // a break with nothing else is U+000B = \v
                                            TextString += "\v";
                                        }
                                    }
                                    // Now look for a section break
                                    XmlNode theSectionBreak = theParagraphData.SelectSingleNode("w:pPr/w:sectPr", nsManager);
                                    if (theSectionBreak != null)
                                    {
                                        TextString += "\f";
                                    }
                                }
                            }
                            if (TextString != "")
                            {
                                // We have some text to process
                                CharacterCount = AnalyseString(theDictionary, FontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                                TextString = "";
                                // Now add the end of line marker
                            }
                            CharacterCount = AnalyseString(theDictionary, theParagraphFont, "\r", RangeCharacterCount, CharacterCount, theStopwatch);


                        }
                        // Now look for any more sections
                        theNodeList = theRoot.SelectNodes(@"//w:body/wx:sect/w:sectPr", nsManager);
                        if (theNodeList != null)
                        {
                            TextString = "";
                            for (int Counter = 0; Counter < theNodeList.Count; Counter++)
                            {
                                TextString += "\f";  // section/page break gives a form feed.
                            }
                            CharacterCount = AnalyseString(theDictionary, FontName, TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                            TextString = "";
                        }
                    }
                    catch (Exception Ex)
                    {

                        System.Windows.Forms.MessageBox.Show(Ex.Message + "\r\r" + Ex.StackTrace, "Error in character counting - analysed by font", MessageBoxButtons.OK);
                        CloseApps(this);

                    }


                }
                else
                {
                    // We aren't analysing by font
                    try
                    {
                        theNodeList = theRoot.SelectNodes(@"//w:body//w:p", nsManager);  // Find the paragraphs
                        theFirstFont = theStyleDictionary["DefaultParagraphFont"]; // The default font
                        TextString = "";
                        foreach (XmlNode theParagraphData in theNodeList)
                        {
                            // Determine the paragraph's type

                            XmlNodeList theRanges = theParagraphData.SelectNodes("w:r", nsManager);
                            foreach (XmlNode theRangeData in theRanges)
                            {
                                XmlNode theSymbol = theRangeData.SelectSingleNode("w:sym", nsManager);
                                if (theSymbol != null)
                                {
                                    // we have a symbol
                                    string theSymbolValue = theSymbol.Attributes["w:char"].Value;
                                    char theChar = Convert.ToChar(Convert.ToUInt16(theSymbolValue, 16));  // get the character number
                                    TextString += Convert.ToString(theChar); // make it a string

                                }
                                else
                                {
                                    // Look for text
                                    XmlNode theText = theRangeData.SelectSingleNode("w:t", nsManager);
                                    if (theText != null)
                                    {
                                        TextString += theText.InnerText;
                                    }
                                    for (int Counter = 0; Counter < SpecialCharacterKeys.Count(); Counter++)
                                    {
                                        XmlNode theSpecialChar = theRangeData.SelectSingleNode(SpecialCharacterKeys[Counter], nsManager);
                                        if (theSpecialChar != null)
                                        {
                                            // We've found a special character
                                            TextString += SpecialCharacterValues[Counter];
                                        }
                                    }
                                    //
                                    // Look for break characters
                                    //
                                    XmlNode theBreak = theRangeData.SelectSingleNode("w:br", nsManager);
                                    if (theBreak != null)
                                    {
                                        if (theBreak.Attributes.Count > 0)
                                        {
                                            try
                                            {
                                                TextString += theBreakDictionary[theBreak.Attributes["w:type"].Value];
                                            }
                                            catch
                                            {
                                            }
                                        }
                                        else
                                        {
                                            // a break with nothing else is U+000B = \v
                                            TextString += "\v";
                                        }
                                    }
                                }
                            }
                            // Now look for a section break
                            XmlNode theSectionBreak = theParagraphData.SelectSingleNode("w:pPr/w:sectPr", nsManager);
                            if (theSectionBreak != null)
                            {
                                TextString += "\f"; // the break character
                            }

                            // Now add the end of line marker
                            TextString += "\r";

                        }
                        /*
                        // Now look for sections and page breaks
                        theNodeList = theRoot.SelectNodes(@"//w:body/wx:sect", nsManager);
                        if (theNodeList != null)
                        {
                            for (int Counter = 0; Counter < theNodeList.Count - 1; Counter++)
                            {
                                TextString += "\f";  // section/page break gives a form feed.
                            }
                        }
                        */
                        CharacterCount = AnalyseString(theDictionary, "", TextString, RangeCharacterCount, CharacterCount, theStopwatch);
                        TextString = "";  // clear the text string.

                    }
                    catch (Exception Ex)
                    {

                        System.Windows.Forms.MessageBox.Show(Ex.Message + "\r\r" + Ex.StackTrace, "Error in character counting - not analysed by font", MessageBoxButtons.OK);
                        CloseApps(this);

                    }

                }
            }

            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Error in analysing text", MessageBoxButtons.OK);
                CloseApps(this);
            }
            listNormalisedErrors.Sort(listNormalisedErrors.Columns[0], ListSortDirection.Ascending);  // Sort first
            btnPause.Enabled = false;
            return CharacterCount;
        }

        private string XmlLookup(XmlNode theNode, string theSearchPath, XmlNamespaceManager nsManager, string theValueID, string InputString = "")
        {
            /*
             * This looks up something in Xml and updates the input string with the returned value.  Otherwise
             * it returns the input string.  The idea is to update some data with new information
             */
            XmlNode theChildNode = theNode.SelectSingleNode(theSearchPath, nsManager);
            if (theChildNode == null)
            {
                // We didn't find anything, so return the input string
                return InputString;
            }
            else
            {
                try
                {
                    string tmpString = theChildNode.Attributes[theValueID].Value;
                    return tmpString;
                }
                catch (Exception Ex)
                {
                    // Something went wrong
                    string theError = Ex.Message;
                    return InputString;
                }
            }
        }
        private string GetDefaultFont(Dictionary<string, string> theStyleDictionary, XmlNode theNode)
        {
            string NodePath = GetNodePath(theNode, "");
            string theDefaultID = "DefaultParagraphFont";  // we assume a normal paragraph
            if (NodePath.Contains("w:tbl"))
            {
                // we have a table
                theDefaultID = "Default Table";
            }

            return theStyleDictionary[theDefaultID];
        }
        private void GetStylesInUse(XmlNode theRoot, XmlNamespaceManager nsManager, Dictionary<string, string> theStyleDictionary)
        {                // Load a list of current styles and their fonts
            XmlNodeList theNodeList = theRoot.SelectNodes(@"//w:styles/w:style", nsManager);
            theStyleDictionary.Clear();  // Empty the style dictionary
            // First look for the styles that have fonts
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    string theFontName = theFont.Attributes["wx:val"].Value;
                    theStyleDictionary.Add(theStyleID, theFontName);
                }
            }
            // Now look for the default fonts - we do this as a second pass in case they don't appear first
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                // For some we can't search on wx:font so we have to iterate
                if (theFont != null)
                {
                    if (theStyle.Attributes.Count == 3)
                    {
                        // check to see if this is the default.
                        bool IsDefault = false;
                        try
                        {
                            IsDefault = (theStyle.Attributes[@"w:default"].Value == "on" /*&& theStyle.Attributes[@"w:type"].Value == "paragraph"*/);
                        }
                        catch
                        {
                        }
                        if (IsDefault)
                        {


                            string theDefaultStyle = theStyle.Attributes[@"w:styleId"].Value;
                            // We have found a default style so we look up its font and add to the nominal styles.
                            switch (theStyle.Attributes[@"w:type"].Value)
                            {
                                case "paragraph":
                                    theStyleDictionary["DefaultParagraphFont"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "table":
                                    theStyleDictionary["Default Table"] = theStyleDictionary[theDefaultStyle];
                                    break;
                                case "character":
                                    theStyleDictionary["Default Character"] = theStyleDictionary[theDefaultStyle];
                                    break;

                            }
                        }
                    }
                }
            }
            // Now the styles that don't have fonts- we have to get the font of the style on which they are based.
            // but first we load those that aren't based on a style 
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                XmlNode theBasedOnStyle = theStyle.SelectSingleNode("w:basedOn", nsManager);
                if (theFont == null && theBasedOnStyle == null)
                {
                    // Use the default paragraph font
                    theStyleDictionary[theStyleID] = theStyleDictionary["DefaultParagraphFont"];
                }
            }
            // Now look at the styles that don't have fonts but are based on other styles, and give them the fonts from the styles on which they were based
            foreach (XmlNode theStyle in theNodeList)
            {
                string theStyleID = theStyle.Attributes["w:styleId"].Value;
                XmlNode theFont = theStyle.SelectSingleNode("w:rPr/wx:font", nsManager);
                XmlNode theBasedOnStyle = theStyle.SelectSingleNode("w:basedOn", nsManager);
                if (theFont == null && theBasedOnStyle != null)
                {
                    // Use the default paragraph font
                    string theBasedOnStyleID = theBasedOnStyle.Attributes["w:val"].Value;
                    theStyleDictionary[theStyleID] = theStyleDictionary[theBasedOnStyleID];
                }
            }

        }

        private string GetNodePath(XmlNode theNode, string InputType)
        {
            // Iteratively walk up the nodes.
            XmlNode theParent = theNode.ParentNode;
            if (theParent != null)
            {
                string tmpString = theParent.Name + "/" + InputType;
                GetNodePath(theParent, tmpString);
                return tmpString;
            }
            else
            {
                return InputType;
            }
        }
        private int AnalyseString(Dictionary<CharacterDescriptor, int> theDictionary, string FontName, string TextString, int RangeCharacterCount, int CharacterCount, Stopwatch theStopwatch)
        {
            /*
             * We shall first use the data for legacy decomposed characters to count them before we use the built-in functions that handle
             * decomposed Unicode characters.
             */

            CharacterDescriptor theKey = null;
            string tmpString = "";
            if (AnalyseByFont.Checked && CombDecomposedChars.Checked && theGlyphDictionary.Keys.Contains(FontName))
            {
                // We will count the glyphs loaded as single characters if we have the data
                try
                {
                    Regex theGlyphs = new Regex(theGlyphDictionary[FontName]);

                    MatchCollection theMatches = theGlyphs.Matches(TextString);
                    foreach (Match theMatch in theMatches)
                    {
                        string theString = theMatch.Value.ToString();
                        theKey = new CharacterDescriptor(FontName, theString);
                        IncrementCount(theDictionary, theKey);
                        CharacterCount += theString.Length;
                        ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);

                    }

                    // now remove all those characters
                    tmpString = theGlyphs.Replace(TextString, "");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message + "\r" + ex.StackTrace);
                    string tmpMessage = ex.Message;
                    CloseApps(null, null);
                }
            }
            else
            {
                tmpString = TextString;
            }
            /*
             * We shall now use the built-in Unicode functions to find Unicode decomposed characters
             */
            if (CombDecomposedChars.Checked)
            {
                TextElementEnumerator theTextElements = StringInfo.GetTextElementEnumerator(tmpString);
                while (theTextElements.MoveNext())
                {
                    string theString = theTextElements.GetTextElement();
                    if (AnalyseByFont.Checked)
                    {
                        theKey = new CharacterDescriptor(FontName, theString);
                    }
                    else
                    {
                        theKey = new CharacterDescriptor(theString);
                    }
                    CharacterCount += theString.Length;
                    IncrementCount(theDictionary, theKey);
                    ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);
                    if (theString.Length > 1)
                    {
                        CheckNormalisation(theString);  // Check to see if the string is normalised
                    }
                }
            }
            else
            {
                for (int i = 0; i < tmpString.Length; i++)
                {
                    if (AnalyseByFont.Checked)
                    {
                        theKey = new CharacterDescriptor(FontName, tmpString[i].ToString());
                    }
                    else
                    {
                        theKey = new CharacterDescriptor(tmpString[i].ToString());
                    }
                    IncrementCount(theDictionary, theKey);
                    ReportProgress(CharacterCount, RangeCharacterCount, theStopwatch);
                    CharacterCount++;
                }
            }

            return CharacterCount;
        }
        private void CheckNormalisation(string theString)
        {
            if (theString.IsNormalized())
            {
                return;  // We need do no more
            }
            string theNormalisedString = theString.Normalize(NormalizationForm.FormC);  // Full canonical normalisation
            theString = GetCodes(theString, USV);
            // Look to see if we have found it already
            bool Found = false;
            foreach (DataGridViewRow theViewRow in listNormalisedErrors.Rows)
            {
                if (theViewRow.Cells[0].Value.ToString() == theString)
                {
                    Found = true;
                    break;
                }
            }
            if (!Found)
            {
                theNormalisedString = GetCodes(theNormalisedString, USV);
                string[] theRow = new string[] { theString, theNormalisedString };
                listNormalisedErrors.Rows.Add(theRow);
                tabControl1.SelectedTab = tabControl1.TabPages[2];
                System.Windows.Forms.Application.DoEvents();
            }

        }

        private string GetCodes(string theString, int theCode)
        {
            /*
             * Return the character code for a character
             */
            string tmpString = "";
            foreach (var theChar in theString)
            {
                /*
                 * We loop through each characer in the string in case we get a composed character.
                 * This aspect hasn't been tested, so I don't know if it will work, but worth a try.
                 */
                int temp = Convert.ToUInt16(theChar);
                string tempstring = "";

                switch (theCode)
                {
                    case dec:
                        tempstring = temp.ToString();
                        break;
                    case hex:
                        tempstring = temp.ToString("X");  // Upper case hex
                        break;
                    case USV:
                        tempstring = String.Format("U+{0:X4}", temp);  // USV
                        break;
                }
                tmpString += tempstring + " ";
            }

            return tmpString.Trim(); ;
        }

        private void documentationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "CharacterCounter.docx");
            System.Diagnostics.Process.Start(HelpPath);
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 About = new AboutBox1();
            About.Show();
        }

        private void LicenseMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "gpl.txt");
            System.Diagnostics.Process.Start("Wordpad.exe", '"' + HelpPath + '"');
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            Paused = !Paused; // Toggle the pause flag
            btnClose.Enabled = Paused;
            if (Paused)
            {
                toolStripStatusLabel1.Text = "Pausing...";
                btnPause.Text = "Resume";
            }
            else
            {
                toolStripStatusLabel1.Text = "Resuming...";
                btnPause.Text = "Pause";
            }

            System.Windows.Forms.Application.DoEvents();
        }


        private void IncrementCount(Dictionary<CharacterDescriptor, int> theDictionary, CharacterDescriptor theKey)
        {
            /*
             * Increment the value of the relevant key, and handle a non-existent value.
             */
            if (theDictionary.Keys.Contains(theKey))
            {
                /*
                 * Increment the character count
                 */
                //DateTime Start = DateTime.Now;
                theDictionary[theKey]++;
                //toolStripStatusLabel2.Text = DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            else
            {
                /*
                 * If haven't met the character/font combination before, the increment operation fails
                 * so we come here and generate a new entry with a count of 1
                 */
                //DateTime Start = DateTime.Now;
                theDictionary.Add(theKey, 1);
                // toolStripStatusLabel2.Text = " Key added in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString();
            }
            return;
        }
        private string LoadAggregateStats(Dictionary<CharacterDescriptor, int> theMultiDictionary, Dictionary<CharacterDescriptor, int>
            theDictionary, string FileList, string InputFile)
        {
            // Here is where we load the aggregate statistics dictionary.
            string tmpString = FileList;
            int Counter = Convert.ToInt16(FileCounter.Text);
            foreach (CharacterDescriptor theKey in theDictionary.Keys)
            {
                CharacterDescriptor tmpKey = new CharacterDescriptor(theKey);
                tmpKey.FileName = InputFile;
                if (theMultiDictionary.Keys.Contains(tmpKey))
                {
                    theMultiDictionary[tmpKey] += theDictionary[theKey];
                }
                else
                {
                    theMultiDictionary.Add(tmpKey, theDictionary[theKey]);
                }

            }
            if (tmpString != "")
            {
                tmpString += ", ";
            }
            tmpString += InputFile;
            Counter++;
            FileCounter.Text = Counter.ToString();
            return tmpString;
        }

        private void ReportProgress(int CharacterCount, int RangeCharacterCount, Stopwatch theStopwatch)
        {
            if ((CharacterCount % 100) == 0)
            {
                // report progress
                TimeSpan TimeToFinish = TimeSpan.FromTicks((long)((RangeCharacterCount - CharacterCount) * ((float)theStopwatch.ElapsedTicks / CharacterCount)));
                toolStripStatusLabel1.Text = CharacterCount.ToString() + " of about " + RangeCharacterCount.ToString()
                    + " chars. Approx time to finish analysis: " + TimeToFinish.ToString(@"hh\:mm\:ss"); ;
                toolStripProgressBar1.Value = Math.Min(CharacterCount, toolStripProgressBar1.Maximum);
                System.Windows.Forms.Application.DoEvents();
                while (Paused)
                {
                    if (theStopwatch.IsRunning)
                    {
                        theStopwatch.Stop();
                    }
                    System.Threading.Thread.Sleep(1000); // wait a second
                    toolStripStatusLabel1.Text = "Paused, click Resume to restart or Close to close";
                    if (CloseApp)
                    {
                        /*
                            * We have clicked Close, so we exit from here and stop doing the processing
                            */
                        theDocument.Close(false);  //Close the document
                        return;  // exit
                    }
                    System.Windows.Forms.Application.DoEvents();

                }
                if (!theStopwatch.IsRunning)
                {
                    theStopwatch.Start();
                }

            }
            return;

        }

        private void btnListFonts_Click(object sender, EventArgs e)
        {
            // list the fonts in the documnent
            Control theControl = (Control)sender;
            theControl.Enabled = false;
            System.Windows.Forms.Application.DoEvents();
            Stopwatch theStopwatch = new Stopwatch();
            theStopwatch.Start();
            btnClose.Enabled = false;  // Disable the close - we are analysing
            btnAnalyse.Enabled = false;
            btnListFonts.Enabled = false;
            AnalyseByFont.Enabled = false;
            CombDecomposedChars.Enabled = false;
            CombDecomposedChars.Enabled = false;
            toolStripProgressBar1.Value = 0;
            XmlNode theRoot = null;
            List<string> theFontTable = new List<string>();
            toolStripStatusLabel1.Text = "Listing fonts in " + Path.GetFileName(InputFileBox.Text) + "...";
            System.Windows.Forms.Application.DoEvents();
            /*
             * Open the Word file(s)
             *
             */
            theFontTable.Clear();
            foreach (string theFile in openInputDialogue.FileNames)
            {
                if (GetFileType(theFile) == WordDoc)  // only get the data if the files are Word files.
                {

                    if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                    {
                        // only run if successful
                        string theFileName = Path.GetFileName(theFile);
                        theRoot = theXMLDictionary[theFileName].DocumentElement;
                        XmlNodeList theFontList = theRoot.SelectNodes(@"w:fonts/w:font", nsManager);
                        foreach (XmlNode theFont in theFontList)
                        {
                            string theFontName = theFont.Attributes["w:name"].Value;
                            if (!theFontTable.Contains(theFontName))
                            {
                                theFontTable.Add(theFont.Attributes["w:name"].Value);  // Add the font if we don't have it already
                            }
                        }
                        theFontTable.Sort();
                        FontList.Items.Clear();  // Clear so we don't load it more than once.
                        foreach (string theFont in theFontTable)
                        {
                            FontList.Items.Add(theFont);
                        }
                        toolStripStatusLabel1.Text = "Finished loading fonts from " + theFile;
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            btnClose.Enabled = true;  // Disable the close - we are analysing
            btnAnalyse.Enabled = true;
            btnListFonts.Enabled = true;
            AnalyseByFont.Enabled = true;
            CombDecomposedChars.Enabled = true;
            toolStripStatusLabel1.Text = "Finished loading fonts in " +
                ((float)theStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            if (theControl.Name == "btnSaveFontList")
            {
                string theFontListFile = "";
                if (Individual)
                {
                    theFontListFile = FontListFileBox.Text;
                }
                else
                {
                    theFontListFile = BulkFontListFileBox.Text;
                }
                WriteFontList(theFontListFile, "List of fonts", FontList);
            }
            theStopwatch.Stop();
            theStopwatch = null;
            theControl.Enabled = true;

        }

        private void InputFileBox_TextChanged(object sender, EventArgs e)
        {
            // A different file so we need to reload it and clear lots of things.
            TextBox theBox = (TextBox)sender;
            if (theBox.Text != "" && !File.Exists(theBox.Text))
            {
                System.Windows.Forms.MessageBox.Show("File does not exist", "Error", MessageBoxButtons.OK);
                theBox.Select();
                return;
            }
            InputFileName = Path.GetFileName(theBox.Text);
            if (InputFileName != null)
            {
                theXMLDictionary.Remove(InputFileName);
                theTextDictionary.Remove(InputFileName);
            }
            else
            {
                return;
            }
            InputFileName = Path.GetFileName(InputFileBox.Text);  // Remember the file name for later.
            // Suggest names for the output files
            OutputFileBox.Text = Path.Combine(OutputDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + ".xlsx";
            InputDir = Path.GetDirectoryName(openInputDialogue.FileName);
            Registry.SetValue(keyName, "InputDir", InputDir);
            btnAnalyse.Enabled = true;
            btnPause.Enabled = false;
            saveExcelDialogue.FileName = OutputFileBox.Text;
            FileType = GetFileType(InputFileBox.Text, false);
            UpdateFolders(); // update the output folders if relevant box checked.
            if (FileType == WordDoc)
            {
                // Only enable these if we are analysing Word documents.
                StyleListFileBox.Text = Path.Combine(StyleDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Styles).xlsx";
                FontListFileBox.Text = Path.Combine(FontDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Fonts).xlsx";
                XMLFileBox.Text = Path.Combine(XMLDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + ".xml";
            }

            ErrorListBox.Text = Path.Combine(ErrorDir, Path.GetFileNameWithoutExtension(InputFileBox.Text)) + " (Suggested Chars).xlsx";

            listStyles.Rows.Clear(); // clear the styles
            FontList.Items.Clear(); // and the fonts
            theStyleDictionary.Clear();
            listNormalisedErrors.Rows.Clear();
            theStyleDictionary.Clear();
            FileType = GetFileType(theBox.Text);  // redetermine the file type.
            openInputDialogue.FileName = theBox.Text;
            btnAnalyse.Enabled = true;  // Let you analyse the new file.


        }
        private void UpdateFolders()
        {
            // Update the output folders if desired
            if (checkUpdateOutputFolders.Checked)
            {
                // Update all the output folders and their registry settings
                StyleDir = SetRegistry(keyName, "StyleDir", OutputDir);
                FontDir = SetRegistry(keyName, "FontDir", OutputDir);
                XMLDir = SetRegistry(keyName, "XMLDir", OutputDir);
                ErrorDir = SetRegistry(keyName, "ErrorDir", OutputDir);
                AggregateDir = SetRegistry(keyName, "AggregateDir", OutputDir);
                Registry.SetValue(keyName, "OutputDir", OutputDir); 
                checkUpdateOutputFolders.Checked = false;  // Clear the checkbox so we aren't continually updating.
            }

        }
        private string SetRegistry(string keyName, string valueName, string Value)
        {
            Registry.SetValue(keyName, valueName, Value);
            return Value;
        }

        private void btnGetStyles_Click(object sender, EventArgs e)
        {
            Control theControl = (Control)sender;
            theControl.Enabled = false;
            XmlNode theRoot = null;
            listStyles.Rows.Clear();  // Empty the list
            //List<DataGridViewRow> theStyleList = new List<DataGridViewRow>(10);
            foreach (string theFile in openInputDialogue.FileNames)
            {
                if (GetFileType(theFile) == WordDoc)  //Only analyse Word documents
                {
                    if (DocumentLoader(theFile, theXMLDictionary, ref nsManager))
                    {
                        string theFileName = Path.GetFileName(theFile); // We just use the file name, not the full path
                        theRoot = theXMLDictionary[theFileName].DocumentElement;
                        GetStylesInUse(theRoot, nsManager, theStyleDictionary);
                        XmlNode theStylesNode = theRoot.SelectSingleNode("w:styles", nsManager);
                        foreach (string theStyleID in theStyleDictionary.Keys)
                        {
                            // Look up the font name rather than the ID
                            string tmpStyle = theStyleID;
                            XmlNode theStyleNode = theStylesNode.SelectSingleNode("w:style[@w:styleId = \"" + theStyleID + "\"]", nsManager);
                            if (theStyleNode != null)
                            {
                                XmlNode theNameNode = theStyleNode.SelectSingleNode("w:name", nsManager);
                                if (theStyleNode.Attributes != null)
                                {
                                    tmpStyle = theNameNode.Attributes["w:val"].Value;
                                }
                            }

                            // Create a new row, and add it if we haven't already got it.
                            DataGridViewRow theRow = new DataGridViewRow();
                            string[] theStringArray = { tmpStyle, theStyleDictionary[theStyleID] };
                            theRow.CreateCells(listStyles, theStringArray);

                            //if (!theStyleList.Contains(theRow))
                            //{
                            //    theStyleList.Add(theRow);
                            //    listStyles.Rows.Add(theRow);
                            //}
                            if (!listStyles.Rows.Contains(theRow))
                            {
                                listStyles.Rows.Add(theRow);
                            }
                        }
                    }
                }

            }
            if (listStyles.Rows.Count > 0)
            {
                listStyles.Sort(listStyles.Columns[0], ListSortDirection.Ascending);  // Sort the list
            }
            if (theControl.Name == "btnSaveStyles")
            {
                // Write the list to Excel
                string theStyleListFile = "";
                if (Individual)
                {
                    theStyleListFile = StyleListFileBox.Text;
                }
                else
                {
                    theStyleListFile = BulkStyleListBox.Text;
                }
                WriteDataGridView(theStyleListFile, listStyles);
                Registry.SetValue(keyName, "StyleDir", StyleDir);

            }
            theControl.Enabled = true;
            System.Windows.Forms.Application.DoEvents();

        }
        private XmlDocument LoadWordDocument(string WordFile)
        {
            // Load the Word document into XML.
            Stopwatch theStopWatch = new Stopwatch();
            theStopWatch.Start();
            XmlDocument theXMLDocument = new XmlDocument();
            try
            {
                bool Retry = true;
                while (Retry == true)
                {
                    try
                    {
                        wrdApp.Documents.Open(WordFile, missing, true);
                        theDocument = wrdApp.ActiveDocument;
                        Retry = false;
                    }
                    catch (COMException ComEx)
                    {
                        if (ComEx.ErrorCode == -2147023174) // RPC Server Unavailable
                        {
                            wrdApp = new Word();
                            Retry = true;
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show(WordFile + " failed to open \r" + ComEx.Message + "\r" + ComEx.StackTrace, "Word failed to open!", MessageBoxButtons.OK);
                            CloseApps(); // Shut down
                        }
                    }
                }
                theDocument.Select();
                try
                {
                    toolStripStatusLabel1.Text = "Loading " + Path.GetFileName(WordFile) + "...";
                    //string XMLDoc = wrdApp.Selection.get_XML(false);
                    string XMLDoc = wrdApp.Selection.XML[false];
                    theXMLDocument.LoadXml(XMLDoc);
                    XMLDoc = null;
                }
                catch (XmlException Ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error loading into " + WordFile + " into XML\r" + Ex.Message + "\r Hex error code :" + Ex.HResult.ToString("X")
                         + "\r Source " + Ex.Source + "\r Line number: " + Ex.LineNumber + "\r Line position: " + Ex.LinePosition, "Error loading XML", MessageBoxButtons.OK);
                    CloseApps(); // Shut down
                }
                theDocument.Close();  // We no longer need it.
                theDocument = null;
            }
            catch (Exception Ex)
            {
                System.Windows.Forms.MessageBox.Show("Error opening " + WordFile + "\r" + Ex.Message + "\r" + Ex.StackTrace, "Error opening document", MessageBoxButtons.OK);
                toolStripStatusLabel1.Text = "Failed to open document after\r" +
                    ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
                theStopWatch.Stop();
                theStopWatch = null;
                return null;
            }
            toolStripStatusLabel1.Text = "Loaded" + Path.GetFileName(WordFile) + " after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            theStopWatch.Stop();
            theStopWatch = null;
            return theXMLDocument;
        }
        private bool DocumentLoader(string WordFile, Dictionary<string, XmlDocument> theXMLDictionary, ref XmlNamespaceManager nsManager)
        {
            /*
             *  If we've already loaded the document, we don't need to do it again
             *
             */
            string theFileName = Path.GetFileName(WordFile);  // We just use the file name as a key, not the whole path.
            if (theXMLDictionary.Keys.Contains(theFileName))
            {
                return true;
            }
            else
            {
                /*
                 * Open the Word file
                */
                XmlDocument theXMLDocument = LoadWordDocument(WordFile);

                if (theXMLDocument != null)
                {
                    nsManager = new XmlNamespaceManager(theXMLDocument.NameTable);
                    nsManager.AddNamespace("wx", wordmlxNamespace);
                    nsManager.AddNamespace("w", wordmlNamespace);
                    // If successful add the root to the dictionary
                    theXMLDictionary.Add(theFileName, theXMLDocument);
                    return true;
                }
                else
                {
                    // We failed
                    return false;
                }
            }
        }
        private bool DocumentLoader(string TextFile, Dictionary<string, string> theTextDictionary)
        {
            /*
             *  If we've already loaded the document, we don't need to do it again
             *
             */
            string theFileName = Path.GetFileName(TextFile);  // We just use the file name as a key, not the whole path.
            if (theTextDictionary.Keys.Contains(theFileName))
            {
                return true;
            }
            else
            {
                /*
                 * Open the Word file
                */
                string theFile = LoadTextDocument(TextFile);
                if (theFile != null)
                {
                    // If successful add the root to the dictionary
                    theTextDictionary.Add(theFileName, theFile);
                    return true;
                }
                else
                {
                    // We failed
                    return false;
                }
            }
        }

        private string LoadTextDocument(string TextFile)
        {
            // Load a text document into a string.
            Stopwatch theStopWatch = new Stopwatch();
            string theTextDocument = null;
            theStopWatch.Start();
            DialogResult Retry = DialogResult.Retry;
            while (Retry == DialogResult.Retry)
            {
                try
                {
                    theTextDocument = File.ReadAllText(TextFile, theEncoding);
                    Retry = DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retry = System.Windows.Forms.MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Word failed to open!", MessageBoxButtons.RetryCancel);
                    if (Retry == DialogResult.Cancel)
                    {
                        toolStripStatusLabel1.Text = "Document load cancelled after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
                        theStopWatch.Stop();
                        theStopWatch = null;
                        return null;
                    }
                }
            }
            toolStripStatusLabel1.Text = "Loaded document after " + ((float)theStopWatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds";
            theStopWatch.Stop();
            theStopWatch = null;
            return theTextDocument;
        }

        private void btnSaveXML_Click(object sender, EventArgs e)
        {
            DialogResult Retrying = System.Windows.Forms.DialogResult.Retry;
            bool DocumentLoaded = DocumentLoader(InputFileBox.Text, theXMLDictionary, ref nsManager);


            while (Retrying == System.Windows.Forms.DialogResult.Retry)
            {
                try
                {
                    theXMLDictionary[InputFileName].Save(XMLFileBox.Text);
                    Retrying = System.Windows.Forms.DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retrying = System.Windows.Forms.MessageBox.Show(Ex.Message, "Failed to save XML file", MessageBoxButtons.RetryCancel);
                    if (Retrying == System.Windows.Forms.DialogResult.Cancel)
                    {
                        toolStripStatusLabel1.Text = "Failed to save XML file " + Path.GetFileName(XMLFileBox.Text);
                        return;  // We do no more
                    }
                }
            }
            Registry.SetValue(keyName, "XMLDir", XMLDir); // Save the output directory
            toolStripStatusLabel1.Text = Path.GetFileName(XMLFileBox.Text) + " saved";
            System.Media.SystemSounds.Beep.Play();
            return;


        }

        private bool LoadDecomposedGlyphs(Dictionary<string, string> theGlyphDictionary, ExcelApp theApp)
        {
            DialogResult Retrying = DialogResult.Retry;
            int theRow = 2;
            ExcelRoot.Workbook theWorkbook = null;
            theGlyphDictionary.Clear();  // Make sure it is empty
            while (Retrying == DialogResult.Retry)
            {
                try
                {
                    theWorkbook = theApp.Workbooks.Open(DecompGlyphBox.Text, missing, true);
                    Retrying = DialogResult.OK;
                }
                catch (Exception Ex)
                {
                    Retrying = System.Windows.Forms.MessageBox.Show("Error opening glyph file "  + DecompGlyphBox.Text + " " + Ex.Message + "\r" + Ex.StackTrace, "Error opening glyph file", MessageBoxButtons.RetryCancel);
                    if (Retrying == DialogResult.Cancel)
                    {
                        return false;
                    }
                }
            }
            //
            //  We have opened the Excel file, so read it
            //
            ExcelRoot.Range theRange = theWorkbook.ActiveSheet.Cells[theRow, 1];
            while (theRange.Value != null)
            {
                string FontName = theRange.Font.Name;
                string theChar = theRange.Value;  // Escape out certain significant Regex characters.

                if (theGlyphDictionary.Keys.Contains(FontName))
                {
                    theGlyphDictionary[FontName] += theChar;
                }
                else
                {
                    theGlyphDictionary[FontName] = "(.[" + theChar;
                }
                theRow++;
                theRange = theWorkbook.ActiveSheet.Cells[theRow, 1];
            }
            theWorkbook.Close(false); // Close the workbook, we don't need it again.
            for (int KeyCounter = 0; KeyCounter < theGlyphDictionary.Keys.Count; KeyCounter++)
            {
                // we now close off the regular expression
                string theKeyName = theGlyphDictionary.Keys.ElementAt(KeyCounter);
                theGlyphDictionary[theKeyName] += "])";
            }
            // We end up with a regular expression of the form (.[<combchar>|<combchar>!..])
            // I.e. we match any character followed by a combining character.
            Registry.SetValue(keyName, "AggregateDir", AggregateDir);
            return true;
        }

        private void DecompGlyphBox_TextChanged(object sender, EventArgs e)
        {
            // The file name of the glyph file has changed so we need to reload
            // If the file doesn't exist, then we won't even try to load it
            TextBox theTextBox = (TextBox)sender;
            theGlyphDictionary.Clear();
            GlyphsLoaded = !File.Exists(theTextBox.Text);
            return;
        }

        private void btnSaveErrorList_Click(object sender, EventArgs e)
        {
            string theErrorListFile = "";
            if (Individual)
            {
                theErrorListFile = ErrorListBox.Text;
            }
            else
            {
                theErrorListFile = BulkErrorListBox.Text;
            }

            WriteDataGridView(theErrorListFile, listNormalisedErrors);
            Registry.SetValue(keyName, "ErrorDir", ErrorDir);

        }

        private void btnGetFont_Click(object sender, EventArgs e)
        {
            if (fontDialog1.ShowDialog() == DialogResult.OK)
            {
                FontBox.Text = fontDialog1.Font.Name;
                Registry.SetValue(keyName, "Font", FontBox.Text); //remember it
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load the encodings
            foreach (EncodingInfo theEncodingInfo in Encoding.GetEncodings())
            {
                Encoding theEncoding = theEncodingInfo.GetEncoding();
                theEncodingDictionary[theEncodingInfo.DisplayName] = theEncoding; // save it in the dictionary
            }
            this.SetEncoding(Registry.GetValue(keyName, "Encoding", "Western European (Windows)").ToString());  // the default is ANSI Code Page 1252
            Enable_Controls(this.IndivOrBulk.TabPages["IndividualFile"], true); // Enable the individual controls
            Enable_Controls(this.IndivOrBulk.TabPages["Bulk"], false); // and disable the bulk controls


        }

        private void btnGetEncoding_Click(object sender, EventArgs e)
        {
            EncodingForm theEncodingForm = new EncodingForm();
            DialogResult theResult = theEncodingForm.ShowDialog(this);
        }
        public void SetEncoding(string theEncodingName)
        {
            theEncoding = theEncodingDictionary[theEncodingName];
            EncodingTextBox.Text = theEncodingName;
            Registry.SetValue(keyName, "Encoding", theEncodingName); // Remember the encoding
        }
        public string GetEncoding()
        {
            if (theEncoding == null)
            {
                return "";
            }
            else
            {
                return theEncoding.EncodingName;
            }
        }

        private void AggregateStatsBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            if (theBox.Text == "")
            {
                AggregateStats.Enabled = false;
                AggregateStats.Checked = false;
                btnSaveAggregateStats.Enabled = false;
            }
            else
            {
                AggregateStats.Enabled = true;
                AggregateStats.Checked = true;

            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (AggregateStats.Checked && !AggregateSaved)
            {
                AggregateSaved = WriteOutput(theAggregateDictionary, "", AggregateDir, AggregateStatsBox.Text, AggregateFileList, true);
            }

        }

        private void btnSaveAggregateStats_Click(object sender, EventArgs e)
        {
            AggregateSaved = WriteOutput(theAggregateDictionary, "", AggregateDir, AggregateStatsBox.Text, AggregateFileList, true);
            btnSaveAggregateStats.Enabled = !AggregateSaved;
        }

        private void CombDecomposedChars_CheckStateChanged(object sender, EventArgs e)
        {
            btnDecompGlyph.Enabled = CombDecomposedChars.Checked;
        }

        private void CombiningCharacters_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "CombiningCharacters.xlsm");
            System.Diagnostics.Process.Start(HelpPath);

        }

        private void Individual_Entered(object sender, EventArgs e)
        {
            // Clear things if we have switched from bulk to individual
            if (Individual == false)
            {
                Individual = true;
                ClearLists();
                btnSaveXML.Enabled = XMLFileBox.Text != ""; ;  //reveal the XML file button
                btnSaveXML.Visible = true;
                Enable_Controls(sender, true);
                Enable_Controls(((TabControl)((Control)sender).Parent).TabPages["Bulk"], false);  // Disable bulk
            }
        }
        private void Enable_Controls(object sender, bool Enable)
        {
            Control theSender = (Control)sender;
            foreach (Control theControl in theSender.Controls)
            {
                Type theType = theControl.GetType();
                if (theType.Name == "TextBox" || theType.Name == "Button" )
                {
                    theControl.Enabled = Enable;
                    theControl.Visible = Enable;
                    theControl.TabStop = Enable;
                }
            }

        }
        private void Bulk_Entered(object sender, EventArgs e)
        {
            if (Individual == true)
            {
                Individual = false;
                InputFolderBox.Text = InputDir;
                OutputFolderBox.Text = OutputDir;
                btnSaveFontList.Enabled = BulkFontListFileBox.Text != "";
                btnSaveStyles.Enabled = BulkStyleListBox.Text != "";
                btnSaveErrorList.Enabled = BulkErrorListBox.Text != "";
                btnGetFont.Enabled = true;
                btnGetEncoding.Enabled = true;
                btnSaveXML.Enabled = false;  // hide the XML file button.
                btnSaveXML.Visible = false;
                Enable_Controls(sender, true);
                Enable_Controls(((TabControl)((Control)sender).Parent).TabPages["IndividualFile"], false);  // Disable the individual.
                ClearLists();
            }
        }
        private void ClearLists()
        {
            //  Clear lists when switching back and forth between individual and bulk processing
            openInputDialogue.FileName = ""; // Forget any files we selected.
            FontList.Items.Clear();
            listStyles.Rows.Clear();
            listNormalisedErrors.Rows.Clear();
            btnAnalyse.Enabled = false;

        }
        private void btnInputFolder_Click(object sender, EventArgs e)
        {
            FolderDialogue.RootFolder = Environment.SpecialFolder.Desktop;
            FolderDialogue.ShowNewFolderButton = false;
            FolderDialogue.SelectedPath = GetDirectory("InputDir");
            if (FolderDialogue.ShowDialog() == DialogResult.OK)
            {
                InputFolderBox.Text = FolderDialogue.SelectedPath;
            }

        }

        private void InputFolderBox_TextChanged(object sender, EventArgs e)
        {
            // A different folder so we need to clear everything it and clear lots of things.

            theXMLDictionary.Clear();
            theTextDictionary.Clear();
            listStyles.Rows.Clear(); // clear the styles
            FontList.Items.Clear(); // and the fonts
            theStyleDictionary.Clear();
            listNormalisedErrors.Rows.Clear();
            theStyleDictionary.Clear();
            if (InputFolderBox.Text != "")
            {
                InputDir = InputFolderBox.Text;
                Registry.SetValue(keyName, "InputDir", InputDir);
                btnSelectFiles.Enabled = true;
            }

        }

        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            openInputDialogue.Multiselect = true;
            openInputDialogue.InitialDirectory = GetDirectory("InputDir");
            if (openInputDialogue.ShowDialog() == DialogResult.OK && openInputDialogue.FileNames.Count() > 0)
            {
                theAggregateDictionary.Clear();  // Make sure it is empty before we analyse in bulk
                FileCounter.Text = "0";

                if (OutputFolderBox.Text != "")
                {
                    btnAnalyse.Enabled = true;
                    FontList.Items.Clear();
                    listStyles.Rows.Clear();
                    WriteIndividualFile.Checked = true;
                    WriteIndividualFile.Enabled = true;
                    btnListFonts.Enabled = openInputDialogue.FileNames.Count() > 0;
                    btnGetStyles.Enabled = openInputDialogue.FileNames.Count() > 0;

                    btnSaveFontList.Enabled = BulkFontListFileBox.Text != "" && btnListFonts.Enabled;
                    btnSaveStyles.Enabled = BulkStyleListBox.Text != "" && btnGetStyles.Enabled;
                    btnSaveErrorList.Enabled = BulkErrorListBox.Text != "" && openInputDialogue.FileNames.Count() > 0;
                }
                else
                {
                    WriteIndividualFile.Enabled = false;
                    WriteIndividualFile.Checked = false;
                }
            }


        }

        private void btnCharStatFolder_Click(object sender, EventArgs e)
        {
            FolderDialogue.Description = "Select the directory to receive the individual files";
            if (OutputFolderBox.Text != "")
            {
                FolderDialogue.SelectedPath = OutputFolderBox.Text;
            }
            else
            {
                FolderDialogue.SelectedPath = GetDirectory("OutputDir");
            }
            FolderDialogue.ShowNewFolderButton = true;
            if (FolderDialogue.ShowDialog() == DialogResult.OK)
            {
                OutputFolderBox.Text = FolderDialogue.SelectedPath;
                OutputDir = OutputFolderBox.Text;
                Registry.SetValue(keyName, "OutputDir", OutputDir);
            }
        }
        private void OutputFileSuffixBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void OutputDirBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void FontListFileBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveFontList.Enabled = theBox.Text != "";
        }

        private void StyleListFileBox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveStyles.Enabled = theBox.Text != "";

        }

        private void BulkErrorListbox_TextChanged(object sender, EventArgs e)
        {
            TextBox theBox = (TextBox)sender;
            btnSaveErrorList.Enabled = theBox.Text != "";

        }

       




    }
    class CharacterDescriptor
    {
        public string FileName;
        public string Font;
        public string Text;

        // Constructor
        public CharacterDescriptor(string FileName, string Font, string Text)
        {
            this.FileName = FileName;
            this.Font = Font;
            this.Text = Text;
        }
        public CharacterDescriptor(string Font, string Text)
        {
            this.FileName = null;
            this.Font = Font;
            this.Text = Text;
        }

        public CharacterDescriptor(string Text)
        {
            this.FileName = null;
            this.Text = Text;
            this.Font = null;
        }
        public CharacterDescriptor(CharacterDescriptor theCharacterDescriptor)
        {
            this.FileName = theCharacterDescriptor.FileName;
            this.Font = theCharacterDescriptor.Font;
            this.Text = theCharacterDescriptor.Text;
        }
        public CharacterDescriptor()
        {
            this.FileName = null;
            this.Font = null;
            this.Text = null;
        }
    }
    class CharacterEqualityComparer : EqualityComparer<CharacterDescriptor>
    {
        override public bool Equals(CharacterDescriptor key1, CharacterDescriptor key2)
        {
            return (key1.FileName == key2.FileName) & (key1.Font == key2.Font) & (key1.Text == key2.Text);
        }
        override public int GetHashCode(CharacterDescriptor key)
        {
            return (key.FileName + "\r" + key.Text + "\r" + key.Font).GetHashCode();
        }
    }

}






