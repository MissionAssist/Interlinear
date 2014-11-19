/*
 * Interlinear - a program to take two Word documents, segment them into paragraphs of up to 20 words in length and write them to Excel
 * with the first (legacy) file in odd rows and the second (Unicode) file in even rows. This enables visual checking without the need to try
 * to do side-by-side comparisons.  It depends on both Word and Excel being installed on the computer.
 * 
 * It was writting as part of a MissionAssist project to convert documents in legacy fonts to Unicode.  Much of the logic is attributable to
 * Dennis Pepler, but the code here was written by Stephen Palmstrom.
 * 
 * Copyright © MissionAssist 2014 and distributed under the terms of the GNU General Public License (http://www.gnu.org/licenses/gpl.html)
 * 
 * Last modified on 9 September 2013 by Stephen Palmstrom (stephen.palmstrom@outlook.com) who asserts the right to be regarded as the author of this program
 * 
 * Acknowledgement is due to Dennis Pepler who worked out how to scan stories etc.
*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.Common;
using System.Threading;
using System.Diagnostics;
using Microsoft.Win32;
using System.Xml;
using System.Xml.XPath;
using WordApp = Microsoft.Office.Interop.Word._Application;
using WordRoot = Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word._Document;
using ExcelApp = Microsoft.Office.Interop.Excel._Application;
using ExcelRoot = Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel.Application;
using WorkBook = Microsoft.Office.Interop.Excel._Workbook;


using Office = Microsoft.Office.Core;


namespace Interlinear
{
    public partial class Form1 : Form
    {
        private WordApp wrdApp;
        WordAppOptions theOptions;
        private Document InputDoc;
        private Document OutputDoc;
        private ExcelApp excelApp;
        private ExcelAppOptions theExcelOptions;
        object missing = Type.Missing;
        private const string theSpace = " ";
        private string[] theMessage = new string[2] {"Legacy text is in odd rows from file ", "Unicode text is in even rows from file "};
        private ExcelRoot.XlRgbColor[] CellColour = new ExcelRoot.XlRgbColor[2] {ExcelRoot.XlRgbColor.rgbYellow, ExcelRoot.XlRgbColor.rgbLightBlue};
        private int MaxParagraphs = 0;
        private bool Paused = false;
        private bool CloseApp = false;
        //  Directories
        private string LegacyInputDir = "";
        private string LegacyOutputDir = "";
        private string UnicodeInputDir = "";
        private string UnicodeOutputDir = "";
        private string ExcelDir = "";

        const string userRoot = "HKEY_CURRENT_USER";
        const string subkey = "Software\\MissionAssist\\Interlinear";
        const string keyName = userRoot + "\\" + subkey;


        public Form1()
        {
            InitializeComponent();
            wrdApp = new Word();
            wrdApp.Visible = false;
            theOptions = new WordAppOptions(wrdApp);  // Save Word setting
            excelApp = new Excel();  // open Excel
            //excelApp.Visible = false; 
            saveLegacyFileDialog.SupportMultiDottedExtensions = true;
            saveUnicodeFileDialog.SupportMultiDottedExtensions = true;
            //Wordcount.SetToolTip(WordsPerLine, "If you want more than eight words per line, they must be in multiples of four");
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
                LegacyInputDir = GetDirectory("LegacyInputDir");
                LegacyOutputDir = GetDirectory("LegacyOutputDir", LegacyInputDir);
                UnicodeInputDir = GetDirectory("UnicodeInputDir");
                UnicodeOutputDir = GetDirectory("UnicodeOutputDir", UnicodeInputDir);
                ExcelDir = GetDirectory("ExcelDir", UnicodeOutputDir);
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace, "Failed to get directories", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            // The output directories
            saveLegacyFileDialog.InitialDirectory = LegacyOutputDir;
            saveUnicodeFileDialog.InitialDirectory = UnicodeOutputDir;

 

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
                MessageBox.Show(Ex.Message + "\r" + Ex.StackTrace + "\rkeyName " + keyName + "\rValueName " + ValueName +
                "\rDefaultPath " + DefaultPath, "Can't read registry", MessageBoxButtons.OK);
                Application.Exit();
            }

            return theDirectory;
        }


        private void btnGetInputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                LegacyInputDir = HandleInputFile(txtLegacyInput, txtLegacyOutput, btnSegmentLegacy, openLegacyFileDialog, saveLegacyFileDialog,
                    btnLegacyToExcel, chkLegacyToExcel, LegacyInputDir, "LegacyInputDir");
            }
            else
            {
                UnicodeInputDir = HandleInputFile(txtUnicodeInput, txtUnicodeOutput, btnSegmentUnicode, openUnicodeFileDialog,saveUnicodeFileDialog,
                    btnUnicodeToExcel, chkLegacyToExcel, UnicodeInputDir, "UnicodeInputDir");
            }

        }
        private void chkSendtoExcel_Change(object sender, EventArgs e)
        {
            // handle a change in the just send to Excel buttons
            CheckBox theCheckBox = (CheckBox)sender;
            if (theCheckBox.Parent.Text == "Legacy")
            {
                HandleCheckBoxChange(txtLegacyOutput.Text, btnLegacyToExcel, theCheckBox.Checked);
            }
            else
            {
                HandleCheckBoxChange(txtUnicodeOutput.Text, btnUnicodeToExcel, theCheckBox.Checked);
            }
        }
        private void HandleCheckBoxChange(string OutputText, Button SendtoExcel, bool Checked)
        {
            SendtoExcel.Enabled = File.Exists(OutputText) && Checked;
        }
        private string HandleInputFile(TextBox InputText, TextBox OutputText, Button SegmentButton, OpenFileDialog theOpenFileDialog, SaveFileDialog theSaveFileDialog,
            Button ExcelButton, CheckBox SendtoExcel, string DefaultDir, string ValueName)
        {
            /*
             * Handle the input file dialog.
             */
            string tmpString = DefaultDir;
            theOpenFileDialog.InitialDirectory = DefaultDir;
            if (theOpenFileDialog.ShowDialog() == DialogResult.OK )
            {
                InputText.Text = theOpenFileDialog.FileName;
                SegmentButton.Enabled = true & OutputText.Text.Length > 0;
                if (Path.GetExtension(InputText.Text) == ".doc")
                {
                    theSaveFileDialog.FilterIndex = 1; // .doc
                }
                else
                {
                    theSaveFileDialog.FilterIndex = 2; // .docx
                }
                if (OutputText.Text == "")
                {
                    // the text box is empty so we fill it.
                    OutputText.Text = Path.Combine(theSaveFileDialog.InitialDirectory, Path.GetFileNameWithoutExtension(InputText.Text) +
                        " (Segmented)" + Path.GetExtension(InputText.Text));
                }
                theSaveFileDialog.FileName = OutputText.Text;
                HandleOutputFile1(InputText.Text, OutputText.Text, SegmentButton, ExcelButton, SendtoExcel);  // process further
                btnSegmentBoth.Enabled = btnSegmentLegacy.Enabled && btnSegmentUnicode.Enabled;
                btnBothToExcel.Enabled = File.Exists(txtLegacyOutput.Text) && File.Exists(txtUnicodeOutput.Text) && txtExcelOutput.Text.Length > 0;
                tmpString = Path.GetDirectoryName(theOpenFileDialog.FileName);
                Registry.SetValue(keyName, ValueName, tmpString);
             };
            return tmpString;
        }
        private void btnGetOutputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                LegacyOutputDir = HandleOutputFile(txtLegacyInput, txtLegacyOutput, saveLegacyFileDialog, btnSegmentLegacy, btnLegacyToExcel, chkLegacyToExcel,
                    LegacyOutputDir, "LegacyOutputDir");
            }
            else
            {
                UnicodeOutputDir = HandleOutputFile(txtUnicodeInput, txtUnicodeOutput, saveUnicodeFileDialog, btnSegmentUnicode, btnUnicodeToExcel, chkUnicodeToExcel,
                    UnicodeOutputDir, "UnicodeOutputDir");

            }
        }
        private string HandleOutputFile (TextBox theInputBox, TextBox theOutputBox, SaveFileDialog theDialog, Button SegmentButton,
            Button ExcelButton, CheckBox SendtoExcel, string DefaultDir, string ValueName )
        {
            theDialog.InitialDirectory = DefaultDir;
            string tmpString = DefaultDir;
            if (theDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                theOutputBox.Text = theDialog.FileName;
                HandleOutputFile1(theInputBox.Text, theOutputBox.Text, SegmentButton, ExcelButton, SendtoExcel);  // process further
                tmpString = Path.GetDirectoryName(theDialog.FileName);
                Registry.SetValue(keyName, ValueName, tmpString);
            }
            return tmpString; 
        }
        private void HandleOutputFile1(string InputText, string OutputText, Button SegmentButton, Button ExcelButton, CheckBox SendtoExcel)
        {
                SegmentButton.Enabled = OutputText.Length > 0 && File.Exists(InputText);  // only enable if both boxes filled in
                ExcelButton.Enabled = OutputText.Length > 0 && txtExcelOutput.Text.Length > 0 && File.Exists(OutputText) && SendtoExcel.Checked;
                /*
                 * If both individual segment buttons are enabled, we enable the segment both button, too.
                 */
                btnSegmentBoth.Enabled = btnSegmentLegacy.Enabled && btnSegmentUnicode.Enabled;
                btnBothToExcel.Enabled = File.Exists(txtLegacyOutput.Text) && File.Exists(txtUnicodeOutput.Text) && txtExcelOutput.Text.Length > 0;


        }
        private void btnGetExcelOutput_Click(object sender, EventArgs e)
        {
            /*
             * Handle the input file dialog.
             */
            saveExcelFileDialog.InitialDirectory = ExcelDir;
            if (saveExcelFileDialog.ShowDialog() == DialogResult.OK)
            {
                Button theButton = (Button)sender;  // A button click triggers this
                txtExcelOutput.Text = saveExcelFileDialog.FileName;
                chkLegacyToExcel.Enabled = true & txtExcelOutput.Text.Length > 0;
                chkUnicodeToExcel.Enabled = true & txtExcelOutput.Text.Length > 0;
                saveExcelFileDialog.FileName = txtExcelOutput.Text;
                // Enable both write to Excel operations.
                chkLegacyToExcel.Checked = true;
                chkUnicodeToExcel.Checked = true;
                // If the segmented files exist we can send them to Excel without resegmenting
                btnLegacyToExcel.Enabled = File.Exists(txtLegacyOutput.Text);
                btnUnicodeToExcel.Enabled = File.Exists(txtUnicodeOutput.Text);
                btnBothToExcel.Enabled = File.Exists(txtLegacyOutput.Text) && File.Exists(txtUnicodeOutput.Text);
                ExcelDir = Path.GetDirectoryName(saveExcelFileDialog.FileName); // Remember the directory
                Registry.SetValue(keyName, "ExcelDir", ExcelDir); // for future reference, too.
            }

        }
        private void btnSegmentInput_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            theButton.Enabled = false;  // Disable as we have started running.
            btnClose.Enabled = false;
            boxProgress.Items.Clear();  // empty the progress box

            tabControl1.SelectTab("Progress");
            Application.DoEvents();
            if (theButton.Parent.Text == "Legacy")
            {
                SegmentFile(txtLegacyInput.Text, txtLegacyOutput.Text, txtLegacyWordCount, chkLegacyToExcel, chkLegacyAddSpace, false);
            }
            else
            {
                SegmentFile(txtUnicodeInput.Text, txtUnicodeOutput.Text, txtUnicodeWordCount, chkUnicodeToExcel, chkUnicodeAddSpace, true);
            }
            theButton.Enabled = true;  // enable it again
            btnClose.Enabled = true;
            Application.DoEvents();
            System.Media.SystemSounds.Beep.Play();  // and beep

        }
        private void btnSegmentBoth_Click(object sender, EventArgs e)
        {
            //  Segment both files in one go
            Button theButton = (Button)sender;
            tabControl1.SelectTab("Progress");
            theButton.Enabled = false;
            btnClose.Enabled = false;
            boxProgress.Items.Clear();  // empty the progress box
            SegmentFile(txtLegacyInput.Text, txtLegacyOutput.Text, txtLegacyWordCount, chkLegacyToExcel, chkLegacyAddSpace, false);
            SegmentFile(txtUnicodeInput.Text, txtUnicodeOutput.Text, txtUnicodeWordCount, chkUnicodeToExcel, chkUnicodeAddSpace, true);
            if (chkLegacyToExcel.Checked || chkUnicodeToExcel.Checked)
            {
                //MakeInterlinear(excelApp);  // Make the interlinear worksheet, too
            }
            theButton.Enabled = true;
            btnClose.Enabled = true;
            System.Media.SystemSounds.Beep.Play();  // and beep

        }
        private void FinalCatch(Exception e)
        {
            MessageBox.Show(e.Message + "\r" + e.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            QuitWord(false);  // don't save the output
            this.Close();
        }
        private void SegmentFile(String theInputFile, String theOutputFile, TextBox txtNumberOfWords, CheckBox SendToExcel, CheckBox AddSpaceAfterRange, bool EvenRows)
        {
            /*
             * This is where we do all the segmentation and, if desired, writing to Excel
             */
            try
            {
                System.Diagnostics.Stopwatch theStopwatch = new System.Diagnostics.Stopwatch();
                theStopwatch.Start();
                toolStripStatusLabel1.Text = "Starting...";
                int NumberOfWords;
                AddSpaceAfterRange.Enabled = false;  // We don't want this changing during our run.
                int RowCounter = 0;
                progressBar1.Value = 0;

                Application.DoEvents();
                try
                {
                    InputDoc = wrdApp.Documents.OpenNoRepairDialog(theInputFile, missing, true);  // Read only, and we don't want the repair dialog
                    File.Delete(theOutputFile); // delete the output file
                    OutputDoc = wrdApp.Documents.Add();  // a new blank document
                    OutputDoc.SaveAs2(theOutputFile, InputDoc.SaveFormat);  // Save the output document
                }
                catch (Exception e)
                {
                    DialogResult theResult = MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tabControl1.SelectTab("Setup");
                    return;

                }


                // process Excel if desired
                if (SendToExcel.Checked)
                {
                    RowCounter = InitialiseExcel(excelApp, EvenRows, theInputFile);
                    if (RowCounter == 0)
                    {
                        return;  // we couldn't open the file
                    }

                }
                // Size the progressbar depending on how many replacments we do
                if (WordsPerLine.Value > 7)
                {
                    progressBar1.Maximum = 3;
                }
                else
                {
                    progressBar1.Maximum = 2;
                }
                /*
                 * Set various Word options to optimise performance
                 * 
                 */
                boxProgress.Items.Add("**** Starting processing " + Path.GetFileName(theInputFile));
                btnPauseResume.Enabled = true;
                Application.DoEvents();
                //OptimiseDoc(InputDoc);
                OptimiseDoc(OutputDoc);
                theOptions.OptimiseApp(wrdApp);  // Optimise the application
                wrdApp.ScreenUpdating = false; // Turn off updating the screen
                wrdApp.ActiveWindow.ActivePane.View.ShowAll = false;  // Don't show special marks
                //wrdApp.Selection.WholeStory(); // Make sure we've selected everything
                wrdApp.ScreenUpdating = false; // Turn off screen updating


                NumberOfWords = InputDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticWords, false);
                System.Collections.Generic.List<WordRoot.Range> TextFrames = new System.Collections.Generic.List<WordRoot.Range>();

                boxProgress.Items.Add("The document contains " + NumberOfWords.ToString() + " words");

                txtNumberOfWords.Text = NumberOfWords.ToString(); // the number of words in the document
                 /*
                  * Start copying the text from the input document to the output document
                  */
                /*
                 * This makes sure we pick up all Stories in the document
                 */
                WordRoot.WdStoryType StoryJunk = InputDoc.Sections[1].Headers[(WordRoot.WdHeaderFooterIndex)1].Range.StoryType;
                /*
                 * Now go through each story and write it to the output document
                 */
                boxProgress.Items.Add("Starting to copy the document");
                System.Diagnostics.Stopwatch theStopwatch2 = new System.Diagnostics.Stopwatch();
                theStopwatch2.Start();
                progressBar1.Value = 0;
                int TotalCharacters = InputDoc.Characters.Count;
                progressBar1.Maximum = TotalCharacters;
                boxProgress.Items.Add("Estimated total of " + TotalCharacters.ToString() + " characters");
                // Select the beginning of the
                OutputDoc.ActiveWindow.Selection.WholeStory();  // Select the whole document to start with.
                int RangeCounter = 0;
                int CharacterCounter = 0;
                //
                //  First pass
                //
                foreach (WordRoot.Range rngStory in InputDoc.StoryRanges)
                {
                    WordRoot.Range tmpStory = rngStory;
                    do
                    {
                        CharacterCounter = InsertAfter(tmpStory, CharacterCounter, AddSpaceAfterRange.Checked, theStopwatch, theStopwatch2);
                        if (tmpStory.StoryType == WordRoot.WdStoryType.wdTextFrameStory)
                        {
                            TextFrames.Add(tmpStory);  // Remember the text frame
                        }
                        tmpStory = tmpStory.NextStoryRange;  // trace through a link of substories
                    }
                    while (tmpStory != null);
                    RangeCounter++;
                    progressBar1.Value = RangeCounter;
                }
                //
                //  Second pass
                 foreach (WordRoot.Range rngStory in InputDoc.StoryRanges)
                {
                    WordRoot.Range tmpStory = rngStory;
                    do
                    {
                    for (int i = 1; i <= tmpStory.ShapeRange.Count; i++)
                    {
                        //
                        // we've remembered the text frames and check on fonts and text - if both agree with previously found frames, we skpt them.
                        //
                        if (tmpStory.ShapeRange[i].TextFrame.HasText != 0)
                        {
                            try
                            {
                                WordRoot.Range theRange = tmpStory.ShapeRange[i].TextFrame.TextRange;
                                bool NotFound = true;  // Assume we didn't find it.
                                foreach (WordRoot.Range theOldRange in TextFrames)
                                {

                                    if (CompareRanges(theOldRange, theRange))
                                    {
                                        // We've already copied this range
                                        NotFound = false;
                                        break;
                                    }
                                }
                                if (NotFound)
                                {
                                    CharacterCounter = InsertAfter(theRange, CharacterCounter, AddSpaceAfterRange.Checked, theStopwatch, theStopwatch2);  // Add it to the document
                                }
           

                            }
                            catch (Exception theException)
                            {
                                // to help us debug errors
                                DialogResult theResult = MessageBox.Show(theException.Message + "\r Index is" + i.ToString(),
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    tmpStory = tmpStory.NextStoryRange;
                        //tmpStory = null;  // end the looop
                }
                while (tmpStory != null);
                 }

                InputDoc.Close(false);  // and close the input document as we no longer need it.
                InputDoc = null;  // and free up memory
                OutputDoc.Save();  // and save the output document
                long ElapsedTime = theStopwatch2.ElapsedMilliseconds;
                toolStripStatusLabel1.Text = "Copy complete.";
                progressBar1.Value = 0;
                boxProgress.Items.Add("Document (" + CharacterCounter.ToString() + " characters) copied in " + (ElapsedTime/1000.0).ToString("f2") + " seconds or " + (CharacterCounter*1000.0/ElapsedTime).ToString("f2") + " cps");
                theStopwatch2.Stop();
                theStopwatch2 = null;
                btnPauseResume.Enabled = false;
                Application.DoEvents();

                /*
                 * Clean up the document
                 */
                CleanWordText(wrdApp, OutputDoc);
                
                /*
                  * Now start splitting into a number of space-separated words, i.e. segmenting it.
                  */
                Segment(wrdApp, OutputDoc.ActiveWindow.Selection, (int)WordsPerLine.Value, NumberOfWords);
                OutputDoc.Save();
                boxProgress.Items.Add(Path.GetFileName(theOutputFile) + " saved in " + theStopwatch.Elapsed.ToString("hh\\:mm\\:ss\\.f"));
                 if (SendToExcel.Checked)
                 {
                     // We'll send the information to Excel
                     FillExcel(excelApp, wrdApp, OutputDoc, RowCounter);
                 }
                 else
                 {
                     OutputDoc.Close();  // and close it
                     OutputDoc = null;  // and free up the memory
                 }
                wrdApp.ScreenUpdating = true; // turn on screen updating
                btnPauseResume.Enabled = false;
                progressBar1.Value = 0;
                theOptions.RestoreApp(wrdApp); // Restore the settings
                boxProgress.Items.Add("Completed in " + theStopwatch.Elapsed.ToString("hh\\:mm\\:ss\\.f"));
                toolStripStatusLabel1.Text = "Completed";
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
            AddSpaceAfterRange.Enabled = true; // Enable us to change settings again.
        }
        private bool CompareRanges(WordRoot.Range RangeOne, WordRoot.Range RangeTwo)
        {
            // Compare two ranges
            return (RangeOne.Font.Name == RangeTwo.Font.Name && RangeOne.Text == RangeTwo.Text);
        }
        private int InsertAfter(WordRoot.Range theRange, int CharacterCounter, bool AddSpaceAfterRange, Stopwatch theStopwatch, Stopwatch theStopwatch2)
        {   
            /*
             * If the paragraph has a single font, insert the whole story followed by a space with the single font.
             * If not, insert paragraph by paragraph
             * If the paragraph doesn't have a single font, insert word by word followed by a space.
             * If the word doesn't have a signle font, insert character by character
             * Thiw will, I hope, avoid the need for sophisticated cleanup operations.
             */
            int tmpCounter = CharacterCounter;
            Stopwatch theStopwatch3 = new Stopwatch();
            theStopwatch3.Start();
            /*
             * See if we have any symbols that we need to look for
             */
            bool FoundSymbol = theRange.get_XML(false).Contains("w:sym");
            //boxProgress.Items.Add("Looked for symbol in " + theStopwatch3.ElapsedMilliseconds.ToString("f2") + " and found " + FoundSymbol.ToString());
            if (theRange.Font.Name != "")
            {
                CharacterCounter = InsertAfter2(theRange, AddSpaceAfterRange, CharacterCounter, ref tmpCounter, theStopwatch, theStopwatch2, theStopwatch3, FoundSymbol);
                return CharacterCounter;
            }
            else
            {

                foreach (WordRoot.Paragraph theParagraph in theRange.Paragraphs)
                {
                    if (theParagraph.Range.Font.Name != "")
                    {
                        CharacterCounter = InsertAfter2(theParagraph.Range, AddSpaceAfterRange, CharacterCounter, ref tmpCounter, theStopwatch, theStopwatch2, 
                            theStopwatch3, FoundSymbol);
                    }
                    else
                    {
                        foreach (WordRoot.Range theWord in theParagraph.Range.Words)
                        {
                            if (theWord.Font.Name != "")
                            {
                                CharacterCounter = InsertAfter2(theWord, AddSpaceAfterRange, CharacterCounter, ref tmpCounter, theStopwatch, theStopwatch2,
                                    theStopwatch3, FoundSymbol);
                            }
                            else
                            {
                                foreach (WordRoot.Range theCharacter in theWord.Characters)
                                {
                                    int CharCount = theWord.Characters.Count;
                                    CharacterCounter = InsertAfter2(theCharacter, false, CharacterCounter, ref tmpCounter, theStopwatch,
                                        theStopwatch2, theStopwatch3, FoundSymbol);
                                }
                            }
                        }

                    }
                }

            }
         return CharacterCounter;
        }


        private int InsertAfter2(WordRoot.Range theRange, bool AddSpace, int CharacterCounter, ref int tmpCounter, 
            Stopwatch theStopwatch, Stopwatch theStopwatch2, Stopwatch theStopwatch3, bool FoundSymbol)
        {
            //Stopwatch theStopwatch4 = new Stopwatch();
            //theStopwatch4.Start();
            /*
             * Make sure we retrieve all text
             */
            const string wordmlNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
            theRange.TextRetrievalMode.IncludeFieldCodes = false;
            theRange.TextRetrievalMode.IncludeHiddenText = true;
            string tmpString = theRange.Text;
            string XMLText = "";
            bool Inserted = false;
            WordRoot.Font theFont = new WordRoot.Font();
            theFont = theRange.Font;
            if (theFont.Color != WordRoot.WdColor.wdColorAutomatic)
            {
                theFont.Color = WordRoot.WdColor.wdColorAutomatic; // Make the colour automatic
            }
 
            XmlDocument theXMLDocument;
            XmlNamespaceManager nsManager;
            XmlNodeList theNodeList;
            if (FoundSymbol && tmpString != "\r\a") // no point in looking if the range doesn't have a symbol
            {
                try
                {
                    XMLText = theRange.get_XML(false);
                    if (XMLText.Contains("<w:sym"))
                    {
                        // We have  symbol in here and we need to process it
                        theXMLDocument = new XmlDocument();
                        theXMLDocument.LoadXml(XMLText);
                        nsManager = new XmlNamespaceManager(theXMLDocument.NameTable);
                        nsManager.AddNamespace("w", wordmlNamespace);
                        nsManager.AddNamespace("wx", wordmlNamespace);
                        XmlNode theRoot = theXMLDocument.DocumentElement;
                        /*
                         * Look for text or symbols in the range
                         */
                        theNodeList = theRoot.SelectNodes(@"(//w:r/w:t | //w:r/w:sym)", nsManager);
                       foreach (XmlNode theData in theNodeList)
                        {
                            OutputDoc.ActiveWindow.Selection.EndKey(WordRoot.WdUnits.wdStory, false);  // Make sure we are at the end of the story.
                            // we look the range structures
                            switch (theData.Name)
                            {
                                case "w:t":
                                    // we have text
                                    
                                    tmpString = theData.InnerText;
                                    theFont = theRange.Font;
                                    CharacterCounter += tmpString.Length;          
                                    OutputDoc.ActiveWindow.Selection.InsertAfter(tmpString);
                                    OutputDoc.ActiveWindow.Selection.EndKey(WordRoot.WdUnits.wdStory, true);  // Make sure we are at the end of the story.
                                    OutputDoc.ActiveWindow.Selection.Range.Font = theFont; // Set the font of the text we have just inserted.
 
                                    break;
                                case "w:sym":
                                    // We have a symbol so we shall insert it
                                    string FontName = theData.Attributes["w:font"].Value;
                                    string theSymbolValue = theData.Attributes["w:char"].Value;
                                    int theChar = Convert.ToInt16(theSymbolValue, 16);  // get the character number
                                    OutputDoc.ActiveWindow.Selection.InsertSymbol(theChar, FontName, true);  // insert the symbol
                                    break;
                            }
                           Inserted = true;
                        }
                    }
                }
                catch (Exception Ex)
                {
                    string theMessage = Ex.Message;
                    MessageBox.Show("Error", theMessage, MessageBoxButtons.OK);
                    
                }
            }
            if (!Inserted)
            {
                // We have normal text or are in a table
                OutputDoc.ActiveWindow.Selection.EndKey(WordRoot.WdUnits.wdStory, false);  // Make sure we are at the end of the story.
                CharacterCounter += tmpString.Length;
                OutputDoc.ActiveWindow.Selection.InsertAfter(tmpString);
                OutputDoc.ActiveWindow.Selection.EndKey(WordRoot.WdUnits.wdStory, true);  // Make sure we are at the end of the story.
                OutputDoc.ActiveWindow.Selection.Range.Font = theFont; // Set the font of the text we have just inserted.
 

            }
            //boxProgress.Items.Add("Copied text after " + theStopwatch4.ElapsedTicks);
            if (AddSpace)
            {
               OutputDoc.ActiveWindow.Selection.InsertAfter(theSpace);
               //OutputDoc.ActiveWindow.Selection.Range.Font = theRange.Font; // Set the font of the text we have just inserted.

            }
           

 
             //OutputDoc.ActiveWindow.Selection.EndKey(WordRoot.WdUnits.wdStory, 1);  // Move the selection to the end.
            //boxProgress.Items.Add("Inserted " + theRange.Text + " " + theStopwatch3.Elapsed);
            //theStopwatch4.Stop();
            //theStopwatch4 = null;
            int DeltaChars = CharacterCounter - tmpCounter;
            if (DeltaChars > 200)
            {
                long ElapsedTime = theStopwatch3.ElapsedTicks; 
                toolStripStatusLabel1.Text = "Copied " + DeltaChars.ToString() + " characters at " + ((float)DeltaChars * Stopwatch.Frequency / ElapsedTime).ToString("f2") + " per sec";
                theStopwatch3.Restart();
                progressBar1.Value = (int)Math.Min(progressBar1.Maximum, CharacterCounter);
                tmpCounter = CharacterCounter;
                OutputDoc.UndoClear();  // Clear the undo buffer lest it is slowing things down.
                Application.DoEvents();
            }
            PauseForThought(theStopwatch, theStopwatch2, theStopwatch3);  // if Pause is clicked we wait in this routine.

            return CharacterCounter;
        }

        private void OptimiseDoc(Document theDoc)
        {
            // Turn off various options to speed up Word
            theDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
            theDoc.ActiveWindow.View.Draft = true;  // Draft View
            theDoc.ShowSpellingErrors = false;  // Don't show spelling errors
            theDoc.ShowGrammaticalErrors = false; // Don't show grammar errors
            theDoc.AutoHyphenation = false;
            

        }

        private void CleanWordText(WordApp theApp, Document theDoc )
        {
            try
            {
                System.Diagnostics.Stopwatch theStopWatch = new System.Diagnostics.Stopwatch();
                theStopWatch.Start();
                //int Counter;
                theDoc.Activate();
                boxProgress.Items.Add("Starting to clean the document...");
                progressBar1.Value = 0;
                theApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);
                /* Make sure we are in the active pane of the Document
                 * rather than headers, footers, or other spots
                 */

                if (theDoc.ActiveWindow.View.Type == WordRoot.WdViewType.wdPrintView)
                {
                    if (theDoc.ActiveWindow.View.SeekView != WordRoot.WdSeekView.wdSeekMainDocument)
                    {
                        theDoc.ActiveWindow.View.SeekView = WordRoot.WdSeekView.wdSeekMainDocument;
                    }
                }
                /*
                 * Remove all shapes.  We seem to need several passes to remove them all for some reason.
                 * 
                 */
                while (RemoveShapes(theApp, theDoc) > 0) ;
                /*
                 * Remove all frames
                 */
                System.Diagnostics.Stopwatch theStopWatch2 = new System.Diagnostics.Stopwatch();
                theStopWatch2.Start();
                int Counter = 0;
                foreach (WordRoot.Frame theFrame in theDoc.Frames)
                {
                    theFrame.TextWrap = false; // Make it no longer wrap text
                    theFrame.Borders.OutsideLineStyle = WordRoot.WdLineStyle.wdLineStyleNone;
                    theFrame.Delete(); // and delete the frame
                    Counter++;
                    if (Counter % 100 == 0)
                    {
                        boxProgress.Items.Add("Deleted " + Counter.ToString() + " frames");
                        Application.DoEvents();
                    }


                }

                boxProgress.Items.Add("Removed " + Counter.ToString() + " frames in " + (theStopWatch2.ElapsedMilliseconds/1000.0).ToString("f2") + " seconds");
                Application.DoEvents();
               /*
                  * Convert tables to text
                  */
                theStopWatch2.Restart();
                Counter = 0;
                foreach (WordRoot.Table theTable in theDoc.Tables)
                {

                    theTable.Rows.ConvertToText(WordRoot.WdTableFieldSeparator.wdSeparateByTabs, true);
                    Counter++;
                    if (Counter % 100 == 0)
                    {
                        boxProgress.Items.Add("Converted " + Counter.ToString() + " tables");
                        Application.DoEvents();
                    }

                }
                boxProgress.Items.Add("Converted " + Counter.ToString() + " tables in " + (theStopWatch2.ElapsedMilliseconds/1000.0).ToString() 
                    + " seconds");
                Application.DoEvents();
  
                // Go to the beginning
                theApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);
                //  Make one column
                OneColumn(theApp);
                // Clear all tabs, paragraph markers, section breaks, manual line feeds, column breaks and manual page breaks.
                // ^m also deals with section breaks when wildcards are on.
                
                GlobalReplace(theApp.Selection, "[^3^4^9^11^13^14^12^m]", theSpace, false, true);
                // And this character found in some documents:  (F020) or a symbol space.
                GlobalReplace(theApp.Selection, "", theSpace, false, false);
                // Clear all multiple spaces or symbol spaces
                GlobalReplace(theApp.Selection, "[| ]{2}", theSpace, true, true);

                /*
              * Now left align everything
              */
                foreach (WordRoot.Paragraph theParagraph in theDoc.Paragraphs)
                {
                    theParagraph.Format.Alignment = WordRoot.WdParagraphAlignment.wdAlignParagraphLeft;
                }




                boxProgress.Items.Add("Cleaned the text in " + (theStopWatch.ElapsedMilliseconds/1000.0).ToString("f2")  + " seconds");
                Application.DoEvents();
                progressBar1.Value += 1;
          
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
         }
        private void QuitWord(bool Save)
        {
            if (wrdApp != null)
            {
                try
                {
                    if (InputDoc != null)
                    {
                        InputDoc.Close(false);
                        InputDoc = null;
                    }
                    if (OutputDoc != null)
                    {
                        OutputDoc.Close(Save);
                        OutputDoc = null;
                    }
                    // Shut down Word
                    theOptions.RestoreApp(wrdApp); // restore the settings
                    wrdApp.Quit(ref missing, ref missing, ref missing);
                }
                catch
                { 
                }
                try
                {
                    // Shut down Excel
                    if (excelApp.ActiveWorkbook != null)
                    {
                        theExcelOptions.RestoreApp(excelApp);  // Restore the settings
                        if (Save)
                        {
                            excelApp.ActiveWorkbook.Save();
                        }
                        excelApp.ActiveWorkbook.Close(Save);
                    }
                    excelApp.Quit();
                }
                catch
                {
                }
                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch
                { // ignore and continue
                }
                wrdApp = null;
                excelApp = null;
            }

        }
        private void btnClose_Click(object sender, EventArgs e)
        {
            /*
            * Exit
            */
            CloseApp = true;            
            QuitWord(!Paused);  // Save the output if we did'nt come here from a paused state
            this.Close();
        }
        private void GlobalReplace(WordRoot.Selection theSelection, string SearchChars, string ReplacementChars, bool Repeat, bool Wildcards)
        {
            // Do a global replacement
            System.Diagnostics.Stopwatch theStopwatch = new System.Diagnostics.Stopwatch();
            theStopwatch.Start();
            bool Found = true;  // Assume success
            theSelection.Find.Text = SearchChars;
            theSelection.Find.Replacement.Text = ReplacementChars;
            theSelection.Find.Wrap = WordRoot.WdFindWrap.wdFindContinue;
            theSelection.Find.MatchWildcards = Wildcards;
        //
            // If we want to keep searching, we'll do so
            //
            while (Found)
            {
                Found = theSelection.Find.Execute(missing, false, false, missing, false, false, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
                missing, missing, missing, missing);
                Found = Repeat && Found;  // If repeat not set, then we only execute once.
                Application.DoEvents();
            }
            theSelection.Find.MatchWildcards = false;  // the default
            boxProgress.Items.Add("Globally replaced " + SearchChars + " in " + (theStopwatch.ElapsedMilliseconds/1000.0).ToString("f2")  + " seconds");
            Application.DoEvents();


         }
        private void Segment(WordApp  theApp, WordRoot.Selection theSelection, int WordCount, int NumberofWords)
        {
             /*
             * Now segment into the number of words specified by the WordCount paramenter
             */
            try
            {
                boxProgress.Items.Add("Starting segmentation...");
                System.Diagnostics.Stopwatch theStopwatch = new System.Diagnostics.Stopwatch();
                theStopwatch.Start();
                bool Found;
                /*
                        * Use wildcards to add the paragraph markers
                        * 
                        */
                theSelection.Find.ClearFormatting();
                theSelection.Find.Replacement.Text = "^&^p";  // Replace with what we just found and a paragraph marker
                theSelection.Find.MatchWildcards = true;
                theSelection.Find.Forward = true;
                theSelection.Find.Wrap = WordRoot.WdFindWrap.wdFindContinue;
                theSelection.Find.Format = false;
                theSelection.Find.MatchCase = false;
                theSelection.Find.MatchWholeWord = false;
                theSelection.Find.MatchKashida = false;
                theSelection.Find.MatchDiacritics = false;
                theSelection.Find.MatchAlefHamza = false;
                theSelection.Find.MatchControl = false;
                theSelection.Find.MatchAllWordForms = false;
                theSelection.Find.MatchSoundsLike = false;
                const string WildCards = "([! ]@[ |])"; // Word ending in a space
                theSelection.Find.Text = "";  // Clear the find string
                /*
                * Build up the search string
                 * 
                 * If the words per line we want at the end are more than three, we need to do the replacement
                 * in two stages as otherwise the wildcard expression gets too complicated.
                */
                int MaxWordPerLine = 2;
                theSelection.Find.Text = WildCards + WildCards;  // We can only handle two or three words at a time


                // Now do the first replacement
                boxProgress.Items.Add("Starting segmentation first pass");
                Application.DoEvents();
                theApp.ActiveDocument.UndoClear();  // Clear the undo stack

                Found = theSelection.Find.Execute(missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
                missing, missing, missing, missing);
                boxProgress.Items.Add("First pass complete in " + (theStopwatch.ElapsedMilliseconds / 1000.0).ToString("f2") + " seconds");
                Application.DoEvents();
                progressBar1.Value += 1;
                Application.DoEvents();


                /*
                 * If the WordCount > 2, we assume 4, 6, 8 etc.
                 */
                if (WordCount > MaxWordPerLine)
                {
                    const string Paragraphs = "(*)^13";  // Match anything ending with a paragraph
                    theSelection.Find.Text = "";
                    theSelection.Find.Replacement.Text = "";
                    /*
                     * Add trailing paragraphs to make sure we have Wordperline/2 paragraphs at the end.
                     */
                    // Go to the end
                    theSelection.EndKey(WordRoot.WdUnits.wdStory);

                    for (int i = 1; i <= WordCount / MaxWordPerLine; i++)
                    {
                        theSelection.Find.Text += Paragraphs; // build up the search string
                        theSelection.Find.Replacement.Text += "\\" + i.ToString();
                        /*
                        * Add trailing paragraphs to make sure we have Wordperline/2 paragraphs at the end.
                        */

                        theSelection.TypeParagraph();

                    }
                    // Go to the beginning
                    theSelection.HomeKey(WordRoot.WdUnits.wdStory);

                    theSelection.Find.Replacement.Text += "^p"; // ending with one paragraph
                    // and do the second paragraph
                    boxProgress.Items.Add("Starting segmentation second pass");
                    System.Diagnostics.Stopwatch theStopwatch2 = new System.Diagnostics.Stopwatch();
                    theStopwatch2.Start();
                    theApp.ActiveDocument.UndoClear();  // Clear the undo stack
                    Found = theSelection.Find.Execute(missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
                        missing, missing, missing, missing);
                    boxProgress.Items.Add("Second pass complete in " + ((float)theStopwatch2.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds");
                    Application.DoEvents();
                    theStopwatch2.Stop();
                    theStopwatch2 = null;
                    progressBar1.Value += 1;
                    Application.DoEvents();


                }

                theSelection.Find.MatchWildcards = false;  // Don't leave wildcards hanging
                /*
                  * Now remove the trailing spaces
                  */

                GlobalReplace(theSelection, " ^p", "^p", false, false);
                /*
                 * And make sure we don't have two consequitive paragraphs
                 */
                GlobalReplace(theSelection, "^p^p", "^p", false, false);
                /*
                 * Delete the double paragraphs at the end
                 */
                GlobalReplace(theSelection, "^p^p", "", true, false);



                theApp.ScreenUpdating = true;  // turn on updating
                progressBar1.Value = progressBar1.Maximum;  // We've finished!
                long ElapsedTicks = theStopwatch.ElapsedTicks;
                boxProgress.Items.Add("Segmentation complete in " + ((float)ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds");
                int LineCounter = NumberofWords / WordCount;
                boxProgress.Items.Add(((float)LineCounter * Stopwatch.Frequency/ ElapsedTicks).ToString("f2") + " lines per second");
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
            return;
         }
        private void OneColumn(WordRoot._Application theApp)
        {
             /*
              * Make the docoument one column
              */
             // If we have a split window, close one of them
             if (theApp.ActiveWindow.View.SplitSpecial != WordRoot.WdSpecialPane.wdPaneNone)
             {
                 theApp.ActiveWindow.Panes[2].Close(); // Close the other window
             }
             // If not print view, make it print view
             if (theApp.ActiveWindow.ActivePane.View.Type != WordRoot.WdViewType.wdPrintView)
             {
                 theApp.ActiveWindow.ActivePane.View.Type = WordRoot.WdViewType.wdPrintView;
             }
             // Now make it one column
             theApp.Selection.PageSetup.TextColumns.SetCount(1); // one column
             theApp.Selection.PageSetup.TextColumns.EvenlySpaced = -1;  // Evenly spaced
             theApp.Selection.PageSetup.TextColumns.LineBetween = 0;  // no lines between
             theApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // Go to beginnng
         }
        private void WordsPerLine_ValueChanged(object sender, EventArgs e)
        {
            NumericUpDown WordsPerLine = (NumericUpDown)sender;
            if (WordsPerLine.Value > 8)
            {
                /* We now do it in multiples of four
                */
                WordsPerLine.Increment = 4;
                if (WordsPerLine.Value % 4 != 0)
                {
                    WordsPerLine.Value = Math.Round((WordsPerLine.Value + 4) / 4) * 4;  // go the next multiple of four
                }
            }
            else
            {
                WordsPerLine.Increment = 1;
            }

        }

        private int InitialiseExcel(ExcelApp excelApp, bool EvenRows, string FileName)
        {
            /*
             * We initialise the file if necessary and clear the relevant sheet.
             */

            bool hasValue = false;
            boxProgress.Items.Add("Clearing worksheet...");
            try
            {
                int theRow;
                string StrippedFileName = Path.GetFileName(FileName);  // Get the file name without the directory
                ExcelRoot.Workbook theWorkbook;
                if (File.Exists(txtExcelOutput.Text))
                {
                    try
                    {
                        theWorkbook = excelApp.Workbooks.Open(txtExcelOutput.Text);  // Open the file
                    }
                    catch (Exception e)
                    {
                        DialogResult theResult = MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tabControl1.SelectTab("Setup");
                        return 0;

                    }

                }
                else
                {
                    theWorkbook = excelApp.Workbooks.Add();  // add it
                    theWorkbook.SaveAs(txtExcelOutput.Text);  // save it
                }
                theWorkbook.Sheets[1].Name = "Interlinear";
                theWorkbook.Sheets[1].Columns("A").ColumnWidth = 100;  // and make the first column wide
                if (EvenRows)
                {
                    theRow = 2;
                }
                else
                {
                    theRow = 1;

                }
               ExcelRoot.Worksheet theWorkSheet = theWorkbook.Sheets[1];
                /*
                 * Clear all non-empty rows apart from the first two
                 */
               int RowCounter = theRow + 2;
               do
                {
                    ExcelRoot.Range theCell;
                    string stringCell = "A" + RowCounter.ToString();
                    theCell = theWorkSheet.Range[stringCell];
                    ExcelRoot.XlRgbColor theCellColour;
                    theCellColour = (ExcelRoot.XlRgbColor)theCell.Interior.Color;
                    hasValue = theCell.Value != null || theCellColour != ExcelRoot.XlRgbColor.rgbWhite;
                    if (hasValue)
                    {
                        theCell.Clear();  // Clear it
                        RowCounter += 2;
                    }
                    

                } while (hasValue);
                return theRow;
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
                return -1;
            }

        }
        private void FillExcel(ExcelApp excelApp, WordApp wrdApp, Document theDoc, int RowCounter)
        {
            /*
             * Here is where we fill the Excel spreadsheet
             */
            try
            {
                Stopwatch theStopwatch = new Stopwatch();
                theStopwatch.Start();
                //excelApp.Visible = true;
                theExcelOptions = new ExcelAppOptions(excelApp);  // save settings
                theExcelOptions.OptimiseApp(excelApp);  // Optimise Excel before filling
                boxProgress.Items.Add("Starting to fill Excel worksheet");
                ExcelRoot.Workbook theWorkBook = excelApp.ActiveWorkbook;  // remember the document.
                excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationManual; // Don't calculate automatically.
                Application.DoEvents();
                // Get document and worksheet
                theDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure it isn't in reading layout.
                ExcelRoot.Worksheet theWorkSheet = theWorkBook.Sheets[1];
                /*
                 * Generate the header messages
                 */
                string HeaderText = theMessage[RowCounter - 1] + Path.GetFileName(theDoc.FullName);
            
            
                int ParagraphCount = theDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticParagraphs);
                MaxParagraphs = Math.Max(MaxParagraphs, ParagraphCount * 2);
                boxProgress.Items.Add("There are " + ParagraphCount.ToString() + " paragraphs");
                btnPauseResume.Enabled = true;
                Application.DoEvents();
                // Initialise the progress bar
                progressBar1.Value = 0;
                progressBar1.Maximum = ParagraphCount;
                Stopwatch theStopwatch2 = new Stopwatch();
                theStopwatch2.Start();
                boxProgress.Items.Add("Copying document to Excel...");
                theWorkSheet = theWorkBook.Sheets[1];
                // The header text
                theWorkSheet.Range["A" + RowCounter.ToString()].Value = HeaderText;
                theWorkSheet.Range["A" + RowCounter.ToString()].Interior.Color = CellColour[RowCounter - 1];
                int theRow = RowCounter + 2;
                int Counter = 0;
                Stopwatch CopyStopwatch = new Stopwatch();
                CopyStopwatch.Start();
                foreach (WordRoot.Paragraph theParagraph in theDoc.Paragraphs)
                {
                    string theCellRef = "A" + theRow.ToString();
                    /*
                     * Sometimes the paste fails, so we try again if that is the case
                     */
                    bool Failure = true;  // Assume failure
                    int ErrorCounter = 0;
                    theParagraph.Range.Copy();  // copy the range
                    PauseForThought(CopyStopwatch, theStopwatch, theStopwatch2);
                    while (Failure && ErrorCounter < 5)
                    {
                        try
                        {
                            theWorkSheet.Paste(theWorkSheet.Range[theCellRef]);  // Paste to it.
                            Clipboard.Clear();  // clear the clipboard
                            Failure = false;
                        
                        }
                        catch (Exception e)
                        {
                            boxProgress.Items.Add("Paste error " + e.Message + " in row " + theRow.ToString() + ". Retrying...");
                            Thread.Sleep(5);  // wait 10 milliseconds
                            ErrorCounter++;
                            if (ErrorCounter >= 5)
                            {
                                boxProgress.Items.Add("*****  Failed to paste " + theRow.ToString() + " " + theParagraph.Range.ToString());
                            }
                            Application.DoEvents();
                        }
                    }
                    theWorkSheet.Range[theCellRef].Font.Size = 11;  // But make it just 11 
                    theWorkSheet.Range[theCellRef].Interior.Color = CellColour[RowCounter - 1];
                    theRow += 2;
                    if (Counter % 10 == 0)
                    {
                        progressBar1.Value = Math.Min(Counter, progressBar1.Maximum); // to make sure we don't try to go beyond the maximum
                        Application.DoEvents();
                    }
                    if (Counter % 50 == 0 && Counter > 0)
                    {
                        toolStripStatusLabel1.Text = "Copied " + Counter.ToString() + " paragraphs in " + ((float)CopyStopwatch.ElapsedTicks/Stopwatch.Frequency).ToString("f2") + " seconds or " + 
                            ((float)Counter*Stopwatch.Frequency/CopyStopwatch.ElapsedTicks).ToString("f2") + " paragraphs/second";
                        Application.DoEvents();
                    }
                    Counter++;

                }
                progressBar1.Value = ParagraphCount;
                boxProgress.Items.Add("Copied " + Counter.ToString() + " paragraphs in " + ((float)CopyStopwatch.ElapsedTicks / Stopwatch.Frequency).ToString("f2") + " seconds or " +
                            ((float)Counter * Stopwatch.Frequency / CopyStopwatch.ElapsedTicks).ToString("f2") + " paragraphs/second");
                CopyStopwatch.Stop();
                CopyStopwatch = null;
                theDoc.Close(false);
                theWorkSheet.Range["A1"].Select();  // go to the start of the worksheet
                theExcelOptions.RestoreApp(excelApp); // Restore the Excel settings we saved earlier
                theWorkBook.Save();
                boxProgress.Items.Add("Excel interlinear worksheet filled in " + (theStopwatch2.Elapsed).ToString("hh\\:mm\\:ss\\.f"));
                theStopwatch2.Stop();
                theStopwatch2 = null;
                btnPauseResume.Enabled = false;
                Application.DoEvents();
                }
            catch (Exception Ex)
                {
                    FinalCatch(Ex);
                }

        }
        private void SendToExcel_Click(object sender, EventArgs e)
        {
            try
            {
            Button theButton = (Button)sender;
            bool EvenRows;
            Document theDoc;
            string FileName;
            btnClose.Enabled = false;
            tabControl1.SelectTab("Progress");
            try
            {
                if (theButton.Parent.Text == "Legacy")
                {
                    EvenRows = false;
                    theDoc = wrdApp.Documents.OpenNoRepairDialog(txtLegacyOutput.Text, true);
                    FileName = txtLegacyOutput.Text;
                }
                else
                {
                    EvenRows = true;
                    theDoc = wrdApp.Documents.OpenNoRepairDialog(txtUnicodeOutput.Text, true);
                    FileName = txtUnicodeInput.Text;
                 }

            }
            catch (Exception ex)
            {
                DialogResult theResult = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tabControl1.SelectTab("Setup");
                return;
            }
            
            // We'll send the information to Excel
            int RowCounter = InitialiseExcel(excelApp, EvenRows, FileName);
            if (RowCounter > 0)
            {
                FillExcel(excelApp, wrdApp, theDoc, RowCounter);
                excelApp.ActiveWorkbook.Close();  // Close the workbook

                boxProgress.Items.Add("Finished sending to Excel.");
            }
            else
            {
                boxProgress.Items.Add("Could not send to Excel");
            }

            theDoc = null;
            btnInterlinear.Enabled = true;
            btnClose.Enabled = true;
            System.Media.SystemSounds.Beep.Play();  // and beep

            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }


        }
        private void BothToExcel_Click(object sender, EventArgs e)
        {
            // We'll send the information to Excel
            try
            {
                Stopwatch theStopwatch = new Stopwatch();
                theStopwatch.Start();
                Document theDoc;
                tabControl1.SelectTab("Progress");
                btnClose.Enabled = false;
                try
                {
                    theDoc = wrdApp.Documents.Open(txtLegacyOutput.Text);
                }
                catch (Exception ex)
                {
                    DialogResult theResult = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tabControl1.SelectTab("Setup");
                    return;
                }
                int RowCounter = InitialiseExcel(excelApp, false, txtLegacyOutput.Text);
                FillExcel(excelApp, wrdApp, theDoc, RowCounter);
                try
                {
                    theDoc = wrdApp.Documents.Open(txtUnicodeOutput.Text);
                }
                catch (Exception ex)
                {
                    DialogResult theResult = MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tabControl1.SelectTab("Setup");
                    return;
                }

                RowCounter = InitialiseExcel(excelApp, true, txtUnicodeOutput.Text);
                FillExcel(excelApp, wrdApp, theDoc, RowCounter);
                //MakeInterlinear(excelApp); // Make the interlinear worksheet, too.
                excelApp.ActiveWorkbook.Close(); // Close the workbook
                boxProgress.Items.Add("Finished sending both files to Excel in " + theStopwatch.Elapsed.ToString("HH.mm.ss.f"));
                theStopwatch.Stop();
                theStopwatch = null;
                theDoc = null;
                btnClose.Enabled = true;
                System.Media.SystemSounds.Beep.Play();  // and beep
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }

        }

        private void btnInterlinear_Click(object sender, EventArgs e)
        {
            boxProgress.Items.Clear();  // empty the progress box

            //MakeInterlinear(excelApp);
        }

             

        private int RemoveShapes(WordApp theApp, Document theDoc)
        {
            try
            {
                Stopwatch theStopwatch2 = new Stopwatch();
                theStopwatch2.Start();
                int TotalCounter = 0;
                int Counter = 0;
                int TableCounter = 0;
                // Remove shapes
                foreach (WordRoot.Shape theShape in theDoc.Shapes)
                {

                    if (theShape.Type == Office.MsoShapeType.msoTextBox || theShape.Type == Office.MsoShapeType.msoGroup)
                    {
                        if (theShape.Type == Office.MsoShapeType.msoTextBox)
                        {
                            theShape.ConvertToInlineShape();
                            theShape.Delete();
                        }
                        else
                        {
                            theShape.Select();
                            theApp.Selection.Delete();
                        }
                        Counter++;
                        if (Counter % 100 == 0)
                        {
                            boxProgress.Items.Add("Deleted " + Counter.ToString() + " shapes");
                            Application.DoEvents();
                        }
                    }
                    else
                    {
                        if (theShape.Type == Office.MsoShapeType.msoTable)
                        // Convert the table to text
                        {
                            WordRoot.Table theTable = (WordRoot.Table)theShape;
                            theTable.Rows.ConvertToText(WordRoot.WdTableFieldSeparator.wdSeparateByTabs, true);
                            TableCounter++;
                            if (Counter % 100 == 0)
                            {
                                boxProgress.Items.Add("Converted " + TableCounter.ToString() + " tables");
                                Application.DoEvents();
                            }

                        }
                    }
 
                }
 
                boxProgress.Items.Add("Removed or converted  " + (TableCounter + Counter).ToString() + " shapes in " + (theStopwatch2.ElapsedMilliseconds/1000.0).ToString("f2") + " seconds");
                Application.DoEvents();
                TotalCounter += Counter + TableCounter;

                /*
                * Remove all inline shapes
                */
                theStopwatch2.Restart();
                Counter = 0;
                foreach (WordRoot.InlineShape theShape in theDoc.InlineShapes)
                {
                    theShape.Delete(); // and delete the frame
                    Counter++;
                    if (Counter % 100 == 0)
                    {
                        boxProgress.Items.Add("Deleted " + Counter.ToString() + " inline shapes");
                        Application.DoEvents();
                    }


                }
                boxProgress.Items.Add("Removed " + Counter.ToString() + " inline shapes in " + (theStopwatch2.ElapsedMilliseconds / 1000.0).ToString("f2") + " seconds");
                Application.DoEvents();
                TotalCounter += Counter;
                return TotalCounter;
            }
            catch (Exception Ex)
            {
                 FinalCatch(Ex);
                 return 0;
            }

        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "UserGuide.docx");
            WordApp HelpApp = new Word();
            try
            {
                HelpApp.Visible = true;
                HelpApp.Documents.Open(HelpPath, missing, true);
            }
            catch (Exception theException)
            {
                MessageBox.Show("Failed to open help file at " + HelpPath + "\r" + theException.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                HelpApp.Quit();
            }

         }

        private void documentationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "Interlinear.docx");
            System.Diagnostics.Process.Start(HelpPath);

        }

        private void licenseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string HelpPath = Path.Combine(Application.StartupPath, "gpl.txt");
            System.Diagnostics.Process.Start("Wordpad.exe", '"' + HelpPath + '"');
 
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 About = new AboutBox1();
            About.Show();

        }

        private void btnPauseResume_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            Paused = !Paused; // Toggle the pause flag
            btnClose.Enabled = Paused;
            if (Paused)
            {
                toolStripStatusLabel1.Text = "Pausing...";
                theButton.Text = "Resume";
            }
            else
            {
                toolStripStatusLabel1.Text = "Resuming...";
                theButton.Text = "Pause";
            }

            Application.DoEvents();
        }

        private void PauseForThought(Stopwatch theStopwatch, Stopwatch theStopwatch2, Stopwatch theStopwatch3)
        {
                while (Paused)
                {
                    if (theStopwatch.IsRunning)
                    {
                        theStopwatch.Stop();
                    }
                    if (theStopwatch2.IsRunning)
                    {
                        theStopwatch2.Stop();
                    }
                    if (theStopwatch3.IsRunning)
                    {
                        theStopwatch3.Stop();
                    }
                    System.Threading.Thread.Sleep(1000); // wait a second
                    toolStripStatusLabel1.Text = "Paused, click Resume to restart or Close to close";
                    if (CloseApp)
                    {
                        /*
                            * We have clicked Close, so we exit from here and stop doing the processing
                            */
                      
                        QuitWord(false);  // don't save the output when we quit.
                        this.Close();  // die
                        return;  // exit
                    }
                    Application.DoEvents();

                }
                if (!theStopwatch.IsRunning)
                {
                    theStopwatch.Start();
                }
                if (!theStopwatch2.IsRunning)
                {
                    theStopwatch2.Start();
                }
                if (!theStopwatch3.IsRunning)
                {
                    theStopwatch3.Start();
                }

         }


        
                                              
    }

    class WordAppOptions
    {
        /*
         * Allows us to save and set application options in Word and then restore the original settings
         */


        private bool AutoFormatAsYouTypeApplyBorders;
        private bool AutoFormatAsYouTypeApplyBulletedLists;
        private bool AutoFormatAsYouTypeApplyHeadings;
        private bool AutoFormatAsYouTypeApplyNumberedLists;
        private bool AutoFormatAsYouTypeApplyTables;
        private bool AutoFormatAsYouTypeAutoLetterWizard;
        private bool AutoFormatAsYouTypeDefineStyles;
        private bool AutoFormatAsYouTypeFormatListItemBeginning;
        private bool AutoFormatAsYouTypeReplaceFractions;
        private bool AutoFormatAsYouTypeReplaceHyperlinks;
        private bool AutoFormatAsYouTypeReplaceOrdinals;
        private bool AutoFormatAsYouTypeReplacePlainTextEmphasis;
        private bool AutoFormatAsYouTypeReplaceQuotes;
        private bool AutoFormatAsYouTypeReplaceSymbols;
        private bool CheckGrammarAsYouType;
        private bool CheckSpellingAsYouType;
        private bool CorrectCapsLock;
        private bool CorrectDays;
        private bool CorrectInitialCaps;
        private bool CorrectKeyboardSetting;
        private bool CorrectSentenceCaps;
        private bool CorrectTableCells;
        private bool DisplayAutoCorrectOptions;
        private bool LabelSmartTags;
        private bool Pagination;
        private bool RepeatWord;
        private bool ReplaceText;
        private bool ReplaceTextFromSpellingChecker;
        private bool TabIndentKey;


        public WordAppOptions(WordApp theApp)
        {
            // Save the current Word settings
            AutoFormatAsYouTypeApplyBorders = theApp.Options.AutoFormatAsYouTypeApplyBorders;
            AutoFormatAsYouTypeApplyBulletedLists = theApp.Options.AutoFormatAsYouTypeApplyBulletedLists;
            AutoFormatAsYouTypeApplyHeadings = theApp.Options.AutoFormatAsYouTypeApplyHeadings;
            AutoFormatAsYouTypeApplyNumberedLists = theApp.Options.AutoFormatAsYouTypeApplyNumberedLists;
            AutoFormatAsYouTypeApplyTables = theApp.Options.AutoFormatAsYouTypeApplyTables;
            AutoFormatAsYouTypeAutoLetterWizard = theApp.Options.AutoFormatAsYouTypeAutoLetterWizard;
            AutoFormatAsYouTypeDefineStyles = theApp.Options.AutoFormatAsYouTypeDefineStyles;
            AutoFormatAsYouTypeFormatListItemBeginning = theApp.Options.AutoFormatAsYouTypeFormatListItemBeginning;
            AutoFormatAsYouTypeReplaceFractions = theApp.Options.AutoFormatAsYouTypeReplaceFractions;
            AutoFormatAsYouTypeReplaceHyperlinks = theApp.Options.AutoFormatAsYouTypeReplaceHyperlinks;
            AutoFormatAsYouTypeReplaceOrdinals = theApp.Options.AutoFormatAsYouTypeReplaceOrdinals;
            AutoFormatAsYouTypeReplacePlainTextEmphasis = theApp.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis;
            AutoFormatAsYouTypeReplaceQuotes = theApp.Options.AutoFormatAsYouTypeReplaceQuotes;
            AutoFormatAsYouTypeReplaceSymbols = theApp.Options.AutoFormatAsYouTypeReplaceSymbols;
            CheckGrammarAsYouType = theApp.Options.CheckGrammarAsYouType;
            CheckSpellingAsYouType = theApp.Options.CheckSpellingAsYouType;
            CorrectKeyboardSetting = theApp.AutoCorrect.CorrectKeyboardSetting;
            CorrectCapsLock = theApp.AutoCorrect.CorrectCapsLock;
            CorrectDays = theApp.AutoCorrect.CorrectDays;
            CorrectInitialCaps = theApp.AutoCorrect.CorrectInitialCaps;
            CorrectSentenceCaps = theApp.AutoCorrect.CorrectSentenceCaps;
            CorrectTableCells = theApp.AutoCorrect.CorrectTableCells;
            DisplayAutoCorrectOptions = theApp.AutoCorrect.DisplayAutoCorrectOptions;
            LabelSmartTags = theApp.Options.LabelSmartTags;
            Pagination = theApp.Options.Pagination;
            RepeatWord = theApp.Options.RepeatWord;
            ReplaceText = theApp.AutoCorrect.ReplaceText;
            ReplaceTextFromSpellingChecker = theApp.AutoCorrect.ReplaceTextFromSpellingChecker;
            TabIndentKey = theApp.Options.TabIndentKey;
        }
        public void OptimiseApp(WordApp theApp)
        {
            theApp.Options.AutoFormatAsYouTypeApplyBorders = false;
            theApp.Options.AutoFormatAsYouTypeApplyBulletedLists = false;
            theApp.Options.AutoFormatAsYouTypeApplyHeadings = false;
            theApp.Options.AutoFormatAsYouTypeApplyNumberedLists = false;
            theApp.Options.AutoFormatAsYouTypeApplyTables = false;
            theApp.Options.AutoFormatAsYouTypeAutoLetterWizard = false;
            theApp.Options.AutoFormatAsYouTypeDefineStyles = false;
            theApp.Options.AutoFormatAsYouTypeFormatListItemBeginning = false;
            theApp.Options.AutoFormatAsYouTypeReplaceFractions = false;
            theApp.Options.AutoFormatAsYouTypeReplaceHyperlinks = false;
            theApp.Options.AutoFormatAsYouTypeReplaceOrdinals = false;
            theApp.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = false;
            theApp.Options.AutoFormatAsYouTypeReplaceQuotes = false;
            theApp.Options.AutoFormatAsYouTypeReplaceSymbols = false;
            theApp.Options.CheckGrammarAsYouType = false;
            theApp.Options.CheckSpellingAsYouType = false;
            theApp.AutoCorrect.CorrectCapsLock = false;
            theApp.AutoCorrect.CorrectDays = false;
            theApp.AutoCorrect.CorrectInitialCaps = false;
            theApp.AutoCorrect.CorrectKeyboardSetting = false;
            theApp.AutoCorrect.CorrectSentenceCaps = false;
            theApp.AutoCorrect.CorrectTableCells = false;
            theApp.AutoCorrect.DisplayAutoCorrectOptions = false;
            theApp.AutoCorrect.CorrectKeyboardSetting = false;
            theApp.Options.LabelSmartTags = false;
            theApp.Options.Pagination = false;  // turn off background pagination
            theApp.Options.RepeatWord = false;
            theApp.AutoCorrect.ReplaceText = false;
            theApp.AutoCorrect.ReplaceTextFromSpellingChecker = false;
            theApp.Options.TabIndentKey = false;


        }
        public void RestoreApp(WordApp theApp)
        {
            theApp.Options.AutoFormatAsYouTypeApplyBorders = AutoFormatAsYouTypeApplyBorders;
            theApp.Options.AutoFormatAsYouTypeApplyBulletedLists = AutoFormatAsYouTypeApplyBulletedLists;
            theApp.Options.AutoFormatAsYouTypeApplyHeadings = AutoFormatAsYouTypeApplyHeadings;
            theApp.Options.AutoFormatAsYouTypeApplyNumberedLists = AutoFormatAsYouTypeApplyNumberedLists;
            theApp.Options.AutoFormatAsYouTypeApplyTables = AutoFormatAsYouTypeApplyTables;
            theApp.Options.AutoFormatAsYouTypeAutoLetterWizard = AutoFormatAsYouTypeAutoLetterWizard;
            theApp.Options.AutoFormatAsYouTypeDefineStyles = AutoFormatAsYouTypeDefineStyles;
            theApp.Options.AutoFormatAsYouTypeFormatListItemBeginning = AutoFormatAsYouTypeFormatListItemBeginning;
            theApp.Options.AutoFormatAsYouTypeReplaceFractions = AutoFormatAsYouTypeReplaceFractions;
            theApp.Options.AutoFormatAsYouTypeReplaceHyperlinks = AutoFormatAsYouTypeReplaceHyperlinks;
            theApp.Options.AutoFormatAsYouTypeReplaceOrdinals = AutoFormatAsYouTypeReplaceOrdinals;
            theApp.Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = AutoFormatAsYouTypeReplacePlainTextEmphasis;
            theApp.Options.AutoFormatAsYouTypeReplaceQuotes = AutoFormatAsYouTypeReplaceQuotes;
            theApp.Options.AutoFormatAsYouTypeReplaceSymbols = AutoFormatAsYouTypeReplaceSymbols;
            theApp.Options.CheckGrammarAsYouType = CheckGrammarAsYouType;
            theApp.Options.CheckSpellingAsYouType = CheckSpellingAsYouType;
            theApp.AutoCorrect.CorrectCapsLock = CorrectCapsLock;
            theApp.AutoCorrect.CorrectDays = CorrectDays;
            theApp.AutoCorrect.CorrectInitialCaps = CorrectInitialCaps;
            theApp.AutoCorrect.CorrectKeyboardSetting = CorrectKeyboardSetting;
            theApp.AutoCorrect.CorrectSentenceCaps = CorrectSentenceCaps;
            theApp.AutoCorrect.CorrectTableCells = CorrectTableCells;
            theApp.AutoCorrect.DisplayAutoCorrectOptions = DisplayAutoCorrectOptions;
            theApp.Options.LabelSmartTags = LabelSmartTags;
            theApp.Options.Pagination = Pagination;
            theApp.Options.RepeatWord = RepeatWord;
            theApp.AutoCorrect.ReplaceText = ReplaceText;
            theApp.AutoCorrect.ReplaceTextFromSpellingChecker = ReplaceTextFromSpellingChecker;
            theApp.Options.TabIndentKey = TabIndentKey;

        }
    };
    class ExcelAppOptions
    {
        /*
         * Allows Excel options to be saved, optimised for copying and restored later.
         */
        private bool TwoInitialCapitals = false;
        private bool CorrectCapsLock = false;
        private bool CorrectSentenceCap = false;
        private bool CapitalizeNamesOfDays = false;
        private bool ReplaceText = false;
        private bool AutoExpandListRange = false;
        private bool AutoFillFormulasInLists = false;
        private bool AutoFormatAsYouTypeReplaceHyperlinks;
        private ExcelRoot.XlCalculation Calculation;

        public ExcelAppOptions(ExcelApp theApp)
        {
            TwoInitialCapitals = theApp.AutoCorrect.TwoInitialCapitals;
            CorrectCapsLock = theApp.AutoCorrect.CorrectCapsLock;
            CorrectSentenceCap = theApp.AutoCorrect.CorrectSentenceCap;
            CapitalizeNamesOfDays = theApp.AutoCorrect.CapitalizeNamesOfDays;
            ReplaceText = theApp.AutoCorrect.ReplaceText;
            AutoExpandListRange = theApp.AutoCorrect.AutoExpandListRange;
            AutoFormatAsYouTypeReplaceHyperlinks = theApp.AutoFormatAsYouTypeReplaceHyperlinks;
            Calculation = theApp.Calculation;
        }
        public void OptimiseApp(ExcelApp theApp)
        {
            theApp.AutoCorrect.TwoInitialCapitals = false;
            theApp.AutoCorrect.CorrectCapsLock = false;
            theApp.AutoCorrect.CorrectSentenceCap = false;
            theApp.AutoCorrect.CapitalizeNamesOfDays = false;
            theApp.AutoCorrect.ReplaceText = false;
            theApp.AutoCorrect.AutoExpandListRange = false;
            theApp.AutoCorrect.AutoFillFormulasInLists = false;
            theApp.Calculation = ExcelRoot.XlCalculation.xlCalculationManual;
        }
        public void RestoreApp(ExcelApp theApp)
        {
            theApp.AutoCorrect.TwoInitialCapitals = TwoInitialCapitals;
            theApp.AutoCorrect.CorrectCapsLock = CorrectCapsLock;
            theApp.AutoCorrect.CorrectSentenceCap = CorrectSentenceCap;
            theApp.AutoCorrect.CapitalizeNamesOfDays = CapitalizeNamesOfDays;
            theApp.AutoCorrect.ReplaceText = ReplaceText;
            theApp.AutoCorrect.AutoExpandListRange = AutoExpandListRange;
            theApp.AutoCorrect.AutoFillFormulasInLists = AutoFillFormulasInLists;
            theApp.Calculation = (ExcelRoot.XlCalculation)Calculation;
            theApp.AutoFormatAsYouTypeReplaceHyperlinks = AutoFormatAsYouTypeReplaceHyperlinks;
        }
    }


};