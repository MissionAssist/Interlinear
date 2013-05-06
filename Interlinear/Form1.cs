/*
 * Interlinear - a program to take two Word documents, segment them into paragraphs of up to 20 words in length and write them to Excel
 * with the first (legacy) file in odd rows and the second (Unicode) file in even rows. This enables visual checking without the need to try
 * to do side-by-side comparisons.  It depends on both Word and Excel being installed on the computer.
 * 
 * It was writting as part of a MissionAssist project to convert documents in legacy fonts to Unicode.  Much of the logic is attributable to
 * Dennis Pepler, but the code here was written by Stephen Palmstrom.
 * 
 * Copyright © MissionAssist 2013 and distributed under the terms of the GNU General Public License (http://www.gnu.org/licenses/gpl.html)
 * 
 * Last modified on 22 April 2013 by Stephen Palmstrom (stephen.palmstrom@btinternet.com) who asserts the right to be regarded as the author of this program
 * 
 * Acknowledgement is due to Dennis Pepler who worked out how to scan stories etc.
*/
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Data.Common;
using System.Threading;
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
        private Document InputDoc;
        private Document OutputDoc;
        private ExcelApp   excelApp;
        object missing = Type.Missing;
        private const string theSpace = " ";
        private string[] theMessage = new string[2] {"Legacy text is in odd rows from file ", "Unicode text is in even rows from file "};
        private ExcelRoot.XlRgbColor[] CellColour = new ExcelRoot.XlRgbColor[2] {ExcelRoot.XlRgbColor.rgbLightGreen, ExcelRoot.XlRgbColor.rgbOrange};
        private int MaxParagraphs = 0;
        public Form1()
        {
            InitializeComponent();
            wrdApp = new Word();
            wrdApp.Visible = false;
            excelApp = new Excel();  // open Excel
            //excelApp.Visible = false; 
            saveLegacyFileDialog.SupportMultiDottedExtensions = true;
            saveUnicodeFileDialog.SupportMultiDottedExtensions = true;
            Wordcount.SetToolTip(WordsPerLine, "If you want more than eight words per line, they must be in multiples of four");
            ToolTip HelpTip = new ToolTip();
            HelpTip.SetToolTip(btnHelp, "Display the User Guide");
            if (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed)
            {
                lblVersion.Text = "Version " + System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }


        }

        private void btnGetInputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                HandleInputFile(txtLegacyInput, txtLegacyOutput, btnSegmentLegacy, openLegacyFileDialog, saveLegacyFileDialog,
                    btnLegacyToExcel, chkLegacyToExcel);
            }
            else
            {
                HandleInputFile(txtUnicodeInput, txtUnicodeOutput, btnSegmentUnicode, openUnicodeFileDialog,saveUnicodeFileDialog,
                    btnUnicodeToExcel, chkLegacyToExcel);
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
        private void HandleInputFile(TextBox InputText, TextBox OutputText, Button SegmentButton, OpenFileDialog theOpenFileDialog, SaveFileDialog theSaveFileDialog,
            Button ExcelButton, CheckBox SendtoExcel)
        {
            /*
             * Handle the input file dialog.
             */
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

                OutputText.Text = Path.Combine(Path.GetDirectoryName(InputText.Text), Path.GetFileNameWithoutExtension(InputText.Text) +
                    " (Segmented)" + Path.GetExtension(InputText.Text));
                theSaveFileDialog.FileName = OutputText.Text;
                HandleOutputFile1(InputText.Text, OutputText.Text, SegmentButton, ExcelButton, SendtoExcel);  // process further
                btnSegmentBoth.Enabled = btnSegmentLegacy.Enabled && btnSegmentUnicode.Enabled;
                btnBothToExcel.Enabled = File.Exists(txtLegacyOutput.Text) && File.Exists(txtUnicodeOutput.Text) && txtExcelOutput.Text.Length > 0;


        };
        }
        private void btnGetOutputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                HandleOutputFile(txtLegacyInput, txtLegacyOutput, saveLegacyFileDialog, btnSegmentLegacy, btnLegacyToExcel, chkLegacyToExcel);
            }
            else
            {
                HandleOutputFile(txtUnicodeInput, txtUnicodeOutput, saveUnicodeFileDialog, btnSegmentUnicode, btnUnicodeToExcel, chkUnicodeToExcel);
            }
        }
        private void HandleOutputFile (TextBox theInputBox, TextBox theOutputBox, SaveFileDialog theDialog, Button SegmentButton,
            Button ExcelButton, CheckBox SendtoExcel)
        {
            if (theDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                theOutputBox.Text = theDialog.FileName;
                HandleOutputFile1(theInputBox.Text, theOutputBox.Text, SegmentButton, ExcelButton, SendtoExcel);  // process further
            }
 
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
                SegmentFile(txtLegacyInput.Text, txtLegacyOutput.Text, txtLegacyWordCount, chkLegacyToExcel, false);
            }
            else
            {
                SegmentFile(txtUnicodeInput.Text, txtUnicodeOutput.Text, txtUnicodeWordCount, chkUnicodeToExcel, true);
            }
            theButton.Enabled = true;  // enable it again
            btnClose.Enabled = true;
            Application.DoEvents();

        }
        private void btnSegmentBoth_Click(object sender, EventArgs e)
        {
            //  Segment both files in one go
            Button theButton = (Button)sender;
            tabControl1.SelectTab("Progress");
            theButton.Enabled = false;
            btnClose.Enabled = false;
            boxProgress.Items.Clear();  // empty the progress box
            SegmentFile(txtLegacyInput.Text, txtLegacyOutput.Text, txtLegacyWordCount, chkLegacyToExcel, false);
            SegmentFile(txtUnicodeInput.Text, txtUnicodeOutput.Text, txtUnicodeWordCount, chkUnicodeToExcel, true);
            if (chkLegacyToExcel.Checked || chkUnicodeToExcel.Checked)
            {
                //MakeInterlinear(excelApp);  // Make the interlinear worksheet, too
            }
            theButton.Enabled = true;
            btnClose.Enabled = true;
        }
        private void FinalCatch(Exception e)
        {
            MessageBox.Show(e.Message + "\r" + e.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            QuitWord();
            this.Close();
        }
        private void SegmentFile(String theInputFile, String theOutputFile, TextBox txtNumberOfWords, CheckBox SendToExcel, bool EvenRows)
        {
            /*
             * This is where we do all the segmentation and, if desired, writing to Excel
             */
            try
            {
                DateTime StartTime = DateTime.Now;  // Get the start time
                int NumberOfWords;
                int RowCounter = 0;
                progressBar1.Value = 0;

                Application.DoEvents();
                try
                {
                    InputDoc = wrdApp.Documents.Open(theInputFile, missing, true);  // Read only
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
                Application.DoEvents();
                OptimiseDoc(InputDoc);
                OptimiseDoc(OutputDoc);
                wrdApp.Options.Pagination = false;  // turn off background pagination
                wrdApp.Options.CheckGrammarAsYouType = false;   // Don't check grammar either
                wrdApp.Options.CheckSpellingAsYouType = false;  // Don't try to check spelling
                wrdApp.ScreenUpdating = false; // Turn off updating the screen
                wrdApp.ActiveWindow.ActivePane.View.ShowAll = false;  // Don't show special marks
                wrdApp.Selection.WholeStory(); // Make sure we've selected everything
                wrdApp.ScreenUpdating = false; // Turn off screen updating
                NumberOfWords = InputDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticWords, false);
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
                progressBar1.Value = 0;
                progressBar1.Maximum = InputDoc.StoryRanges.Count;
                foreach (WordRoot.Range rngStory in InputDoc.StoryRanges)
                {
                        rngStory.Copy();
                        //OutputDoc.ActiveWindow.Selection.MoveEnd(WordRoot.WdUnits.wdStory);  // move to the end of the story
                        OutputDoc.ActiveWindow.Selection.PasteAndFormat(WordRoot.WdRecoveryType.wdFormatOriginalFormatting); // We copy and paste to preserve formats etc.
                        OutputDoc.ActiveWindow.Selection.InsertAfter(" "); // and a space
               }
                InputDoc.Close(false);  // and close the input document as we no longer need it.
                OutputDoc.Save();  // and save the output document
                boxProgress.Items.Add("Document copied");
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
                boxProgress.Items.Add(Path.GetFileName(theOutputFile) + " saved after " +
                    DateTime.Now.Subtract(StartTime).TotalSeconds.ToString());
                if (SendToExcel.Checked)
                {
                    // We'll send the information to Excel
                    FillExcel(excelApp, wrdApp, OutputDoc, RowCounter);
                }
                wrdApp.ScreenUpdating = true; // turn on screen updating
                boxProgress.Items.Add("Completed in " + DateTime.Now.Subtract(StartTime).ToString());
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
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
                DateTime StartTime = DateTime.Now;
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
                DateTime StartTime2 = DateTime.Now;
                int Counter = 0;
                foreach (WordRoot.Frame theFrame in theDoc.Frames)
                {
                    theFrame.TextWrap = false; // Make it no longer wrap text
                    theFrame.Borders.OutsideLineStyle = WordRoot.WdLineStyle.wdLineStyleNone;
                    theFrame.Delete(); // and delete the frame
                    Counter++;
                    if (Counter % 100 == 0)
                    {
                        boxProgress.Items.Add("Deleted " + Counter.ToString() + " shapes");
                        Application.DoEvents();
                    }


                }
                DateTime EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed " + Counter.ToString() + " frames in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();                /*
                 /*
                  * Convert tables to text
                  */
                StartTime2 = DateTime.Now;
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
                EndTime = DateTime.Now;
                boxProgress.Items.Add("Converted " + Counter.ToString() + " tables in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();


                // Go to the beginning
                theApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);
                //  Make one column
                OneColumn(theApp);
                // Clear all tabs, paragraph markers, section breaks, manual line feeds, column breaks and manual page breaks.
                // ^m also deals with section breaks when wildcards are on.
                GlobalReplace(theApp.Selection, "[^9^11^13^14^12^m]", theSpace, false, true);


                // Clear all multiple spaces
                GlobalReplace(theApp.Selection, "  ", theSpace, true, false);
                // And this strange character found in some documents:  (F020)
                GlobalReplace(theApp.Selection, "", theSpace, true, false);

                // Clear the final space
                GlobalReplace(theApp.Selection, " ^p", "", false, false);
                /*
              * Now left align everything
              */
                foreach (WordRoot.Paragraph theParagraph in theDoc.Paragraphs)
                {
                    theParagraph.Format.Alignment = WordRoot.WdParagraphAlignment.wdAlignParagraphLeft;
                }



                EndTime = DateTime.Now;  //

                boxProgress.Items.Add("Cleaned the text in " + EndTime.Subtract(StartTime).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();
                progressBar1.Value += 1;
          
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
         }
        private void QuitWord()
        {
            if (wrdApp != null)
            {
                try
                {
                    // Shut down Word
                    wrdApp.Quit(ref missing, ref missing, ref missing);
                }
                catch
                { 
                }
                try
                {
                    // Shut down Excel
                    excelApp.ActiveWorkbook.Save();
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
            QuitWord();
            this.Close();
        }
        private void GlobalReplace(WordRoot.Selection theSelection, string SearchChars, string ReplacementChars, bool Repeat, bool Wildcards)
        {
            // Do a global replacement
            DateTime StartTime = DateTime.Now;
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
            DateTime EndTime = DateTime.Now;
            boxProgress.Items.Add("Globally replaced " + SearchChars + " in " + EndTime.Subtract(StartTime).TotalSeconds.ToString() + " seconds");
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
                DateTime StartTime = DateTime.Now;  // Start
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
                DateTime EndTime = DateTime.Now;
                TimeSpan ElapsedTime = EndTime.Subtract(StartTime);
                boxProgress.Items.Add("First pass complete in " + ElapsedTime.TotalSeconds.ToString() + " seconds");
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
                    DateTime StartTime2 = DateTime.Now;
                    theApp.ActiveDocument.UndoClear();  // Clear the undo stack
                    Found = theSelection.Find.Execute(missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
                        missing, missing, missing, missing);
                    EndTime = DateTime.Now;
                    ElapsedTime = EndTime.Subtract(StartTime2);
                    boxProgress.Items.Add("Second pass complete in " + ElapsedTime.TotalSeconds.ToString() + " seconds");
                    Application.DoEvents();
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


                theApp.ScreenUpdating = true;  // turn on updating
                EndTime = DateTime.Now;
                ElapsedTime = EndTime.Subtract(StartTime);
                progressBar1.Value = progressBar1.Maximum;  // We've finished!
                boxProgress.Items.Add("Segmentation complete in " + ElapsedTime.TotalSeconds.ToString() + " seconds");
                int LineCounter = NumberofWords / WordCount;
                boxProgress.Items.Add((LineCounter / ElapsedTime.TotalSeconds).ToString() + " lines per second");
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
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
            DateTime StartTime = DateTime.Now;
            //excelApp.Visible = true;
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
            // Go to the beginning of the document
            Application.DoEvents();
            // Initialise the progress bar
            progressBar1.Value = 0;
            progressBar1.Maximum = ParagraphCount;
            DateTime Start = DateTime.Now;
            boxProgress.Items.Add("Copying document to Excel...");
            theWorkSheet = theWorkBook.Sheets[1];
            // The header text
            theWorkSheet.Range["A" + RowCounter.ToString()].Value = HeaderText;
            theWorkSheet.Range["A" + RowCounter.ToString()].Interior.Color = CellColour[RowCounter - 1];
            int theRow = RowCounter + 2;
            int Counter = 0;
            DateTime StartCopy = DateTime.Now;
            foreach (WordRoot.Paragraph theParagraph in theDoc.Paragraphs)
            {
                string theCellRef = "A" + theRow.ToString();
                /*
                 * Sometimes the paste fails, so we try again if that is the case
                 */
                bool Failure = true;  // Assume failure
                int ErrorCounter = 0;
                theParagraph.Range.Copy();  // copy the range
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
                    boxProgress.Items.Add("Copied " + Counter.ToString() + " paragraphs in " + DateTime.Now.Subtract(StartCopy).TotalSeconds.ToString() + " seconds");
                    Application.DoEvents();
                }
                Counter++;

            }
            progressBar1.Value = ParagraphCount;
            boxProgress.Items.Add("Copied " + Counter.ToString() + " paragraphs in " +DateTime.Now.Subtract(Start).TotalSeconds.ToString() + " seconds");
            excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationAutomatic; // restore to automatic calculations
            excelApp.CalculateBeforeSave = true;
            theDoc.Close(false);
            theWorkSheet.Range["A1"].Select();  // go to the start of the worksheet
            theWorkBook.Save();
            boxProgress.Items.Add("Excel interlinear worksheet filled in " + DateTime.Now.Subtract(Start).ToString());
            Application.DoEvents();
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }

        }
        /*
        private void MakeInterlinear(ExcelApp theApp)
        {
            /*
             * Set up the interlinear worksheet
             
            try{
            DateTime Start = DateTime.Now;
            tabControl1.SelectTab("Progress");
            boxProgress.Items.Add("Generating interlinear worksheet");
            Application.DoEvents();
            WorkBook theWorkBook = theApp.ActiveWorkbook;
            // Clear the worksheet 
            theWorkBook.Sheets[1].Activate();
            ExcelRoot.Worksheet theWorkSheet = theWorkBook.Sheets[1];
            //excelApp.Visible = true;
            theWorkSheet.Cells.Select();
            excelApp.Selection.Clear();
            // now fill it in
            ExcelRoot.Range theCells;
            ExcelRoot.Range SourceCells;
            for (int theRow = 1; theRow <= 2; theRow++)
            {
                theCells = theWorkSheet.Cells[theRow, 1];
                theCells.Value = HeaderText[theRow - 1];
                theCells.Interior.Color = CellColour[theRow - 1];
            }
            for (int theRow = 1; theRow <= 2; theRow++)
            {
                theCells = theWorkSheet.Cells[theRow + 2, 1];
                theCells.Select();
                theCells.FormulaR1C1 = "=" + theWorkBook.Sheets[theRow + 1].Name + "!RC";
                // Copy the font name from the source cells.
                SourceCells = theWorkBook.Sheets[theRow + 1].cells[theRow + 2, 1];
                theCells.Font.Name = SourceCells.Font.Name;
                theCells.Interior.Color = CellColour[theRow - 1];
                theCells.HorizontalAlignment = ExcelRoot.XlHAlign.xlHAlignLeft;
            }
            /*
             * Now copy and paste the formulae
      

            theCells = theWorkSheet.Range["A3:A4"];
            theCells.Select();
            excelApp.Selection.Copy();
            theCells = theWorkSheet.Range["A5:A" + (MaxParagraphs + 4).ToString()];
            theCells.Select();
            theWorkSheet.Paste();
            theWorkSheet.Range["A1:A1"].Select();  // select cell A1
            theWorkBook.Save();  // Save it
            boxProgress.Items.Add("Finished creating interlinear worksheet in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString());
            theWorkBook.Close(); // and close the workbook
            Application.DoEvents();
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }

        }
        */
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
                    theDoc = wrdApp.Documents.Open(txtLegacyOutput.Text);
                    FileName = txtLegacyOutput.Text;
                }
                else
                {
                    EvenRows = true;
                    theDoc = wrdApp.Documents.Open(txtUnicodeOutput.Text);
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
            DateTime Start = DateTime.Now;
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
            boxProgress.Items.Add("Finished sending both files to Excel in " + DateTime.Now.Subtract(Start).ToString());
            theDoc = null;
            btnClose.Enabled = true;
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
                DateTime StartTime2 = DateTime.Now;
                DateTime EndTime;
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
 
                EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed or converted  " + (TableCounter + Counter).ToString() + " shapes in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();
                TotalCounter += Counter + TableCounter;

                /*
                * Remove all inline shapes
                */
                StartTime2 = DateTime.Now;
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
                EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed " + Counter.ToString() + " inline shapes in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
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

        
                                              
    }
};
