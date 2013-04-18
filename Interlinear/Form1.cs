/*
 * Interlinear - a program to take two Word documents, segment them into paragraphs of up to 20 words in length and write them to Excel
 * with the first (legacy) file in odd rows and the second (Unicode) file in even rows. This enables visual checking without the need to try
 * to do side-by-side comparisons.  It depends on both Word and Excel being installed on the computer.
 * 
 * It was writting as part of a MissionAssist project to convert documents in legacy fonts to Unicode.  Much of the logic is attributable to
 * Dennis Pepler, but the code here was written by Stephen Palmstrom.
 * 
 * Copyright © MissionAssist 2013
 * 
 * Last modified on 16 April 2013 by Stephen Palmstrom (stephen.palmstrom@btinternet.com).
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
        private ExcelApp   excelApp;
        object missing = Type.Missing;
        private const string theSpace = " ";
        private string[] theMessage = new string[2] {"Legacy text is in odd rows from file ", "Unicode text is in even rows from file "};
        private string[] HeaderText = new string[2];
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
                MakeInterlinear(excelApp);  // Make the interlinear worksheet, too
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
                    wrdApp.Documents.Open(theInputFile);
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
                InputDoc = wrdApp.ActiveDocument;
                InputDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
                InputDoc.ActiveWindow.View.Draft = true;  // Draft View
                InputDoc.ShowSpellingErrors = false;  // Don't show spelling errors
                InputDoc.ShowGrammaticalErrors = false; // Don't show grammar errors
                InputDoc.AutoHyphenation = false;
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
                 * Now remove text boxes, etc. from the document to clean it up.
                 * We end with a single, huge paragraph
                 */
                CleanWordText(wrdApp, InputDoc); // Clean the document

                /*
                  * Now start splitting into a number of space-separated words, i.e. segmenting it.
                  */
                Segment(wrdApp, wrdApp.Selection, (int)WordsPerLine.Value, NumberOfWords);

                InputDoc.SaveAs2(theOutputFile, InputDoc.SaveFormat); // Save in the same format as the input file
                boxProgress.Items.Add(Path.GetFileName(theOutputFile) + " saved after " +
                    DateTime.Now.Subtract(StartTime).TotalSeconds.ToString());
                if (SendToExcel.Checked)
                {
                    // We'll send the information to Excel
                    FillExcel(excelApp, wrdApp, RowCounter);
                }
                wrdApp.ScreenUpdating = true; // turn on screen updating
                wrdApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
                InputDoc.Close(false);
                boxProgress.Items.Add("Completed in " + DateTime.Now.Subtract(StartTime).ToString());
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }
        }
        private void CleanWordText(WordApp theApp, Document theDoc )
        {
            try
            {
                DateTime StartTime = DateTime.Now;
                int Counter;
                theDoc.Activate();
                boxProgress.Items.Add("Starting to clean the document...");
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
                Counter = 0;
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
                Application.DoEvents();
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

                    theSelection.Find.Replacement.Text += "^p^p"; // ending with two paragraph
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
                 * And any more than two paragraph markers together
                 */
                GlobalReplace(theSelection, "^p^p^p", "^p^p", false, false);


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
                if (theWorkbook.Sheets.Count < 3)
                {
                    theWorkbook.Sheets.Add(missing, missing, 3 - theWorkbook.Sheets.Count); // Add two sheets.
                }
                theWorkbook.Sheets[1].Name = "Interlinear";
                theWorkbook.Sheets[2].Name = "Legacy";
                theWorkbook.Sheets[3].Name = "Unicode";
                theWorkbook.Sheets[1].Columns("A").ColumnWidth = 100;  // and make the first column wide
                theWorkbook.SaveAs(txtExcelOutput.Text);  // save it
            }
            if (EvenRows)
            {
                theRow = 2;
            }
            else
            {
                theRow = 1;

            }
            // Clear the individual sheets
            if (theWorkbook.Sheets.Count < 3)
            {
                MessageBox.Show("Not enough worksheets - you haven't chosen the right workbook", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                chkLegacyToExcel.Checked = false;
                chkUnicodeToExcel.Checked = false;
                theWorkbook.Close(false);  // and close the workbook
                return -1;
            }
            else
            {
                theWorkbook.Sheets[theRow + 1].Activate();
                ExcelRoot.Worksheet theWorkSheet = theWorkbook.ActiveSheet;
                theWorkSheet.Cells.Select();
                excelApp.Selection.Clear();
                return theRow;
            }
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
                return -1;
            }

        }
        private void FillExcel(ExcelApp excelApp, WordApp wrdApp, int RowCounter)
        {
            try
            {
            DateTime StartTime = DateTime.Now;
            //excelApp.Visible = true;
            boxProgress.Items.Add("Starting to fill Excel worksheet");
            ExcelRoot.Workbook theWorkBook = excelApp.ActiveWorkbook;  // remember the document.
            excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationManual; // Don't calculate automatically.
            Application.DoEvents();
            // Get document and worksheet
            WordRoot.Document theDoc = wrdApp.ActiveDocument;
            theDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure it isn't in reading layout.
            ExcelRoot.Worksheet theWorkSheet = theWorkBook.Sheets[RowCounter + 1];
            int ErrorCounter = 0;
            bool Failure;
            /*
             * Generate the header messages
             */
            HeaderText[RowCounter - 1] = theMessage[RowCounter - 1] + Path.GetFileName(theDoc.FullName);
              /*
            System.Text.RegularExpressions.Regex NonBreakingHyphen = new System.Text.RegularExpressions.Regex("\x1E", 
                System.Text.RegularExpressions.RegexOptions.Multiline);  // Non-breaking hyphen
             */
            int ParagraphCount = theDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticParagraphs);
            MaxParagraphs = Math.Max(MaxParagraphs, ParagraphCount * 2);
            boxProgress.Items.Add("There are " + ParagraphCount.ToString() + " paragraphs");
            // Go to the beginning of the document
            Application.DoEvents();
            // Initialise the progress bar
            progressBar1.Value = 0;
            theDoc.ActiveWindow.Selection.WholeStory();
            GlobalReplace(theDoc.ActiveWindow.Selection, "^~", "-", false, false); // replace non-breaking hyphens with hyphens
            theDoc.ActiveWindow.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
            theDoc.ActiveWindow.Selection.WholeStory();  // Select it all
            DateTime Start = DateTime.Now;
            boxProgress.Items.Add("Copying document...");
            theDoc.ActiveWindow.Selection.Copy(); // copy it.
            boxProgress.Items.Add("Document copied in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString());
            //string Message = "";
            Start = DateTime.Now;
            theWorkBook.Sheets[RowCounter + 1].Activate();
            theWorkSheet = theWorkBook.ActiveSheet;
            theWorkSheet.Cells[RowCounter + 2, 1].Select();
            Failure = true;  // Assume failure so we go into the loop.
            //excelApp.Visible = true;
            while (Failure && ErrorCounter < 3)
            {
                try
                    {
                    theWorkSheet.Paste();
                    Failure = false;
                    }
                catch (Exception e)
                {
                    boxProgress.Items.Add("Copy error " + e.Message + " in row " + RowCounter.ToString());
                    ErrorCounter++;
                    Application.DoEvents();
                }
            }
            //excelApp.Visible = false;
            boxProgress.Items.Add("Document pasted in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString());
            Application.DoEvents();
            Start = DateTime.Now;

            excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationAutomatic; // restore to automatic calculations
            excelApp.CalculateBeforeSave = true;
            theWorkBook.Save();
            boxProgress.Items.Add("Excel interlinear worksheet pasted in " + DateTime.Now.Subtract(Start).TotalSeconds.ToString());
            boxProgress.Items.Add("Finished filling Excel in " + DateTime.Now.Subtract(StartTime).ToString());
            Application.DoEvents();
            }
            catch (Exception Ex)
            {
                FinalCatch(Ex);
            }

        }
        private void MakeInterlinear(ExcelApp theApp)
        {
            /*
             * Set up the interlinear worksheet
             */
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
             */

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
                FillExcel(excelApp, wrdApp, RowCounter);
                boxProgress.Items.Add("Finished sending to Excel.  Go to the setup tab to create the interlinear worksheet");
            }
            else
            {
                boxProgress.Items.Add("Could not send to Excel - you chose the wrong workbook");
            }
            theDoc.Close(false);
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
            bool Continue = false;
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
            if (RowCounter > 0)
            {
                Continue = true;
                FillExcel(excelApp, wrdApp, RowCounter);
            }
            else
            {
                boxProgress.Items.Add("Could not send to Excel  - you chose the wrong worksheet");
            }
            theDoc.Close(false);
            if (Continue)
            {
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
                if (RowCounter > 0)
                {
                    FillExcel(excelApp, wrdApp, RowCounter);
                    MakeInterlinear(excelApp); // Make the interlinear worksheet, too.
                    boxProgress.Items.Add("Finished sending both files to Excel in " + DateTime.Now.Subtract(Start).ToString());
                }
                else
                {
                    boxProgress.Items.Add("Could not send to Excel - you chose the wrong worksheet");

                }
            }
            theDoc.Close(false);
            theDoc = null;
            btnInterlinear.Enabled = true;
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

            MakeInterlinear(excelApp);
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
