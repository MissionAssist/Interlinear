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
        public Form1()
        {
            InitializeComponent();
            wrdApp = new Word();
            wrdApp.Visible = false;
            excelApp = new Excel();  // open Excel
            excelApp.Visible = false; 
            saveLegacyFileDialog.SupportMultiDottedExtensions = true;
            saveUnicodeFileDialog.SupportMultiDottedExtensions = true;
            Wordcount.SetToolTip(WordsPerLine, "If you want more than eight words per line, they must be in multiples of four");
        }

        private void btnGetInputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                HandleInputFile(txtLegacyInput, txtLegacyOutput, btnSegmentLegacy, openLegacyFileDialog, saveLegacyFileDialog);
            }
            else
            {
                HandleInputFile(txtUnicodeInput, txtUnicodeOutput, btnSegmentUnicode, openUnicodeFileDialog,saveUnicodeFileDialog);
            }

        }
        private void HandleInputFile(TextBox InputText, TextBox OutputText, Button SegmentButton, OpenFileDialog theOpenFileDialog, SaveFileDialog theSaveFileDialog)
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
                    " (Segmented)", Path.GetExtension(InputText.Text));

                if (File.Exists(OutputText.Text))
                {
                    SegmentButton.Enabled = true;
                }

        };
        }
        private void btnGetOutputFile_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            if (theButton.Parent.Text == "Legacy")
            {
                HandleOutputFile(txtLegacyInput, txtLegacyOutput, saveLegacyFileDialog, btnSegmentLegacy);
            }
            else
            {
                HandleOutputFile(txtUnicodeInput, txtUnicodeOutput, saveUnicodeFileDialog, btnSegmentUnicode);
            }
        }
        private void HandleOutputFile (TextBox theInputBox, TextBox theOutputBox, SaveFileDialog theDialog, Button SegmentButton)
        {
            if (theDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                theOutputBox.Text = theDialog.FileName;
                SegmentButton.Enabled = theOutputBox.Text.Length > 0 && File.Exists(theInputBox.Text);  // only enable if both boxes filled in
                /*
                 * If both individual segment buttons are enabled, we enable the segment both button, too.
                 */
                btnSegmentBoth.Enabled = btnSegmentLegacy.Enabled && btnSegmentUnicode.Enabled;
                btnBothToExcel.Enabled = File.Exists(txtLegacyOutput.Text) && File.Exists(txtUnicodeOutput.Text);


            }
 
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
            Application.DoEvents();

        }
        private void btnSegmentBoth_Click(object sender, EventArgs e)
        {
            //  Segment both files in one go
            Button theButton = (Button)sender;
            tabControl1.SelectTab("Progress");
            theButton.Enabled = false;
            SegmentFile(txtLegacyInput.Text, txtLegacyOutput.Text, txtLegacyWordCount, chkLegacyToExcel, false);
            SegmentFile(txtUnicodeInput.Text, txtUnicodeOutput.Text, txtUnicodeWordCount, chkUnicodeToExcel, true); 
            theButton.Enabled=true;
        }
        private void SegmentFile(String theInputFile, String theOutputFile, TextBox txtNumberOfWords, CheckBox SendToExcel, bool EvenRows)
        {
            /*
             * This is where we do all the segmentation and, if desired, writing to Excel
             */

            DateTime StartTime = DateTime.Now;  // Get the start time
            int NumberOfWords;
            int RowCounter = 0;
            ExcelRoot.XlThemeColor CellColour = ExcelRoot.XlThemeColor.xlThemeColorAccent1;
            wrdApp.Documents.Open(theInputFile);
            // process Excel if desired
            if (SendToExcel.Checked)
            {
                RowCounter = InitialiseExcel(excelApp, EvenRows,  ref CellColour, theInputFile);
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
            progressBar1.Value = 0;
            /*
             * Set various Word options to optimise performance
             * 
             */
            boxProgress.Items.Add("**** Starting processing " + Path.GetFileName(theInputFile));
            Application.DoEvents();
            wrdApp.Options.Pagination = false;  // turn off background pagination
            wrdApp.Options.CheckGrammarAsYouType = false;   // Don't check grammar either
            wrdApp.Options.CheckSpellingAsYouType = false;  // Don't try to check spelling
            wrdApp.ScreenUpdating = false; // Turn off updating the screen
            wrdApp.ActiveWindow.ActivePane.View.ShowAll = false;  // Don't show special marks
            wrdApp.Selection.WholeStory(); // Make sure we've selected everything
            wrdApp.ScreenUpdating = false; // Turn off screen updating
            InputDoc = wrdApp.ActiveDocument;
            InputDoc.ActiveWindow.View.Draft = true;  // Draft View
            InputDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
            InputDoc.ShowSpellingErrors = false;  // Don't show spelling errors
            InputDoc.ShowGrammaticalErrors = false; // Don't show grammar errors
            InputDoc.AutoHyphenation = false;
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
            DateTime EndTime = DateTime.Now;
            boxProgress.Items.Add(Path.GetFileName(theOutputFile) + " saved after " + 
                EndTime.Subtract(StartTime).TotalSeconds.ToString());
            if (SendToExcel.Checked)
            {
                // We'll send the information to Excel
                FillExcel(excelApp, wrdApp, RowCounter, CellColour);
            }
            wrdApp.ScreenUpdating = true; // turn on screen updating
            wrdApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
            InputDoc.Close(false);
            boxProgress.Items.Add("Completed in " + EndTime.Subtract(StartTime).ToString());
 
        }
        private void CleanWordText(WordApp theApp, Document theDoc )
        {
                DateTime StartTime = DateTime.Now;
                int Counter;
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
                 * Remove all shapes
                 * 
                 */
                Counter = 0;
                DateTime StartTime2 = DateTime.Now;
                foreach (WordRoot.Shape theShape in theDoc.Shapes)
                {
                
                    if (theShape.Type == Office.MsoShapeType.msoTextBox)
                    {
                        theShape.ConvertToInlineShape();
 
                        theShape.Delete();
                        Counter++;
                        if (Counter % 100 == 0)
                        {
                            boxProgress.Items.Add("Deleted " + Counter.ToString() + " shapes");
                            Application.DoEvents();
                        }
                    }

                }
                DateTime EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed " + Counter.ToString() + " textboxes in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();
 
                /*
                 * Remove all frames
                 */
                StartTime2 = DateTime.Now;
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
                EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed " + Counter.ToString() + " frames in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
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
            // Shut down Excle
            excelApp.Quit();
        }
        catch
        {
        }
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
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

            // Go to the beginning
             theSelection.HomeKey(WordRoot.WdUnits.wdStory);
             boxProgress.Items.Add("Starting segmentation...");
             DateTime StartTime = DateTime.Now;  // Start
             bool Found;
      /*
              * Use wildcards to add the paragraph markers
              * 
              */
            theSelection.Find.ClearFormatting();
            theSelection.Find.Replacement.Text = "^&^p";  // Replace with what we just found and a paragraph marker
            theSelection.Find.MatchWildcards =true;
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
            const string WildCards = "([! ]@ )";
            theSelection.Find.Text = "";  // Clear the find string
            /*
            * Build up the search string
             * 
             * If the words per line we want at the end are more than seven, we need to do the replacement
             * in two stages as otherwise the wildcard expression gets too complicated.
            */
            int MaxWordPerLine = 4;
            if (WordCount <= 7)
            {
                MaxWordPerLine = 7; 
            }
            for (int i = 1; i <= Math.Min(WordCount, MaxWordPerLine); i++)
            {
                theSelection.Find.Text += WildCards;

            }

            // Now do the first replacement
            boxProgress.Items.Add("Starting segmentation first pass");
            Application.DoEvents();
            theApp.ActiveDocument.UndoClear();  // Clear the undo stack

            Found = theSelection.Find.Execute(missing,  missing, missing, missing, missing, missing, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
            missing, missing, missing, missing);
            DateTime EndTime = DateTime.Now;
            TimeSpan ElapsedTime = EndTime.Subtract(StartTime);
            boxProgress.Items.Add("First pass complete in " + ElapsedTime.TotalSeconds.ToString() + " seconds");
            Application.DoEvents();
            progressBar1.Value += 1;
            Application.DoEvents();


            /*
             * If the WordCount > 4, we assume 8 etc.
             */
            if (WordCount > 7)
            {
                const string Paragraphs = "(*)^13";  // Match anything ending with a paragraph
                theSelection.Find.Text = "";
                theSelection.Find.Replacement.Text = "";
                for (int i = 1; i <= WordCount / 4; i++)
                {
                    theSelection.Find.Text += Paragraphs; // build up the search string
                    theSelection.Find.Replacement.Text += "\\" + i.ToString();
                }
                theSelection.Find.Replacement.Text += "^p"; // ending with paragraph
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
             
             theApp.ScreenUpdating = true;  // turn on updating
             EndTime = DateTime.Now;
             ElapsedTime =  EndTime.Subtract(StartTime);
             progressBar1.Value = progressBar1.Maximum;  // We've finished!
             boxProgress.Items.Add("Segmentation complete in " +ElapsedTime.TotalSeconds.ToString() + " seconds");
             int LineCounter = NumberofWords / WordCount;
             boxProgress.Items.Add((ElapsedTime.TotalSeconds/LineCounter).ToString() + " seconds per line");
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

        private int InitialiseExcel(ExcelApp excelApp, bool EvenRows, ref ExcelRoot.XlThemeColor CellColour, string FileName)
        {
            string HeaderText;
            int theRow;
            int RowCounter;
            bool CellFilled = true;
            string StrippedFileName;
            StrippedFileName = Path.GetFileName(FileName);  // Get the file name without the directory
            if (File.Exists(txtExcelOutput.Text))
            {
                excelApp.Workbooks.Open(txtExcelOutput.Text);  // Open the file
            }
            else
            {
                excelApp.Workbooks.Add();  // add it
                excelApp.ActiveWorkbook.ActiveSheet.Columns("A").ColumnWidth = 100;  // and make the first column wide
                excelApp.ActiveWorkbook.SaveAs(txtExcelOutput.Text);  // save it
            }
            // Now initialise the cells
            ExcelRoot.Worksheet theWorkSheet = excelApp.ActiveSheet;
            /*
             * We start at row 2 if even rows, otherwise row 1
             */
            if (EvenRows)
            {
                theRow = 2;
                HeaderText = "Unicode text in even rows from " + StrippedFileName ;
                CellColour = ExcelRoot.XlThemeColor.xlThemeColorAccent2;
            }
            else
            {
                theRow = 1;
                HeaderText = "Legacy text in odd rows from " + StrippedFileName;
                CellColour = ExcelRoot.XlThemeColor.xlThemeColorAccent6;

            }
            // Write the header row
            theWorkSheet.Range["A" + theRow.ToString()].Select();
            excelApp.ActiveCell.FormulaR1C1 = HeaderText;
            excelApp.Selection.Interior.ThemeColor = CellColour;
            //
            // Clear the remaining rows
            //
            RowCounter = theRow + 2;  // start filling two rows down from the header
            while (CellFilled)
            {
                ExcelRoot.Range theCells = theWorkSheet.Cells[RowCounter, 1];
                CellFilled = theCells.Value != null;
                if (CellFilled)
                {
                    theCells.Value = null;
                    theCells.ClearFormats(); // Clear the formats
                }
                RowCounter += 2;  // Increment by 2
            }
            return theRow + 2;
        }
        private void FillExcel(ExcelApp excelApp, WordApp wrdApp, int RowCounter, ExcelRoot.XlThemeColor CellColour)
        {
            DateTime StartTime = DateTime.Now;
            boxProgress.Items.Add("Starting to fill Excel worksheet");
            Application.DoEvents();
            // Get document and worksheet
            WordRoot.Document theDoc = wrdApp.ActiveDocument;
            theDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure it isn't in reading layout.
            ExcelRoot.Worksheet theWorkSheet = excelApp.ActiveSheet;
            excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationManual; // Don't calculate automatically.
            int CharactersMoved = 2;
            int ErrorCounter = 0;
            bool Failure = true;
            string theText;
            System.Text.RegularExpressions.Regex NonBreakingHyphen = new System.Text.RegularExpressions.Regex("\x1E", 
                System.Text.RegularExpressions.RegexOptions.Multiline);  // Non-breaking hyphen
            int ParagraphCount = theDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticParagraphs);
            boxProgress.Items.Add("There are " + ParagraphCount.ToString() + " paragraphs");
            // Go to the beginning of the document
            Application.DoEvents();
            wrdApp.Selection.WholeStory();
            wrdApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
            int cellCounter = 0;
            //string Message = "";
            while (CharactersMoved > 1)
            {
                CharactersMoved = wrdApp.Selection.EndOf(WordRoot.WdUnits.wdParagraph, WordRoot.WdMovementType.wdExtend); // select to end of paragraph
                //CharactersMoved = wrdApp.Selection.MoveRight(WordRoot.WdUnits.wdParagraph, 1, WordRoot.WdMovementType.wdExtend);
                if (CharactersMoved > 0)
                {
                    //wrdApp.Selection.Copy(); // copy to clipboard
                    ExcelRoot.Range theCells = theWorkSheet.Cells[RowCounter, 1];  // get the cell
                    //
                    //  We'll retry pasting if we hit an error
                    //
                    Failure = true;  // Assume failure so we go into the loop.
                    while (Failure && ErrorCounter < 3)
                    try
                    {
                        theText = NonBreakingHyphen.Replace(wrdApp.Selection.Text, "-");
                        theCells.Font.Name = wrdApp.Selection.Font.Name;  // and the font
                        theCells.Value = theText;  // copy the Word selection to Excel
                        theCells.Interior.ThemeColor = CellColour;
                        ErrorCounter = 0;
                        Failure = false;
                    }
                    catch (Exception e)
                    {
                        boxProgress.Items.Add("Copy error " + e.Message + " in row " + RowCounter.ToString());
                        ErrorCounter++;
                    }
                   
                    //
                    //  This manoeuver should detect the end of the document.
                    CharactersMoved = wrdApp.Selection.MoveRight(WordRoot.WdUnits.wdCharacter, 2, WordRoot.WdMovementType.wdMove); // move to the next two characters
                    wrdApp.Selection.MoveLeft(WordRoot.WdUnits.wdCharacter, 1, WordRoot.WdMovementType.wdMove); // and back one
                    RowCounter += 2;  // Increment the row
                    cellCounter++;  // and a counter
                    if (ParagraphCount > 10 && cellCounter % (ParagraphCount/10) == 0)
                    {
                        boxProgress.Items.Add("Written " + cellCounter.ToString() + " rows");
                        Application.DoEvents();
                    }
                    if (cellCounter % 50 == 0)
                    {
                        Application.DoEvents();
                    }

                }
            }
            excelApp.Calculation = ExcelRoot.XlCalculation.xlCalculationAutomatic; // restore to automatic calculations
            excelApp.CalculateBeforeSave = true;
            excelApp.ActiveWorkbook.Save();
            DateTime EndTime = DateTime.Now;
            boxProgress.Items.Add("Finished filling " + cellCounter.ToString() + " rows in Excel in " + EndTime.Subtract(StartTime).TotalSeconds.ToString());
        }

        private void SendToExcel_Click(object sender, EventArgs e)
        {
            Button theButton = (Button)sender;
            bool EvenRows;
            ExcelRoot.XlThemeColor CellColour = ExcelRoot.XlThemeColor.xlThemeColorAccent2;
            Document theDoc;
            string FileName;
            btnClose.Enabled = false;
            tabControl1.SelectTab("Progress");
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
            
            // We'll send the information to Excel
            int RowCounter = InitialiseExcel(excelApp, EvenRows, ref CellColour, FileName);
            FillExcel(excelApp, wrdApp, RowCounter, CellColour);
            theDoc.Close(false);
            theDoc = null;
            btnClose.Enabled = true;

        }
        private void BothToExcel_Click(object sender, EventArgs e)
        {
            ExcelRoot.XlThemeColor CellColour = ExcelRoot.XlThemeColor.xlThemeColorAccent2;
            // We'll send the information to Excel
            Document theDoc;
            tabControl1.SelectTab("Progress");
            btnClose.Enabled = false;
            theDoc = wrdApp.Documents.Open(txtLegacyOutput.Text);

            int RowCounter = InitialiseExcel(excelApp, false, ref CellColour, txtLegacyOutput.Text);
            FillExcel(excelApp, wrdApp, RowCounter, CellColour);
            theDoc.Close(false);
            theDoc = wrdApp.Documents.Open(txtUnicodeOutput.Text);
            RowCounter = InitialiseExcel(excelApp, true, ref CellColour, txtUnicodeOutput.Text);
            FillExcel(excelApp, wrdApp, RowCounter, CellColour);
            theDoc.Close(false);
            theDoc = null;
            btnClose.Enabled = true;
        }

             
    }
    
};
