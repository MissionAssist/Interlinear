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
using Word = Microsoft.Office.Interop.Word.Application;
using WordRoot = Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word._Document;

using Office = Microsoft.Office.Core;


namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private WordApp wrdApp;
        private Document InputDoc;
        object missing = Type.Missing;
        private const string theSpace = " ";
        public Form1()
        {
            InitializeComponent();
            wrdApp = new Word();
            wrdApp.Visible = true;
            saveFileDialog1.SupportMultiDottedExtensions = true;
        }

        private void btnGetInputFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK )
            {
                txtInput.Text = openFileDialog1.FileName;
                btnSegmentInput.Enabled = true & txtOutput.Text.Length > 0;
                if (Path.GetExtension(txtInput.Text) == ".doc")
                {
                    saveFileDialog1.FilterIndex = 1; // .doc
                }
                else
                {
                    saveFileDialog1.FilterIndex = 2; // .docx
                }
                
                txtOutput.Text = Path.GetDirectoryName(txtInput.Text) + "\\" + Path.GetFileNameWithoutExtension(txtInput.Text) + " (Segmented)" + 
                    Path.GetExtension(txtInput.Text); // Add segmented to the file name by default.
                saveFileDialog1.FileName = txtOutput.Text;

        };
        }

        private void btnSegmentInput_Click(object sender, EventArgs e)
        {
            DateTime StartTime = DateTime.Now;  // Get the start time
            int NumberOfWords;
            wrdApp.Documents.Open(txtInput.Text);

            boxProgress.Items.Clear();
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
            boxProgress.Items.Add("Starting processing");
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
            wrdApp.Visible = false;  // Hide
            NumberOfWords = InputDoc.ComputeStatistics(WordRoot.WdStatistic.wdStatisticWords, false);

            txtWordCount.Text = NumberOfWords.ToString(); // the number of words in the document
            txtExpectedLines.Text = (NumberOfWords / WordsPerLine.Value).ToString();  // and the number of lines
            /*
             * Now remove text boxes, etc. from the document to clean it up.
             * We end with a single, huge paragraph
             */
            CleanWordText(wrdApp, InputDoc); // Clean the document

            /*
              * Now start splitting into a number of space-separated words, i.e. segmenting it.
              */
            Segment(wrdApp, wrdApp.Selection, (int)WordsPerLine.Value, NumberOfWords);

            InputDoc.SaveAs2(txtOutput.Text, InputDoc.SaveFormat); // Save in the same format as the input file
            DateTime EndTime = DateTime.Now;
            boxProgress.Items.Add("Completed in " + EndTime.Subtract(StartTime).ToString());
            wrdApp.ScreenUpdating = true; // turn on screen updating
            wrdApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
            wrdApp.Visible = true;  // show the finished document
            MessageBox.Show("Finished");
            Application.DoEvents();

            InputDoc.Close(false);
 
        }
        private void CleanWordText(WordApp theApp, Document theDoc )
        {
                DateTime StartTime = DateTime.Now;
                int Counter;
                boxProgress.Items.Add("Starting to clean the document...");
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
                    //Application.DoEvents();
                }
                EndTime = DateTime.Now;
                boxProgress.Items.Add("Removed " + Counter.ToString() + " frames in " + EndTime.Subtract(StartTime2).TotalSeconds.ToString() + " seconds");
                Application.DoEvents();

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

                // Go to the beginning
                theApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);
                //  Make one column
                OneColumn(theApp);
                // Clear all tabs, paragraph markers, section breaks, manual line feeds, column breaks and manual page breaks.
                // ^m also deals with section breaks when wildcards are on.
                GlobalReplace(theApp.Selection, "[^9^11^13^14^12^m]", theSpace, false, true);

                /*
                // Clear all paragraph markers
                GlobalReplace(theApp.Selection, "^p", theSpace, false);
                // Clear all section breaks
                GlobalReplace(theApp.Selection, "^b", theSpace, false);
                // Clear all manual line feeds
                GlobalReplace(theApp.Selection, "^l", theSpace, false);
                // Clear all column breaks
                GlobalReplace(theApp.Selection, "^n", theSpace, false);
                // Clear all manual page breaks
                GlobalReplace(theApp.Selection, "^m", theSpace, false);
                 */
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
             wrdApp.Quit(ref missing, ref missing, ref missing);
         }
         catch
         { 
         }

         GC.Collect();
         GC.WaitForPendingFinalizers();
         GC.Collect();
         GC.WaitForPendingFinalizers();
         wrdApp = null;

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
             int Lines = NumberofWords/WordCount;  // The number of lines rounded down
             theSelection.Find.Text = theSpace;
             theSelection.Find.Forward = true;
             theSelection.Find.Wrap = WordRoot.WdFindWrap.wdFindStop;  // Stop at end of document
             int LineCounter = 0;
             bool Found = true;  // Assume success
             // Now add paragraph markers
             for (LineCounter = 1; LineCounter <= Lines && Found; LineCounter++)  //  Keep going till we find no more spaces.
             {
                 int Counter = 0;
                 for (Counter = 0; Found && Counter < WordCount; Counter++ )
                 /*
                  * Keep going until we find the right number of spaces,  find no more
                  * spaces.
                  */
             /*
                 {
                     int ErrorCounter = 0;
                     bool Failure = true; // assume failure
                     while (Failure && ErrorCounter < 3) // Keep going until success or failure count >= 3
                     {
                         // retry twice on failure
                         try
                         {
                             Found = theSelection.Find.Execute();
                             Failure = false;  // we succeeded
                         }
                         catch (Exception e)
                         {
                             boxProgress.Items.Add("Find Error " + e.Message + " at line " + LineCounter.ToString());
                             ErrorCounter++;
                         }
                     }
                  }
                  LineCounter++;
                 if (Found)  // we still have some way to go so we add a paragraph marker
                 {
                     theSelection.InsertParagraphAfter();  // Add a paragraph mark
                     //theSelection.MoveRight(WordRoot.WdUnits.wdCharacter);  // and move beyond it
                     //theSelection.TypeText("\n");
                 }
                 if (LineCounter % 50 == 0 || ! Found)  //  Also write this at the end of the document
                 {
                     txtLineCount.Text = LineCounter.ToString();  // Mark progress
                     progressBar1.Value = Math.Min(LineCounter * WordCount, progressBar1.Maximum);
                     Application.DoEvents();
                 }
              }
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

         private void btnBrowseOutput_Click(object sender, EventArgs e)
         {
             if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
             {
                 
                 txtOutput.Text = saveFileDialog1.FileName;
                 btnSegmentInput.Enabled = true &  txtInput.Text.Length > 0;

             }
         }

         private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
         {

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

      

            
             
             
    }
    
};
