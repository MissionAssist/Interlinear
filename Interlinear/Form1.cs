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
                txtOutput.Text = Path.GetFileNameWithoutExtension(txtInput.Text) + " (Segmented)" + 
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
            /*
             * Set various Word options to optimise performance
             * 
             */
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
            boxProgress.Items.Add("Starting the cleanup...");
            Application.DoEvents();
            /*
             * Now remove text boxes, etc. from the document to clean it up.
             * We end with a single, huge paragraph
             */
            CleanWordText(wrdApp, InputDoc); // Clean the document
            DateTime EndTime = DateTime.Now;  //
            
            boxProgress.Items.Add("Cleaned the text in " + EndTime.Subtract(StartTime).ToString());
            Application.DoEvents();

            /*
              * Now start splitting into a number of space-separated words, i.e. segmenting it.
              */
            Segment(wrdApp, wrdApp.Selection, (int)WordsPerLine.Value, NumberOfWords);

            InputDoc.SaveAs2(txtOutput.Text, InputDoc.SaveFormat); // Save in the same format as the input file
            EndTime = DateTime.Now;
            boxProgress.Items.Add("Completed in " + EndTime.Subtract(StartTime).ToString());
            wrdApp.ScreenUpdating = true; // turn on screen updating
            wrdApp.Selection.HomeKey(WordRoot.WdUnits.wdStory);  // go to the beginning
            wrdApp.Visible = true;  // show the finished document
            MessageBox.Show("Finished");

            InputDoc.Close(false);
 
        }
        private void CleanWordText(WordApp theApp, Document theDoc )
        {
         
            /*
             * Remove all shapes
             * 
             */
            foreach (WordRoot.Shape theShape in theDoc.Shapes)
            {
                
                if (theShape.Type == Office.MsoShapeType.msoTextBox)
                {
                    theShape.ConvertToInlineShape();
 
                    theShape.Delete();
                    Application.DoEvents();
                }
            }
            /*
             * Remove all frames
             */
            foreach (WordRoot.Frame theFrame in theDoc.Frames)
            {
                theFrame.TextWrap = false; // Make it no longer wrap text
                theFrame.Borders.OutsideLineStyle = WordRoot.WdLineStyle.wdLineStyleNone;
                theFrame.Delete(); // and delete the frame
                Application.DoEvents();
            }
            /*
             * Now left align everything
             */
            foreach (WordRoot.Paragraph theParagraph in theDoc.Paragraphs)
            {
                theParagraph.Format.Alignment = WordRoot.WdParagraphAlignment.wdAlignParagraphLeft;
            }

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
            // Clear all tabs
            GlobalReplace(theApp.Selection, "^t", theSpace, true);
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
            // Clear all multiple spaces
            GlobalReplace(theApp.Selection, "  ", theSpace, true);
            // Clear the final space
            GlobalReplace(theApp.Selection, " ^p", "", false);
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
         private void GlobalReplace(WordRoot.Selection theSelection, string SearchChars, string ReplacementChars, bool Repeat)
         {
             // Do a global replacement
             bool Found = true;  // Assume success
             theSelection.Find.Text = SearchChars;
             theSelection.Find.Replacement.Text = ReplacementChars;
             theSelection.Find.Wrap = WordRoot.WdFindWrap.wdFindContinue;
            //
             // If we want to keep searching, we'll do so
             //
             while (Found)
             {
                 Found = theSelection.Find.Execute(missing, false, false, false, false, false, missing, missing, missing, missing, WordRoot.WdReplace.wdReplaceAll,
                 missing, missing, missing, missing);
                 Found = Repeat && Found;  // If repeat not set, then we only execute once.
                 Application.DoEvents();
             }
         }
         private void Segment(WordApp  theApp, WordRoot.Selection theSelection, int WordCount, int NumberofWords)
         {
             /*
             * Now segment into the number of words specified by the WordCount paramenter
             */

            // Go to the beginning
             theSelection.HomeKey(WordRoot.WdUnits.wdStory);
             // Size the progressbar
             progressBar1.Maximum = NumberofWords;
             boxProgress.Items.Add("Starting segmentation...");
             DateTime StartTime = DateTime.Now;  // Start
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
              * Now remove the trailing spaces
              */
             GlobalReplace(theSelection, " ^p", "^p", false);
             
             theApp.ScreenUpdating = true;  // turn on updating
             DateTime EndTime = DateTime.Now;
             TimeSpan ElapsedTime =  EndTime.Subtract(StartTime);
             progressBar1.Value = progressBar1.Maximum;  // We've finished!
             boxProgress.Items.Add("Segmented in " +ElapsedTime.TotalSeconds.ToString() + " seconds");
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

      

            
             
             
    }
    
};
