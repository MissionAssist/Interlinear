using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
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
        private Document LegacyDoc;
        object missing = Type.Missing;
        public Form1()
        {
            InitializeComponent();
            wrdApp = new Word();
            wrdApp.Visible = true;
            saveFileDialog1.SupportMultiDottedExtensions = true;
        }

        private void btnGetLegacyFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK )
            {
                txtLegacy.Text = openFileDialog1.FileName;
                btnSegmentLegacy.Enabled = true & txtOutput.Text.Length > 0;

        };
        }

        private void btnSegmentLegacy_Click(object sender, EventArgs e)
        {
            DateTime StartTime = DateTime.Now;  // Get the start time
            wrdApp.Documents.Open(txtLegacy.Text);
            wrdApp.Visible = false;  // Hide
            boxProgress.Items.Clear();
            LegacyDoc = wrdApp.ActiveDocument;
            LegacyDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
            wrdApp.Selection.WholeStory(); // Make sure we've selected everything
            CleanWordText(wrdApp, LegacyDoc); // Clean the document
            DateTime EndTime = DateTime.Now;  //
            
            boxProgress.Items.Add("Cleaned the text in " + EndTime.Subtract(StartTime).ToString());

            /*
              * Now start splitting into a number of space-separated words
              */
            Segment(wrdApp, wrdApp.Selection, (int)WordsPerLine.Value);

            LegacyDoc.SaveAs2(txtOutput.Text, LegacyDoc.SaveFormat); // Save in the same format as the input file
            EndTime = DateTime.Now;
            boxProgress.Items.Add("Completed in " + EndTime.Subtract(StartTime).ToString());
            wrdApp.Visible = true;  // show the finished document
            MessageBox.Show("Finished");

            LegacyDoc.Close(false);
 
        }
        private void CleanWordText(WordApp theApp, Document theDoc )
        {
            theApp.ScreenUpdating = false; // Turn off screen updating
         
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
            // Clear all tabs
            GlobalReplace(theApp.Selection, "^t", " ", true);
            // Clear all paragraph markers
            GlobalReplace(theApp.Selection, "^p", " ", false);
            // Clear all section breaks
            GlobalReplace(theApp.Selection, "^b", " ", false);
            // Clear all manual line feeds
            GlobalReplace(theApp.Selection, "^l", " ", false);
            // Clear all manual page breaks
            GlobalReplace(theApp.Selection, "^m", " ", false);
            // Clear all multiple spaces
            GlobalReplace(theApp.Selection, "  ", " ", true);
            theApp.ScreenUpdating = true; // turn on screen updating
         }
  
         private void QuitWord()
            {
                if (wrdApp != null)
                {
                    wrdApp.Quit(ref missing, ref missing, ref missing);
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
                 Found = Found && Repeat;  // If repeat not set, then we only execute once.
                 Application.DoEvents();
             }
         }
         private void Segment(WordApp  theApp, WordRoot.Selection theSelection, int WordCount)
         {
             /*
             * Now segment into the number of words specified by the WordCount paramenter
             */
             DateTime StartTime = DateTime.Now;  // Start
             // Go to the beginning
             theSelection.HomeKey(WordRoot.WdUnits.wdStory);
             theApp.ScreenUpdating = false; // Turn off updating the screen
             // Size the progressbar
             progressBar1.Maximum = theApp.ActiveDocument.Words.Count/2; 
 
             theSelection.Find.Text = " ";
             theSelection.Find.Forward = true;
             theSelection.Find.Wrap = WordRoot.WdFindWrap.wdFindStop;  // Stop at end of document
             int LineCounter = 0;
            // Now add paragraph markers
             while (theApp.WordBasic.AtEndofDocument() == 0)
             {
                 int Counter = 0;
                 bool Found = true;
                 while (Counter < WordCount & theApp.WordBasic.AtEndofDocument() == 0 & Found)
                    /*
                     * Keep going until we find the right number of spaces, the end of document or find no more
                     * spaces.
                     */
                 {
                     int ErrorCounter = 0;
                     bool Failure = true; // assume failure
                     while (Failure & ErrorCounter < 3) // Keep going until success or failure count >= 3
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
                     Counter++; // increment
                 }
                 theSelection.InsertParagraphAfter();  // Add a paragraph mark
                 theSelection.MoveRight(WordRoot.WdUnits.wdCharacter);  // and move beyond it
                 LineCounter++;
                 if (LineCounter % 10 == 0)
                 {
                     txtLineCount.Text = LineCounter.ToString();  // Mark progress
                     progressBar1.Value = Math.Min(LineCounter * WordCount, progressBar1.Maximum);
                 }
                 Application.DoEvents();
              }
             theApp.ScreenUpdating = true;  // turn on updating
             DateTime EndTime = DateTime.Now;
             boxProgress.Items.Add("Segmented in " + EndTime.Subtract(StartTime).ToString());
         }

         private void btnBrowseOutput_Click(object sender, EventArgs e)
         {
             if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
             {
                 
                 txtOutput.Text = saveFileDialog1.FileName;
                 btnSegmentLegacy.Enabled = true &  txtLegacy.Text.Length > 0;

             }
         }

 
             
    }

};
