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
            wrdApp.Documents.Open(txtLegacy.Text);
            LegacyDoc = wrdApp.ActiveDocument;
            LegacyDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
            wrdApp.Selection.WholeStory(); // Make sure we've selected everything
            CleanWordText(wrdApp, LegacyDoc); // Clean the document

            /*
              * Now start splitting into 12 space-separated words
              */
            Segment(wrdApp, wrdApp.Selection, 12);

            LegacyDoc.SaveAs2(txtOutput.Text, LegacyDoc.SaveFormat); // Save in the same format as the input file
            LegacyDoc.Close(false);
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
                 Found = Found && Repeat;
                 Application.DoEvents();
             }
         }
         private void Segment(WordApp  theApp, WordRoot.Selection theSelection, int WordCount)
         {
             /*
             * Now segment into the number of words specified by the WordCount paramenter
             */
             
             // Go to the beginning
             progressBar1.Maximum = theApp.ActiveDocument.Words.Count; 
             theSelection.HomeKey(WordRoot.WdUnits.wdStory);
             theSelection.Find.Text = " ";
             theSelection.Find.Forward = true;
             int LineCounter = 0;
            // Now add paragraph markers
             while (theApp.WordBasic.AtEndofDocument() == 0)
             {
                 int Counter = 0;
                 while (Counter < WordCount & theApp.WordBasic.AtEndofDocument() == 0)
                 {
                     theSelection.Find.Execute();
                     Counter++; // increment
                 }
                 theSelection.InsertParagraphAfter();  // Add a paragraph mark
                 theSelection.MoveRight(WordRoot.WdUnits.wdCharacter);  // and move beyond it
                 LineCounter++;
                 if (LineCounter % 10 == 0)
                 {
                     txtLineCount.Text = LineCounter.ToString();  // Mark progress
                     progressBar1.Value = LineCounter * WordCount;
                 }
                 Application.DoEvents();
              }
             
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
