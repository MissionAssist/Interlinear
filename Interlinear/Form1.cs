using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Common;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private Word.Application wrdLegacyDoc;
 
        public Form1()
        {
            InitializeComponent();
        }

        private void btnGetLegacyFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK )
            {
                txtLegacy.Text = openFileDialog1.FileName;
                btnSegmentLegacy.Enabled = true;

        };
        }

        private void btnSegmentLegacy_Click(object sender, EventArgs e)
        {
            wrdLegacyDoc = new Word.Application();
            wrdLegacyDoc.Documents.Open(txtLegacy.Text);
            wrdLegacyDoc.Visible = true;
            wrdLegacyDoc.ActiveWindow.View.ReadingLayout = false;  // Make sure we are in edit mode
            wrdLegacyDoc.Documents[1].Activate();
            //wrdLegacyDoc.ActiveDocument.ReadOnlyRecommended = false;
            //wrdLegacyDoc.ActiveDocument.Select();

            CleanWordText(wrdLegacyDoc); // delete the shape

            QuitWord();
        }
        private void CleanWordText(Microsoft.Office.Interop.Word.Application theDoc)
        {
         
            theDoc.Selection.HomeKey(Word.WdUnits.wdStory, Word.WdMovementType.wdMove);

            for (int ShapeCounter =1; ShapeCounter <= theDoc.ActiveDocument.Shapes.Count; ShapeCounter++)
            {
                
                if (theDoc.ActiveDocument.Shapes[ShapeCounter].Type == Office.MsoShapeType.msoTextBox)
                {
                    theDoc.ActiveDocument.Shapes[ShapeCounter].ConvertToInlineShape();
                    theDoc.ActiveDocument.Shapes[ShapeCounter].Delete();
                }
            }

        }
         private void QuitWord()
            {
                if (wrdLegacyDoc != null)
                {
                    object missing = Type.Missing;
                    wrdLegacyDoc.Quit(ref missing, ref missing, ref missing);
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    wrdLegacyDoc = null;

                }

            }
 
    }

};
