using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace CharacterCounter
{
    public partial class EncodingForm : Form
    {
        public EncodingForm()
        {
            InitializeComponent();
        }

        private void EncodingForm_Load(object sender, EventArgs e)
        {
            // Load the encodings
            int Counter = 0;
            string theEncodingName = ((Form1)Owner).GetEncoding();
            if (theEncodingName == "")
            {
                theEncodingName = "Western European (Windows)";  // Equivalent to ANSI
            }
            foreach (EncodingInfo theEncodingInfo in Encoding.GetEncodings())
            {
                EncodingListBox.Items.Add(theEncodingInfo.DisplayName);
                if (theEncodingInfo.DisplayName == theEncodingName)
                {
                    EncodingListBox.SelectedItem = EncodingListBox.Items[Counter];
                }
                Counter++;
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Form1 theOwner = (Form1)this.Owner;
            theOwner.SetEncoding(EncodingListBox.SelectedItem.ToString());
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

  
    }
}
