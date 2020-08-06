using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;

namespace OpenXmlFileViewer
{
    /// <summary>
    /// CUSTOM USER CONTROL
    /// </summary>
    public partial class LineNumberTextBox : UserControl
    {
        /// <summary>
        /// DECLARATIONS
        /// </summary>
        public delegate void ViewChangeHandler(int first, int last);
        public event ViewChangeHandler ViewChange;
        public delegate void TextChangeHandler();
        public event TextChangeHandler TextChanged;
        private System.Windows.Forms.Timer tmrWheel = new System.Windows.Forms.Timer();
        private int MintLastWheelLine = 0;
        private bool MbolDisbaleUpdates = false;

        /// <summary>
        /// CTOR
        /// </summary>
        public LineNumberTextBox()
        {
            InitializeComponent();
            label1.Text = "1";
            richTextBox1.Text = "";
            richTextBox1.MouseWheel += new MouseEventHandler(richTextBox1_MouseWheel);
            tmrWheel.Interval = 1000;
            // hook the timer event
            tmrWheel.Tick += (PobjO, PobjE) =>
            {
                Point LobjPos = new Point(0, 0);
                LobjPos.X = ClientRectangle.Width;
                LobjPos.Y = ClientRectangle.Height;
                int LintLastIndex = richTextBox1.GetCharIndexFromPosition(LobjPos);
                int LintLastLine = richTextBox1.GetLineFromCharIndex(LintLastIndex);
                if (LintLastLine != MintLastWheelLine)
                {
                    tmrWheel.Stop();
                    MintLastWheelLine = LintLastLine;
                    updateNumberLabel();
                }
                else
                {
                    tmrWheel.Stop();
                    updateNumberLabel();
                }
            };
            // hook the text change event
            richTextBox1.TextChanged += (PobjO, PobjE) =>
            {
                if (TextChanged != null)
                    TextChanged();
            };
        }

        /// <summary>
        /// TEXT PROPERTY
        /// </summary>
        public string Text
        {
            get
            {
                return richTextBox1.Text;
            }
            set
            {
                richTextBox1.Text = value;
            }
        }

        /// <summary>
        /// RTF PROPERTY
        /// </summary>
        public string Rtf
        {
            get
            {
                return richTextBox1.Rtf;
            }
            set
            {
                richTextBox1.Rtf = value;
            }
        }

        /// <summary>
        /// RICH TEXTBOX REFERENCE PROPERTY
        /// </summary>
        public RichTextBox RTB
        {
            get
            {
                return richTextBox1;
            }
        }

        /// <summary>
        /// DISABLES UPDATES TO THE CONTROL
        /// </summary>
        public bool DisableUpdates
        {
            get
            {
                return MbolDisbaleUpdates;
            }
            set
            {
                MbolDisbaleUpdates = value;
            }
        }

        /// <summary>
        /// Allows caller to specify a selection
        /// </summary>
        /// <param name="PintPos"></param>
        /// <param name="LintLen"></param>
        public void Select(int PintPos, int LintLen)
        {
            richTextBox1.Select(PintPos, LintLen);
            richTextBox1.ScrollToCaret();
        }

        /// <summary>
        /// Mouse wheel moved
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void richTextBox1_MouseWheel(object sender, MouseEventArgs e)
        {
            tmrWheel.Start();
        }

        /// <summary>
        /// Updates the number lables on the side
        /// </summary>
        private void updateNumberLabel()
        {
            // for large changes we want to disable this until later
            // so we skip it here and redraw line numbers when we
            // have this flag turned off
            if (MbolDisbaleUpdates)
                return;

            label1.Visible = false;
            //label1.Font = new Font(richTextBox1.Font.FontFamily, richTextBox1.Font.Size + 1.019f);
            //we get index of first visible char and number of first visible line
            Point LobjPos = new Point(0, 0);
            int LintFirstIndex = richTextBox1.GetCharIndexFromPosition(LobjPos);
            int LintFirstLine = richTextBox1.GetLineFromCharIndex(LintFirstIndex);

            //now we get index of last visible char and number of last visible line
            LobjPos.X = ClientRectangle.Width;
            LobjPos.Y = ClientRectangle.Height;
            int LintLastIndex = richTextBox1.GetCharIndexFromPosition(LobjPos);
            int LintLastLine = richTextBox1.GetLineFromCharIndex(LintLastIndex);

            //this is point position of last visible char, we'll use its Y value for calculating numberLabel size
            LobjPos = richTextBox1.GetPositionFromCharIndex(LintLastIndex);

            //finally, renumber label
            label1.Text = "";
            for (int LintIdx = LintFirstLine; LintIdx <= LintLastLine + 1; LintIdx++)
            {
                label1.Text += LintIdx + 1 + "\n";
            }

            label1.Visible = true;
            // fire the event
            if (ViewChange != null)
                ViewChange(LintFirstLine, LintLastLine);
        }

        /// <summary>
        /// User scrolled vertically
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void richTextBox1_VScroll(object sender, EventArgs e)
        {
            ////move location of numberLabel for amount of pixels caused by scrollbar
            int LintD = richTextBox1.GetPositionFromCharIndex(0).Y % (richTextBox1.Font.Height + 1);
            label1.Location = new Point(0, LintD);
            tmrWheel.Start();
        }

        /// <summary>
        /// User updated text - update line numbers
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            updateNumberLabel();
        }

        /// <summary>
        /// User changed the size of the form, update
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void richTextBox1_Resize(object sender, EventArgs e)
        {
            richTextBox1_VScroll(null, null);
        }

        /// <summary>
        /// User changed the font size
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void richTextBox1_FontChanged(object sender, EventArgs e)
        {
            updateNumberLabel();
            richTextBox1_VScroll(null, null);
        }
    }
}
