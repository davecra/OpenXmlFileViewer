using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OpenXmlFileViewer
{
    public partial class FindDialog : Form
    {
        public delegate void FindNextHandler(string item);
        public event FindNextHandler FindNext;
        public delegate void ResetHandler();
        public event ResetHandler Reset;
        public FindDialog()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(FindNext != null)
                FindNext(textBox1.Text);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (FindNext != null)
                    FindNext(textBox1.Text);
            }
        }

        public void FindNextPoke()
        {
            if (FindNext != null)
                FindNext(textBox1.Text);
        }

        private void FindDialog_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if(Reset != null)
                Reset();
        }
    }
}
