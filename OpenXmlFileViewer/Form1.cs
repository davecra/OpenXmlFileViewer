using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Packaging;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Threading;

namespace OpenXmlFileViewer
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// DECLARATIONS
        /// </summary>
        private TreeNode MobjPreviousNode = null;
        private bool MbolInSelect = false;
        private string MstrPath = "";
        private int MintLastFindPos = 0;
        private FindDialog MobjFd = null;

        /// <summary>
        /// CTOR
        /// </summary>
        /// <param name="PstrArgs"></param>
        public Form1(string[] PstrArgs)
        {
            InitializeComponent();
/*            lineNumberTextBox1.TextChanged += () =>
            {
                toolStripButton3.Enabled = true;
            };*/
            this.Text = "OpenXml File Viewer";
            if (PstrArgs.Length > 0)
            {
                MstrPath = PstrArgs[0];
                openFile();
            }
            removeTabPages();
            webBrowser1.Navigate("about:blank");
        }

        /// <summary>
        /// Removes/hides all the tab pages from the control
        /// </summary>
        private void removeTabPages()
        {
            // hide all the tags
            foreach (TabPage LobjPage in tabControl1.TabPages)
                tabControl1.TabPages.Remove(LobjPage);
        }

        /// <summary>
        /// The user clicked the OPEN button
        /// </summary>
        /// <param name="PobjSender"></param>
        /// <param name="PobjEventArgs"></param>
        private void toolStripButton1_Click(object PobjSender, EventArgs PobjEventArgs)
        {
            OpenFileDialog LobjOfd = new OpenFileDialog();
            LobjOfd.Filter = "All|*.doc*;*.xls*;*.ppt*|Word Documents|*.doc*|Excel Workbooks|*.xls*|PowerPoint Presentations|*.ppt*";
            if (LobjOfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                MstrPath = LobjOfd.FileName;
                openFile();
            }
        }

        /// <summary>
        /// CORE - here is where we open the file denoted by the MstrPath value
        /// </summary>
        private void openFile()
        {
            toolStripButton1.Enabled = false;
            toolStripButton2.Enabled = true;
            this.Text = "[" + new FileInfo(MstrPath).Name + "]";
            // open the package
            using (ZipPackage LobjZip = (ZipPackage)ZipPackage.Open(MstrPath, FileMode.Open, FileAccess.Read))
            {
                // setup the root node
                TreeNode LobjRoot = treeView1.Nodes.Add("/", "/");
                // read all the parts
                foreach (ZipPackagePart LobjPart in LobjZip.GetParts())
                {
                    // build a path string and...
                    int LintLen = LobjPart.Uri.OriginalString.LastIndexOf("/");
                    string LstrKey = LobjPart.Uri.OriginalString;
                    string LstrPath = LstrKey.Substring(0, LintLen);
                    string LstrName = LstrKey.Substring(LintLen + 1);
                    // set the parent, then
                    TreeNode LobjParent = LobjRoot;
                    if (LstrPath.Length != 0)
                        LobjParent = FindClosestNode(LstrPath);
                    // add the node to the tree control
                    LobjParent.Nodes.Add(LstrKey, LstrName);
                }
            }
            MobjPreviousNode = treeView1.Nodes[0];
            treeView1.Nodes[0].Expand();
        }

        /// <summary>
        /// Locates the closest node to the given key
        /// </summary>
        /// <param name="PstrKey"></param>
        /// <returns></returns>
        private TreeNode FindClosestNode(string PstrKey)
        {
            // get the parts of the path
            string[] LstrParts = PstrKey.Split(new string[1] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            string LstrPath = "";
            // grab the root node
            TreeNode LobjLastNode = treeView1.Nodes[0];
            // search through all the parts
            foreach (string LstrPart in LstrParts)
            {
                // build th path
                LstrPath += "/" + LstrPart;
                // get the node with that path
                TreeNode LobjNode = NodeWithPath(treeView1.Nodes[0], LstrPath);
                if (LobjNode != null)
                    LobjLastNode = LobjNode;
                else
                    LobjLastNode = LobjLastNode.Nodes.Add(LstrPath, LstrPart);
            }
            // return the found node
            return LobjLastNode;
        }

        /// <summary>
        /// Get the node with the given path
        /// </summary>
        /// <param name="PobjNode"></param>
        /// <param name="PstrPath"></param>
        /// <returns></returns>
        private TreeNode NodeWithPath(TreeNode PobjNode, string PstrPath)
        {
            TreeNode LobjRetVal = null;
            foreach (TreeNode LobjNode in PobjNode.Nodes)
            {
                LobjRetVal = NodeWithPath(LobjNode, PstrPath);
                if(LobjRetVal != null)
                    return LobjRetVal;
            }
            string LstrNodePath = PobjNode.FullPath.Substring(1).Replace("\\","/");
            if (LstrNodePath == PstrPath)
                return PobjNode;
            else
                return null;
        }

        /// <summary>
        /// Does the key for the given node exist
        /// </summary>
        /// <param name="PobjNode"></param>
        /// <param name="PstrKey"></param>
        /// <returns></returns>
        private bool HasKey(TreeNode PobjNode, string PstrKey)
        {
            foreach (TreeNode LobjNode in PobjNode.Nodes)
                if (HasKey(LobjNode, PstrKey))
                    return true;
            if (PobjNode.Nodes.ContainsKey(PstrKey))
                return true;
            else
                return false;
        }

        /// <summary>
        /// Take the input XML and parse it so that it is
        /// indented in a readable format. We use this so
        /// that we can load the XML into the text editor
        /// in a clean format the user can read
        /// </summary>
        /// <param name="PstrInputXml"></param>
        /// <returns></returns>
        public string FormatXml(string PstrInputXml)
        {
            XmlDocument LobjDocument = new XmlDocument();
            // loas the string into the XML document
            LobjDocument.Load(new StringReader(PstrInputXml));
            MemoryStream LobjMemoryStream = new MemoryStream();
            StringBuilder LobjStringBuilder = null;
            // open the memory stream for input
            using (XmlTextWriter LobjWriter = new XmlTextWriter(LobjMemoryStream, Encoding.UTF8))
            {
                LobjWriter.Formatting = Formatting.Indented;
                // input the XML Document into the memory stream - formatted
                LobjDocument.Save(LobjWriter);
                LobjMemoryStream.Position = 0;
                // red the memory stram into the string builder
                StreamReader sr = new StreamReader(LobjMemoryStream);
                sr.BaseStream.Position = 0;
                LobjStringBuilder = new StringBuilder(sr.ReadToEnd());
            }
            // return the formatted - indented string text
            return LobjStringBuilder.ToString();
        }

        /// <summary>
        /// The user clicked CLOSE
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            // clean up everything
            treeView1.Nodes.Clear();
            webBrowser1.DocumentText = "";
            lineNumberTextBox1.Text = "";
            this.Text = "OpenXml File Viewer";
            label1.Text = "[path]";
            toolStripButton1.Enabled = true;
            toolStripButton2.Enabled = false;
            toolStripButton3.Enabled = false;
            toolStripButton4.Enabled = false;
        }

        /// <summary>
        /// The user clicked a node inthe TreeView
        /// </summary>
        /// <param name="PobjSender"></param>
        /// <param name="PobjEventArgs"></param>
        private void treeView1_AfterSelect_1(object PobjSender, TreeViewEventArgs PobjEventArgs)
        {
            if (MbolInSelect)
                return;

            // is the document dirty? And are we on the Edit tab
            if (toolStripButton3.Enabled && tabControl1.SelectedTab == tabPage2)
            {
                DialogResult LobjDr = MessageBox.Show("Are you sure you want to switch? \n\n" +
                                                      "The currently loaded part [" + label1.Text + "] has not been saved.",
                                                      "Save Loaded Part",
                                                      MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (LobjDr == DialogResult.Yes)
                {
                    toolStripButton3_Click(null, null);
                }
                else if (LobjDr == System.Windows.Forms.DialogResult.Cancel)
                {
                    MbolInSelect = true;
                    treeView1.SelectedNode = MobjPreviousNode;
                    MbolInSelect = false;
                    return;
                }
            }

            // reset the form
            MobjPreviousNode.BackColor = SystemColors.Window;
            MobjPreviousNode = PobjEventArgs.Node;
            MobjPreviousNode.BackColor = SystemColors.MenuHighlight;
            lineNumberTextBox1.Text = "";
            webBrowser1.DocumentText = "";
            toolStripButton3.Enabled = false;
            toolStripButton4.Enabled = false;
            label1.Text = PobjEventArgs.Node.FullPath.Substring(1).Replace("\\", "/");
            toolStripButton5.Enabled = true; // always allow export

            // determine the PART type - if it is not VML or XML, then do not try to
            // read it.
            if (!PobjEventArgs.Node.FullPath.ToLower().EndsWith(".xml") && !PobjEventArgs.Node.FullPath.ToLower().EndsWith(".rels") &&
                !PobjEventArgs.Node.FullPath.ToLower().EndsWith(".vml"))
            {
                // hide the text panes - since this part cannot be shown
                removeTabPages();

                // if the type is an image, then we will show it
                if (PobjEventArgs.Node.FullPath.ToLower().EndsWith("jpeg") ||
                   PobjEventArgs.Node.FullPath.ToLower().EndsWith("jpg") ||
                   PobjEventArgs.Node.FullPath.ToLower().EndsWith("png") ||
                   PobjEventArgs.Node.FullPath.ToLower().EndsWith("bmp") ||
                   PobjEventArgs.Node.FullPath.ToLower().EndsWith("wmf") ||
                   PobjEventArgs.Node.FullPath.ToLower().EndsWith("emf"))
                {
                    tabControl1.TabPages.Add(tabPage3);
                    tabPage3.Select();
                    using (ZipPackage LobjZip = (ZipPackage)ZipPackage.Open(MstrPath, FileMode.Open, FileAccess.Read))
                    {
                        // get the URI for th part
                        string LstrUri = PobjEventArgs.Node.FullPath.Substring(1).Replace("\\", "/");
                        // grab the part
                        ZipPackagePart LobjPart = (ZipPackagePart)LobjZip.GetPart(new Uri(LstrUri, UriKind.Relative));
                        Stream LobjBaseStream = LobjPart.GetStream(FileMode.Open, FileAccess.Read);
                        pictureBox1.Image = new Bitmap(LobjBaseStream);
                        LobjBaseStream.Close();
                    }
                }
                return;
            }
            else
            {
                // show the text panes
                removeTabPages();
                tabControl1.TabPages.Add(tabPage1);
                tabControl1.TabPages.Add(tabPage2);
            }

            try
            {
                // open the part to read the XML
                using (ZipPackage LobjZip = (ZipPackage)ZipPackage.Open(MstrPath, FileMode.Open, FileAccess.Read))
                {
                    // get the URI for th part
                    string LstrUri = PobjEventArgs.Node.FullPath.Substring(1).Replace("\\", "/");
                    // grab the part
                    ZipPackagePart LobjPart = (ZipPackagePart)LobjZip.GetPart(new Uri(LstrUri, UriKind.Relative));
                    Stream LobjBaseStream = LobjPart.GetStream(FileMode.Open, FileAccess.Read);
                    MemoryStream LobjMemoryStream = new MemoryStream();
                    LobjBaseStream.CopyTo(LobjMemoryStream);
                    LobjBaseStream.Close();
                    // load the stream into a string
                    LobjMemoryStream.Position = 0;
                    string LstrXml = new StreamReader(LobjMemoryStream, Encoding.UTF8).ReadToEnd();
                    webBrowser1.DocumentText = LstrXml;
                    LobjMemoryStream.Position = 0;
                    // format the string
                    lineNumberTextBox1.Text = FormatXml(LstrXml);
                    //Highlight();
                }
                toolStripButton3.Enabled = false;
                toolStripButton4.Enabled = true;
                this.Refresh();
            }
            catch { }
        }

        /// <summary>
        /// Highlight all the text
        /// </summary>
        private void Highlight()
        {
            //reset any exisiting formatting
            lineNumberTextBox1.DisableUpdates = true;
            lineNumberTextBox1.RTB.SelectAll();
            lineNumberTextBox1.RTB.SelectionBackColor = Color.White;
            lineNumberTextBox1.RTB.SelectionColor = Color.Black;
            lineNumberTextBox1.RTB.DeselectAll();
            setFormatting("<", Color.Blue);
            setFormatting(">", Color.Blue);
            setFormatting("\"", Color.Red);

            lineNumberTextBox1.DisableUpdates = false;
        }

        /// <summary>
        /// Formats the text RTF
        /// </summary>
        /// <param name="PstrTag"></param>
        /// <param name="LobjColor"></param>
        private void setFormatting(string PstrTag, Color LobjColor)
        {
            MatchCollection LobjMatch;
            Regex LobjReg = new Regex(PstrTag, RegexOptions.Compiled);
            LobjMatch = LobjReg.Matches(lineNumberTextBox1.RTB.Text);
            for (int LintIdx = 0; LintIdx < LobjMatch.Count - 1; LintIdx += 1)
            {
                lineNumberTextBox1.RTB.Select(LobjMatch[LintIdx].Index, LobjMatch[LintIdx].Length);
                lineNumberTextBox1.RTB.SelectionColor = LobjColor;
            }
        }

        /// <summary>
        /// User selected to SAVE the selected open part
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            // open the package
            using (ZipPackage LobjZip = (ZipPackage)ZipPackage.Open(MstrPath, FileMode.Open, FileAccess.ReadWrite))
            {
                string LstrUri = label1.Text;
                ZipPackagePart LobjPart = (ZipPackagePart)LobjZip.GetPart(new Uri(LstrUri, UriKind.Relative));
                Stream LobjStream = LobjPart.GetStream(FileMode.Open, FileAccess.ReadWrite);
                LobjStream.SetLength(0);
                LobjStream.Flush();
                StreamWriter LobjSw = new StreamWriter(LobjStream);
                LobjSw.Write(lineNumberTextBox1.Text);
                LobjSw.Close();
            }
            toolStripButton3.Enabled = false;
        }

        /// <summary>
        /// User clicked FIND
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            // find dialog
            MobjFd = new FindDialog();
            MobjFd.TopMost = true;
            MobjFd.FindNext += (string PstrItem) =>
            {
                int LintIdx = lineNumberTextBox1.Text.IndexOf(PstrItem, MintLastFindPos);
                int LintLen = PstrItem.Length;
                if (LintIdx >= 0)
                {
                    lineNumberTextBox1.Select(LintIdx, LintLen);
                    MintLastFindPos = LintIdx + LintLen;
                }
                else
                {
                    MintLastFindPos = 0;
                }
                lineNumberTextBox1.Select();
                lineNumberTextBox1.Focus();
                this.Focus();
            };
            MobjFd.Reset += () =>
            {
                MintLastFindPos = 0;
            };
            MobjFd.Show();
        }

        /// <summary>
        /// The user pressed a Key, was it F3 - for Find Next
        /// </summary>
        /// <param name="PobjSender"></param>
        /// <param name="PobjKeyArgs"></param>
        private void Form1_KeyUp(object PobjSender, KeyEventArgs PobjKeyArgs)
        {
            if (PobjKeyArgs.KeyCode == Keys.F3 && toolStripButton4.Enabled)
            {
                if (MobjFd != null && MobjFd.Visible)
                    MobjFd.FindNextPoke();
                else
                    toolStripButton4_Click(null, null);
            }
        }

        /// <summary>
        /// User clicked EXPORT to export a part
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            // get the filename for the part
            string LstrFn = label1.Text.Substring(label1.Text.LastIndexOf('/')).Replace("/","");
            // ask the user
            SaveFileDialog LobjSfd = new SaveFileDialog();
            LobjSfd.Filter = "All Files (*.*)|*.*";
            LobjSfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            LobjSfd.FileName = LobjSfd.InitialDirectory + "\\" + LstrFn;
            if(LobjSfd.ShowDialog() == System.Windows.Forms.DialogResult.Cancel)
                return;

            // open the package
            using (ZipPackage LobjZip = (ZipPackage)ZipPackage.Open(MstrPath, FileMode.Open, FileAccess.Read))
            {
                if (new FileInfo(LobjSfd.FileName).Exists)
                    new FileInfo(LobjSfd.FileName).Delete();

                // get the uri
                string LstrUri = label1.Text;
                // grab the part
                ZipPackagePart LobjPart = (ZipPackagePart)LobjZip.GetPart(new Uri(LstrUri, UriKind.Relative));
                // write the part to disk
                StreamReader LobjSr = new StreamReader(LobjPart.GetStream(FileMode.Open, FileAccess.Read));
                BinaryWriter LobjSw = new BinaryWriter(new FileInfo(LobjSfd.FileName).Create());
                while(!LobjSr.EndOfStream)
                    LobjSw.Write(LobjSr.Read());
                LobjSw.Close();
                LobjSr.Close();
            }
        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
    }
}
