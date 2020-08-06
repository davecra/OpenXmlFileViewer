using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OpenXmlFileViewer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] PstrArgs)
        {
            // attempt to register the types in the registry
            try
            {
                // Register Word Extensions
                RegisterKey("Word.Document.12");
                RegisterKey("Word.DocumentMacroEnabled.12");
                RegisterKey("Word.Template.12");
                RegisterKey("Word.TemplateMacroEnabled.12");
                // Register Excel Extensions
                RegisterKey("Excel.Sheet.12");
                RegisterKey("Excel.SheetMacroEnabled.12");
                RegisterKey("Excel.Template.8");
                // Register PowerPoint Extensions
                RegisterKey("PowerPoint.Show.12");
                RegisterKey("PowerPoint.ShowMacroEnabled.12");
                RegisterKey("PowerPoint.Template.12");
                RegisterKey("PowerPoint.TemplateMacroEnabled.12");
            }
            catch (Exception LobjEx)
            {
                //MessageBox.Show("Unable to register");
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(PstrArgs));
        }

        /// <summary>
        /// Write the registry keys
        /// </summary>
        /// <param name="PstrKey"></param>
        static void RegisterKey(string PstrKey)
        {
            RegistryKey LobjKey = Registry.ClassesRoot.OpenSubKey(PstrKey + "\\shell", true);
            if (LobjKey != null)
            {
                LobjKey = LobjKey.CreateSubKey("OpenXml");
                LobjKey = LobjKey.CreateSubKey("command");
                LobjKey.SetValue("", Application.ExecutablePath + " \"%1\"");
            }
        }
    }
}
