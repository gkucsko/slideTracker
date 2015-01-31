using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new SlideTrackerRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace SlideTracker
{
    [ComVisible(true)]
    public class SlideTrackerRibbon : Office.IRibbonExtensibility
    {
        bool startup = false; // starts as false. after initializing will be true. for setting default options
        bool displayStopButton = false; //should we display the stop button (true) or broadcast button (false)
        bool displayOptionsGroup = false; //is the options group displayed
        private Office.IRibbonUI ribbon; //the ribbon object

        public SlideTrackerRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SlideTracker.SlideTrackerRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public bool IsStopButtonVisible(Office.IRibbonControl control)
        {
            return displayStopButton;
        }

        public bool IsExportButtonVisible(Office.IRibbonControl control)
        {
            return !displayStopButton;
        }

        public bool DisplayOptionsGroup(Office.IRibbonControl control)
        {
            return displayOptionsGroup;
        }

        public bool OptionsVisible(Office.IRibbonControl contro)
        {
            return !displayOptionsGroup;
        }

        public bool OptioinsNotVisible(Office.IRibbonControl control)
        {
            return displayOptionsGroup;
        }

        public void ToggleDisplay(Office.IRibbonControl control)
        {
            displayOptionsGroup = !displayOptionsGroup;
            this.ribbon.InvalidateControl("OptionsGroup");
            GetToggleDisplayLabel(control);
        }

        public string GetToggleDisplayLabel(Office.IRibbonControl control)
        {
            string ret;
            if (displayOptionsGroup)
            {
                ret = "Hide Options";
            }
            else
            {
                ret = "Show Options";
            }
            this.ribbon.InvalidateControl("HideOptionsButton");
            return ret;
        }
        
        public void OnExportButton(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.uploadSuccess = true;
            Globals.ThisAddIn.Application.ActivePresentation.Export(Globals.ThisAddIn.SlideDir, Globals.ThisAddIn.fmt);
            if (Globals.ThisAddIn.allowDownload)
            {
                Globals.ThisAddIn.Application.ActivePresentation.ExportAsFixedFormat(
                    Globals.ThisAddIn.SlideDir + "/presentation.pdf", PPT.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            System.Windows.Forms.Form progressForm = new System.Windows.Forms.Form();
            try
            {
                System.Windows.Forms.Label lab = new System.Windows.Forms.Label();
                progressForm.Size = new System.Drawing.Size(350, 150);
                progressForm.Text = "Uploade Progress";
                lab.Text = "Contacting server... This may take a moment.";
                lab.Font = new System.Drawing.Font("Arial", 12);
                lab.Size = new System.Drawing.Size(340, 140);
                progressForm.Controls.Add(lab);
                progressForm.Show();
                progressForm.Update();
                string resp = Globals.ThisAddIn.CreateRemotePresentation();
                lab.Text = "uploading remote presentation...";
                progressForm.Update();
                string resp2 = Globals.ThisAddIn.UploadRemotePresentation();
                //progressForm.Close();
                System.Windows.Forms.MessageBox.Show("Done!");
                displayStopButton = true;
                this.ribbon.InvalidateControl("BroadcastButton"); //updates the display for this control
                this.ribbon.InvalidateControl("StopBroadcast"); //update display
            }
            catch
            {
                Globals.ThisAddIn.uploadSuccess = false;
                System.Windows.Forms.MessageBox.Show("Problem communicating with server. Check internet connection and try again");
                //progressForm.Close();
            }
            finally
            {
                if (!progressForm.IsDisposed) { progressForm.Close(); }
            }

        }

        public void OnStopBroadcast(Office.IRibbonControl control)
        {
            //gets called when the StopBroadcast button is pressed
            //delete remote pres, delete all slide files in slideDir, update button
            Globals.ThisAddIn.DeleteRemotePresentation();
            DirectoryInfo dirInfo = new DirectoryInfo(Globals.ThisAddIn.SlideDir);
            foreach(FileInfo fi in dirInfo.GetFiles("*." + Globals.ThisAddIn.fmt)) //dont delete log file
            {
                fi.Delete();
            }
            //now delete the pdf file (if exists)
            foreach (FileInfo fi in dirInfo.GetFiles("*.pdf"))
            {
                fi.Delete();
            }
            displayStopButton = false;
            this.ribbon.InvalidateControl("BroadcastButton");
            this.ribbon.InvalidateControl("StopBroadcast");
        }

        public void OnAllowDownload(Office.IRibbonControl control, bool isClicked)
        {
            //gets called when the AllowDownload button is checked/unchecked
            Globals.ThisAddIn.allowDownload = isClicked;
        }

        public void OnDropDownShowIP(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            Globals.ThisAddIn.showOnAll = ("all" == selectedId);
        }

        public string GetSelectedShowIP(Office.IRibbonControl control)
        {
            //set default dropdown menu to "all"
            //this is a hack and will break if we change the order of things
            //relies on the fact that this one loads before the next dropdown menu
            if (startup)
            {
                return control.Id;
            }
            else
            {
                return "all";
            }
        }

        public void OnBannerLocation(Office.IRibbonControl control, string selectedID, int selectedIndex)
        {
            float width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - (float)Globals.ThisAddIn.width;
            float height = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - (float)Globals.ThisAddIn.height;
            float offset = 8;
            switch (selectedIndex)
            {
                case 0: // BL
                    Globals.ThisAddIn.left = offset;
                    Globals.ThisAddIn.top = height - offset;
                    break;
                case 1: //BR
                    Globals.ThisAddIn.left = width - offset;
                    Globals.ThisAddIn.top = height - offset;
                    break;
                case 2: //TL
                    Globals.ThisAddIn.left = offset;
                    Globals.ThisAddIn.top = offset;
                    break;
                case 3: //TR
                    Globals.ThisAddIn.left = width - offset;
                    Globals.ThisAddIn.top = offset;
                    break;
            }
        }

        public string GetSelectedShowBanner(Office.IRibbonControl control)
        //this is a hack. relies on the fact that this downdown loads second
        {
            if (startup)
            {
                return control.Id;
            }
            else
            {
                startup = true;
                //terrible hack to hard code the corret values at startup
                float width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - (float)Globals.ThisAddIn.width;
                float height = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - (float)Globals.ThisAddIn.height;
                float offset = 8;
                Globals.ThisAddIn.left = width - offset;
                Globals.ThisAddIn.top = offset;

                return "TR";
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
