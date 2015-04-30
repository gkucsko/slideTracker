using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PPT = Microsoft.Office.Interop.PowerPoint;
using System.ComponentModel;

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
        public bool displayStopButton = false; //should we display the stop button (true) or broadcast button (false)
        public bool displayOptionsGroup = false; //is the options group displayed
        public bool showRibbon = true; //should the ribbon be shown at all
        private Office.IRibbonUI ribbon; //the ribbon object
        internal static Office.IRibbonUI ribbon1; //for access from other functions
        //private System.Windows.Forms.Form successForm; //form to notify success
        public static TrackerForm tForm; // one trackerForm ribbon for the ribbon
        public SlideTrackerRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SlideTracker.SlideTrackerRibbon.xml");
        }

        #endregion

        
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            int i = Globals.ThisAddIn.CheckVersion(); //1=bad version, 0=good, -1=no connection
            if (i==1) //bad version
            {
                System.Windows.Forms.MessageBox.Show("Your slideTracker version, " +
                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() +
                    " is out of date on no longer compatible. Please visit www.slidetracker.org" +
                    " for the latest version. ","slideTracker Error");
                showRibbon = false;

            }
            this.ribbon = ribbonUI;
            ribbon1 = ribbonUI; // to expose this to globals.ribbons
            Globals.ThisAddIn.ribbon = this;

        }

        #region visibility helpers

        public bool DisplayRibbon(Office.IRibbonControl control) // show/hide the whole ribbon
        {
            return showRibbon;
        }

        public bool IsStopButtonVisible(Office.IRibbonControl control) // show/Hide StopBroadcast button
        {
            return displayStopButton;
        }

        public bool IsExportButtonVisible(Office.IRibbonControl control) // show/hide export button, opposite of stop button
        {
            return !displayStopButton;
        }

        public bool DisplayOptionsGroup(Office.IRibbonControl control) // show/hide options group
        {
            return displayOptionsGroup;
        }

        /*public bool OptioinsNotVisible(Office.IRibbonControl control)
        {
            return displayOptionsGroup;
        }*/

        public void ToggleDisplay(Office.IRibbonControl control)
        {
            displayOptionsGroup = !displayOptionsGroup;
            this.ribbon.InvalidateControl("OptionsGroup");
            GetToggleDisplayLabel(control);
        }

        public string GetToggleDisplayLabel(Office.IRibbonControl control) 
            // text for button to display/hide the options
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
        #endregion
        #region Ribbon Callbacks

        public void OnExportButton(Office.IRibbonControl control) //export to png, make remote pres, upload it. 
        {
            // first check to see that presentation isn't read only and that there is one that is active
            Office.MsoTriState state = Office.MsoTriState.msoTrue;
            try
            {
                state = Globals.ThisAddIn.Application.ActivePresentation.ReadOnly;
            }
            catch {}
            if (state != Office.MsoTriState.msoFalse)
            {
                System.Windows.Forms.MessageBox.Show("Current Presentation is Read only. Please enable editing to proceed with SlideTracker", "Permission Error");
                return;
            }
            tForm = new TrackerForm();
            tForm.SetOKVisible(false);
            tForm.SetCancelVisible(true);
            //now check network connection
            if (!System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable())
            {
                System.Windows.Forms.MessageBox.Show("Cannot connect to internet. Please fix connection and try again", "Connection error");
                return;
            }
            Globals.ThisAddIn.uploadSuccess = true;
            tForm.ChangeLabelText("Exporting Files to " + Globals.ThisAddIn.fmt);
            tForm.InitProgressBar(Globals.ThisAddIn.GetNumSlides());
            tForm.Show();

            Globals.ThisAddIn.MakeLUT();
            Globals.ThisAddIn.Application.ActivePresentation.Export(Globals.ThisAddIn.SlideDir, Globals.ThisAddIn.fmt);
            SlideTrackerRibbon.tForm.Focus();
            Globals.ThisAddIn.DeleteHiddenSlides();
            if (Globals.ThisAddIn.allowDownload)
            {
                Globals.ThisAddIn.Application.ActivePresentation.ExportAsFixedFormat(
                    Globals.ThisAddIn.SlideDir + "/presentation.pdf", PPT.PpFixedFormatType.ppFixedFormatTypePDF);
            }
            try
            {
                if (!Globals.ThisAddIn.CheckFileRequirements())
                {
                    tForm.ChangeLabelText("Sorry, total file size too big for slideTracker.");
                    Globals.ThisAddIn.uploadSuccess = false;
                    //System.Windows.Forms.MessageBox.Show("Sorry, total file size too big for slideTracker.");
                    return;
                }

                string resp = Globals.ThisAddIn.CreateRemotePresentation();
                tForm.ChangeLabelText("uploading remote presentation...");
                string resp2 = Globals.ThisAddIn.UploadRemotePresentation();

                //FIXME: the next if statement will essentially always return true;
                // need to find a way to make it wait for background worker to finish
                // for now, HACK: worker will just change ribbon itself if cancelled/error
                if (!tForm.cancelledForm && Globals.ThisAddIn.uploadSuccess) { displayStopButton = true; }
                UpdateDisplay();
            }
            catch (Exception e)
            {
                if (Globals.ThisAddIn.debug) { Globals.ThisAddIn.logWrite(e.ToString()); }
                Globals.ThisAddIn.uploadSuccess = false;
                System.Windows.Forms.MessageBox.Show("Problem communicating with server. Check internet connection and try again");
                tForm.done = true;
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
            Globals.ThisAddIn.broadcastPresentationName = null;
            UpdateDisplay(); // go back to start broadcast button, remove pres_ID, etc. 
            Globals.ThisAddIn.uploadSuccess = false;
            Globals.ThisAddIn.maxClients = 0;
        }

        public void UpdateDisplay() //update the controls that may get changed
        {
            this.ribbon.InvalidateControl("BroadcastButton"); //updates the display for this control
            this.ribbon.InvalidateControl("StopBroadcast"); //update display
            this.ribbon.InvalidateControl("PresID");
            this.ribbon.InvalidateControl("PresIDLink");
            this.ribbon.InvalidateControl("PresIDGroup");
            this.ribbon.InvalidateControl("NumViewers");
            this.ribbon.InvalidateControl("AllowDownload");
        }

        public void OnAllowDownload(Office.IRibbonControl control, bool isClicked)
        {
            //gets called when the AllowDownload button is checked/unchecked
            Globals.ThisAddIn.allowDownload = isClicked;
        }

        public bool EnableAllowDownload(Office.IRibbonControl control) //callback for clicking "allow Downloads"
        {
            return !Globals.ThisAddIn.uploadSuccess;
        }

        public void OnDropDownShowIP(Office.IRibbonControl control, string selectedId, int selectedIndex)
        // callback for selecting which slides to display tracking ID 
        {
            Globals.ThisAddIn.showOnAll = ("all" == selectedId);
        }

        public string GetSelectedShowIP(Office.IRibbonControl control) //return list item for which slides to show tracking ID
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
        // callback for banner location dropdown
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

        public string GetSelectedShowBanner(Office.IRibbonControl control) //returns list item for where on slide to show tracking ID
        {
            //this is a hack. relies on the fact that this downdown loads second
            if (startup)
            {
                return control.Id;
            }
            else
            {
                startup = true;
                //System.Runtime.InteropServices.COMException
                //terrible hack to hard code the corret values at startup
                float width = 400;
                float height = 200;
                try
                {
                    width = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideWidth - (float)Globals.ThisAddIn.width;
                    height = Globals.ThisAddIn.Application.ActivePresentation.PageSetup.SlideHeight - (float)Globals.ThisAddIn.height;
                }
                catch (System.Runtime.InteropServices.COMException) { }
                float offset = 8;
                Globals.ThisAddIn.left = width - offset;
                Globals.ThisAddIn.top = offset;

                return "TR";
            }
        }

        public string GetPresLink(Office.IRibbonControl control) //returns the link text for presentation
        {
            if (Globals.ThisAddIn.uploadSuccess)
            {
                return System.Environment.NewLine + Globals.ThisAddIn.GetLinkURL();
            }
            else
            {
                return "";
            }
        }

        public void FollowPresLink(Office.IRibbonControl control) //executed when link pressed in ribbon
        {
            if (Globals.ThisAddIn.uploadSuccess)
            {
                System.Diagnostics.Process.Start(Globals.ThisAddIn.GetLinkURL());
            }
        }

        public string GetPresID(Office.IRibbonControl control) //return the text for pres_ID to ribbon
        {
            if (Globals.ThisAddIn.uploadSuccess)
            {
                return "Presentation ID:  " + Globals.ThisAddIn.pres_ID + 
                    System.Environment.NewLine + "   Presentation: " + Globals.ThisAddIn.broadcastPresentationName;
            }
            else
            {
                return "";
            }
        }

        public string GetNumViewers(Office.IRibbonControl control) //return the text for the max num viewers for ribbon
        {
            if (Globals.ThisAddIn.maxClients > 0 && Globals.ThisAddIn.uploadSuccess)
            {
                return "Maximum viewers: " + Globals.ThisAddIn.maxClients;
            }
            else
            {
                return "";
            }
        }

        private void LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e) //callback for clicking link
        {
            System.Diagnostics.Process.Start(Globals.ThisAddIn.GetLinkURL());
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
