using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Net; //for HTTPWebRequest
using System.IO; //for Stream
using System.Drawing;
using System.Net.NetworkInformation;

namespace SlideTracker
{
    public partial class ThisAddIn
    {
        public string SlideDir = @"C:\"; //won't get used. assigned a random temp directory upon exporting
        public string fmt = "png"; //export the slides to
        public string postURL = "http://www.slidetracker.org/api/v1/presentations"; // production server
        //public string postURL = "http://54.208.192.158/api/v1/presentations"; //dev server
        private string userAgent = ""; //not really used. could be anything. for future development
        public string privateHash = "foobar"; //will get set when creating remote pres
        public string pres_ID = "123"; //will be overwritten by info from server
        public string userName = ""; //will be taken from mac address of computer
        private string[] textBoxIds; //ids for text boxes with ip address
        private string[] rectangleIds; //for box behind text
        public bool showOnAll = true; //show banner on all slides? first slide?
        public bool allowDownload = false;// allow others to download pdf from website
        public bool debug = false; //write stuff to log file
        public float left = 0; // points away from left edge of slide for IP text box
        public float top = 0; // points away from top edge of slide for IP text box
        public float width = 85; // width in points of text box
        public float height = 30; // height in points of text box
        public bool uploadSuccess = false; // set to true upon success in upload
        private bool failedDuringPresentation = false; //will be set to true if things fail during pres
        private string logFile = @""; //file to write log notes 
        public int maxClients = 0; //max number of viewers ever

        #region Slide Show Functions
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationNewSlide +=
                new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            this.Application.SlideShowBegin +=
                new PowerPoint.EApplication_SlideShowBeginEventHandler(Application_SlideShowBegin);
            this.Application.SlideShowEnd +=
                new PowerPoint.EApplication_SlideShowEndEventHandler(Application_SlideEnd);
            this.Application.SlideShowNextSlide +=
                new PowerPoint.EApplication_SlideShowNextSlideEventHandler(Application_SlideShowNextSlide);
            GenerateTempDir();
            this.logFile = this.SlideDir + "\\log.txt";
            System.IO.File.Delete(this.logFile);
            File.Create(this.logFile).Dispose(); //makes an empty file.
            if (this.debug) { logWrite("Starting up"); }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                DeleteRemotePresentation();
            }
            catch
            {
                if (this.debug) { logWrite("problems deleting remote presentation"); }
            }
            if (this.debug) { logWrite("deleting remote presentation"); }
        }

        public void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
        }

        void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            if (this.uploadSuccess)
            {
                AddBannerToAll("slidetracker.org" + System.Environment.NewLine + "# " + this.pres_ID);
                if (this.debug) { logWrite("Started Show "); }
            }

        }

        void Application_SlideEnd(PowerPoint.Presentation Pr)
        {
            if (this.textBoxIds.Length >0)
            {
                DeleteBannerFromAll();
                if (this.debug) { logWrite("ending Show "); }
            }
            if (this.maxClients>=0 && this.uploadSuccess)
            {
                SlideTrackerRibbon.ribbon1.InvalidateControl("NumViewers");
            }
            if (this.failedDuringPresentation)
            {
                System.Windows.Forms.MessageBox.Show("Contact with server lost during presentation.");
            }
        }

        void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            if (!this.uploadSuccess) { return; }
            int curSlide = Wn.View.CurrentShowPosition;
            if (this.debug) { logWrite(("went to next slide " + curSlide)); }
            UpdateCurrentSlide(curSlide);
        }
        #endregion

        #region Communication with server
        public string CreateRemotePresentation()
        {
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            if (nics.Length > 0)
            {
                this.userName = nics[0].GetPhysicalAddress().ToString();
            }
            else
            {
                this.userName = "Gorg";
                if (this.debug) { logWrite("messed up mac address. assigning some other username"); }
            }
       
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            postParameters.Add("pres_ID", this.pres_ID);
            postParameters.Add("creator", this.userName);
            postParameters.Add("n_slides", "" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Range().Count);
            //leaving out optional operation string in MultipartForDataPost. default operation = "POST"
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL, this.userAgent, postParameters);

            //now process response
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            webResponse.Close();
            this.pres_ID = GetInfoFromJson(fullResponse,"pres_ID");
            this.privateHash = GetInfoFromJson(fullResponse,"passHash");
            if (this.debug)
            {
                logWrite("pres_ID: " + this.pres_ID);
                logWrite("privateHash: " + this.privateHash); //needed for all future communication
            }
            return fullResponse;
        }

        public bool CheckFileRequirements()
        {
            bool allGood = true;
            string[] files = System.IO.Directory.GetFiles(this.SlideDir, "*." + this.fmt);
            Int64 totalSize = 0;
            for (int i = 0; i < files.Length; i++)
            {
                FileInfo fi = new FileInfo(files[i]);
                totalSize += fi.Length;
                if (fi.Length > 2000000)
                {
                    allGood = false;
                    break;
                }
            }
            if (totalSize > 20000000) { allGood = false; }
            string[] pdfFiles = System.IO.Directory.GetFiles(this.SlideDir, "*.pdf");
            if (pdfFiles.Length > 1) { allGood = false; }
            if (pdfFiles.Length > 0)
            {
                FileInfo pdfInfo = new FileInfo(pdfFiles[0]);
                if (pdfInfo.Length > 20000000) { allGood = false; }
            }
            return allGood;
        }
        
        public string UploadRemotePresentation()
        {
            //upload all slides and, if allowed pdf presentation to server
            int count = 1;
            string[] files = System.IO.Directory.GetFiles(this.SlideDir, "*." + this.fmt);
            for (int fileInd = 0; fileInd < files.Length; fileInd++)
            {
                string file = new FileInfo(files[fileInd]).Name;
                string sldNum = System.Text.RegularExpressions.Regex.Match(file, @"\d+").Value;
                string resp = UploadRemoteSlide(int.Parse(sldNum), file);
                if (String.Compare(resp, "\"upload succeeded!\"") < 0)
                {
                    this.uploadSuccess = false;
                    break;
                }
                if (this.debug)
                {
                    logWrite("uploaded " + file + " response = " + resp);
                }
                count++;
            }
            //upload pdf if wanted
            if (this.allowDownload) //gets set in ribbon
            {
                string presName = "presentation.pdf";
                Dictionary<string, object> postParameters = new Dictionary<string, object>();
                FileStream fs = new FileStream(this.SlideDir + "/" + presName, FileMode.Open, FileAccess.Read);
                byte[] data = new byte[fs.Length];
                fs.Read(data, 0, data.Length);
                fs.Close();
                postParameters.Add("pres", new FormUpload.FileParameter(data, presName, "application/pdf"));
                HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID + "/presentation/",
                 this.userAgent, postParameters); //leaving out optional operation string. defaults to "POST"
            }

            string readyResp = this.MarkAsReady();
            if (this.debug) { logWrite(readyResp); }
            return "done uploading files";
        }

        private string UploadRemoteSlide(int slide_ID, string fileName)
        {
            //upload a single slide to server
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            postParameters.Add("slide_ID", "" + slide_ID);
            FileStream fs = new FileStream(this.SlideDir + "/" + fileName, FileMode.Open, FileAccess.Read);
            if (this.debug) { logWrite("uploading file  " + this.SlideDir + "/" + fileName); }
            byte[] data = new byte[fs.Length];
            fs.Read(data, 0, data.Length);
            fs.Close();
            postParameters.Add("slide", new FormUpload.FileParameter(data, fileName, "image/" + this.fmt));
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID + "/slides/",
                  this.userAgent, postParameters); //leaving out optional operation string. defaults to "POST"

            //now process response
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            webResponse.Close();
            logWrite(fullResponse);
            return fullResponse;
        }

        public string MarkAsReady()
        {
            //turns active to 'true' on the remote server.
            if (!this.uploadSuccess)
            {
                return "";
            }
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            postParameters.Add("n_slides", "" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Range().Count);
            postParameters.Add("cur_slide", "" + 1);
            postParameters.Add("active", "true");
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID,
                  this.userAgent, postParameters, "PUT");
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            webResponse.Close();
            return fullResponse;
        }

        public void UpdateCurrentSlide(int slideNumber)
        {
            //try to update the current slide on the server.
            // if it fails, deletes the ID banners from the presentation
            if (!this.uploadSuccess) { return; } //shouldn't be necessary but doesn't hurt
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            postParameters.Add("n_slides", "" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Range().Count);
            postParameters.Add("cur_slide", "" + slideNumber);
            postParameters.Add("active", "true");
            string fullResponse;
            try
            {
                HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID,
                this.userAgent, postParameters, "PUT");
                StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
                fullResponse = responseReader.ReadToEnd();
                if (this.debug) { logWrite(fullResponse); }
                webResponse.Close();
            }
            catch
            {
                this.uploadSuccess = false; //should prevent this function from ever getting called again
                this.failedDuringPresentation = true;
                DeleteBannerFromAll();
                if (this.debug) { logWrite("Problems on slide " + slideNumber); }
                return;
            }
            //do this statistics here. failing this shouldn't ruin the presentation tracking
            int temp;
            bool parsed = Int32.TryParse(GetInfoFromJson(fullResponse, "clients"), out temp);
            if (parsed && temp > this.maxClients) { this.maxClients = temp; }
            if (this.debug) { logWrite("Current Slide = " + slideNumber + "  number of viewers = " + temp); }
            
        }

        public void DeleteRemotePresentation()
        {
            //deletes the presentation on the remote server. No longer viewable.
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            try
            {
                HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID + "/delete",
                     this.userAgent, postParameters, "PUT");
                StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
                string fullResponse = responseReader.ReadToEnd();
                webResponse.Close();
                if (this.debug) { logWrite("deleteing presentation. response:  " + fullResponse); }
            }
            catch
            {
                if (this.debug) { logWrite("Problems deleting presentation "+  this.pres_ID + "orphan files likely"); }
            }
        }
        #endregion

        private string GetInfoFromJson(string json, string field)
        {
            //hack around actually parsing json. returns the next item (string) after finding field
            string[] separators = { ",", ".", "!", "?", ";", ":", " ", "{", "}" };
            string[] words = json.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = words[i].Replace("\"", ""); //remove quotes
                //logWrite(words[i]);
            }
            int idx = Array.IndexOf(words, field);
            //if (this.debug) { logWrite("found" +  words[idx] + ":  " + words[idx + 1]); }
            return words[idx + 1];
        }

        public static string GetTempDir()
        {
            string tempFolder = System.IO.Path.GetTempPath() + "slideShare_" + System.IO.Path.GetRandomFileName();
            return tempFolder; //won't actually create it. gets created in GenerateTempDir
        }

        public void GenerateTempDir()
        {
            //get full path without extension of temp dir
            string dirName = GetTempDir();
            int lastPeriod = dirName.LastIndexOf(".");
            dirName = dirName.Substring(0, lastPeriod);
            this.SlideDir = dirName;
            System.IO.Directory.CreateDirectory(dirName);
        }

        private void AddBannerToAll(string banner)
        {
            PowerPoint.SlideRange allSlides = this.Application.ActivePresentation.Slides.Range(); //no argument = all slides
            int numSlides = allSlides.Count;
            int numBanners;
            if (this.debug) { logWrite("show on all is " + this.showOnAll); }
            if (this.showOnAll)
            {
                numBanners = numSlides;
            }
            else
            {
                numBanners = 1;
            }

            PowerPoint.Shape textBox;
            String[] shapeIds = new string[numBanners];
            String[] boxIds = new string[numBanners];
            //Office.MsoAutoShapeType tp = Office.MsoAutoShapeType.msoShapeRectangle; //ugly version
            Office.MsoAutoShapeType tp = Office.MsoAutoShapeType.msoShapeRoundedRectangle; //pretty version
            for (int i = 1; i <= numBanners; i++)
            {
                PowerPoint.Shape bx = allSlides[i].Shapes.AddShape(tp, this.left, this.top, this.width, this.height);
                bx.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(37, 37, 37).ToArgb(); //careful, really BGR vals
                bx.Line.ForeColor.RGB = System.Drawing.Color.FromArgb(37, 37, 37).ToArgb();
                boxIds[i - 1] = bx.Name;

                textBox = allSlides[i].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, this.left, this.top, this.width, this.height);
                textBox.TextFrame.TextRange.InsertAfter(banner);
                textBox.TextFrame.TextRange.Font.Size = 10;
                textBox.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                shapeIds[i - 1] = textBox.Name;
            }
            if (this.debug)
            {
                logWrite("adding ip address" + banner);
                logWrite("detected " + shapeIds.Length + " slides");
            }
            this.textBoxIds = shapeIds;
            this.rectangleIds = boxIds;

        }

        private void DeleteBannerFromAll()
        {
            //deletes the text boxes and rectangles. sets the this.rectangledIds and this.texBoxIds to empty
            for (int i = 0; i < this.textBoxIds.Length; i++)
            {
                this.Application.ActivePresentation.Slides[i + 1].Shapes[this.rectangleIds[i]].Delete();
                this.Application.ActivePresentation.Slides[i + 1].Shapes[this.textBoxIds[i]].Delete();
            }
            this.rectangleIds = new string[0];
            this.textBoxIds = new string[0]; 
            if (this.debug) { logWrite("deleted IP address banners"); }
        }
  
        public void logWrite(string msg)
        {
            if (System.IO.File.Exists(this.logFile))
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(this.logFile, true))
                    {
                        file.WriteLine(msg);
                    }
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Problem with log file");
                }

            }

        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
          {
              return new SlideTrackerRibbon();
          }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
