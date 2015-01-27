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
        public string SlideDir = @"C:\";
        public string fmt = "png";
        private string postURL = "http://www.slidetracker.org/api/v1/presentations"; // production server
        //private string postURL = "http://www.dangerzone.elasticbeanstalk.com/api/v1/presentations";
        private string userAgent = "";
        private string pres_ID = "123"; //will be overwritten by info from server
        public string userName = ""; //will be taken from mac address of computer
        private string[] textBoxIds; //ids for text boxes with ip address
        private string[] rectangleIds; //for box behind text
        public bool showOnAll = true; //show banner on all slides? first slide?
        public bool debug = true; //write stuff to log file
        public float left = 0; // points away from left edge of slide for IP text box
        public float top = 0; // points away from top edge of slide for IP text box
        public float width = 85; // width in points of text box
        public float height = 30; // height in points of text box
        public bool uploadSuccess = false;
        private string logFile = @"";

        #region Slide Show Functions
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationNewSlide +=
                new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            this.Application.SlideShowBegin +=
                new PowerPoint.EApplication_SlideShowBeginEventHandler(Application_SlideShowBegin);
            this.Application.SlideShowEnd +=
                new PowerPoint.EApplication_SlideShowEndEventHandler(Application_SlideEnd);
            //Application.SlideShowEnd += Application_SlideEnd;
            this.Application.SlideShowNextSlide +=
                new PowerPoint.EApplication_SlideShowNextSlideEventHandler(Application_SlideShowNextSlide);
            GenerateTempDir();
            this.logFile = this.SlideDir + "\\log.txt";
            System.IO.File.Delete(this.logFile);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            DeleteRemotePresentation();
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
            if (this.uploadSuccess)
            {
                DeleteBannerFromAll();
                if (this.debug) { logWrite("ending Show "); }
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
            this.pres_ID = GetPresIDFromJson(fullResponse);
            return fullResponse;
        }

        public string UploadRemotePresentation()
        {
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
                //lab.Text = "Done with " + count + "of " + files.Length;
                //f.Show();
                if (this.debug)
                {
                    logWrite("uploaded " + file + " response = " + resp);
                }
                count++;
            }
            string readyResp = this.MarkAsReady();
            if (this.debug) { logWrite(readyResp); }
            //f.Close();
            return "done uploading files";
        }

        private string UploadRemoteSlide(int slide_ID, string fileName)
        {
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

        public void DeleteRemotePresentation()
        {
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            //postParameters.Add("key", "N3sN7AiWTFK9XNwSCn7um35joV6OFslL");
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID + "/delete",
                 this.userAgent, postParameters, "PUT");
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            if (this.debug)
            {
                logWrite("deleteing presentation. response:  " + fullResponse);
            }
            webResponse.Close();
        }

        private string GetPresIDFromJson(string json)
        {
            string[] separators = { ",", ".", "!", "?", ";", ":", " ", "{", "}" };
            string[] words = json.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = words[i].Replace("\"", ""); //remove quotes
            }
            int idx = Array.IndexOf(words, "pres_ID");
            return words[idx + 1];
        }

        public static string GetTempDir()
        {
            string tempFolder = System.IO.Path.GetTempPath() + "slideShare_" + System.IO.Path.GetRandomFileName();
            return tempFolder;
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

        public void UpdateCurrentSlide(int slideNumber)
        {
            Dictionary<string, object> postParameters = new Dictionary<string, object>();
            postParameters.Add("n_slides", "" + Globals.ThisAddIn.Application.ActivePresentation.Slides.Range().Count);
            postParameters.Add("cur_slide", "" + slideNumber);
            postParameters.Add("active", "true");
            //postParameters.Add("key", "N3sN7AiWTFK9XNwSCn7um35joV6OFslL");
            HttpWebResponse webResponse = FormUpload.MultipartFormDataPost(this.postURL + "/" + this.pres_ID,
                 this.userAgent, postParameters, "PUT");
            StreamReader responseReader = new StreamReader(webResponse.GetResponseStream());
            string fullResponse = responseReader.ReadToEnd();
            if (this.debug) { logWrite(fullResponse); }
            webResponse.Close();
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
            for (int i = 0; i < this.textBoxIds.Length; i++)
            {
                this.Application.ActivePresentation.Slides[i + 1].Shapes[this.rectangleIds[i]].Delete();
                this.Application.ActivePresentation.Slides[i + 1].Shapes[this.textBoxIds[i]].Delete();
            }
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
