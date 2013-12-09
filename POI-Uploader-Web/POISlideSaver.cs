using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using POILibCommunication;
using System.Threading;
using System.IO;

using System.Web.Script.Serialization;

namespace POI_Uploader_Web
{
    class POISlideSaver
    {

        int pptID;
        string folderPath;
        POIPresentation presentation;
        string name;
        string presentor;
        Dictionary<string, string> keywordDict;
        
        public string FolderPath
        {
            get { return folderPath; }
        }
        public POISlideSaver(string presName, string description, int presId)
        {
            //Register the content to the content server and retrieve its ID
            //pptID = POIWebService.UploadPresentation(presName, description);
            pptID = presId;

            folderPath = Path.Combine(POIArchive.ArchiveHome, pptID.ToString()); 
            Directory.CreateDirectory(folderPath);

            name = presName;
            presentor = description;

            presentation = new POIPresentation(pptID, name, description);
            keywordDict = new Dictionary<string, string>();
        }
        public  void saveSlideImageToPresentation(int slideIndex)
        {
            POIStaticSlide slide = new POIStaticSlide(slideIndex, presentation);
            presentation.Insert(slide);

            //Upload the image to the content server
            string savedFileName = Path.Combine(FolderPath, slideIndex.ToString() + ".PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);
        }

        public void saveCoverPageToPresentation(int slideIndex)
        {
            POIStaticSlide slide = new POIStaticSlide(slideIndex, presentation);
            presentation.Insert(slide);

            //Upload the image to the content server
            string savedFileName = Path.Combine(FolderPath, "cover.PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);

            savedFileName = Path.Combine(FolderPath, slideIndex + ".PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);
        }

        public void saveQuestionPageToPresentation(int slideIndex)
        {
            POIStaticSlide slide = new POIStaticSlide(slideIndex, presentation);
            presentation.Insert(slide);

            //Upload the image to the content server
            string savedFileName = Path.Combine(FolderPath, "question.PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);

            savedFileName = Path.Combine(FolderPath, slideIndex + ".PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);
        }

        public void saveAnswerPageToPresentation(int slideIndex)
        {
            POIStaticSlide slide = new POIStaticSlide(slideIndex, presentation);
            presentation.Insert(slide);

            //Upload the image to the content server
            string savedFileName = Path.Combine(FolderPath, "answer.PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);

            savedFileName = Path.Combine(FolderPath, slideIndex + ".PNG");
            POIContentServerHelper.uploadContent(presentation.PresID, savedFileName);
        }

        public  void saveSlideAnimationToPresentation(int slideIndex, List<int> durationList)
        {
            POIAnimationSlide slide = new POIAnimationSlide(durationList, slideIndex, presentation);
            presentation.Insert(slide);

            string savedAnimationName = Path.Combine(FolderPath, slideIndex.ToString() + ".mp4");
            string savedCoverName = Path.Combine(FolderPath, slideIndex.ToString() + ".PNG");

            POIContentServerHelper.uploadContent(presentation.PresID, savedAnimationName);
            POIContentServerHelper.uploadContent(presentation.PresID, savedCoverName);
        }

        public void saveSlideKewordIntoPresentation(int slideIndex, string keyword)
        {
            try
            {
                keywordDict.Add(pptID + "_" + slideIndex, keyword);
            }
            catch (Exception e)
            {
                POIGlobalVar.POIDebugLog(e);
            }
        }

        public void uploadSlideKeywordsToServer()
        {
            string keywordsName = Path.Combine(folderPath, presentation.PresID + POIGlobalVar.KeywordsFileType);
            POIContentServerHelper.uploadContent(presentation.PresID, keywordsName);
        }
        public  void saveToPOIFile()
        {
            byte[] fileBuffer = new byte[presentation.Size];
            int offset = 0;
            presentation.serialize(fileBuffer, ref offset);

            string fileName = Path.Combine(folderPath, pptID.ToString() + ".POI");

            FileStream writeStream = new FileStream(fileName, FileMode.OpenOrCreate);

            BinaryWriter bw = new BinaryWriter(writeStream);

            bw.Write(fileBuffer);

            bw.Close();

            writeStream.Close();

            //Upload the byte data to the content server
            POIContentServerHelper.uploadContent(presentation.PresID, fileName);
            POIWebService.UploadKeyword(keywordDict);
            POIGlobalVar.POIDebugLog("presID of slides is" + presentation.PresID);

            //Upload the json data to the content server
            try
            {
                POIGlobalVar.POIDebugLog("Uploading POI.json");
                string jsonFn = Path.Combine(folderPath, pptID + ".POI.json");
               
                using (StreamWriter writer = new StreamWriter(jsonFn))
                {
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    writer.Write(js.Serialize(presentation));
                }

                //Upload the .json to the content server
                POIContentServerHelper.uploadContent(presentation.PresID, jsonFn);
            }
            catch (Exception e)
            {
                POIGlobalVar.POIDebugLog("In writing .POI.json: " + e.Message);
            }
            

            //Create an empty session
            //Get a new session ID from the DNS server
            Dictionary<string, string> reqDict = new Dictionary<string, string>();
            reqDict["creator"] = "system_default";
            reqDict["presId"] = presentation.PresID.ToString();
            reqDict["type"] = "public";
            reqDict["status"] = "waiting";
            int sessionId = POIWebService.CreateSession(reqDict);

            POIGlobalVar.POIDebugLog("sessionID is " + sessionId);

            reqDict = new Dictionary<string, string>();
            reqDict["userId"] = "system_default";
            reqDict["sessionId"] = sessionId.ToString();
            POIWebService.EndSession(reqDict);
        }
   
    }
}
