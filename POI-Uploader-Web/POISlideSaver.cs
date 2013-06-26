using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using POILibCommunication;
using System.Threading;
using System.IO;

namespace POI_Uploader_Web
{
    class POISlideSaver
    {

        int pptID;
        string folderPath;
        POIPresentation presentation;
        string name;
        string presentor;
        
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
        }
        public  void saveSlideImageToPresentation(int slideIndex)
        {
            POIStaticSlide slide = new POIStaticSlide(slideIndex, presentation);
            presentation.Insert(slide);

            //Upload the image to the content server
            string savedFileName = Path.Combine(FolderPath, slideIndex.ToString() + ".PNG");
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

            //Upload to the content server
            POIContentServerHelper.uploadContent(presentation.PresID, fileName);

            POIGlobalVar.POIDebugLog("presID of slides is" + presentation.PresID);
        }
   
    }
}
