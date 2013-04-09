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
        static int slideNumberCounter = 0;

        int pptID;
        string folderPath;
        POIPresentation presentation;
        string name;
        string presentor;
        
        public string FolderPath
        {
            get { return folderPath; }
        }
        public POISlideSaver(string presName, string presPresentor)
        {
            
            //pptID = Properties.Settings.Default.SlideNumberCounter;
            //Properties.Settings.Default.SlideNumberCounter++;
            //Properties.Settings.Default.Save();

            pptID = slideNumberCounter++;

            folderPath = Path.Combine(POIArchive.ArchiveHome, pptID.ToString()); 
            Directory.CreateDirectory(folderPath);

            name = presName;
            presentor = presPresentor;

            presentation = new POIPresentation(pptID, name, presentor);
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

            Console.WriteLine("presID of slides is" + presentation.PresID);
        }
   
    }
}
