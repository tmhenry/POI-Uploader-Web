using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using POILibCommunication;
namespace POI_Uploader_Web
{
    class POIFileReader
    {
        string fileName;
        byte[] buffer;
        public POIFileReader(string fileNameFromTextBox)
        {
            fileName = fileNameFromTextBox;
            getByteFromFile();

        }

        private void getByteFromFile()
        {
            FileStream poiFile = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                            BinaryReader br = new BinaryReader(poiFile);
            try
            {
             

                int length = (int)poiFile.Length;

                buffer = new byte[length];



                br.Read(buffer, 0, length);
            }
            finally
            {
                poiFile.Close();
                br.Close();
            }
        }
        public void GetImageAndAnimationFromFile()
        {
            int offset = 0;
            POIPresentation presentationForRead = new POIPresentation();
            presentationForRead.deserialize(buffer, ref offset);
            POIGlobalVar.POIDebugLog("Number of Slide is " + presentationForRead.Count);
        }
        
        
    }
}
