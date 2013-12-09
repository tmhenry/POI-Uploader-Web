using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


using System.Threading;
using POILibCommunication;
using System.Diagnostics;

namespace POI_Uploader_Web
{
    public class POIUploadHandler
    {
        public static void HandleUploadedFile(object arg)
        {
            if (arg == null)
            {
                POIGlobalVar.POIDebugLog("Arg is null when handling uploaded file!");
                return;
            }

            String[] param = arg as String[];
            String extName = param[0];
            String savedFn = param[1];
            String name = param[2];
            String description = param[3];
            int presId = Int32.Parse(param[4]);
            String uploadType = param[5];

            Stopwatch uploadTime = new Stopwatch();
            uploadTime.Start();

            switch (extName)
            {
                case @".PDF":
                case @".pdf":
                    POIPDFProcessor.Process(savedFn, name, description, presId);
                    break;
                case @".PPT":
                case @".ppt":
                case @".PPTX":
                case @".pptx":
                    POIPPTProcessor.Process(savedFn, name, description, presId, uploadType);
                    break;
                case @".POI":
                    POIFileReader reader = new POIFileReader(savedFn);
                    reader.GetImageAndAnimationFromFile();
                    break;
            }

            uploadTime.Stop();
            POIGlobalVar.POIDebugLog("Time used for uploading:" + uploadTime.Elapsed);
        }
    }
}