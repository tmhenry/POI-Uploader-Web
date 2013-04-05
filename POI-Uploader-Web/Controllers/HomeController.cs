using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using POILibCommunication;
using Communication;
using System.Threading;

namespace POI_Uploader_Web.Controllers
{
    public class HomeController : Controller
    {
        
        public ActionResult UploadPresentation()
        {
            String presentor = "unknown";
            String description = "unknown";
            
            if (Request.Form.AllKeys.Contains("presentor"))
            {
                presentor = Request.Form["presentor"];
            }

            if (Request.Form.AllKeys.Contains("description"))
            {
                description = Request.Form["description"];
            }

            foreach (string file in Request.Files)
            {
                
                HttpPostedFileBase hpf = Request.Files[file] as HttpPostedFileBase;

                //Ignore the file if the length is zero
                if (hpf.ContentLength == 0) continue;

                String presFn = Path.GetFileName(hpf.FileName);

                String savedFn = Path.Combine(POIArchive.ArchiveHome, presFn);

                hpf.SaveAs(savedFn);

                //Create a new thread and start handling
                String[] param = new String[4];
                param[0] = Path.GetExtension(presFn);
                param[1] = savedFn;
                param[2] = presentor;
                param[3] = description;

                Thread fileHandler = new Thread(HandleUploadedFile);
                fileHandler.Start(param);
            }

            return Json("Presentation files processed!");
        }

        private void HandleUploadedFile(object arg)
        {
            String[] param = arg as String[];
            String extName = param[0];
            String savedFn = param[1];
            String presentor = param[2];
            String description = param[3];

            switch (extName)
            {
                case @".PDF":
                case @".pdf":
                    POIPDFProcessor.Process(savedFn, presentor, description);
                    break;
                case @".PPT":
                case @".ppt":
                case @".PPTX":
                case @".pptx":
                    POIPPTProcessor.Process(savedFn, presentor, description);
                    break;
                case @".POI":
                    POIFileReader reader = new POIFileReader(savedFn);
                    reader.GetImageAndAnimationFromFile();
                    break;
            }
        }
    }
}
