using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using POILibCommunication;
using Communication;

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

                switch(Path.GetExtension(presFn))
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

            return Json("Presentation files processed!");
        }
    }
}
