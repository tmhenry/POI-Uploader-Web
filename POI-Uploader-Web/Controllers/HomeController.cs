using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

using System.IO;
using POILibCommunication;
using Communication;
using System.Threading;
using System.Diagnostics;

namespace POI_Uploader_Web.Controllers
{
    public class HomeController : Controller
    {
        
        public ActionResult UploadPresentation()
        {
            /*
            Dictionary<string, string> presInfo = new Dictionary<string, string>();

            presInfo["name"] = "unknown";
            presInfo["description"] = "unknown";
            //presInfo["creator"] = "unknown";
            presInfo["type"] = "public";
            
            if (Request.Form.AllKeys.Contains("name"))
            {
                presInfo["name"] = Request.Form["name"];
            }

            if (Request.Form.AllKeys.Contains("description"))
            {
                presInfo["description"] = Request.Form["description"];
            }

            if (Request.Form.AllKeys.Contains("creator"))
            {
                //presInfo["creator"] = Request.Form["creator"];
            }

            if (Request.Form.AllKeys.Contains("type"))
            {
                presInfo["type"] = Request.Form["type"];
            }*/

            //int pptID = POIWebService.UploadPresentation(presInfo);

            int pptID = -1;
            if (Request.Form.AllKeys.Contains("pid"))
            {
                pptID = Int32.Parse(Request.Form["pid"]);
            }

            if (pptID > 0)
            {
                foreach (string file in Request.Files)
                {

                    HttpPostedFileBase hpf = Request.Files[file] as HttpPostedFileBase;

                    //Ignore the file if the length is zero
                    if (hpf.ContentLength == 0) continue;

                    String presFn = Path.GetFileName(hpf.FileName);

                    String savedFn = Path.Combine(POIArchive.ArchiveHome, presFn);

                    hpf.SaveAs(savedFn);

                    //Create a new thread and start handling
                    String[] param = new String[5];
                    param[0] = Path.GetExtension(presFn);
                    param[1] = savedFn;
                    //param[2] = presInfo["name"];
                    param[2] = "name";
                    //param[3] = presInfo["description"];
                    param[3] = "description";
                    param[4] = pptID.ToString();

                    //Enqueue the request to the queue
                    ProcessQueue.EnqueueRequest(param);
                }
            }

            Dictionary<string, string> response = new Dictionary<string, string>();
            response["presId"] = pptID.ToString();

            return Json(response);
        }

        
    }
}
