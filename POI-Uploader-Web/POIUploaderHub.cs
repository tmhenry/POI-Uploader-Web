using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using Microsoft.AspNet.SignalR;
using Microsoft.AspNet.SignalR.Hubs;
using Owin;

using POILibCommunication;

namespace POI_Uploader_Web
{
    [HubName("poiUploader")]
    public class POIUploaderHub : Hub
    {
        public void Log(string msg)
        {
            POIGlobalVar.POIDebugLog(msg);
        }
    }
}