using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using POILibCommunication;
using Microsoft.AspNet.SignalR;

namespace POI_Uploader_Web
{
    public class POIUploaderLogHandler: LogMessageDelegate
    {
        public void logMessage(string msg)
        {
            //Notify all the server web end about the message
            var context = GlobalHost.ConnectionManager.GetHubContext<POIUploaderHub>();
            context.Clients.All.logMessage(msg);
        }
    }
}