using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

using POILibCommunication;
using System.Web.Configuration;

namespace POI_Uploader_Web
{
    // Note: For instructions on enabling IIS6 or IIS7 classic mode, 
    // visit http://go.microsoft.com/?LinkId=9394801

    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();

            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            loadConfigFile();

            ProcessQueue.Run(POIUploadHandler.HandleUploadedFile);
        }

        protected void loadConfigFile()
        {
            //string fn = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"poi_config");

            try
            {
                POIGlobalVar.ContentServerHome = WebConfigurationManager.AppSettings["ContentServer"];
                POIGlobalVar.DNSServerHome = WebConfigurationManager.AppSettings["DNSServer"];
                POIGlobalVar.ProxyServerIP = WebConfigurationManager.AppSettings["ProxyServerIP"];
                POIGlobalVar.ProxyServerPort = Int32.Parse(WebConfigurationManager.AppSettings["ProxyServerPort"]);

                //POIGlobalVar.POIDebugLog(POIGlobalVar.ContentServerHome);

            }
            catch (Exception e)
            {
                POIGlobalVar.POIDebugLog(e);
            }
        }
    }
}