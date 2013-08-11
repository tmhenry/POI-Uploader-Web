using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace POI_Uploader_Web.Models
{
    public class HomeViewModel
    {
        public string Name { get; set; }

        public string Uploader { get; set; }
        public string ContentServer { get; set; }
        public string DNSServer { get; set; }
    }
}