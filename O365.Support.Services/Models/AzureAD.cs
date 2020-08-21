using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace O365.Support.Services.Models
{
    public class AzureAD
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
        public string TenantId { get; set; }
        public string Instance { get; set; }
        public string GraphResource { get; set; }
        public string GraphResourceEndPoint { get; set; }
    }
}
