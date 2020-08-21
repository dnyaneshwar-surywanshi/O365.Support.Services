using Microsoft.Graph;
using System.Collections.Generic;

namespace O365.Support.Services.Models
{
    public class DistributionGroup
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string description { get; set; }
        public List<string> groupTypes { get; set; }
        public bool mailEnabled { get; set; }
        public string mailNickname { get; set; }
        public bool securityEnabled { get; set; }        
        public List<string> owners { get; set; }
        public List<string> members { get; set; }

    }   
}
