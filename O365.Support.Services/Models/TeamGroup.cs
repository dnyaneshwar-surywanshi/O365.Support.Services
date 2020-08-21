using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace O365.Support.Services.Models
{
    public class TeamGroup
    {
        public string GroupId { get; set; }
        public string TeamName { get; set; }
        public string TeamOwner { get; set; }
        public List<string> Members { get; set; }
        public string Description { get; set; }
    }
}
