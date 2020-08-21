using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace O365.Support.Services.Models
{    
    public class Groups
    {
        public int itemsPerPage { get; set; }
        public int startIndex { get; set; }
        public int totalResults { get; set; }
        public List<DistributionGroup> resources { get; set; }
    }
}
