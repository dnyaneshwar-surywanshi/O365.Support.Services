using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace O365.Support.Services.Models
{
    public class Users
    {
        public int itemsPerPage { get; set; }
        public int startIndex { get; set; }
        public int totalResults { get; set; }
        public List<User> resources { get; set; }
    }
}
