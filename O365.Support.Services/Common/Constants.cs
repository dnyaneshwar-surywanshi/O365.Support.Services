using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace O365.Support.Services.Common
{
    public static class Constants
    {
        public static string ASPNETCORE_ENVIRONMENT = "ASPNETCORE_ENVIRONMENT";
        public static string GRAPH_MEMBERS_URL = @"https://graph.microsoft.com/v1.0/users";
        public static string GRAPH_USERS_URL = @"https://graph.microsoft.com/v1.0/directoryObjects";
        public static string OWNERS_ODATA_BIND = "\"owners@odata.bind\"";
        public static string MEMBERS_ODATA_BIND = "\"members@odata.bind\"";
        public static string UNIFIED = "Unified";
        public static string ID = "id";
        public static string TEMPLATE_DOCUMENT_LIBRARY = "documentLibrary";

    }
}
