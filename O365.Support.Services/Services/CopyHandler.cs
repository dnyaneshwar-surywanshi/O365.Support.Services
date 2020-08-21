using O365.Support.Services.Models;

namespace O365.Support.Services.Services
{
    public class CopyHandler
    {
        public static User UserProperty(Microsoft.Graph.User graphUser)
        {
            User user = new User();
            user.id = graphUser.Id;
            user.givenName = graphUser.GivenName;
            user.surname = graphUser.Surname;
            user.userPrincipalName = graphUser.UserPrincipalName;
            user.email = graphUser.Mail;

            return user;
        }

        public static DistributionGroup GroupProperty(Microsoft.Graph.Group graphGroup)
        {
            DistributionGroup group = new DistributionGroup();
            group.id = graphGroup.Id;
            group.displayName = graphGroup.DisplayName;

            return group;
        }
    }
}
