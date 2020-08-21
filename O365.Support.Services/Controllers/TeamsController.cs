using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using O365.Support.Services.Models;
using O365.Support.Services.Services;


namespace O365.Support.Services.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    [ApiController]
    public class TeamsController : ControllerBase
    {
        [HttpPost("CreateTeamsGroup")]
        public async Task<IActionResult> CreateTeamGroup([FromBody] TeamGroup teamGroup)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = await MicrosoftGraphClient.GetGraphServiceClient();
                StringBuilder strMember = new StringBuilder();
                
                // adding members to the Team Group.
                if (teamGroup.Members != null && teamGroup.Members.Count > 0)
                {
                    int i = 0;
                    strMember.AppendFormat(@"[");
                    foreach (string member in teamGroup.Members)
                    {
                        if (i < teamGroup.Members.Count)
                        {
                            var memberUser = await graphClient.Users[member].Request().GetAsync();

                            if (memberUser != null)
                            {
                                if (i == teamGroup.Members.Count - 1)
                                {
                                    var addMember = await graphClient.Users[member].Request().Select(Common.Constants.ID).GetAsync();
                                    await graphClient.Groups[teamGroup.GroupId].Members.References.Request().AddAsync(addMember);
                                    //strMember.AppendFormat("\"{0}/{1}\"", Common.Constants.GRAPH_USERS_URL, memberUser.Id);
                                }
                                else
                                {
                                    var addMember = await graphClient.Users[member].Request().Select(Common.Constants.ID).GetAsync();
                                    await graphClient.Groups[teamGroup.GroupId].Members.References.Request().AddAsync(addMember);
                                    //strMember.AppendFormat("\"{0}/{1}\",", Common.Constants.GRAPH_USERS_URL, memberUser.Id);
                                }
                            }
                        }

                        i++;
                    }
                    strMember.AppendFormat(@"]");
                }

                // Create team 
                Team newTeam = new Team()
                { 
                    GuestSettings = new TeamGuestSettings
                    {
                        AllowCreateUpdateChannels = false,
                        AllowDeleteChannels = false,
                        ODataType = null
                    },
                    MemberSettings = new TeamMemberSettings
                    {
                        AllowCreateUpdateChannels = false,
                        ODataType = null
                    },
                    MessagingSettings = new TeamMessagingSettings
                    {
                        AllowUserEditMessages = true,
                        AllowUserDeleteMessages = true,
                        ODataType = null
                    },
                    FunSettings = new TeamFunSettings
                    {
                        AllowGiphy = true,
                        GiphyContentRating = GiphyRatingType.Strict,
                        ODataType = null
                    },
                    DisplayName = teamGroup.TeamName,
                    ODataType = null
                };
                
                await graphClient.Groups[teamGroup.GroupId].Team
                    .Request()
                    .PutAsync(newTeam);

                return Ok(newTeam);
            }
            catch (ServiceException ex)
            {

                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest();
                }
                else
                {
                    return NotFound();
                }
            }
		}
    }
}
