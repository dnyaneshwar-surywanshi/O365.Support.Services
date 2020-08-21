using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using O365.Support.Services.Models;
using O365.Support.Services.Services;
using Common = O365.Support.Services.Common;

namespace O365.Support.Services.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    [ApiController]
    public class ExchangeOnlineController : ControllerBase
    {
        internal static class RouteNames
        {
            public const string Users = nameof(Users);
            public const string UserById = nameof(UserById);
            public const string Groups = nameof(Groups);
            public const string GroupById = nameof(GroupById);
        }

        [HttpGet("users/{id}", Name = RouteNames.UserById)]
        public async Task<IActionResult> GetUser(string id)
        {
            Models.User objUser = new Models.User();
            try
            {
                if (string.IsNullOrEmpty(id) || string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest();
                }


                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load user profile.
                var user = await client.Users[id].Request().GetAsync();

                // Copy Microsoft-Graph User to DTO User
                objUser = CopyHandler.UserProperty(user);

                return Ok(objUser);
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

        [HttpGet("users/")]
        public async Task<IActionResult> GetUsers()
        {
            O365.Support.Services.Models.Users users = new O365.Support.Services.Models.Users();
            try
            {
                users.resources = new List<Models.User>();

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load users profiles.
                var userList = await client.Users.Request().GetAsync();

                // Copy Microsoft User to DTO User
                foreach (var user in userList)
                {
                    var objUser = CopyHandler.UserProperty(user);
                    users.resources.Add(objUser);
                }
                users.totalResults = users.resources.Count;

                return Ok(users);
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

        [HttpPost("createuser")]
        public async Task<IActionResult> CreateUser([FromBody] Microsoft.Graph.User userInfo)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load group profile.
                var _user = await client.Users.Request().AddAsync(userInfo);

                return Ok(_user);
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

        [HttpDelete("delete/{id}")]
        public async Task<IActionResult> DeleteUser(string id)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                if (!string.IsNullOrEmpty(id))
                {
                    var _user = client.Users[id].Request().DeleteAsync();
                    return Ok(_user);
                }
                else
                {
                    return NotFound();
                }

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

        [HttpGet("contacts")]
        public async Task<IActionResult> GetContacts(string ownerId)
        {
            IEnumerable<Microsoft.Graph.Contact> contacts = null;
            try
            {
                contacts = new List<Microsoft.Graph.Contact>();

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load users profiles.
                var contactList = await client.Users[ownerId].Contacts.Request().GetAsync();

                return Ok(contactList);
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

        [HttpPost("addcontact/{id}")]
        public async Task<IActionResult> AddContact(string id, [FromBody] Microsoft.Graph.Contact contactInfo)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();
                Microsoft.Graph.Contact user = new Microsoft.Graph.Contact();

                var _contact = await client.Users[id].Contacts.Request().AddAsync(contactInfo);

                return Ok(_contact);
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

        [HttpPost("deletecontact")]
        public async Task<IActionResult> DeleteContact(string ownerId, string email)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();
                string contactId = string.Empty;
                var contactList = await client.Users[ownerId].Contacts.Request().GetAsync();
                if (contactList != null && contactList.Count > 0)
                {
                    contactId = contactList.Where(i => i.EmailAddresses.FirstOrDefault().Address.ToLower() == email.ToLower()).FirstOrDefault().Id;
                    var _contact = client.Users[ownerId].Contacts[contactId].Request().DeleteAsync();

                    return Ok(_contact);
                }
                else
                {
                    return NotFound();
                }
            }
            catch (ServiceException ex)
            {
                if (ex.StatusCode == HttpStatusCode.BadRequest)
                {
                    return BadRequest(ex);
                }
                else
                {
                    return NotFound();
                }
            }
        }

        [HttpGet("groups/{id}", Name = RouteNames.GroupById)]
        public async Task<IActionResult> GetGroup(string id)
        {
            Models.DistributionGroup objGroup = new Models.DistributionGroup();
            try
            {

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load group profile.
                var group = await client.Groups[id].Request().GetAsync();

                // Copy Microsoft-Graph Group to DTO Group
                objGroup = CopyHandler.GroupProperty(group);

                return Ok(objGroup);
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

        [HttpGet("groups")]
        public async Task<IActionResult> GetGroups()
        {
            Models.Groups groups = new Models.Groups();
            try
            {
                groups.resources = new List<Models.DistributionGroup>();

                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();

                // Load groups profiles.
                var groupList = await client.Groups.Request().GetAsync();

                // Copy Microsoft-Graph Group to DTO Group
                foreach (var group in groupList)
                {
                    var objGroup = CopyHandler.GroupProperty(group);
                    groups.resources.Add(objGroup);
                }

                groups.totalResults = groups.resources.Count;

                return Ok(groups);
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

        [HttpPost]
        [Route("creategroup")]
        public async Task<IActionResult> CreateGroup([FromBody] Models.DistributionGroup groupInfo)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient client = await MicrosoftGraphClient.GetGraphServiceClient();
                StringBuilder strOwner = new StringBuilder();
                StringBuilder strMember = new StringBuilder();
               
                if (groupInfo != null)
                {
                    int i = 0;
                    int maxOwnerCount = 10;
                    if (groupInfo.owners != null && groupInfo.owners.Count > 0)
                    {
                        strOwner.AppendFormat(@"[");
                        foreach (string owner in groupInfo.owners)
                        {
                           var  ownerUser = await client.Users[owner].Request().GetAsync();

                            if (ownerUser != null)
                            {
                                if (i < groupInfo.owners.Count && i < maxOwnerCount)
                                {
                                    if (i == groupInfo.owners.Count - 1)
                                    {
                                        strOwner.AppendFormat("\"{0}/{1}\"", Common.Constants.GRAPH_MEMBERS_URL, ownerUser.Id);
                                    }
                                    else
                                    {
                                        strOwner.AppendFormat("\"{0}/{1}\",", Common.Constants.GRAPH_MEMBERS_URL, ownerUser.Id);
                                    }
                                }
                            }

                            i++;
                        }
                        strOwner.AppendFormat(@"]");
                    }
                    
                    if (groupInfo.members != null && groupInfo.members.Count > 0)
                    {
                        i = 0;
                        strMember.AppendFormat(@"[");
                        foreach (string member in groupInfo.members)
                        {
                            if (i < groupInfo.members.Count)
                            {
                                var memberUser = await client.Users[member].Request().GetAsync();

                                if (memberUser != null)
                                {
                                    if (i == groupInfo.members.Count - 1)
                                    {
                                        strMember.AppendFormat("\"{0}/{1}\"", Common.Constants.GRAPH_USERS_URL, memberUser.Id);
                                    }
                                    else
                                    {
                                        strMember.AppendFormat("\"{0}/{1}\",", Common.Constants.GRAPH_USERS_URL, memberUser.Id);
                                    }
                                }
                            }

                            i++;
                        }
                        strMember.AppendFormat(@"]");
                    }

                    if(groupInfo.groupTypes == null)
                    {
                        groupInfo.groupTypes = new List<string>()
                        {
                            Common.Constants.UNIFIED
                        };
                    }

                    var group = new Group
                    {
                        Description = groupInfo.description,
                        DisplayName = groupInfo.displayName,
                        GroupTypes = groupInfo.groupTypes,
                        MailEnabled = groupInfo.mailEnabled,
                        MailNickname = groupInfo.mailNickname,
                        SecurityEnabled = false,
                        AdditionalData = new Dictionary<string, object>()
                        {
                            
                            {"\"members@odata.bind\"", strMember.ToString()}
                        }
                    };

                    var _group = await client.Groups.Request().AddAsync(group);
                    //var _gowner = await client.Groups[_group.Id].Owners.References.Request().AddAsync()
                    //var directoryObject = new DirectoryObject
                    //{
                    //    Id = ownerUser.Id
                    //};
                    //var mem = await client.Users[groupInfo.owners.FirstOrDefault()].Request().GetAsync();
                    //await client.Groups[group.Id].Owners.References.Request().AddAsync(mem);
                    await client.Groups[group.Id].Request().UpdateAsync(group);
                    return Ok(_group);
                }
                else
                {
                    return BadRequest();
                }

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
