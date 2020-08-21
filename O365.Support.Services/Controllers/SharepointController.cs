using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph;
using Neo4jClient;
using O365.Support.Services.Models;
using O365.Support.Services.Services;

namespace O365.Support.Services.Controllers
{
    [Produces("application/json")]
    [Route("api/[controller]")]
    [ApiController]
    public class SharepointController : ControllerBase
    {
        [HttpPost("createdocumentlibrary")]
        public async Task<IActionResult> CreateDocumentLibrary([FromBody] DocumentLibrary documentLibrary)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = await MicrosoftGraphClient.GetGraphServiceClient();

               // var rootSite = await graphClient.Sites.Root.Request().GetAsync();
                
                // Get site ID based on SiteCollectionURL
                var site = await graphClient.Sites[documentLibrary.SiteCollectionURL]
                                      .Request().GetAsync();

                var list = new List
                {
                    DisplayName = documentLibrary.LibraryName,
                    ListInfo = new ListInfo
                    {
                        Template = Common.Constants.TEMPLATE_DOCUMENT_LIBRARY
                    }
                };
               
                await graphClient.Sites[site.Id].Lists
                    .Request()
                    .AddAsync(list);
                return Ok(list);
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

        [HttpPost("restoreItem")]
        public async Task<IActionResult> RestoreItem([FromBody] DocumentLibrary documentLibrary)
        {
            try
            {
                // Initialize the GraphServiceClient.
                GraphServiceClient graphClient = await MicrosoftGraphClient.GetGraphServiceClient();
                var drive = await graphClient.Sites[documentLibrary.SiteCollectionURL].r
                var parentReference = new ItemReference
                {
                    Id = drive.
                };

                var name = "NewDeleteItem";

                await graphClient.Me.Drive.Items[drive.Id]
                    .Restore(parentReference, name)
                    .Request()
                    .PostAsync();
                return Ok();
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
