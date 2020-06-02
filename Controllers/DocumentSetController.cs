using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using SMBLibrary;
using SMBLibrary.Client;

using SharePointAPI.Middleware;
using SharePointAPI.Models;

namespace SharePointAPI.Controllers
{
    [ApiController]
    [Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public class DocumentSetController : ControllerBase
    {
        private readonly ILogger<SharePointController> _logger;
        public string _username, _password, _baseurl;
        public DocumentSetController(ILogger<SharePointController> logger)
        {
            _logger = logger;
            
            if(Environment.GetEnvironmentVariable("baseurl") != null){
                Console.WriteLine("Baseline url: " + Environment.GetEnvironmentVariable("baseurl"));
                _baseurl = Environment.GetEnvironmentVariable("baseurl");
                _username = Environment.GetEnvironmentVariable("username");
                _password = Environment.GetEnvironmentVariable("password");

            }
            else{
                using (var file = System.IO.File.OpenText("test.json"))
                {
                    var reader = new JsonTextReader(file);
                    var jObject = JObject.Load(reader);
                    _baseurl = jObject.GetValue("baseurl").ToString();
                    _username = jObject.GetValue("username").ToString();
                    _password = jObject.GetValue("password").ToString();
                }
            }
            
        }
        [HttpGet]
        [Produces("application/json")]
        [Consumes("application/json")]

        /// <summary>
        /// Create a new doc
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/folders?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> FolderCollection([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                Console.WriteLine(url);
                
                List list = cc.Web.Lists.GetByTitle(listname);
                
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                
                var root = list.RootFolder;
                cc.Load(root, 
                    r => r.Folders.Include(
                        folder => folder.ProgID,
                        folder => folder.Name,
                        folder => folder.ListItemAllFields
                    )
                );
                await cc.ExecuteQueryAsync();
                

                List<Dictionary<string, object>> SPDocs = new List<Dictionary<string, object>>();

                var folders = root.Folders;
                foreach (var folder in folders)
                {
                    SPDocs.Add(folder.ListItemAllFields.FieldValues);
                }
                return new OkObjectResult(SPDocs);
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }
        [HttpGet]
        [Produces("application/json")]
        [Consumes("application/json")]

        /// <summary>
        /// Get all folders
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/folders?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> Metadata([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                Console.WriteLine(url);
                
                List list = cc.Web.Lists.GetByTitle(listname);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = @"<View Scope='RecursiveAll'><Where>
                                        <BeginsWith>
                                            <FieldRef Name='ContentTypeId' />
                                            <Value Type='ContentTypeId'>0x0120</Value>
                                        </BeginsWith>
                                    </Where><RowLimit>5000</RowLimit></View>";
                
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                
                List<Dictionary<string, object>> SPDocs = new List<Dictionary<string, object>>();
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                do
                {
                    ListItemCollection listItemCollection = list.GetItems(camlQuery);
                    cc.Load(listItemCollection);
                    await cc.ExecuteQueryAsync();

                    //Adding the current set of ListItems in our single buffer
                    foreach (var listitem in listItemCollection)
                    {
                        SPDocs.Add(listitem.FieldValues);
                        
                    }
                    //Reset the current pagination info
                    camlQuery.ListItemCollectionPosition = listItemCollection.ListItemCollectionPosition;
                    
                } while (camlQuery.ListItemCollectionPosition != null);


                return new OkObjectResult(SPDocs);
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }
        
        
    }
}