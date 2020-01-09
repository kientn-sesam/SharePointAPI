using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using SharePointAPI.Middleware;

namespace SharePointAPI.Controllers
{
    [ApiController]
    [Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public class SharePointController : ControllerBase
    {
         public string _username, _password, _baseurl;
        //private ClientContext cc;

        public SharePointController()
        {
            if(Environment.GetEnvironmentVariable("baseurl") != null){
                Console.WriteLine("Baseline url: " + Environment.GetEnvironmentVariable("baseurl"));
                _baseurl = Environment.GetEnvironmentVariable("baseurl");
                _username = Environment.GetEnvironmentVariable("username");
                _password = Environment.GetEnvironmentVariable("password");

            }
            else{
                using (var file = System.IO.File.OpenText("helpers.json"))
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
        ///     GET /api/sharepoint/documentswithfields?site=sitename&list=listname
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> DocumentsWithFields([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            var SPDocs = new List<JObject>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                
                Web web = cc.Web;
                cc.Load(web);
                var Lists = web.Lists;
                cc.Load(Lists);
                await cc.ExecuteQueryAsync();

                List list = web.Lists.GetByTitle(listname);
                cc.Load(list, l => l.Fields);
                await cc.ExecuteQueryAsync();
                List<string> fieldNames = new List<string>();
   

                foreach (var tmpfield in list.Fields)
                {
                    if((!tmpfield.FromBaseType && tmpfield.Hidden == false)){
                        if (tmpfield.InternalName.Equals("ContentType"))
                        {
                            continue;
                        }
                        Console.WriteLine(tmpfield.InternalName);
                        
                        fieldNames.Add(tmpfield.InternalName);
                    }
                }
                
                FolderCollection folders = list.RootFolder.Folders;
                cc.Load(folders);
                await cc.ExecuteQueryAsync();

                

                foreach (var folder in folders)
                {
                    
                    var items = folder.Files;
                    Console.WriteLine("Folder Name: " + folder.Name);
                    
                    // Skip unecessary folder
                    if(string.IsNullOrEmpty(folder.ProgID)){
                        continue;
                    }

                    cc.Load(items);
                    await cc.ExecuteQueryAsync();

                    foreach (var file in items)
                    {
                        
                        ListItem item = file.ListItemAllFields;
                        cc.Load(item);
                        await cc.ExecuteQueryAsync();
                        
                        var json = new JObject();
                        json.Add(new JProperty("filename", file.Name));
                        json.Add(new JProperty("folder", folder.Name));
                        json.Add(new JProperty("uri", file.LinkingUri));
                        foreach (var fieldname in fieldNames)
                        {
                            Console.WriteLine("Name: " + fieldname + "Value: " + item[fieldname]);
                            if (item[fieldname] != null)
                            {
                                json.Add(new JProperty(fieldname, item[fieldname]));
                                
                                Console.WriteLine("Name: " + fieldname + "Value: " + item[fieldname]);
                            }
                        }
                        
                        SPDocs.Add(json);

                    }
                }



                return new OkObjectResult(SPDocs);
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        /// <summary>
        /// Create a new doc
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/document
        ///     {
        ///         "list":"Documents",
	    ///         "file_url":"https://sesam.io/images/Cyan.svg",
	    ///         "foldername":"My first document set",
        ///         "site": "Sesamsitewithdocumentsets",
        ///         "filename": "Cyan.svg",
	    ///         "fields":{
	    ///         		"BLAD":"4",
	    ///         		"BESKRIVELSE":"Beskrivelse test",
	    ///         		"DOC_NO": "12345",
	    ///         		"DATO":"2019-11-12 00:00:00"
	    ///         }
        ///         
        ///     }
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        /// <response code="201">Returns success with the new site title</response>
        /// <response code="404">Returns resource not found if the ID of the new site is empty</response>
        /// <response code="500">If the input parameter is null or empty</response>
        public async Task<IActionResult> NewDocument([FromBody] JObject doc)
        {

            string site = doc["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{
                
                Web web = cc.Web;
                cc.Load(web);
                var lists = web.Lists;
                cc.Load(lists);
                await cc.ExecuteQueryAsync();
                
                List list = web.Lists.GetByTitle(doc["list"].ToString());

                FileCreationInformation newFile = new FileCreationInformation();
                //byte[] imageBytes = webClient.DownloadData("https://sesam.io/images/howitworks.jpg");
                using (var webClient = new WebClient()){
                    byte[] imageBytes = webClient.DownloadData(doc["file_url"].ToString());
                    newFile.Content = imageBytes;
                }

                newFile.Url = doc["filename"].ToString();
                
                //Folder folder = list.RootFolder.Folders.GetByUrl("My first document set");
                Folder folder = list.RootFolder.Folders.GetByUrl(doc["foldername"].ToString());
                File uploadFile = folder.Files.Add(newFile);
                
                ListItem item = uploadFile.ListItemAllFields;
                
                JObject fields = doc["fields"] as JObject;
                //Add metadata
                foreach (KeyValuePair<string, JToken> field in fields)
                {
                    string fieldValue = (string)field.Value;
                
                    if (int.TryParse(fieldValue, out int n)){
                        item[field.Key] = n;
                    }
                    else if(DateTime.TryParse(fieldValue, out DateTime dt))
                    {
                        item[field.Key] = dt;
                    }
                    else
                    {
                        item[field.Key] = fieldValue;
                    }
                    
                    
                }
                item.SystemUpdate();
                

                cc.Load(list);
                cc.Load(uploadFile);
                await cc.ExecuteQueryAsync();

                Console.WriteLine("Done");

                return new NoContentResult();



            }
            catch (Exception ex)
            {
                Console.WriteLine("Error message: " + ex);
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }

        /// <summary>
        /// Delete a site
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     DELETE /api/sharepoint/deletesite
        ///     {
        ///        "site": <"site name">
        ///     }
        /// </remarks>
        /// <param name="url">The site url to delete</param>
        /// <returns></returns>
        /// <response code="204">Returns success with No-content result</response>
        /// <response code="500">If the input parameter is null or empty</response>
        [HttpDelete("{url}")]
        [Produces("application/json")]
        [Consumes("application/json")]
        [ProducesResponseType((int)HttpStatusCode.NotFound)]
        [ProducesResponseType((int)HttpStatusCode.RequestTimeout)]
        [ProducesResponseType((int)HttpStatusCode.NoContent)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> DeleteSite([FromBody] string site)
        {
    
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{    
                //Web web = context.Web;
                //context.Load(web);
                //context.Credentials = new NetworkCredential("khteh", "", "dddevops.onmicrosoft.com");
                Web web = cc.Web;
                // Retrieve the new web information. 
                cc.Load(web);
                //context.Load(newWeb);
                await cc.ExecuteQueryAsync();
                web.DeleteObject();
                await cc.ExecuteQueryAsync();
                return new NoContentResult();
            } catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }
        
        /// <summary>
        /// Create a document site
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/documentset
        ///     {
        ///         "site": <"site name">,
        ///         "list" :<"list name">,
        ///         "sitecontent" : <"site content name">,
        ///         "documentset" : <"name of the new document set">,
        ///      }
        /// </remarks>
        /// <returns></returns>
        /// <response code="201">Returns success with the new site title</response>
        /// <response code="404">Returns resource not found if the ID of the new site is empty</response>
        /// <response code="500">If the input parameter is null or empty</response>
        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        [ProducesResponseType((int)HttpStatusCode.NotFound)]
        [ProducesResponseType((int)HttpStatusCode.RequestTimeout)]
        [ProducesResponseType((int)HttpStatusCode.Created)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> DocumentSet([FromBody] JObject param)
        {
            string url = _baseurl + "sites/" + param["site"];

            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{
                Web web = cc.Web;
                cc.Load(web);
                // Example: List list = cc.Web.Lists.GetByTitle("Documents");
                List list = cc.Web.Lists.GetByTitle(param["list"].ToString());
                ContentTypeCollection listContentTypes = list.ContentTypes;
                cc.Load(listContentTypes, types => types.Include(type => type.Id, type => type.Name, type => type.Parent));
                
                string SiteContentName = param["sitecontent"].ToString();
                // Example: var result = cc.LoadQuery(listContentTypes.Where(c => c.Name == "document set 2"));
                var result = cc.LoadQuery(listContentTypes.Where(c => c.Name == SiteContentName));
                
                await cc.ExecuteQueryAsync();

                ContentType targetDocumentSetContentType = result.FirstOrDefault();
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();

                newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                //newItemInfo.LeafName = "Document Set Kien2";
                newItemInfo.LeafName = param["documentset"].ToString();
                
                //newItemInfo.FolderUrl = list.RootFolder.ServerRelativeUrl.ToString();
                
                ListItem newListItem = list.AddItem(newItemInfo);
                newListItem["ContentTypeId"] = targetDocumentSetContentType.Id.ToString();
                newListItem.Update();
                list.Update();
                await cc.ExecuteQueryAsync();

                return new NoContentResult();
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }


    }
}