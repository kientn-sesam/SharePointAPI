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
    public class SharePointController : ControllerBase
    {
        private readonly ILogger<SharePointController> _logger;
         public string _username, _password, _baseurl;
        //private ClientContext cc;

        public SharePointController(ILogger<SharePointController> logger)
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
        /// get all documents in library
        /// 
        /// NB! This endpoint does not support 5000+ entities. 
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/documents?site=<sitename>&list=<listname>

        /// </remarks>
        /// <returns></returns>
        public async Task<IActionResult> Documents([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {

                List list = cc.Web.Lists.GetByTitle(listname);
                
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                
                var root = list.RootFolder;
                cc.Load(root, 
                    r => r.Folders.Include(
                        folder => folder.ProgID,
                        folder => folder.Name,
                        files => files.Files.Include(
                            file => file.Name,
                            file => file.LinkingUri,
                            file => file.ListItemAllFields
                        )
                    ),
                    r => r.Files.Include(
                        file => file.Name,
                        file => file.LinkingUri,
                        file => file.ListItemAllFields
                        )
                    );
                await cc.ExecuteQueryAsync();
                

                List<JObject> SPDocs = new List<JObject>();
                if (root.Files.Count > 0)
                {
                    var files = root.Files;

                    SPDocs.AddRange(SharePointHelper.GetDocuments(cc, files, null));
                }
                if(root.Folders.Count > 0)
                {
                    FolderCollection folders = root.Folders;
                    for (int fs = 0; fs < folders.Count; fs++)
                    {
                        FileCollection files = folders[fs].Files;
                        string foldername = folders[fs].Name;
                        
                        // Skip unecessary folder
                        if(string.IsNullOrEmpty(folders[fs].ProgID)){
                            continue;
                        }
                        SPDocs.AddRange(SharePointHelper.GetDocuments(cc, files, foldername));
                    }
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
        /// List of libraries from a SharePoint site
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/lists?site=<sitename>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> Lists([FromQuery(Name = "site")] string sitename)
        {
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List<ListModel> listSP = new List<ListModel>();

                ListCollection lists = cc.Web.Lists;
                cc.Load(lists);
                await cc.ExecuteQueryAsync();

                for (int i = 0; i < lists.Count; i++)
                {
                    List list = lists[i];
                    listSP.Add(new ListModel(){id = list.Id.ToString(), title = list.Title, templateurl = list.DocumentTemplateUrl});
                }

                return new OkObjectResult(listSP);
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
        /// List of folders/documentsets from a sharepoint library
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/folders?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> Folders([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
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
                    if (folder.Name.Equals("Forms"))
                        continue;
                    else                    
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
        /// documents with metadata
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/documentswithfields?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        public async Task<IActionResult> DocumentsWithFields([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                Console.WriteLine(_baseurl);
                
                List list = SharePointHelper.GetListItemByTitle(cc, listname);
                
                List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);

                FolderCollection folders = list.RootFolder.Folders;
                cc.Load(folders);
                await cc.ExecuteQueryAsync();

                List<JObject> SPDocs = SharePointHelper.GetItemsFromListByField(cc, folders, fieldNames);

                return new OkObjectResult(SPDocs);
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }

        [HttpGet]
        /// <summary>
        /// List of available fields on specific library
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/fields?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        public List<Metadata> Fields([FromQuery(Name = "site")] string site,[FromQuery(Name = "list")] string listname)
        {

            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);

                return SharePointHelper.GetFields(cc, list);
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
        ///        "list":"Dokumentasjon",
        ///        "file_url":"https://www.bring.no/radgivning/sende-noe/adressetjenester/postnummer/_/attachment/download/c0300459-6555-4833-b42c-4b16496b7cc0:1127fa77303a0347c45d609069d1483b429a36c0/Postnummerregister-Excel.xlsx",
        ///        "foldername":"Lan",
        ///        "site": "sporaevk",
        ///        "filename": "Postnummerregister-Excel.xlsx",
        ///        "sitecontent":"SPOR Dokumentsett",
        ///        "documentsetfields":{
        ///        	"SPORStatus" : {
        ///        		"label":"Under arbeid", 
        ///        		"type":"Text"
        ///        	},
        ///        	"SPORResponsible" : {
        ///        		"label":"Nina.Torjesen@ae.no", 
        ///        		"type" : "User"
        ///        		
        ///        	}
        ///        },
        ///        "fields":{
        ///        		"SPORProjectName": "Smeland, nye kjølere T1",
        ///                 "Title": "Testing",
        ///                 "SPORResponsible": "Nina.Torjesen@ae.no"
        ///        }
        ///        
        ///     }    
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        /// <response code="201">Returns success with the new site title</response>
        /// <response code="404">Returns resource not found if the ID of the new site is empty</response>
        /// <response code="500">If the input parameter is null or empty</response>
        public async Task<IActionResult> NewDocument([FromBody] JArray param)
        {
            
            JObject doc = param.ToObject<List<JObject>>().FirstOrDefault();
            Console.WriteLine(doc);

            string site = doc["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{
                List list = SharePointHelper.GetListItemByTitle(cc, doc["list"].ToString());
                SMBCredential SMBCredential = new SMBCredential(){ 
                    username = Environment.GetEnvironmentVariable("smb_username"), 
                    password = _password, 
                    domain = Environment.GetEnvironmentVariable("domain"),
                    ipaddr = Environment.GetEnvironmentVariable("ipaddr"),
                    share = Environment.GetEnvironmentVariable("share"),
                };
                var serverAddress = System.Net.IPAddress.Parse(SMBCredential.ipaddr);
                SMB2Client client = new SMB2Client();
                bool success = client.Connect(serverAddress, SMBTransportType.DirectTCPTransport);

                NTStatus nts = client.Login(SMBCredential.domain, SMBCredential.username, SMBCredential.password);
                ISMBFileStore fileStore = client.TreeConnect(SMBCredential.share, out nts);
                
                FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(doc["file_url"].ToString(), doc["filename"].ToString(), SMBCredential, client, nts, fileStore);
                
                File uploadFile;
                //Upload file to library/list
                if (SharePointHelper.FolderJObjectExist(doc) == false)
                {
                    uploadFile = list.RootFolder.Files.Add(newFile);
                    Console.WriteLine("folder missing!!!!");
                }
                // upload file into document set 
                else{
                    //Folder folder = list.RootFolder.Folders.GetByUrl("My first document set");
                    Folder folder = SharePointHelper.GetFolder(cc, list, doc["foldername"].ToString());
                    
                    if (folder == null)
                    {
                        JObject documentSetFields = doc["documentsetfields"] as JObject;
                        folder = SharePointHelper.CreateFolder(cc, list, doc["sitecontent"].ToString(), doc["foldername"].ToString(), documentSetFields);
                    }

                    uploadFile = folder.Files.Add(newFile);
                }

                ListItem item = uploadFile.ListItemAllFields;
                cc.Load(item);

                FieldCollection fields = list.Fields;
                cc.Load(fields);
                cc.ExecuteQuery();

                JObject inputFields = doc["fields"] as JObject;
                //Add metadata
                SharePointHelper.SetMetadataFields(cc, inputFields, fields, item);                
                

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
                newListItem.SystemUpdate();
                list.Update();
                await cc.ExecuteQueryAsync();

                return new NoContentResult();
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }
        [HttpPost]
        public async Task<IActionResult> UpdateTest([FromBody] JArray param)
        {
            // last batch is empty
            if (param.ToArray().Length == 0)
            {
                return null;
            }

            //Console.WriteLine(param);
            JObject tmpDoc = param.ToObject<List<JObject>>().FirstOrDefault();

            //string site = doc["site"].ToString();
            string site = tmpDoc["site"].ToString();
            
            string url = _baseurl + "sites/" + site;
            string tmpList = tmpDoc["list"].ToString();
            

            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(tmpList);
                //User user = cc.Web.EnsureUser("nina.torjesen@ae.no");

                List<JObject> docs = param.ToObject<List<JObject>>();
                for (int i = 0; i < docs.Count; i++)
                {
                    string filename = docs[i]["filename"].ToString();
                    string foldername = docs[i]["foldername"].ToString();
                    //JObject inputFields = docs[i]["fields"] as JObject;

                    Folder folder = list.RootFolder.Folders.GetByUrl(foldername);


                    File file = folder.Files.GetByUrl(filename);
                    ListItem item = file.ListItemAllFields;

                    JObject inputFields = docs[i]["fields"] as JObject;

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    foreach (KeyValuePair<string, JToken> inputField in inputFields)
                    {

                        if (inputField.Value == null || inputField.Value.ToString() == "" || inputField.Key.Equals("Modified"))
                        {
                            continue;
                        }
                        
                        string fieldValue = (string)inputField.Value;
                        Match match = regex.Match(fieldValue);

                        if (inputField.Key.Equals("Author") || inputField.Key.Equals("Editor"))
                        {
                           
                            var user = FieldUserValue.FromUser(fieldValue);
                            item[inputField.Key] = user;
                        }
                        else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By"))
                        {
                            StringBuilder sb = new StringBuilder("i:0#.f|membership|");
                            sb.Append(fieldValue);
                            item[inputField.Key] = sb;

                        }
                        else if(match.Success)
                        {
                            fieldValue = fieldValue.Replace("~t","");
                            if(DateTime.TryParse(fieldValue, out DateTime dt))
                            {
                                if(dtMin <= dt){
                                    item[inputField.Key] = dt;
                                    _logger.LogInformation("Set field " + inputField.Key + "to " + dt);
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            int tokenLength = inputField.Value.Count();

                            if(tokenLength >= 1){
                                continue;
                            }
                            else
                            {
                                item[inputField.Key] = fieldValue;
                                _logger.LogInformation("Set " + inputField.Key + " to " + fieldValue);
                                
                            }
                        }

                        
                        item.Update();
                    }

                    //Modified needs to be updated last
                    string strModified = inputFields["Modified"].ToString();
                    Match matchModified = regex.Match(strModified);
                    if(matchModified.Success)
                    {
                        strModified = strModified.Replace("~t","");
                        if(DateTime.TryParse(strModified, out DateTime dt))
                        {
                                item["Modified"] = dt;
                                item.Update();
                        }
                    }

                    
                    await cc.ExecuteQueryAsync();

                }

                return new NoContentResult();
                
            }
            catch (System.Exception)
            {
                
                throw;
            }

        }

        /// <summary>
        /// update metadata
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/updatemetadata
        ///     {
        ///     	"ListName":"Documents",
        ///     	"FileName":"Cyan.svg",
        ///     	"FolderName":"My first document set",
        ///     	"Fields":{
        ///     			"BLAD":"9",
        ///     			"BESKRIVELSE":"Beskrivelse updated",
        ///     			"DOC_NO": "123433334455",
        ///     			"DATO":"2020-01-01 04:00:00"
        ///     
        ///     	}
        ///     }
        /// </remarks>
        /// <returns></returns>
        /// <response code="201"></response>
        /// <response code="404"></response>
        /// <response code="500">If the input parameter is null or empty</response>
        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        [ProducesResponseType((int)HttpStatusCode.NotFound)]
        [ProducesResponseType((int)HttpStatusCode.RequestTimeout)]
        [ProducesResponseType((int)HttpStatusCode.Created)]
        [ProducesResponseType((int)HttpStatusCode.InternalServerError)]
        public async Task<IActionResult> UpdateMetadata([FromBody] JObject param){
            string site = param["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(param["list"].ToString());

                
                Folder folder = list.RootFolder.Folders.GetByUrl(param["foldername"].ToString());
                
                var items = folder.Files;
                var file = items.GetByUrl(param["filename"].ToString());
                
                cc.Load(file);
                
                ListItem item = file.ListItemAllFields;

                cc.Load(item);
                await cc.ExecuteQueryAsync();

                JObject inputFields = param["fields"] as JObject;
                //update metadata
                foreach (KeyValuePair<string, JToken> inputField in inputFields)
                {
                    JObject taxObj = inputField.Value as JObject;
                    //string fieldValue = (string)inputField.Value;

                    var clientRuntimeContext = item.Context;
                    var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                    cc.Load(field);
                    var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);
                    cc.Load(taxKeywordField);
                    await cc.ExecuteQueryAsync();

                    TaxonomyFieldValue termValue = new TaxonomyFieldValue();
                    termValue.Label = taxObj["Label"].ToString();
                    termValue.TermGuid = taxObj["TermGuid"].ToString();
                    termValue.WssId = (int)taxObj["WssId"];

                    taxKeywordField.SetFieldValueByValue(item, termValue);
                    Console.WriteLine(taxKeywordField);
                    taxKeywordField.Update();
                    
                }
                
                item.SystemUpdate();
                
                cc.Load(file);
                
                await cc.ExecuteQueryAsync();

                Console.WriteLine("Update metadata SUCCESS!");

                return new NoContentResult();
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
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
	    ///         		"SPORProjectNameValue":{
        ///                    "label": "Skjerka nytt aggregat - 8026-3",
        ///                    "TermGuid":"e381ccae-bb79-4a35-9dbb-a54638348fc7",
        ///                    "wssid": 25
        ///                  },
	    ///         		"SPORConstruction":{
        ///                    "label": "Skjerka nytt aggregat - 8026-3",
        ///                    "TermGuid":"e381ccae-bb79-4a35-9dbb-a54638348fc7",
        ///                    "wssid": 25
        ///                  }
	    ///         }
        ///         
        ///     }
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        /// <response code="201">Returns success with the new site title</response>
        /// <response code="404">Returns resource not found if the ID of the new site is empty</response>
        /// <response code="500">If the input parameter is null or empty</response>
        public async Task<IActionResult> UpMetadata([FromBody] JArray param)
        {
            // last batch is empty
            if (param.ToArray().Length == 0)
            {
                return null;
            }

            JObject tmpDoc = param.ToObject<List<JObject>>().FirstOrDefault();
            string site = tmpDoc["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{
                
                List list = cc.Web.Lists.GetByTitle(tmpDoc["list"].ToString());
                
                List<JObject> docs = param.ToObject<List<JObject>>();
                for (int i = 0; i < docs.Count; i++)
                {
                    ListItem listItem;
                    if (SharePointHelper.FolderJObjectExist(docs[i]) == false)
                    {
                        listItem = list.RootFolder.ListItemAllFields;
                    }
                    else
                    {    
                        Folder folder = list.RootFolder.Folders.GetByUrl(docs[i]["foldername"].ToString());
                        listItem = folder.ListItemAllFields;
                    }

                    cc.Load(listItem);
                    cc.ExecuteQuery();

                    var clientRuntimeContext = listItem.Context;


                    JObject inputFields = docs[i]["fields"] as JObject;
                    foreach (KeyValuePair<string, JToken> inputField in inputFields)
                    {
                        var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                        cc.Load(field);
                        cc.ExecuteQuery();
                        var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);

                        
                        JObject taxObj = inputField.Value as JObject;

                        Guid _id = taxKeywordField.TermSetId;
                        string _termID = TermHelper.GetTermIdByName(cc, taxObj["Label"].ToString(), _id);

                        TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                        {
                            Label = taxObj["Label"].ToString(),
                            TermGuid = _termID,
                            //WssId = -1
                            //WssId = (int)taxObj["WssId"]
                        };
                        
                        
                        taxKeywordField.SetFieldValueByValue(listItem, termValue);
                        taxKeywordField.Update();
                        //string termValue = "42;#" + taxObj["Label"].ToString() + "|" + taxObj["TermGuid"].ToString();
                        //listItem[inputField.Key] = termValue;
                        

                    }
                    listItem.SystemUpdate();
                    cc.Load(listItem);
                    
                }
                await cc.ExecuteQueryAsync();
                
                return new NoContentResult();



            }
            catch (Exception ex)
            {
                Console.WriteLine("Error message: " + ex);
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }


        /// <summary>
        /// Upload file to sharepoint
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/UploadToSharePoint
        ///     {
        ///         "list":"Dokumentasjon",
        ///         "file_url":"https://www.bring.no/radgivning/sende-noe/adressetjenester/postnummer/_/attachment/download/c0300459-6555-4833-b42c-4b16496b7cc0:1127fa77303a0347c45d609069d1483b429a36c0/Postnummerregister-Excel.xlsx",
        ///         "foldername":"Landskaps og miljøplan",
        ///         "site": "sporaevk",
        ///         "filename": "Postnummerregister-Excel.xlsx"
        ///     }
        /// </remarks>
        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> UploadToSharePoint([FromBody] JObject doc)
        {
            string filename = doc["filename"].ToString();
            string site = doc["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                string foldername = doc["foldername"].ToString();
                //SMBCredential webCredential = new SMBCredential(){ 
                //    username = _username, 
                //    password = _password, 
                //    domain = "AE03PDFS01.a-e.no"
                //};
                FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(doc["file_url"].ToString(), doc["filename"].ToString());
                //FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(doc["file_url"].ToString(), doc["filename"].ToString(), webCredential);
                List list = cc.Web.Lists.GetByTitle(doc["list"].ToString());
                Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                File uploadFile = folder.Files.Add(newFile);
                ListItem item = uploadFile.ListItemAllFields;
                cc.Load(item);
                await cc.ExecuteQueryAsync();
                return new NoContentResult();
            }
            catch (System.Exception)
            {
                
                throw;
            }

        }
        /// <summary>
        /// Array of foldernames
        /// </summary>
        /// <remarks>
        /// Sample request:
        /// 
        /// </remarks>
        [HttpGet]
        [Produces("application/json")]
        [Consumes("application/json")]
        public string[] FolderNames([FromQuery(Name = "site")] string site,[FromQuery(Name = "list")] string listname)
        {
            
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                FolderCollection folders = SharePointHelper.GetFolders(cc, list);
                string[] foldernames = SharePointHelper.GetFolderNames(cc, list, folders);

                return foldernames;

            }
            catch (System.Exception)
            {
                
                throw;
            }

        }

        /// <summary>
        /// Create a new doc
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     POST /api/sharepoint/document
        ///     {
        ///         "list":"Dokumentasjon",
        ///         "file_url":"https://www.bring.no/radgivning/sende-noe/adressetjenester/postnummer/_/attachment/download/c0300459-6555-4833-b42c-4b16496b7cc0:1127fa77303a0347c45d609069d1483b429a36c0/Postnummerregister-Excel.xlsx",
        ///         "foldername":"Landskaps og miljøplan",
        ///         "site": "sporaevk",
        ///         "filename": "Postnummerregister-Excel.xlsx"
        ///     }
        /// </remarks>
        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> Migration([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }

            SMB2Client client = new SMB2Client();
            
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            //Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                
                SMBCredential SMBCredential = new SMBCredential(){ 
                    username = Environment.GetEnvironmentVariable("smb_username"), 
                    password = Environment.GetEnvironmentVariable("smb_password"), 
                    domain = Environment.GetEnvironmentVariable("domain"),
                    ipaddr = Environment.GetEnvironmentVariable("ipaddr"),
                    share = Environment.GetEnvironmentVariable("share"),
                };

                var serverAddress = System.Net.IPAddress.Parse(SMBCredential.ipaddr);
                bool success = client.Connect(serverAddress, SMBTransportType.DirectTCPTransport);

                NTStatus nts = client.Login(SMBCredential.domain, SMBCredential.username, SMBCredential.password);
                ISMBFileStore fileStore = client.TreeConnect(SMBCredential.share, out nts);


                //List list = cc.Web.Lists.GetById(listGuid);
                List list = cc.Web.Lists.GetByTitle(listname);
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);

                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename, SMBCredential, client, nts, fileStore);
                    ///FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename);
                
                    if (newFile == null){
                        _logger.LogError("Failed to upload. Skip: " + filename);
                        continue;
                    }

                    File uploadFile;
                    if(docs[i].foldername == null){
                        uploadFile = list.RootFolder.Files.Add(newFile);
                    }
                    else{
                        string foldername = docs[i].foldername;
                        string sitecontent = docs[i].sitecontent;
                        
                        //Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                        Folder folder = SharePointHelper.GetFolder(cc, list, foldername);
                        if (folder == null && taxFields != null)
                            folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, inputFields, fields, taxFields);
                        else if (folder == null)
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername);
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);

                    ListItem item = uploadFile.ListItemAllFields;


                    DateTime dtMin = new DateTime(1900,1,1);
                    Regex regex = new Regex(@"~t.*");
                    var listItemFormUpdateValueColl = new List <ListItemFormUpdateValue>();
                    if (inputFields != null)
                    {    
                        foreach (KeyValuePair<string, string> inputField in inputFields)
                        {
                            if (inputField.Value == null || inputField.Value == "" || inputField.Key.Equals("Modified"))
                            {
                                continue;
                            }
                            

                            string fieldValue = inputField.Value;
                            Match match = regex.Match(fieldValue);

                            if (inputField.Key.Equals("Author") || inputField.Key.Equals("Editor"))
                            {
                                FieldUserValue user = FieldUserValue.FromUser(fieldValue);
                                
                                item[inputField.Key] = user;

                            }
                            //endre hard koding
                            else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By") || inputField.Key.Equals("Dokumentansvarlig"))
                            {
                                StringBuilder sb = new StringBuilder("i:0#.f|membership|");
                                sb.Append(fieldValue);
                                item[inputField.Key] = sb;
                            }
                            else if(match.Success)
                            {
                                fieldValue = fieldValue.Replace("~t","");
                                if(DateTime.TryParse(fieldValue, out DateTime dt))
                                {
                                    if(dtMin <= dt){
                                        item[inputField.Key] = dt;
                                        _logger.LogInformation("Set field " + inputField.Key + "to " + dt);
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                item[inputField.Key] = fieldValue;
                                _logger.LogInformation("Set " + inputField.Key + " to " + fieldValue);

                            }

                            
                            item.Update();
                            
                        }

                        cc.ExecuteQuery();
                    }

                    if (taxFields != null)
                    {
                        var clientRuntimeContext = item.Context;
                        for (int t = 0; t < taxFields.Count; t++)
                        {
                            var inputField = taxFields.ElementAt(t);
                            var fieldValue = inputField.Value;
                            
                            var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                            cc.Load(field);
                            cc.ExecuteQuery();
                            var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);

                            Guid _id = taxKeywordField.TermSetId;
                            string _termID = TermHelper.GetTermIdByName(cc, fieldValue, _id);

                            TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                            {
                                Label = fieldValue.ToString(),
                                TermGuid = _termID,
                            };
                            
                            
                            taxKeywordField.SetFieldValueByValue(item, termValue);
                            taxKeywordField.Update();
                        }
                        
                    }


                    //Modified needs to be updated last
                    string strModified = inputFields["Modified"];
                    Match matchModified = regex.Match(strModified);


                    if(matchModified.Success)
                    {
                        strModified = strModified.Replace("~t","");
                        

                        if(DateTime.TryParse(strModified, out DateTime dt))
                        {
                                item["Modified"] = dt;

                        }
                        item.Update();
                    }
                    //var ver = uploadFile.Versions;
                    //cc.Load(ver);
                    //cc.ExecuteQuery();

                    //uploadFile.CheckOut();
                    
                    
                    try
                    {
                        await cc.ExecuteQueryAsync();
                        Console.WriteLine("Successfully uploaded " + newFile.Url + " and updated metadata");
                    }
                    catch (System.Exception e)
                    {
                        _logger.LogError("Failed to update metadata.");
                        Console.WriteLine(e);
                        continue;
                    }


                }
            }
            catch (System.Exception)
            {
                
                throw;
            }
            finally
            {
                client.Logoff();
                client.Disconnect();
            }

            return new NoContentResult();

        }


        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> MigrationFolder([FromBody] JArray param)
        {
             // last batch is empty
            if (param.ToArray().Length == 0)
            {
                return new NoContentResult();
            }

            //Console.WriteLine(param);
            JObject tmpDoc = param.ToObject<List<JObject>>().FirstOrDefault();

            //string site = doc["site"].ToString();
            string site = tmpDoc["site"].ToString();
            
            string url = _baseurl + "sites/" + site;
            string tmpList = tmpDoc["list"].ToString();
            SMB2Client client = new SMB2Client();

            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                SMBCredential SMBCredential = new SMBCredential(){ 
                    username = Environment.GetEnvironmentVariable("smb_username"), 
                    password = _password, 
                    domain = Environment.GetEnvironmentVariable("domain"),
                    ipaddr = Environment.GetEnvironmentVariable("ipaddr"),
                    share = Environment.GetEnvironmentVariable("share"),
                };

                var serverAddress = System.Net.IPAddress.Parse(SMBCredential.ipaddr);
                bool success = client.Connect(serverAddress, SMBTransportType.DirectTCPTransport);

                NTStatus nts = client.Login(SMBCredential.domain, SMBCredential.username, SMBCredential.password);
                ISMBFileStore fileStore = client.TreeConnect(SMBCredential.share, out nts);
                
                List list = cc.Web.Lists.GetByTitle(tmpList);
                
            
                //cc.Load(list);
                //cc.ExecuteQuery();

                //List<Metadata> fields = SharePointHelper.GetFields(cc, list);


                //FolderCollection folders = SharePointHelper.GetFolders(cc, list);
                //string[] foldernames = SharePointHelper.GetFolderNames(cc, list, folders);
                List<JObject> docs = param.ToObject<List<JObject>>();

                for (int i = 0; i < docs.Count; i++)
                {
                    string filename = docs[i]["filename"].ToString();
                    string file_url = docs[i]["file_url"].ToString();
                    JObject inputFields = docs[i]["fields"] as JObject;
                    JObject taxFields = docs[i]["taxonomyfields"] as JObject;

                    FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename, SMBCredential, client, nts, fileStore);
                    ///FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename);
                    //FieldCollection fields = list.Fields;
                    
                    if (newFile == null){
                        _logger.LogError("Failed to upload. Skip: " + filename);
                        continue;
                    }

                    File uploadFile;
                    if (SharePointHelper.FolderJObjectExist(docs[i]) == false)
                    {
                        uploadFile = list.RootFolder.Files.Add(newFile);
                    }
                    else
                    {
                        string foldername = docs[i]["foldername"].ToString();
                        string sitecontent = docs[i]["sitecontent"].ToString();
                        
                        //Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                        Folder folder = SharePointHelper.GetFolder(cc, list, sitecontent);
                        if (folder == null)
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername);
                        
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);
                        
                    ListItem item = uploadFile.ListItemAllFields;


                    DateTime dtMin = new DateTime(1900,1,1);
                    Regex regex = new Regex(@"~t.*");
                    foreach (KeyValuePair<string, JToken> inputField in inputFields)
                    {

                        if (inputField.Value == null || inputField.Value.ToString() == "" || inputField.Key.Equals("Modified"))
                        {
                            continue;
                        }
                        
                        string fieldValue = (string)inputField.Value;
                        Match match = regex.Match(fieldValue);

                        if (inputField.Key.Equals("Author") || inputField.Key.Equals("Editor"))
                        {
                            var user = FieldUserValue.FromUser(fieldValue);
                            item[inputField.Key] = user;
                        }
                        else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By"))
                        {
                            StringBuilder sb = new StringBuilder("i:0#.f|membership|");
                            sb.Append(fieldValue);
                            item[inputField.Key] = sb;

                        }
                        else if(match.Success)
                        {
                            fieldValue = fieldValue.Replace("~t","");
                            if(DateTime.TryParse(fieldValue, out DateTime dt))
                            {
                                if(dtMin <= dt){
                                    item[inputField.Key] = dt;
                                    _logger.LogInformation("Set field " + inputField.Key + "to " + dt);
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            int tokenLength = inputField.Value.Count();

                            if(tokenLength >= 1){
                                continue;
                            }
                            else
                            {
                                item[inputField.Key] = fieldValue;
                                _logger.LogInformation("Set " + inputField.Key + " to " + fieldValue);
                                
                            }
                        }

                        
                        item.Update();
                    }

                    var clientRuntimeContext = item.Context;
                    if(taxFields != null)
                    {

                        foreach (KeyValuePair<string, JToken> inputField in taxFields)
                        {
                            var fieldValue = inputField.Value;
                            var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                            cc.Load(field);
                            cc.ExecuteQuery();
                            var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);

                            JObject taxObj = inputField.Value as JObject;

                            Guid _id = taxKeywordField.TermSetId;
                            string _termID = TermHelper.GetTermIdByName(cc, fieldValue.ToString(), _id);

                            TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                            {
                                Label = fieldValue.ToString(),
                                TermGuid = _termID,
                            };
                            
                            
                            taxKeywordField.SetFieldValueByValue(item, termValue);
                            taxKeywordField.Update();


                        }
                    }



                    //Modified needs to be updated last
                    string strModified = inputFields["Modified"].ToString();
                    Match matchModified = regex.Match(strModified);
                    if(matchModified.Success)
                    {
                        strModified = strModified.Replace("~t","");
                        if(DateTime.TryParse(strModified, out DateTime dt))
                        {
                                item["Modified"] = dt;
                                item.Update();
                        }
                    }
                    
                    try
                    {
                        await cc.ExecuteQueryAsync();
                        Console.WriteLine("Successfully uploaded " + newFile.Url + " and updated metadata");
                    }
                    catch (System.Exception e)
                    {
                        _logger.LogError("Failed to update metadata.");
                        Console.WriteLine(e);
                        continue;
                    }
                    

                }
                    
                    
                /// 1. use finally block to close open connection or to clean unused objects.
                /// 2. Use for-loop instead of foreach based on context (refer shared link)
                /// 3. see if you can remove un-nesessary use of try catch block
                /// 4. use timer to find out health of code  and real time consumed 
                /// 5. try to use private methods whenever possible
            
                
                //client.Logoff();
                
                //await cc.ExecuteQueryAsync();
                
                return new NoContentResult();
            }
            catch (System.Exception)
            {
                
                throw;
            }
            finally{
                client.Logoff();
                client.Disconnect();
            }
            

        }

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        /// <summary>
        /// 
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> MigrationOptimize([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }

            SMB2Client client = new SMB2Client();
            
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            Console.WriteLine(url);
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                
                SMBCredential SMBCredential = new SMBCredential(){ 
                    username = Environment.GetEnvironmentVariable("smb_username"), 
                    password = Environment.GetEnvironmentVariable("smb_password"), 
                    domain = Environment.GetEnvironmentVariable("domain"),
                    ipaddr = Environment.GetEnvironmentVariable("ipaddr"),
                    share = Environment.GetEnvironmentVariable("share"),
                };

                var serverAddress = System.Net.IPAddress.Parse(SMBCredential.ipaddr);
                bool success = client.Connect(serverAddress, SMBTransportType.DirectTCPTransport);

                NTStatus nts = client.Login(SMBCredential.domain, SMBCredential.username, SMBCredential.password);
                ISMBFileStore fileStore = client.TreeConnect(SMBCredential.share, out nts);


                List list = cc.Web.Lists.GetById(listGuid);

                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                //List list = cc.Web.Lists.GetByTitle(listname);

                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;
                    var taxListFields = docs[i].taxListFields;


                    FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename, SMBCredential, client, nts, fileStore);
                    ///FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename);
                
                    if (newFile == null){
                        _logger.LogError("Failed to upload. Skip: " + filename);
                        continue;
                    }

                    File uploadFile;
                    if(docs[i].foldername == null){
                        uploadFile = list.RootFolder.Files.Add(newFile);
                    }
                    else{
                        string foldername = docs[i].foldername;
                        string sitecontent = docs[i].sitecontent;
                        
                        //Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                        Folder folder = SharePointHelper.GetFolder(cc, list, foldername);
                        if (folder == null){
                            if(taxFields != null){
                                folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, inputFields, fields, taxFields);
                            }
                            else
                            {
                                folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                            }

                        }
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);

                    ListItem item = uploadFile.ListItemAllFields;
                    cc.Load(item);
                    cc.ExecuteQuery();
                    item["Title"] = filename;


                    DateTime dtMin = new DateTime(1900,1,1);
                    Regex regex = new Regex(@"~t.*");
                    var listItemFormUpdateValueColl = new List <ListItemFormUpdateValue>();
                    if (inputFields != null)
                    {    
                        foreach (KeyValuePair<string, string> inputField in inputFields)
                        {
                            if (inputField.Value == null || inputField.Value == "" || inputField.Key.Equals("Modified") || inputField.Key.Equals("SPORResponsibleRetired"))
                            {
                                continue;
                            }
                            

                            string fieldValue = inputField.Value;
                            Match match = regex.Match(fieldValue);
                            
                            Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                            if (field.TypeAsString.Equals("User"))
                            {
                                int uid = SharePointHelper.GetUserId(cc, fieldValue);

                                if(uid == 0){
                                    //user does not exist in AD. 
                                    //item["SPORResponsibleRetired"] = fieldValue;
                                    continue;
                                }
                                else
                                {
                                    item[inputField.Key] = new FieldUserValue{LookupId = uid};
                                    
                                }
                            }
                            //endre hard koding
                            else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By") || inputField.Key.Equals("Dokumentansvarlig"))
                            {
                                StringBuilder sb = new StringBuilder("i:0#.f|membership|");
                                sb.Append(fieldValue);
                                item[inputField.Key] = sb;
                            }
                            else if(match.Success)
                            {
                                fieldValue = fieldValue.Replace("~t","");
                                if(DateTime.TryParse(fieldValue, out DateTime dt))
                                {
                                    if(dtMin <= dt){
                                        item[inputField.Key] = dt;
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                item[inputField.Key] = fieldValue;

                            }

                            
                            item.Update();
                            
                            
                        }
                    }
                    cc.ExecuteQuery();
                    if (taxFields != null)
                    {
                        var clientRuntimeContext = item.Context;
                        for (int t = 0; t < taxFields.Count; t++)
                        {
                            var inputField = taxFields.ElementAt(t);
                            var fieldValue = inputField.Value;
                            if (fieldValue == null || fieldValue.Equals(""))
                            {
                                continue;
                            }
                            
                            var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                            cc.Load(field);
                            cc.ExecuteQuery();
                            var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);

                            Guid _id = taxKeywordField.TermSetId;
                            string _termID = TermHelper.GetTermIdByName(cc, fieldValue, _id);

                            TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                            {
                                Label = fieldValue.ToString(),
                                TermGuid = _termID,
                            };
                            
                            
                            taxKeywordField.SetFieldValueByValue(item, termValue);
                            taxKeywordField.Update();
                        }
                        
                    }
                    if (taxListFields != null)
                    {
                        var clientRuntimeContext = item.Context;
                        for (int t = 0; t < taxListFields.Count; t++)
                        {
                            var inputField = taxListFields.ElementAt(t);
                            var fieldListValue = inputFields.Values;

                            if (fieldListValue == null)
                            {
                                continue;
                            }
                            
                            var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                            cc.Load(field);
                            cc.ExecuteQuery();

                            var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);
                            Guid _id = taxKeywordField.TermSetId;
                            foreach (var fieldValue in fieldListValue)
                            {
                                string _termID = TermHelper.GetTermIdByName(cc, fieldValue, _id);

                                TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                                {
                                    Label = fieldValue.ToString(),
                                    TermGuid = _termID,
                                };
                                
                                taxKeywordField.SetFieldValueByValue(item, termValue);
                                taxKeywordField.Update();
                                
                            }


                            
                        }
                    }


                    //Modified needs to be updated last
                    string strModified = inputFields["Modified"];
                    Match matchModified = regex.Match(strModified);


                    if(matchModified.Success)
                    {
                        strModified = strModified.Replace("~t","");
                        

                        if(DateTime.TryParse(strModified, out DateTime dt))
                        {
                                item["Modified"] = dt;

                        }
                        item.Update();
                    }
                    
                    
                    try
                    {
                        await cc.ExecuteQueryAsync();
                        Console.WriteLine("Successfully uploaded " + newFile.Url + " and updated metadata");
                    }
                    catch (System.Exception e)
                    {
                        _logger.LogError("Failed to update metadata.");
                        Console.WriteLine(e);
                        continue;
                    }


                }
            }
            catch (System.Exception)
            {
                
                throw;
            }
            finally
            {
                client.Logoff();
                client.Disconnect();
            }

            return new NoContentResult();

        }

        [HttpGet]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> eDocsDokumentnavn([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname, [FromQuery(Name = "eDocsDokumentnavn")] string eDocsDokumentnavn)
        {
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                var cquery = new CamlQuery();
                    cquery.ViewXml = string.Format(
                        @"<View>  
                            <Query> 
                                <Where>
                                    <Eq><FieldRef Name='eDocsDokumentnavn' />
                                    <Value Type='Text'>{0}</Value></Eq>
                                </Where> 
                            </Query> 
                        </View>", eDocsDokumentnavn);

                var listitems = list.GetItems(cquery);
                cc.Load(listitems);
                await cc.ExecuteQueryAsync();

                return new NoContentResult();
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
        /// Return user id
        /// </summary>
        /// <param name=""site""></param>
        /// <returns></returns>
        public int userId([FromQuery(Name = "site")] string sitename, [FromQuery(Name = "name")] string uname)
        {
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                
                var otheruser = cc.Web.EnsureUser(uname);
                cc.Load(otheruser, u => u.Id);
                cc.ExecuteQuery();
                
                return otheruser.Id;
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
        /// Use only on lists with over 5000 documents
        /// 
        /// eDocsDokumentnavn is a indexed field.
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> documentfix([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return null;
            }
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetById(listGuid);
                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    string eDocsDokumentnavn = inputFields["eDocsDokumentnavn"];
                    var cquery = new CamlQuery();
                    if(string.IsNullOrEmpty(foldername)){
                            cquery.ViewXml = string.Format(
                            @"<View>  
                                <Query> 
                                    <Where>
                                        <Eq><FieldRef Name='eDocsDokumentnavn' />
                                        <Value Type='Text'>{0}</Value></Eq>
                                    </Where> 
                                </Query> 
                            </View>", eDocsDokumentnavn);

                    }
                    else
                    {
                        cquery.ViewXml = string.Format(
                            @"<View>  
                                <Query> 
                                    <Where>
                                        <Eq><FieldRef Name='eDocsDokumentnavn' />
                                        <Value Type='Text'>{0}</Value></Eq>
                                        <Eq><FieldRef Name='FileDirRef' />
                                        <Value Type='Text'>{1}</Value></Eq>
                                    </Where> 
                                </Query> 
                            </View>", eDocsDokumentnavn, foldername);
                        
                    }

                    var listitems = list.GetItems(cquery);
                    cc.Load(listitems);
                    await cc.ExecuteQueryAsync();

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    if (listitems.Count > 0)
                    {
                        //ListItem item = listitems[0];
                        foreach (var item in listitems)
                        {
                            if (filename.Equals(item["FileLeafRef"]))
                            {
                                
                                _logger.LogInformation("fix: " + filename);
                                if (inputFields != null)
                                {
                                    foreach (KeyValuePair<string, string> inputField in inputFields)
                                    {

                                        if (inputField.Value == null || inputField.Value == "")
                                        {
                                            continue;
                                        }

                                        string fieldValue = inputField.Value;
                                        Match match = regex.Match(fieldValue);

                                        if (inputField.Key.Equals("Author") || inputField.Key.Equals("Editor"))
                                        {
                                            if (int.TryParse(fieldValue, out int uid))
                                            {
                                                item[inputField.Key] = new FieldUserValue{LookupId = uid};
                                            }
                                            else
                                            {
                                                FieldUserValue user = FieldUserValue.FromUser(fieldValue);
                                                item[inputField.Key] = user;                                    
                                            }

                                        }
                                        //endre hard koding
                                        else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By") || inputField.Key.Equals("Dokumentansvarlig"))
                                        {
                                            string user = "i:0#.f|membership|" + fieldValue;
                                            item[inputField.Key] = user;
                                        }
                                        else if(match.Success)
                                        {
                                            fieldValue = fieldValue.Replace("~t","");
                                            if(DateTime.TryParse(fieldValue, out DateTime dt))
                                            {
                                                if(dtMin <= dt){
                                                    item[inputField.Key] = dt;
                                                }
                                                else
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            item[inputField.Key] = fieldValue;                                                
                                        }
                                    }
                                    item.Update();
                                }
                            }
                        }
                        
                        await cc.ExecuteQueryAsync();
                        _logger.LogInformation("updated: " + filename);
                    

                    }
                    else
                    {
                        _logger.LogInformation("file not found: " + eDocsDokumentnavn);
                        continue;
                    }
                    
                }
                
            }
            catch (System.Exception)
            {
                
                throw;
            }

            return new NoContentResult();
        }

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        /// <summary>
        /// Update existing document with SystemUpdate() to prevent version increment.
        /// POST /api/sharepoint/document
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> Document([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return null;
            }

            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                var listItemFormUpdateValueColl = new List <ListItemFormUpdateValue>();
                //List list = cc.Web.Lists.GetByTitle(listname);
                List list = cc.Web.Lists.GetById(listGuid);
                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    Folder root = list.RootFolder;
                    File file;

                    if (foldername != null)
                    {
                        Folder folder = root.Folders.GetByUrl(foldername);
                        file = folder.Files.GetByUrl(filename);
                    }
                    else
                        file = root.Files.GetByUrl(filename);

                    //cc.Load(file);
                    //cc.ExecuteQuery();

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    ListItem item = file.ListItemAllFields;
                    if (inputFields != null)
                    {
                        foreach (KeyValuePair<string, string> inputField in inputFields)
                        {

                            if (inputField.Value == null || inputField.Value == "" || inputField.Key.Equals("Modified"))
                            {
                                continue;
                            }

                            string fieldValue = inputField.Value;
                            Match match = regex.Match(fieldValue);

                            if (inputField.Key.Equals("Author") || inputField.Key.Equals("Editor"))
                            {
                                FieldUserValue user = FieldUserValue.FromUser(fieldValue);
                                
                                item[inputField.Key] = user;

                            }
                            //endre hard koding
                            else if (inputField.Key.Equals("Modified_x0020_By") || inputField.Key.Equals("Created_x0020_By") || inputField.Key.Equals("Dokumentansvarlig"))
                            {
                                StringBuilder sb = new StringBuilder("i:0#.f|membership|");
                                sb.Append(fieldValue);
                                item[inputField.Key] = sb;
                            }
                            else if(match.Success)
                            {
                                fieldValue = fieldValue.Replace("~t","");
                                if(DateTime.TryParse(fieldValue, out DateTime dt))
                                {
                                    if(dtMin <= dt){
                                        item[inputField.Key] = dt;
                                        
                                    }
                                    else
                                    {
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                int tokenLength = inputField.Value.Count();
                                item[inputField.Key] = fieldValue;
                                    
                                    
                            }

                            item.SystemUpdate();
                        }
                    }
                    try
                    {
                        
                        await cc.ExecuteQueryAsync();
                    }
                    catch (System.Exception)
                    {
                        _logger.LogDebug("file note found: " + filename);
                        continue;
                    }

                }

                
                
                


            }
            catch (System.Exception)
            {
                
                throw;
            }

            return new NoContentResult();
        }

        [HttpGet]
        public int CountFiles([FromQuery(Name = "site")] string site,[FromQuery(Name = "list")] string listname)
        {

            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                
                var folder = list.RootFolder;
                cc.Load(folder, f => f.ItemCount);
                cc.ExecuteQuery();
                
                return folder.ItemCount;
            }
            catch (System.Exception)
            {
                
                throw;
            }

        }

        



        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> CreateFolder([FromBody] DocumentModel[] docs)
        {
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);

                for (int i = 0; i < docs.Length; i++)
                {
                    string foldername = docs[i].foldername;
                    string sitecontent = docs[i].sitecontent;
                    

                    ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
                    newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                    newItemInfo.LeafName = foldername;

                    ListItem newListItem = list.AddItem(newItemInfo);
                    newListItem.Update();

                    await cc.ExecuteQueryAsync();

                }

            }
            catch (System.Exception)
            {
                
                throw;
            }

            return new NoContentResult();
        }

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> CreateDocumentSet([FromBody] DocumentModel[] docs)
        {
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);

                for (int i = 0; i < docs.Length; i++)
                {
                    string foldername = docs[i].foldername;
                    string sitecontent = docs[i].sitecontent;
                    var taxonomies = docs[i].taxFields;
                    var inputFields = docs[i].fields;
                    
                    //Folder folder;
                    if (taxonomies != null)
                    
                        SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, inputFields, fields, taxonomies);
                    else
                        SharePointHelper.CreateFolder(cc, list, sitecontent, foldername);

                    await cc.ExecuteQueryAsync();
                }

                
            }
            catch (System.Exception)
            {
                
                throw;
            }

            return new NoContentResult();
        }


    }
}