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
                        //test
                        cc.Load(item);
                        await cc.ExecuteQueryAsync();

                        var json = new JObject();
                        json.Add(new JProperty("filename", file.Name));
                        json.Add(new JProperty("folder", folder.Name));
                        json.Add(new JProperty("uri", file.LinkingUri));
                        foreach (var fieldname in fieldNames)
                        {
                            
                            if (item[fieldname] != null)
                            {
                                Regex rg = new Regex(@"Microsoft\.SharePoint\.Client\..*");
                                var match = rg.Match(item[fieldname].ToString());
                                if(match.Success){
                                    if(fieldname.Equals("SPORResponsible") == true){
                                        FieldUserValue fieldUserValue = item[fieldname] as FieldUserValue;
                                        json.Add(new JProperty(fieldname, fieldUserValue.Email));

                                    }
                                    else
                                    {
                                        TaxonomyFieldValue taxonomyFieldValue = item[fieldname] as TaxonomyFieldValue;
                                        json.Add(new JProperty(fieldname, taxonomyFieldValue.Label));
                                    }
                                }
                                else
                                {
                                    json.Add(new JProperty(fieldname, item[fieldname]));  
                                }

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
        ///        		
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
        public async Task<IActionResult> NewDocument([FromBody] JObject doc)
        {
            

            string site = doc["site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{

                List list = SharePointHelper.GetListItemByTitle(cc, doc["list"].ToString());

                FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(doc["file_url"].ToString(), doc["filename"].ToString());
                File uploadFile;
                if (SharePointHelper.FolderJObjectExist(doc) == false)
                {
                    uploadFile = list.RootFolder.Files.Add(newFile);
                    Console.WriteLine("folder missing!!!!");
                }
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

                FieldCollection fields = SharePointHelper.GetFields(cc, list);

                
                JObject inputFields = doc["fields"] as JObject;
                //Add metadata
                SharePointHelper.SetMetadataFields(cc, inputFields, fields, item);                
                
                
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
                Web web = cc.Web;
                await cc.ExecuteQueryAsync();
                cc.Load(web);                
                var lists = web.Lists;
                cc.Load(lists);
                await cc.ExecuteQueryAsync();
                
                List list = web.Lists.GetByTitle(param["list"].ToString());

                Folder folder = list.RootFolder.Folders.GetByUrl(param["foldername"].ToString());
                cc.Load(folder);
                await cc.ExecuteQueryAsync();
                
                var items = folder.Files;
                //cc.Load(items);
                //await cc.ExecuteQueryAsync();
                
                //await cc.ExecuteQueryAsync();
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
                    ///if (int.TryParse(fieldValue, out int n)){
                    ///    item[inputField.Key] = n;
                    ///}
                    ///else if(DateTime.TryParse(fieldValue, out DateTime dt))
                    ///{
                    ///    item[inputField.Key] = dt;
                    ///}
                    ///else
                    ///{
                    ///    item[inputField.Key] = fieldValue;
                    ///}
                    
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
        public async Task<IActionResult> UpMetadata([FromBody] JObject doc)
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

                Folder folder = list.RootFolder.Folders.GetByUrl(doc["foldername"].ToString());
                ListItem listItem = folder.ListItemAllFields;
                cc.Load(listItem);
                cc.ExecuteQuery();

                var clientRuntimeContext = listItem.Context;


                JObject inputFields = doc["fields"] as JObject;
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
                listItem.Update();
                cc.Load(listItem);
                await cc.ExecuteQueryAsync();
                
                return new NoContentResult();



            }
            catch (Exception ex)
            {
                Console.WriteLine("Error message: " + ex);
                return StatusCode(StatusCodes.Status500InternalServerError, ex);
            }
        }



    }
}