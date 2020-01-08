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
        public async Task<IActionResult> DocumentsWithFields()
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            var SPDocs = new List<JObject>();
            string site = "Sesamsitewithdocumentsets";
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                
                Web web = cc.Web;
                cc.Load(web);
                var Lists = web.Lists;
                cc.Load(Lists);
                await cc.ExecuteQueryAsync();

                List list = web.Lists.GetByTitle("Documents");
                cc.Load(list, l => l.Fields);
                await cc.ExecuteQueryAsync();
                List<string> fieldNames = new List<string>();
   

                foreach (var tmpfield in list.Fields)
                {
                    if(!tmpfield.FromBaseType && tmpfield.Hidden == false){
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
                            
                            json.Add(new JProperty(fieldname, item[fieldname]));
                            Console.WriteLine("Name: " + fieldname + "Value: " + item[fieldname]);
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
        ///     POST /api/sharepoint/newdoc
        ///     {
        ///         "ListName":"Documents",
	    ///         "FileUrl":"https://sesam.io/images/Cyan.svg",
	    ///         "FolderName":"My first document set",
        ///         "Site": "Sesamsitewithdocumentsets",
        ///         "Filename": "Cyan.svg",
	    ///         "Fields":{
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
        public async Task<IActionResult> newdoc([FromBody] JObject doc)
        {

            string site = doc["Site"].ToString();
            string url = _baseurl + "sites/" + site;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try{
                
                Web web = cc.Web;
                cc.Load(web);
                var lists = web.Lists;
                cc.Load(lists);
                await cc.ExecuteQueryAsync();
                Console.WriteLine(doc["ListName"].ToString());
                List list = web.Lists.GetByTitle(doc["ListName"].ToString());

                FileCreationInformation newFile = new FileCreationInformation();
                //byte[] imageBytes = webClient.DownloadData("https://sesam.io/images/howitworks.jpg");
                using (var webClient = new WebClient()){
                    byte[] imageBytes = webClient.DownloadData(doc["FileUrl"].ToString());
                    newFile.Content = imageBytes;
                }
                //Regex match /\w+\.\w+$/g
                /*Regex rg = new Regex(@"\w+\.\w+$"); 
                var match = rg.Match(doc["FileUrl"].ToString());
                if(match.Success){
                    //Document filename
                    Console.WriteLine(newFile.Url);
                    newFile.Url = match.Value;    
                }*/

                newFile.Url = doc["Filename"].ToString();
                
                //Folder folder = list.RootFolder.Folders.GetByUrl("My first document set");
                Folder folder = list.RootFolder.Folders.GetByUrl(doc["FolderName"].ToString());
                File uploadFile = folder.Files.Add(newFile);
                
                ListItem item = uploadFile.ListItemAllFields;
                
                JObject fields = doc["Fields"] as JObject;
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
    }
}