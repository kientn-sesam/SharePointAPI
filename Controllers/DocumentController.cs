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
    public class DocumentController : ControllerBase
    {
        private readonly ILogger<SharePointController> _logger;
        public string _username, _password, _baseurl;
        public DocumentController(ILogger<SharePointController> logger)
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

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        public async Task<IActionResult> sesamtest([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }
            
             
            string site = docs[0].site;
            string url = _baseurl + "teams/" + site;
            string listname = docs[0].list;
            
            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                //cc.Load(list);
                //await cc.ExecuteQueryAsync();
                
                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    FileCreationInformation newFile = SharePointHelper.GetFileCreationInformation(file_url, filename);
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
                            folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, taxFields);
                        else if (folder == null)
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    DateTime dtMin = new DateTime(1900,1,1);
                    Regex regex = new Regex(@"~t.*");
                    ListItem item = uploadFile.ListItemAllFields;

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
                            
                            Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                            if (field.TypeAsString.Equals("User"))
                            {
                                int uid = SharePointHelper.GetUserId(cc, fieldValue);
                                item[inputField.Key] = new FieldUserValue{LookupId = uid};
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

                Console.WriteLine("Hellloooooo");
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
        /// Use only on lists with over 5000 documents
        /// 
        /// eDocsDokumentnavn is a indexed field.
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> UpdateOverwriteVersion([FromBody] DocumentModel[] docs)
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
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    string eDocsDokumentnavn = inputFields["eDocsDokumentnavn"];
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
                    cc.Load(listitems, items => items.Include(
                        item => item.File
                    ));
                    await cc.ExecuteQueryAsync();

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    if (listitems.Count > 0)
                    {
                        ListItem item = listitems[0];
                        var file = item.File;

                        Console.WriteLine(file.CheckOutType.ToString());
                        if (file.CheckOutType == CheckOutType.None)
                        {
                            item.File.CheckOut();
                            cc.ExecuteQuery();
                        }

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

                                Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                                if (field.TypeAsString.Equals("User"))
                                {
                                    int uid = SharePointHelper.GetUserId(cc, fieldValue);
                                    item[inputField.Key] = new FieldUserValue{LookupId = uid};
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
                        }
                        item.File.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                        
                        await cc.ExecuteQueryAsync();
                    

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
        public async Task<IActionResult> MigrationWithVersioning([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }

            SMB2Client client = new SMB2Client();
            
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
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
                            folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, taxFields);
                        else if (folder == null)
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);

                    ListItem item = uploadFile.ListItemAllFields;
                    if (uploadFile.CheckOutType == CheckOutType.None)
                    {
                        uploadFile.CheckOut();
                        cc.ExecuteQuery();
                    }


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
                            
                            Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                            if (field.TypeAsString.Equals("User"))
                            {
                                int uid = SharePointHelper.GetUserId(cc, fieldValue);
                                item[inputField.Key] = new FieldUserValue{LookupId = uid};
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

                    }
                    _logger.LogInformation("taxfield: " + taxFields);
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
                    uploadFile.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                    
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
        /// <summary>
        /// Use only on lists with over 5000 documents
        /// 
        /// eDocsDokumentnavn is a indexed field.
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> test([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return null;
            }
            string site = docs[0].site;
            string url = _baseurl + "teams/" + site;
            string listname = docs[0].list;
            //Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetByTitle(listname);
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    //string eDocsDokumentnavn = inputFields["eDocsDokumentnavn"];
                    var cquery = new CamlQuery();
                    cquery.ViewXml = string.Format(
                        @"<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Eq></Where></Query></View>", 5);

                    var listitems = list.GetItems(cquery);
                    cc.Load(listitems);
                    cc.Load(listitems, items => items.Include(
                        item => item.File
                    ));
                    await cc.ExecuteQueryAsync();
                    

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    if (listitems.Count > 0)
                    {
                        ListItem item = listitems[0];
                        var file = item.File;

                        Console.WriteLine(file.CheckOutType.ToString());
                        if (file.CheckOutType == CheckOutType.None)
                        {
                            item.File.CheckOut();
                            cc.ExecuteQuery();
                        }

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
                                

                                Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                                if (field.TypeAsString.Equals("User"))
                                {
                                    int uid = SharePointHelper.GetUserId(cc, fieldValue);
                                    item[inputField.Key] = new FieldUserValue{LookupId = uid};
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
                                    //item.Update();
                                }

                                item.Update();
                            }
                        }
                        item.File.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);
                        
                        await cc.ExecuteQueryAsync();
                    

                    }
                    else
                    {
                        //_logger.LogInformation("file not found: " + eDocsDokumentnavn);
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
        [Produces("application/json")]
        [Consumes("application/json")]

        /// <summary>
        /// Create a new doc
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/sharepoint/documents?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> All([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                Console.WriteLine(_baseurl);
                //Guid listGuid = new Guid(listname);
                //List list = cc.Web.Lists.GetById(listGuid);
                //var lists = cc.Web.Lists;
                //cc.Load(lists);
                //cc.ExecuteQuery();
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




    }
}