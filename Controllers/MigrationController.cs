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
using System.Collections.ObjectModel;

using SharePointAPI.Middleware;
using SharePointAPI.Models;

namespace SharePointAPI.Controllers
{
    [ApiController]
    [Produces("application/json")]
    [Route("api/[controller]/[action]")]
    public class MigrationController : ControllerBase
    {
        private readonly ILogger<SharePointController> _logger;
         public string _username, _password, _baseurl;
        public MigrationController(ILogger<SharePointController> logger)
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
        public async Task<IActionResult> td([FromBody] DocumentModel[] docs)
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
                            folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, inputFields, fields, taxFields);
                        else if (folder == null)
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);

                    ListItem item = uploadFile.ListItemAllFields;


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

                    DateTime dtMin = new DateTime(1900,1,1);
                    Regex regex = new Regex(@"~t.*");
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
        /// 
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> Test([FromBody] DocumentModel[] docs)
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
                
                ///SMBCredential SMBCredential = new SMBCredential(){ 
                ///    username = Environment.GetEnvironmentVariable("smb_username"), 
                ///    password = Environment.GetEnvironmentVariable("smb_password"), 
                ///    domain = Environment.GetEnvironmentVariable("domain"),
                ///    ipaddr = Environment.GetEnvironmentVariable("ipaddr"),
                ///    share = Environment.GetEnvironmentVariable("share"),
                ///};
///
                ///var serverAddress = System.Net.IPAddress.Parse(SMBCredential.ipaddr);
                ///bool success = client.Connect(serverAddress, SMBTransportType.DirectTCPTransport);
///
                ///NTStatus nts = client.Login(SMBCredential.domain, SMBCredential.username, SMBCredential.password);
                ///ISMBFileStore fileStore = client.TreeConnect(SMBCredential.share, out nts);


                List list = cc.Web.Lists.GetById(listGuid);

                var fieldcollection = list.Fields;
                cc.Load(fieldcollection);
                cc.ExecuteQuery();
                var cquery = new CamlQuery();
                cquery.ViewXml = string.Format(@"<View>  
                                <Query> 
                                    <Where>
                                        <Eq><FieldRef Name='Title' />
                                        <Value Type='Text'>{0}</Value></Eq>
                                    </Where> 
                                </Query> 
                            </View>", "dummy.pdf");
                var listitems = list.GetItems(cquery);
                cc.Load(listitems);
                cc.ExecuteQuery();
                ListItem item = listitems[0];
                var clientRuntimeContext = item.Context;
                string[] SPORDocPropertyNo = { 
                    "4223-7/1",
                    "4223-7/3",
                    "4223-7/4",
                    "4223-7/5"};
    
                foreach (var fieldObj in fieldcollection)
                {
                    if (fieldObj.InternalName.Equals("SPORDocPropertyNo"))
                    {
                        var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(fieldObj);
                        Guid _id = taxKeywordField.TermSetId;
                        
                        List<string> ListTermString = new List<string>();
                        for (int i = 0; i < SPORDocPropertyNo.Length; i++)
                        {
                            string _termID = TermHelper.GetTermIdByName(cc, SPORDocPropertyNo[i], _id);
                            
                            ListTermString.Add(string.Format("-1;#{0}{1}{2}",SPORDocPropertyNo[i], "|", _termID));

                            
                        }
                        string tax = string.Join(";#", ListTermString);
                        Console.WriteLine(tax);

                       
                        taxKeywordField.SetFieldValueByValueCollection(item, new TaxonomyFieldValueCollection(cc, tax, taxKeywordField));
                        taxKeywordField.Update();
                        break;
                    }
                }
                item.Update();
                await cc.ExecuteQueryAsync();

            }
            catch (System.Exception)
            {
                
                throw;
            }
            finally
            {
                //client.Logoff();
                //client.Disconnect();
            }

            return new NoContentResult();

        }

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]
        /// <summary>
        /// 
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> sporaen_eiendomsarkivet([FromBody] DocumentModel[] docs)
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

                var fieldcollection = list.Fields;
                cc.Load(fieldcollection);
                cc.ExecuteQuery();
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

                    File uploadFile = list.RootFolder.Files.Add(newFile);
                    
                    _logger.LogInformation("Upload file: " + newFile.Url);

                    ListItem item = uploadFile.ListItemAllFields;
                    item["Title"] = filename;
                    cc.ExecuteQuery();


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
                            
                            //Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                            foreach (var field in fieldcollection)
                            {
                                if (field.InternalName.Equals(inputField.Key))
                                {
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
                                    continue;
                                    
                                }
                            }
                            
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
                            if (fieldValue == null || fieldValue.Equals(""))
                            {
                                continue;
                            }
                            
                            //var field = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                            //cc.Load(field);
                            //cc.ExecuteQuery();
                            foreach (var field in fieldcollection)
                            {
                                if (field.InternalName.Equals(inputField.Key))
                                {
                                    var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);
                                    Guid _id = taxKeywordField.TermSetId;
                                    string _termID = TermHelper.GetTermIdByName(cc, fieldValue, _id);
                                    if (_termID.Equals(""))
                                    {
                                        continue;
                                    }
                                    TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                                    {
                                        Label = fieldValue.ToString(),
                                        TermGuid = _termID,
                                    };
                                    
                                    
                                    taxKeywordField.SetFieldValueByValue(item, termValue);
                                    taxKeywordField.Update();
                                    item.Update();
                                    cc.ExecuteQuery();
                                }
                            }

                        }
                        
                    }
                    if (taxListFields != null)
                    {
                        var clientRuntimeContext = item.Context;
                        for (int t = 0; t < taxListFields.Count; t++)
                        {
                            var inputField = taxListFields.ElementAt(t);
                            var fieldListValue = inputField.Value;

                            if (fieldListValue == null)
                            {
                                continue;
                            }
                            
                            foreach (var field in fieldcollection)
                            {
                                if (field.InternalName.Equals(inputField.Key))
                                {
                                    var taxKeywordField = clientRuntimeContext.CastTo<TaxonomyField>(field);
                                    Guid _id = taxKeywordField.TermSetId;
                                    
                                    List<string> ListTermString = new List<string>();
                                    for (int f = 0; f < fieldListValue.Count; f++)
                                    {
                                        string _termID = TermHelper.GetTermIdByName(cc, fieldListValue[f], _id);
                                        
                                        ListTermString.Add(string.Format("-1;#{0}{1}{2}",fieldListValue[f], "|", _termID));

                                        
                                    }
                                    string tax = string.Join(";#", ListTermString);

                                
                                    taxKeywordField.SetFieldValueByValueCollection(item, new TaxonomyFieldValueCollection(cc, tax, taxKeywordField));
                                    taxKeywordField.Update();
                                }
                            }
                            item.Update();
                            await cc.ExecuteQueryAsync();


                            
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

        
        
        
    }
}