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
        /// <summary>
        /// Enrich metadata on documentset only
        /// 
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> FolderEnrichment([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);

            using (ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                List list = cc.Web.Lists.GetById(listGuid);
                cc.Load(list);
                cc.ExecuteQuery();
                

                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                for (int i = 0; i < docs.Length; i++)
                {

                    //Console.WriteLine(JsonConvert.SerializeObject(docs[i],Formatting.Indented));
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;
                    Console.WriteLine("Updating:" + foldername);


                    Folder folder = list.RootFolder.Folders.GetByUrl(foldername);
                    ListItem item = folder.ListItemAllFields;
                    cc.Load(item);
                    await cc.ExecuteQueryAsync();

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);

                        
                    if (inputFields != null)
                    {
                        var clientRuntimeContext = item.Context;
                        foreach (KeyValuePair<string, string> inputField in inputFields)
                        {

                            if (inputField.Value == null)
                            {
                                continue;
                            }

                            string fieldValue = inputField.Value;
                            Match match = regex.Match(fieldValue);

                            Metadata field = fields.Find(x => x.InternalName.Equals(inputField.Key));
                            if (field.TypeAsString.Equals("User"))
                            {
                                int uid = SharePointHelper.GetUserId(cc, fieldValue);
                                if (uid == 0)
                                {
                                    continue;
                                }
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
                    try
                    {
                        
                        await cc.ExecuteQueryAsync();
                        Console.WriteLine("updated: " + foldername);
                        Console.WriteLine("------------------------------------------------------");
                        
                    }
                    catch (System.Exception e)
                    {
                         _logger.LogError("Failed to update metadata.");
                        Console.WriteLine(e);
                        continue;
                        
                    }

                    if (taxFields.Count > 0)
                    {
                        var clientRuntimeContext = item.Context;
                        for (int t = 0; t < taxFields.Count; t++)
                        {
                            var inputField = taxFields.ElementAt(t);
                            var fieldValue = inputField.Value;
                            if (string.IsNullOrEmpty(fieldValue))
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
                            item.Update();
                            
                            try
                            {
                                
                                await cc.ExecuteQueryAsync();
                            }
                            catch (System.Exception e)
                            {
                                
                                _logger.LogError("Failed to update taxonomy metadata.");
                                Console.WriteLine(e);
                                continue;
                            }
                        }

                        
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
                    ///if(foldername == null){
                    ///        cquery.ViewXml = string.Format(
                    ///        @"<View>  
                    ///            <Query> 
                    ///                <Where>
                    ///                    <Eq><FieldRef Name='eDocsDokumentnavn' />
                    ///                    <Value Type='Text'>{0}</Value></Eq>
                    ///                </Where> 
                    ///            </Query> 
                    ///        </View>", eDocsDokumentnavn);
                    ///
                    ///}
                    ///else
                    ///{
                    ///    cquery.ViewXml = string.Format(
                    ///        @"<View>  
                    ///            <Query> 
                    ///                <Where>
                    ///                    <Eq><FieldRef Name='eDocsDokumentnavn' />
                    ///                    <Value Type='Text'>{0}</Value></Eq>
                    ///                    <Eq><FieldRef Name='FileDirRef' />
                    ///                    <Value Type='Text'>{1}</Value></Eq>
                    ///                </Where> 
                    ///            </Query> 
                    ///        </View>", eDocsDokumentnavn, foldername);
                    /// 
                    ///}

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
        /// <summary>
        /// Use only on lists with over 5000 documents
        /// 
        /// NB! search on indexed field eDocsDokumentnavn.
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> UpdateWithoutVersioning([FromBody] DocumentModel[] docs)
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
                cc.Load(list, 
                    l => l.EntityTypeName,
                    l => l.Fields.Include(
                        f => f.InternalName,
                        f => f.Title,
                        f => f.TypeAsString
                    ));
                cc.ExecuteQuery();
                List<Metadata> fields = SharePointHelper.GetFields(list);
                string entityTypeName = list.EntityTypeName;
                

                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;

                    var cquery = new CamlQuery();
                    cquery.ViewXml = string.Format(
                            @"<View>  
                                <Query> 
                                    <Where>
                                        <Eq><FieldRef Name='FileLeafRef' />
                                        <Value Type='Text'>{0}</Value></Eq>
                                    </Where> 
                                </Query> 
                            </View>", filename);

                    if(foldername != null)
                        cquery.FolderServerRelativeUrl = "/sites/" + site + "/" + entityTypeName + "/" + foldername;
                    
                        
                    var listitems = list.GetItems(cquery);
                    cc.Load(listitems);

                    await cc.ExecuteQueryAsync();

                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    if (listitems.Count > 0)
                    {
                        ListItem item = listitems[0];
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
                        
                        await cc.ExecuteQueryAsync();
                    

                    }
                    else
                    {
                        _logger.LogError("file not found: " + filename);
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
        ///     Migration library for big files
        ///     Using SMB2Client to fetch files from smb file server.
        ///     Adds metadata for normal and taxonomy fields
        ///     [ 
        ///         {
        ///           "fields": {
        ///             "Created_x0020_By": "",
        ///             "eDocsRNLinjer": null,
        ///             "eDocsSaksbehandler": "RXINDEX IMPORT",
        ///             ....
        ///           },
        ///           "file_url": "//.../...",
        ///           "filename": "<string>",
        ///           "foldername": "<string>",
        ///           "list": "<guid>",
        ///           "site": "<site name>",
        ///           "sitecontent": "<site content name for creating document set>"
        ///         }
        ///     ]
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> MigrationXLFiles([FromBody] DocumentModel[] docs)
        {
            if (docs.Length == 0)
            {
                return new NoContentResult();
            }

            SMB2Client client = new SMB2Client();
            
            string site = docs[0].site;
            string url = _baseurl + "teams/" + site;
            Console.WriteLine(url);
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
   
                List list = cc.Web.Lists.GetByTitle(listname);
                cc.Load(list.RootFolder, p => p.ServerRelativeUrl);

                List<Metadata> fields = SharePointHelper.GetFields(cc, list);

                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;
                    var taxListFields = docs[i].taxListFields;
                    var doc = docs[i];

                    Microsoft.SharePoint.Client.File uploadFile = null;
            ClientResult<long> bytesUploaded = null;
            //SMBLibrary.NTStatus actionStatus;
            FileCreationInformation newFile = new FileCreationInformation();
            NTStatus status = nts;
        
            object handle;
            FileStatus fileStatus;
            string tmpfile = System.IO.Path.GetTempFileName();
            status = fileStore.CreateFile(out handle, out fileStatus, doc.file_url, AccessMask.GENERIC_READ, 0, ShareAccess.Read, CreateDisposition.FILE_OPEN, CreateOptions.FILE_NON_DIRECTORY_FILE, null);
            if (status != NTStatus.STATUS_SUCCESS)
            {
                Console.WriteLine(status);
                return null;
            }
            else{
                //string uniqueFileName = String.Empty;
                int blockSize = 8000000; // 8 MB
                long fileSize;
                Guid uploadId = Guid.NewGuid();
                
                byte[] buf;
                var fs = new System.IO.FileStream(tmpfile, System.IO.FileMode.OpenOrCreate);
                var bw = new System.IO.BinaryWriter(fs);
                int bufsz = 64 * 1000;
                int j = 0;
                
                do{
                    status = fileStore.ReadFile(out buf, handle, i * bufsz, bufsz);
                    if (status == NTStatus.STATUS_SUCCESS)
                    {
                        int n = buf.GetLength(0);
                        
                        bw.Write(buf, 0, n);
                        if (n < bufsz) break;
                        i++;
                    }
                
                }
                while (status != NTStatus.STATUS_END_OF_FILE && j < 1000);
                
                if (status == NTStatus.STATUS_SUCCESS)
                {
                    fileStore.CloseFile(handle);
                    bw.Flush();
                    fs.Close();
                    //fs = System.IO.File.OpenRead(tmpfile);
                    
                    //byte[] fileBytes = new byte[fs.Length];
                    //fs.Read(fileBytes, 0, fileBytes.Length);
                    try
                    {
                        
                        fs = System.IO.File.Open(tmpfile, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                        fileSize = fs.Length;
                        //uniqueFileName = System.IO.Path.GetFileName(fs.Name);
                        using (System.IO.BinaryReader br = new System.IO.BinaryReader(fs))
                        {
                            byte[] buffer = new byte[blockSize];
                            byte[] lastBuffer = null;
                            long fileoffset = 0;
                            long totalBytesRead = 0;
                            int bytesRead;
                            bool first = true;
                            bool last = false;

                            while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                totalBytesRead = totalBytesRead + bytesRead;
                                if (totalBytesRead >= fileSize)
                                {
                                    last = true;
                                    lastBuffer = new byte[bytesRead];
                                    Array.Copy(buffer, 0, lastBuffer, 0 , bytesRead);
                                }

                                if (first)
                                {
                                    using (System.IO.MemoryStream contentStream = new System.IO.MemoryStream())
                                    {
                                        newFile.ContentStream = contentStream;
                                        newFile.Url = doc.filename;
                                        newFile.Overwrite = true;
                                        
                                        if (doc.foldername == null)
                                        {
                                            uploadFile = list.RootFolder.Files.Add(newFile);
                                            
                                        }
                                        else
                                        {
                                            string foldername = doc.foldername;
                                            string sitecontent = doc.sitecontent;
                                            
                                            //Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                                            Folder folder = SharePointHelper.GetFolder(cc, list, foldername);
                                            if (folder == null){
                                                if(doc.taxFields != null){
                                                    folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, doc.fields, fields, doc.taxFields);
                                                }
                                                else
                                                {
                                                    folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, doc.fields, fields);
                                                }

                                            }
                                        }

                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                        {
                                            bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                            cc.ExecuteQuery();

                                            fileoffset = bytesUploaded.Value;
                                        }

                                        first = false;
                                    }
                                }
                                else
                                {
                                    string SRUrl = list.RootFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + filename;
                                    uploadFile = cc.Web.GetFileByServerRelativeUrl(SRUrl);
                                    if (last)
                                    {
                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(lastBuffer))
                                        {
                                            uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                            await cc.ExecuteQueryAsync();
                                            

                                        }
                                    }
                                    else
                                    {
                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                        {
                                            bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                            cc.ExecuteQuery();

                                            fileoffset = bytesUploaded.Value;
                                        }
                                        
                                    }
                                }
                            }
                        }

                    }
                    catch{
                        System.IO.File.Delete(tmpfile);
                        throw;
                    }
                    finally 
                    {
                        System.IO.File.Delete(tmpfile);
                        if (fs != null)
                        {
                            fs.Dispose();
                        }
                    }

                    
                }
                else
                {
                    System.IO.File.Delete(tmpfile);
                    return null;
                }
         

                    //File uploadFile = SharePointHelper.GetBigSharePointFile(file_url, filename, SMBCredential, client, nts, fileStore, list, cc, doc, fields);

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
                        Console.WriteLine("Successfully uploaded " + uploadFile.Name + " and updated metadata");
                    }
                    catch (System.Exception e)
                    {
                        _logger.LogError("Failed to update metadata.");
                        Console.WriteLine(e);
                        continue;
                    }


                }
            }}
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
        ///     Migration library with versioning.
        ///     Using SMB2Client to fetch files from smb file server.
        ///     Adds metadata for normal and taxonomy fields
        ///     [ 
        ///         {
        ///           "fields": {
        ///             "Created_x0020_By": "",
        ///             "eDocsRNLinjer": null,
        ///             "eDocsSaksbehandler": "RXINDEX IMPORT",
        ///             ....
        ///           },
        ///           "file_url": "//.../...",
        ///           "filename": "<string>",
        ///           "foldername": "<string>",
        ///           "list": "<guid>",
        ///           "site": "<site name>",
        ///           "sitecontent": "<site content name for creating document set>"
        ///         }
        ///     ]
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
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
            Console.WriteLine(url);
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
                cc.Load(list, l => l.EntityTypeName);
                cc.ExecuteQuery();
                List<Metadata> fields = SharePointHelper.GetFields(cc, list);
                //List list = cc.Web.Lists.GetByTitle(listname);

                for (int i = 0; i < docs.Length; i++)
                {
                    string filename = docs[i].filename;
                    string file_url = docs[i].file_url;
                    var inputFields = docs[i].fields;
                    var taxFields = docs[i].taxFields;
                    string foldername = docs[i].foldername;
                    string sitecontent = docs[i].sitecontent;



                    var qry = new CamlQuery();
                    
                    string FileDirRef = "/sites/" + site + "/" + list.EntityTypeName + "/" + foldername;
                    if(foldername != null)
                        qry.FolderServerRelativeUrl = FileDirRef;
                    qry.ViewXml = string.Format(@"
                    <View>  
                                <Query> 
                                    <Where>
                                        <Eq><FieldRef Name='FileLeafRef' />
                                        <Value Type='Text'>{0}</Value></Eq>
                                    </Where> 
                                </Query> 
                            </View>", filename);
                    
                    //check if file already exist
                    var items = list.GetItems(qry);
                    cc.Load(items);
                    cc.ExecuteQuery();
                    if( items.Count > 0){
                        _logger.LogInformation(filename + " already exist");
                        ListItem oItem = items.FirstOrDefault();
                        File targetFile = oItem.File;
                        targetFile.DeleteObject();
                        cc.ExecuteQuery();

                        _logger.LogInformation(filename + " deleted");
                    }
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
                        
                        //Folder folder = list.RootFolder.Folders.GetByUrl(foldername);

                        Folder folder = SharePointHelper.GetFolder(cc, list, foldername);
                        if (folder == null)
                        {
                            folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                        }
                        //if (folder == null && taxFields != null)
                        //    folder = SharePointHelper.CreateDocumentSetWithTaxonomy(cc, list, sitecontent, foldername, inputFields, fields, taxFields);
                        //else if (folder == null)
                        //    folder = SharePointHelper.CreateFolder(cc, list, sitecontent, foldername, inputFields, fields);
                        
                        //cc.ExecuteQuery();
                        uploadFile = folder.Files.Add(newFile);
                    }

                    _logger.LogInformation("Upload file: " + newFile.Url);
                    

                    ListItem item = uploadFile.ListItemAllFields;
                    uploadFile.CheckOut();
                    cc.ExecuteQuery();
                    //Console.WriteLine("checkout" + uploadFile.Name);
                    
                    ///if (uploadFile.CheckOutType == CheckOutType.None)
                    ///{
                    ///    uploadFile.CheckOut();
                    ///    cc.ExecuteQuery();
                    ///}


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
        /// DELETE AFTERWARDS
        /// Use only on lists with over 5000 documents
        /// 
        /// eDocsDokumentnavn is a indexed field.
        /// </summary>
        /// <param name="docs"></param>
        /// <returns></returns>
        public async Task<IActionResult> test([FromBody] DocumentModel[] docs)
        {
            int blockSize = 8000000; // 8 MB
            string fileName = "aen-per-oddvar.mp4",uniqueFileName = String.Empty;
            long fileSize;
            File uploadFile = null;
            Guid uploadId = Guid.NewGuid();
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
                cc.Load(list.RootFolder, p => p.ServerRelativeUrl);

                for (int i = 0; i < docs.Length; i++)
                {
                    
                    string file_url = docs[i].file_url;
                    string foldername = docs[i].foldername;
                    System.IO.FileStream fs = null;
                    // Use large file upload approach
                    ClientResult<long> bytesUploaded = null;
                    
                        fs = System.IO.File.Open(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite);
                        fileSize = fs.Length;
                        uniqueFileName = System.IO.Path.GetFileName(fs.Name);
                        using (System.IO.BinaryReader br = new System.IO.BinaryReader(fs))
                        {
                            byte[] buffer = new byte[blockSize];
                            byte[] lastBuffer = null;
                            long fileoffset = 0;
                            long totalBytesRead = 0;
                            int bytesRead;
                            bool first = true;
                            bool last = false;

                            // Read data from filesystem in blocks
                            while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                totalBytesRead = totalBytesRead + bytesRead;
                                if (totalBytesRead >= fileSize)
                                {
                                    last = true;
                                    // Copy to a new buffer that has the correct size
                                    lastBuffer = new byte[bytesRead];
                                    Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                                }

                                if (first)
                                {
                                    using(System.IO.MemoryStream contentStream = new System.IO.MemoryStream())
                                    {
                                        FileCreationInformation fileInfo = new FileCreationInformation();
                                        fileInfo.ContentStream = contentStream;
                                        fileInfo.Url = uniqueFileName;
                                        fileInfo.Overwrite = true;
                                        uploadFile = list.RootFolder.Files.Add(fileInfo);

                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                        {
                                            bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                            cc.ExecuteQuery();

                                            fileoffset = bytesUploaded.Value;
                                        }

                                        first = false;
                                    }
                                }
                                else
                                {
                                    uploadFile = cc.Web.GetFileByServerRelativeUrl(list.RootFolder.ServerRelativeUrl + System.IO.Path.AltDirectorySeparatorChar + uniqueFileName);

                                    if (last)
                                    {
                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(lastBuffer))
                                        {
                                            uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                            await cc.ExecuteQueryAsync();
                                        }
                                    }
                                    else
                                    {
                                        using (System.IO.MemoryStream s = new System.IO.MemoryStream(buffer))
                                        {
                                            bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                            cc.ExecuteQuery();

                                            fileoffset = bytesUploaded.Value;
                                        }
                                        
                                    }
                                }

                            }
                        }

                        if (fs != null)
                        {
                            fs.Dispose();
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
        /// Get all documents in library
        /// NOT FINISHED
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/document/all?site=<sitename>&list=<listname>
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> All([FromQuery(Name = "site")] string sitename,[FromQuery(Name = "list")] string listname)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();
            
            string url = _baseurl + "sites/" + sitename;
            Console.WriteLine(url);
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                Console.WriteLine("list: " + listname);
                cc.RequestTimeout = -1;

                List list = cc.Web.Lists.GetByTitle(listname);
                cc.ExecuteQuery();
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>";
                

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

        [HttpPost]
        [Produces("application/json")]
        [Consumes("application/json")]

        /// <summary>
        /// Get all documents in library
        /// NOT FINISHED
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/document/updatebiglibrary
        ///     update documents on big libraries. This was made because none of the fields we needed was not indexed
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> UpdateBigLibrary([FromBody] DocumentModel[] docs)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();

            Console.WriteLine("count: " + docs.Length);
            if (docs.Length == 0)
            {
                return null;
            }
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);
            
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                cc.RequestTimeout = -1;
                Console.WriteLine(_baseurl);
                
                List list = cc.Web.Lists.GetById(listGuid);
                cc.Load(list, 
                    l => l.EntityTypeName,
                    l => l.Fields.Include(
                        f => f.InternalName,
                        f => f.Title,
                        f => f.TypeAsString
                    ));
                cc.ExecuteQuery();
                List<Metadata> fields = SharePointHelper.GetFields(list);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>";
                

                List<ListItem> items = new List<ListItem>();
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                do
                {
                    ListItemCollection listItemCollection = list.GetItems(camlQuery);
                    cc.Load(listItemCollection);
                    await cc.ExecuteQueryAsync();

                    //Adding the current set of ListItems in our single buffer
                    items.AddRange(listItemCollection);
                    //Reset the current pagination info
                    camlQuery.ListItemCollectionPosition = listItemCollection.ListItemCollectionPosition;
                } while (camlQuery.ListItemCollectionPosition != null);

                
                if (items.Count > 0)
                {
                    Regex regex = new Regex(@"~t.*");
                    DateTime dtMin = new DateTime(1900,1,1);
                    for (int i = 0; i < docs.Length; i++)
                    {
                        string filename = docs[i].filename;
                        string file_url = docs[i].file_url;
                        string foldername = docs[i].foldername;
                        var inputFields = docs[i].fields;
                        var taxFields = docs[i].taxFields;
                        Console.WriteLine("searching: " + filename);

                        foreach (var item in items)
                        {
                            if (filename.Equals(item["FileLeafRef"]))
                            {
                                Console.WriteLine("updating: " +  filename);
                                
                                item["Title"] = filename;
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
                                    await cc.ExecuteQueryAsync();
                                    Console.WriteLine("Metadata updated: " + filename);
                                }
                            

                            }
                        }
                    }
                }
                



                return new NoContentResult();
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
        /// Get all documents in library
        /// </summary>
        /// <remarks>
        /// Sample request:
        ///
        ///     GET /api/document/updatetitlebiglibrary
        ///     update documents title field on big libraries. 
        ///     
        /// </remarks>
        /// <param name="param">New document parameters</param>
        /// <returns></returns>
        public async Task<IActionResult> UpdateTitle([FromBody] DocumentModel[] docs)
        {
            //List<SharePointDoc> SPDocs = new List<SharePointDoc>();

            Console.WriteLine("count: " + docs.Length);
            if (docs.Length == 0)
            {
                return null;
            }
            string site = docs[0].site;
            string url = _baseurl + "sites/" + site;
            string listname = docs[0].list;
            Guid listGuid = new Guid(listname);
            
            using(ClientContext cc = AuthHelper.GetClientContextForUsernameAndPassword(url, _username, _password))
            try
            {
                cc.RequestTimeout = -1;
                Console.WriteLine(_baseurl);
                
                List list = cc.Web.Lists.GetById(listGuid);
                cc.Load(list, 
                    l => l.EntityTypeName,
                    l => l.Fields.Include(
                        f => f.InternalName,
                        f => f.Title,
                        f => f.TypeAsString
                    ));
                cc.ExecuteQuery();
                
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = string.Format(
                            @"<View Scope='Recursive'> 
                                <Query> 
                                    <Where>
                                        <IsNull><FieldRef Name='Title' /></IsNull>
                                    </Where> 
                                </Query> 
                                <RowLimit Paged='TRUE'>100</RowLimit>
                            </View>");
                //camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>150</RowLimit></View>";
                
                int counter = 0;
                List<ListItem> items = new List<ListItem>();
                //List<string> fieldNames = SharePointHelper.GetVisibleFieldNames(cc, list);
                do
                {
                    ListItemCollection listItemCollection = list.GetItems(camlQuery);
                    cc.Load(listItemCollection);
                    await cc.ExecuteQueryAsync();
                    
                    foreach (var item in listItemCollection)
                    {
                        counter++;

                        item["Title"] = item["FileLeafRef"];
                        item.SystemUpdate();
                        Console.WriteLine(counter + " " + item["Title"]);

                    
                    }
                    await cc.ExecuteQueryAsync();

                    //Adding the current set of ListItems in our single buffer
                    //items.AddRange(listItemCollection);
                    //Reset the current pagination info
                    camlQuery.ListItemCollectionPosition = listItemCollection.ListItemCollectionPosition;
                } while (camlQuery.ListItemCollectionPosition != null);

               



                return new NoContentResult();
            }
            catch (System.Exception)
            {
                
                throw;
            }


        }


    }
}