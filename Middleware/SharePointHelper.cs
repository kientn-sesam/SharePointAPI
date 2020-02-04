using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using SharpCifs.Smb;
using SMBLibrary;
using SMBLibrary.Client;
using Microsoft.AspNetCore.Http;





using SharePointAPI.Models;


namespace SharePointAPI.Middleware
{
    public class SharePointHelper
    {
        public static List GetListItemByTitle(ClientContext cc, string title)
        {
            try
            {

                //var Lists = cc.Web.Lists;
                //cc.Load(Lists);
                List list = cc.Web.Lists.GetByTitle(title);

                
                return list;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
        }
        public static List<string> GetVisibleFieldNames(ClientContext cc, List list)
        {
            try
            {
                List<string> fieldNames = new List<string>();
                cc.Load(list, l => l.Fields);
                cc.ExecuteQuery();
                foreach (var tmpfield in list.Fields)
                {
                    if((!tmpfield.FromBaseType && tmpfield.Hidden == false)){
                        if (tmpfield.InternalName.Equals("ContentType"))
                        {
                            continue;
                        }
                        
                        fieldNames.Add(tmpfield.InternalName);
                    }
                }
                return fieldNames;
                
            }
            catch (System.Exception)
            {
                throw;
            }

        }

        public static List<JObject> GetItemsFromListByField(ClientContext cc, FolderCollection folders, List<string> fieldNames)
        {
            var SPDocs = new List<JObject>();
            try
            {
                
                foreach (var folder in folders)
                {
                    
                    var items = folder.Files;
                    
                    
                    // Skip unecessary folder
                    if(string.IsNullOrEmpty(folder.ProgID)){
                        continue;
                    }

                    cc.Load(items);
                    cc.ExecuteQuery();
                    
                    foreach (var file in items)
                    {
                        
                        ListItem item = file.ListItemAllFields;
                        //test
                        cc.Load(item);
                        cc.ExecuteQuery();

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
            }
            catch (System.Exception)
            {
                
                throw;
            }

            return SPDocs;
        }

        public static FolderCollection GetFolders(ClientContext cc, List list)
        {
            FolderCollection folders = list.RootFolder.Folders;
            cc.Load(folders);
            cc.ExecuteQuery();

            return folders;
        }

        public static Folder GetFolder(ClientContext cc, List list, string foldername)
        {
            try
            {
                Folder folder = list.RootFolder.Folders.GetByUrl(foldername);
                cc.Load(folder);
                cc.ExecuteQuery();
                return folder;
            }
            catch (System.Exception)
            {
                return null;
            }

            
        }

        public static Boolean FolderJObjectExist(JObject doc)
        {
            try
            {
                if (doc["foldername"].ToString() != null)
                {
                    return true;
                }
            }
            catch (System.Exception)
            {
                return false;
            }

            return true;
        }

        public static Folder CreateFolder(ClientContext cc, List list, string sitecontent, string documentSetName, JObject fields){

            try
            {
                ContentTypeCollection listContentTypes = list.ContentTypes;
                cc.Load(listContentTypes, types => types.Include(type => type.Id, type => type.Name, type => type.Parent));
                //var result = cc.LoadQuery(listContentTypes.Where(c => c.Name == "document set 2"));
                string SiteContentName = sitecontent;
                var result = cc.LoadQuery(listContentTypes.Where(c => c.Name == SiteContentName));
                
                cc.ExecuteQuery();

                ContentType targetDocumentSetContentType = result.FirstOrDefault();
                ListItemCreationInformation newItemInfo = new ListItemCreationInformation();

                newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
                //newItemInfo.LeafName = "Document Set Kien2";
                newItemInfo.LeafName = documentSetName;
                
                
                //newItemInfo.FolderUrl = list.RootFolder.ServerRelativeUrl.ToString();
                
                ListItem newListItem = list.AddItem(newItemInfo);
                newListItem["ContentTypeId"] = targetDocumentSetContentType.Id.ToString();
                foreach (KeyValuePair<string, JToken> field in fields)
                {
                    JObject fieldObj = field.Value as JObject;
                    if (fieldObj["type"].ToString().Equals("User"))
                    {
                        var user = FieldUserValue.FromUser(fieldObj["label"].ToString());
                        newListItem[field.Key] = user;
                    }
                    else
                    {
                        newListItem[field.Key] = fieldObj["label"].ToString();
                    }
                }

                newListItem.SystemUpdate();
                list.Update();
                cc.ExecuteQuery();

                //Folder folder = GetFolder(cc, list, documentSetName);
                Folder folder = newListItem.Folder;
                return folder;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to create document set");
                Console.WriteLine(ex);
                throw;
            }
        }

        public static FileCreationInformation GetFileCreationInformation(string fileurl, string filename, SMBCredential SMBCredential, SMB2Client client, NTStatus nts, ISMBFileStore fileStore)
        {
            
                
            //SMBLibrary.NTStatus actionStatus;
            FileCreationInformation newFile = new FileCreationInformation();
            NTStatus status = nts;
        
            object handle;
            FileStatus fileStatus;
            
            //string path = fileurl;
            
            //string path = "Dokument/ARKIV/RUNSAL/23_02_2011/sz001!.PDF";
            
            string tmpfile = Path.GetTempFileName();
            status = fileStore.CreateFile(out handle, out fileStatus, fileurl, AccessMask.GENERIC_READ, 0, ShareAccess.Read, CreateDisposition.FILE_OPEN, CreateOptions.FILE_NON_DIRECTORY_FILE, null);
            if (status != NTStatus.STATUS_SUCCESS)
            {
                Console.WriteLine(status);
            }
            else{
                
                byte[] buf;
                var fs = new FileStream(tmpfile, FileMode.OpenOrCreate);
                var bw = new BinaryWriter(fs);
                int bufsz = 64 * 1000;
                int i = 0;
                
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
                while (status != NTStatus.STATUS_END_OF_FILE && i < 1000);
                
                if (status == NTStatus.STATUS_SUCCESS)
                {
                    fileStore.CloseFile(handle);
                    bw.Flush();
                    fs.Close();
                    fs = System.IO.File.OpenRead(tmpfile);
                    
                    //byte[] fileBytes = new byte[fs.Length];
                    //fs.Read(fileBytes, 0, fileBytes.Length);
                    
                    newFile.Overwrite = true;
                    newFile.ContentStream = fs;
                    //newFile.Content = fileBytes;
                    newFile.Url = filename;
                    
                    
                }
                else
                {
                    System.IO.File.Delete(tmpfile);
                    return null;
                }
                
                
                
                System.IO.File.Delete(tmpfile);
                
                    
            }
                
            
            
            
            return newFile;
                

        }
        public static FileCreationInformation GetFileCreationInformation(string fileurl, string filename)
        {
            try
            {
                FileCreationInformation newFile = new FileCreationInformation();
                using (var webClient = new WebClient()){
                        byte[] fileBytes = webClient.DownloadData(fileurl);

                        newFile.Overwrite = true;
                        newFile.Content = fileBytes;

                        Console.WriteLine("Download " + filename + " successful.");
                }
                // set filename
                newFile.Url = filename;

                return newFile;
                
            }
            catch (System.Exception)
            {
                Console.WriteLine("failed to download: " + filename);
                throw;
            }
        }
        
        public static List<Metadata> GetFields(ClientContext cc, List list)
        {
            try
            {
                List<Metadata> metadata = new List<Metadata>();

                FieldCollection fields = list.Fields;
                cc.Load(fields, fields => fields.Include(
                    f => f.InternalName,
                    f => f.Title,
                    f => f.TypeAsString
                ));
                cc.ExecuteQuery();

                for (int i = 0; i < fields.Count; i++)
                {
                    metadata.Add(new Metadata(){ Title = fields[i].Title, TypeAsString = fields[i].TypeAsString, InternalName = fields[i].InternalName});
                }
                
                return metadata;                
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);   
                throw;
            }
           
        }

        public static void SetMetadataFields(ClientContext cc, List list, JObject inputFields, List<Metadata> fields, ListItem item)
        {
            try
            {
                DateTime dtMin = new DateTime(1900,1,1);
                foreach (KeyValuePair<string, JToken> inputField in inputFields)
                {

                    if (inputField.Value == null || inputField.Value.ToString() == "" )
                    {
                        //Console.WriteLine(inputField.Key);
                        continue;
                    }
                    
                    var field = fields.Find(f => f.InternalName == inputField.Key);


                    if(field.TypeAsString.Equals("TaxonomyFieldType"))
                    {
                        Field taxField = list.Fields.GetByInternalNameOrTitle(inputField.Key);
                        var taxKeywordField = cc.CastTo<TaxonomyField>(taxField); 
                        

                        Guid _id = taxKeywordField.TermSetId;
                        string _termID = TermHelper.GetTermIdByName(cc, inputField.Value.ToString(), _id);

                        
                        TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                        {
                            Label = inputField.Value.ToString(),
                            TermGuid = _termID,
                            //WssId = -1
                            //WssId = (int)taxObj["WssId"]
                        };

                        taxKeywordField.SetFieldValueByValue(item, termValue);
                        taxKeywordField.Update();
                    }
                    else if(field.TypeAsString.Equals("User"))
                    {
                        var user = FieldUserValue.FromUser(inputField.Value.ToString());
                        item[inputField.Key] = user;
                        Console.WriteLine("Set field " + inputField.Key + " to " + user); 
                        
                    }
                    else if(field.TypeAsString.Equals("DateTime")){
                        
                        string dateTimeStr = inputField.Value.ToString();
                        dateTimeStr = dateTimeStr.Replace("~t","");
                        if(DateTime.TryParse(dateTimeStr, out DateTime dt))
                        {
                            if(dtMin <= dt){
                                item[inputField.Key] = dt;
                                Console.WriteLine("Set field " + inputField.Key + "to " + dt);
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
                            item[inputField.Key] = inputField.Value.ToString();
                            Console.WriteLine("Set " + inputField.Key + " to " + inputField.Value.ToString());
                            
                        }
                        
                    }

                    item.SystemUpdate();
                }
                
                
            }
            catch (System.Exception)
            {
                
                throw;
            }
        }

        public static void SetMetadataFields(ClientContext cc, JObject inputFields, FieldCollection fields, ListItem item)
        {
            foreach (KeyValuePair<string, JToken> inputField in inputFields)
            {   
                
                var field = fields.GetByInternalNameOrTitle(inputField.Key);
                
                cc.Load(field);
                cc.ExecuteQuery();
                Console.WriteLine(field.TypeAsString);
                
                
                if(field.TypeAsString.Equals("TaxonomyFieldType"))
                {
                    var taxKeywordField = cc.CastTo<TaxonomyField>(field);                  

                    Guid _id = taxKeywordField.TermSetId;
                    string _termID = TermHelper.GetTermIdByName(cc, inputField.Value.ToString(), _id);

                    
                    TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                    {
                        Label = inputField.Value.ToString(),
                        TermGuid = _termID,
                        //WssId = -1
                        //WssId = (int)taxObj["WssId"]
                    };

                    taxKeywordField.SetFieldValueByValue(item, termValue);
                    taxKeywordField.Update();
                }
                else if(field.TypeAsString.Equals("User"))
                {
                    var user = FieldUserValue.FromUser(inputField.Value.ToString());
                    item[inputField.Key] = user;
                    
                }
                else if(field.TypeAsString.Equals("DateTime") && inputField.Value.ToString() != ""){
                    
                    string dateTimeStr = inputField.Value.ToString();
                    dateTimeStr = dateTimeStr.Replace("~t","");
                    item[inputField.Key] = Convert.ToDateTime(dateTimeStr);
                }
                else if(inputField.Value.ToString() == ""){
                    continue;
                }
                else
                {
                    item[inputField.Key] = inputField.Value.ToString();
                    
                }
                
                // This method works but not practical
                //string termValue = "-1;#" + taxObj["Label"].ToString() + "|" + taxObj["TermGuid"].ToString();
                //item[inputField.Key] = termValue;
                
                
                item.SystemUpdate();
                //cc.ExecuteQuery();
                
            }
                

        }

        public static string[] GetFolderNames(ClientContext cc, List list, FolderCollection folders)
        {
            string[] foldernames = new string[folders.Count];
                
                for (int i = 0; i < folders.Count; i++)
                {
                    foldernames[i] = folders[i].Name;
                }

                return foldernames;
        }


    }
}