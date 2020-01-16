using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.SharePoint.Client.UserProfiles;
using System;
using System.Text;
using System.Net;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;


namespace SharePointAPI.Middleware
{
    public class SharePointHelper
    {
        public static List GetListItemByTitle(ClientContext cc, string title)
        {
            try
            {
                //Web web = cc.Web;
                //cc.Load(web);
                var Lists = cc.Web.Lists;
                cc.Load(Lists);
                //cc.ExecuteQuery();
                List list = Lists.GetByTitle(title);
                
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
                        Console.WriteLine(tmpfield.InternalName);
                        
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
                    Console.WriteLine("Folder Name: " + folder.Name);
                    
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
                throw;
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

                newListItem.Update();
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

        public static FieldCollection GetFields(ClientContext cc, List list)
        {
            try
            {
                FieldCollection fields = list.Fields;
                cc.Load(fields);
                cc.ExecuteQuery();

                return fields;                
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);   
                throw;
            }
        }
        /*public static List<Object> GetFieldNameAndType(ClientContext cc, FieldCollection fields)
        {
            List<Object> fieldType = new List<object>();

            

        }*/


        public static void SetMetadataFields(ClientContext cc, JObject inputFields, FieldCollection fields, ListItem item)
        {
            foreach (KeyValuePair<string, JToken> inputField in inputFields)
            {   
                
                var field = fields.GetByInternalNameOrTitle(inputField.Key);
                cc.Load(field);
                cc.ExecuteQuery();
                
                
                if(field.TypeAsString.Equals("TaxonomyFieldType"))
                {
                    var taxKeywordField = cc.CastTo<TaxonomyField>(field);                  

                    Guid _id = taxKeywordField.TermSetId;
                    string _termID = TermHelper.GetTermIdByName(cc, inputField.Value.ToString(), _id);

                    
                    TaxonomyFieldValue termValue = new TaxonomyFieldValue()
                    {
                        Label = inputField.Value.ToString().ToString(),
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
                else
                {
                    item[inputField.Key] = inputField.Value.ToString();
                    
                }
                
                // This method works but not practical
                //string termValue = "-1;#" + taxObj["Label"].ToString() + "|" + taxObj["TermGuid"].ToString();
                //item[inputField.Key] = termValue;
                
                
                item.Update();
                //cc.ExecuteQuery();
                
            }
                

        }
    }
}