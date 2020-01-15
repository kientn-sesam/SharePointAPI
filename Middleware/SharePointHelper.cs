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

namespace SharePointAPI.Middleware
{
    public class SharePointHelper
    {
        public static List GetListItemByTitle(ClientContext cc, string title)
        {
            try
            {
                Web web = cc.Web;
                cc.Load(web);
                var Lists = web.Lists;
                cc.Load(Lists);
                cc.ExecuteQuery();
                List list = web.Lists.GetByTitle(title);

                return list;
                
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
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
                        string download = Encoding.ASCII.GetString(fileBytes);
                        Console.WriteLine(download);
                        newFile.Overwrite = true;
                        newFile.Content = fileBytes;

                        Console.WriteLine("Download " + filename + " successful.");
                }
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
    }
}