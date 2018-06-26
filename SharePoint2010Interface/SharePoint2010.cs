using Interfaces;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace SharePoint2010Interface
{
    public class SharePoint2010 : ISource
    {
        #region Properties
        private CamlQuery getLastItemIdQuery { get; } = new CamlQuery()
        {
            ViewXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>1</RowLimit></View>"
        };
        private ICredentials credentials { get; set; }
        private string url { get; set; }
        private string[] itemAttributes { get; } = new string[] { "Title", "Modified", "Created" };
        #endregion
        #region Public
        public SharePoint2010(string url, string username, string password)
        {
            this.url = url;
            if(username == null || password == null)
            { //Use default credentials if none was specified
                this.credentials = CredentialCache.DefaultCredentials;
            }
            else
            { //Use network credentials if values were specified
                this.credentials = new NetworkCredential(username, password);
            }
        }
        #region MethodsNeededToUseThisAsASource
        public IEnumerable<SourceList> GetLists()
        { //Returns lists used for processing
            return PrivateGetLists();
        }
        public IDictionary<string, string> GetItemAttributes(string listTitle, int itemId)
        { //Returns item attributes for given properties
            ListItem item;
            Dictionary<string, string> output = new Dictionary<string, string>();
            item = GetItem(listTitle, itemId); //Get the item using the list title and id
            if (item != default(ListItem))
            {
                //Format properties to return
                foreach(var attributeName in itemAttributes)
                {
                    output.Add(attributeName, item[attributeName].ToString());
                }
            }
            return output; //return any found properties
        }
        public async Task<IEnumerable<string>> GetFolderNames(string url)
        { //Returns subfolder names of a given URL
            return PrivateGetFolderNames(CleanUrl(url));
        }
        public async Task<IEnumerable<string>> GetFileNames(string url)
        { //Returns file names under a given URL
            return PrivateGetFileNames(CleanUrl(url));
        }
        public async Task<Stream> GetFileStream(string url)
        { //Returns a filestream of a given file URL
            return await PrivateGetFileStream(CleanUrl(url));
        }
        public async Task<IEnumerable<string>> GetItemAttachmentPaths(string listTitle, int itemId)
        { //Returns a dictionary of list item attachments
            ListItem item = GetItem(listTitle, itemId); //Get the item
            if (item["Attachments"] as bool? == true) //Make sure the item has attachments
            {
                FileCollection attachmentCollection = GetAttachmentCollection(listTitle, itemId); //Get the collection of attachments
                return attachmentCollection.Select(x => UncleanUrl(x.ServerRelativeUrl));
            }
            return new List<string>();
        }
        #endregion
        #endregion
        #region Helpers
        private ClientContext context
        { //Used to create SharePoint context
            get
            {
                if (string.IsNullOrEmpty(url) || credentials == default(ICredentials))
                {
                    return null;
                }
                return new ClientContext(url)
                {
                    Credentials = credentials
                };
            }
        }
        private string CleanUrl(string url)
        { //Prepends the servers relative URL to make it a valid relative path for this server
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                c.ExecuteQuery();
                return c.Web.ServerRelativeUrl + url;
            }
        }
        private string UncleanUrl(string url)
        { //Converts the server URL back into a relative url
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                c.ExecuteQuery();
                return url.Replace(c.Web.ServerRelativeUrl, string.Empty);
            }

        }
        private IEnumerable<string> PrivateGetFolderNames(string url)
        { //returns a list of folder names under a given URL
            ClientContext c;
            Folder folder;
            using (c = context)
            {
                try
                {
                    folder = c.Web.GetFolderByServerRelativeUrl(url); //Get the passed in folder
                    c.Load(folder, x => x.Folders.Include(y => y.Name)); //Load the folder names
                    c.ExecuteQuery();
                    return folder.Folders.Select(x => x.Name); //Return the folder names
                }
                catch (Exception ex)
                {
                    if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                    {
                        throw ex;
                    }
                }
            }
            return new List<string>();
        }
        private IEnumerable<string> PrivateGetFileNames(string url)
        { //Returns a list of files under passed URL
            ClientContext c;
            Folder folder;
            IDictionary<string, Stream> output = new Dictionary<string, Stream>();
            using (c = context)
            {
                try
                {
                    folder = c.Web.GetFolderByServerRelativeUrl(url); //get the folder using passed URL
                    c.Load(folder, x => x.Files.Include(y => y.Name)); //Get files names under that folder
                    c.ExecuteQuery();
                    return folder.Files.Select(x => x.Name); //Return the files names
                }
                catch (Exception ex)
                {
                    if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                    {
                        throw ex;
                    }
                }
            }
            return new List<string>();
        }
        private async Task<Stream> PrivateGetFileStream(string url)
        { //Returns a stream of the passed in file url
            ClientContext c;
            FileInformation file;
            Stream output = new MemoryStream();
            using (c = context)
            {
                try
                {
                    file = Microsoft.SharePoint.Client.File.OpenBinaryDirect(c, url); //Gets the file
                    await file.Stream.CopyToAsync(output); //Copies the file to a memory stream
                    output.Seek(0, SeekOrigin.Begin); //Resets the stream so it is ready to be used
                }
                catch (Exception ex)
                {
                    if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                    {
                        throw ex;
                    }
                }
            }
            return output; //Returns the stream
        }
        private FileCollection GetAttachmentCollection(string title, int id)
        { //Returns attachments for a given list item
            ClientContext c;
            FileCollection attachmentFiles = default(FileCollection);
            try
            {
                using (c = context)
                {
                    attachmentFiles = c.Web.GetFolderByServerRelativeUrl(string.Format("Lists/{0}/Attachments/{1}", title, id)).Files; //Get the files for a known attachment path folder
                    c.Load(attachmentFiles, x => x.Include(y => y.ServerRelativeUrl, y => y.Name)); //Queue a query to get required fields for files
                    c.ExecuteQuery(); //Execute queued queries
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                {
                    throw ex;
                }
            }
            return attachmentFiles; //Return the attachments
        }
        private ListItem GetItem(string title, int id)
        { //Returns a list item given a title and item id
            ListItem item = default(ListItem);
            ClientContext c;
            try
            {
                using (c = context)
                {
                    item = c.Web.Lists.GetByTitle(title).GetItemById(id); //Get the list and item using parameters
                    c.Load(item); //Load the item properties
                    c.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                {
                    throw ex;
                }
            }
            return item; //Return the item
        }
        private IEnumerable<SourceList> PrivateGetLists()
        { //Returns parsed lists for this SharePoint site
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.Lists.Include(y => y.Title, y => y.BaseTemplate)); //Gets the base fields
                c.ExecuteQuery();
                return c.Web.Lists.Select(x => new SourceList() { Title = x.Title, Type = x.BaseTemplate, ItemCount = GetLastItemId(x) }); //Returns formatted lists
            }
        }
        private int GetLastItemId(List list)
        { //This is used to get the last item id from a list
            ListItemCollection items;
            items = list.GetItems(getLastItemIdQuery); //Get the last Item added
            list.Context.Load(items, x => x.Include(y => y.Id)); //Queue a query to get the last Item Id
            list.Context.ExecuteQuery(); //Execute query to get the last item Id
            if (items.Any())
            {
                return items.First().Id; //return the id
            }
            return 0; //return 0 because there are no items
        }
        #endregion
    }
}
