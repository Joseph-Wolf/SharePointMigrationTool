using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace SharePoint2010Interface
{
    public class SharePoint2010
    {
        #region Properties
        private ICredentials credentials { get; set; }
        private string url { get; set; }
        private string[] itemAttributes { get; } = new string[] { "Title", "Modified", "Created" };
        #endregion
        #region Public
        public SharePoint2010(string url, string username = null, string password = null)
        {
            this.url = url;
            if(username == null || password == null)
            {
                this.credentials = CredentialCache.DefaultCredentials;
            }
            else
            {
                this.credentials = new NetworkCredential(username, password);
            }
        }
        #region MethodsNeededToUseThisAsASource
        public IEnumerable<SharePoint2010List> GetLists()
        {
            return PrivateGetLists();
        }
        public IDictionary<string, string> GetItemAttributes(string listTitle, int itemId)
        {
            ListItem item;
            int attributeIndex;
            Dictionary<string, string> output = new Dictionary<string, string>();
            item = GetItem(listTitle, itemId);
            if (item != default(ListItem))
            {
                //Format properties to return
                for (attributeIndex = 0; attributeIndex < itemAttributes.Length; attributeIndex++)
                {
                    output.Add(itemAttributes[attributeIndex], item[itemAttributes[attributeIndex]].ToString());
                }
            }
            return output;
        }
        public async Task<IEnumerable<string>> GetFolderNames(string url)
        {
            return PrivateGetFolderNames(CleanUrl(url));
        }
        public async Task<IEnumerable<string>> GetFileNames(string url)
        {
            return PrivateGetFileNames(CleanUrl(url));
        }
        public async Task<Stream> GetFileStream(string url)
        {
            return await PrivateGetFileStream(CleanUrl(url));
        }
        public async Task<IDictionary<string, Stream>> GetItemAttachments(string listTitle, int itemId)
        {
            int attachmentIndex;
            IDictionary<string, Stream> output = new Dictionary<string, Stream>();
            ListItem item = GetItem(listTitle, itemId);
            if (item["Attachments"] as bool? == true) //Make sure attachments exist
            {
                FileCollection attachmentCollection = GetAttachmentCollection(listTitle, itemId);
                for (attachmentIndex = 0; attachmentIndex < attachmentCollection.Count; attachmentIndex++)
                {
                    output.Add(attachmentCollection[attachmentIndex].Name, await PrivateGetFileStream(attachmentCollection[attachmentIndex].ServerRelativeUrl));
                }
            }
            return output;
        }
        #endregion
        #endregion
        #region Helpers
        private ClientContext context
        {
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
        {
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                c.ExecuteQuery();
                return c.Web.ServerRelativeUrl + url;
            }
        }
        private IEnumerable<string> PrivateGetFolderNames(string url)
        {
            ClientContext c;
            Folder folder;
            using (c = context)
            {
                try
                {
                    folder = c.Web.GetFolderByServerRelativeUrl(url);
                    c.Load(folder, x => x.Folders.Include(y => y.Name));
                    c.ExecuteQuery();
                    return folder.Folders.Select(x => x.Name);
                }
                catch (ArgumentException ex)
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
        {
            ClientContext c;
            Folder folder;
            IDictionary<string, Stream> output = new Dictionary<string, Stream>();
            using (c = context)
            {
                try
                {
                    folder = c.Web.GetFolderByServerRelativeUrl(url);
                    c.Load(folder, x => x.Files.Include(y => y.Name));
                    c.ExecuteQuery();
                    return folder.Files.Select(x => x.Name);
                }
                catch (ArgumentException ex)
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
        {
            ClientContext c;
            FileInformation file;
            Stream output = new MemoryStream();
            using (c = context)
            {
                try
                {
                    file = Microsoft.SharePoint.Client.File.OpenBinaryDirect(c, url);
                    await file.Stream.CopyToAsync(output);
                    output.Seek(0, SeekOrigin.Begin);
                }
                catch (ArgumentException ex)
                {
                    if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                    {
                        throw ex;
                    }
                }
            }
            return output;
        }
        private FileCollection GetAttachmentCollection(string title, int id)
        {
            ClientContext c;
            FileCollection attachmentFiles = default(FileCollection);
            try
            {
                using (c = context)
                {
                    attachmentFiles = c.Web.GetFolderByServerRelativeUrl(string.Format("Lists/{0}/Attachments/{1}", title, id)).Files;
                    c.Load(attachmentFiles, x => x.Include(y => y.ServerRelativeUrl, y => y.Name)); //Queue a query to get required fields for files
                    c.ExecuteQuery(); //Execute queued queries
                }
            }
            catch (ArgumentException ex)
            {
                if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                {
                    throw ex;
                }
            }
            return attachmentFiles;
        }
        private ListItem GetItem(string title, int id)
        {
            ListItem item = default(ListItem);
            ClientContext c;
            try
            {
                using (c = context)
                {
                    item = c.Web.Lists.GetByTitle(title).GetItemById(id);
                    c.Load(item);
                    c.ExecuteQuery();
                }
            }
            catch (ArgumentException ex)
            {
                if (ex.Message != "Item does not exist. It may have been deleted by another user.")
                {
                    throw ex;
                }
            }
            return item;
        }
        private IEnumerable<SharePoint2010List> PrivateGetLists()
        {
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.Lists.Include(y => y.Title, y => y.BaseTemplate, y => y.ItemCount));
                c.ExecuteQuery();
                return c.Web.Lists.Select(x => new SharePoint2010List() { Title = x.Title, Type = x.BaseTemplate, ItemCount = x.ItemCount });
            }
        }
        #endregion
    }
}
