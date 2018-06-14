using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace SharePointOnlineInterface
{
    public class SharePointOnline
    {
        #region Properties
        private CamlQuery getItemsToDelete { get; } = new CamlQuery()
        {
            ViewXml = "<View><RowLimit>1000</RowLimit></View>"
        };
        private CamlQuery getLastItemIdQuery { get; } = new CamlQuery()
        {
            ViewXml = "<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE'/></OrderBy></Query><RowLimit>1</RowLimit></View>"
        };
        private int MaxInsertCount { get; } = 10;
        private string[] itemAttributes { get; } = new string[] { "Title", "Modified", "Created" };
        private SharePointOnlineCredentials credentials { get; set; }
        private string url { get; set; }
        #region Dependencies
        private Func<string, int, IDictionary<string, string>> GetSourceItemAttributes { get; set; }
        private Func<string, int, Task<IDictionary<string, Stream>>> GetSourceItemAttachments { get; set; }
        private Func<string, Task<IEnumerable<string>>> GetSourceFolderNames { get; set; }
        private Func<string, Task<IEnumerable<string>>> GetSourceFileNames { get; set; }
        private Func<string, Task<Stream>> GetSourceFileStream { get; set; }
        #endregion
        #endregion
        #region Public
        public SharePointOnline(string url, string username, string password, Func<string, int, IDictionary<string, string>> GetSourceItemAttributes, Func<string, int, Task<IDictionary<string, Stream>>> GetSourceItemAttachments, Func<string, Task<IEnumerable<string>>> GetSourceFolderNames, Func<string, Task<IEnumerable<string>>> GetSourceFileNames, Func<string, Task<Stream>> GetSourceFileStream)
        {
            SecureString pass = new SecureString();
            int characterIndex;
            for (characterIndex = 0; characterIndex < password.Length; characterIndex++)
            {
                pass.AppendChar(password[characterIndex]);
            }

            this.credentials = new SharePointOnlineCredentials(username, pass);
            this.url = url;
            this.GetSourceItemAttributes = GetSourceItemAttributes;
            this.GetSourceItemAttachments = GetSourceItemAttachments;
            this.GetSourceFolderNames = GetSourceFolderNames;
            this.GetSourceFileNames = GetSourceFileNames;
            this.GetSourceFileStream = GetSourceFileStream;
        }
        #region MethodsNeededToUseThisAsADestination
        public async Task AddList(string title, int type, int itemCount)
        {
            List list = await GetOrCreateList(title, type);

            if (Enum.TryParse(type.ToString(), out ListTemplateType listType))
            {
                switch (listType)
                {
                    case ListTemplateType.GenericList:
                        await InitGenericList(list, itemCount);
                        break;
                    case ListTemplateType.DocumentLibrary:
                        await InitDocumentLibrary(list);
                        break;
                    default:
                        Console.WriteLine(string.Format("Did not plan for list type {0}", listType));
                        break;
                }
            }
        }
        #endregion
        #endregion
        #region Helpers
        private ClientContext context
        {
            get
            {
                if (string.IsNullOrEmpty(url) || credentials == default(SharePointOnlineCredentials))
                {
                    return null;
                }
                return new ClientContext(url)
                {
                    Credentials = credentials
                };
            }
        }
        private async Task<string> CleanUrl(string url)
        {
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                await c.ExecuteQueryAsync();
                return c.Web.ServerRelativeUrl + url;
            }
        }
        private async Task<List> GetOrCreateList(string title, int type)
        {
            List output;
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web.Lists, x => x.Where(y => y.Title == title));
                await c.ExecuteQueryAsync();

                if (c.Web.Lists.Any()) //Get the existing list
                {
#if DEBUG
                    Console.WriteLine(string.Format("Existing List:'{0}'", title));
#endif
                    output = c.Web.Lists.First();
                }
                else //Add new list
                {
#if DEBUG
                    Console.WriteLine(string.Format("Adding List:'{0}'", title));
#endif
                    output = c.Web.Lists.Add(new ListCreationInformation()
                    {
                        Title = title,
                        TemplateType = type
                    });
                    await c.ExecuteQueryAsync();
                }
                return output;
            }
        }
        private async Task DeleteList(List list)
        {
            ListItemCollection items;
            int itemIndex;
            do
            {
                items = list.GetItems(getItemsToDelete);
                list.Context.Load(items, x => x.ListItemCollectionPosition);
                list.Context.Load(items, x => x.Include(y => y.Id));
                await list.Context.ExecuteQueryAsync(); //Execute query to get the last source item Id
                for (itemIndex = 0; itemIndex < items.Count; itemIndex++)
                {
                    items[itemIndex].DeleteObject();
                }
            }
            while (items.ListItemCollectionPosition != null);

        }
        private async Task<int> GetLastItemId(List list)
        {
            ListItemCollection items;
            items = list.GetItems(getLastItemIdQuery); //Get the last Source Item added
            list.Context.Load(items, x => x.Include(y => y.Id)); //Queue a query to get the last source Item Id
            await list.Context.ExecuteQueryAsync(); //Execute query to get the last source item Id
            if (items.Any())
            {
                return items.First().Id;
            }
            return 0;
        }
        #region GenericList
        private async Task InitGenericList(List list, int itemCount)
        {
#if DEBUG
            Console.WriteLine(string.Format("InitGenericList:'{0}'", list.Title));
#endif

            ListItem item;
            IDictionary<string, string> attributes;
            IDictionary<string, Stream> attachments;
            int itemIndex;
            int attributeIndex;
            int attachmentIndex;
            int currentCount = await GetLastItemId(list);

            //Add items
            for (itemIndex = currentCount + 1; itemIndex < itemCount; itemIndex++)
            {
                try
                {
                    //Add blank item
                    item = list.AddItem(new ListItemCreationInformation());

                    //Update properties
                    attributes = GetSourceItemAttributes(list.Title, itemIndex); //Get properties from source
                    for (attributeIndex = 0; attributeIndex < itemAttributes.Length; attributeIndex++)
                    {
                        item[itemAttributes[attributeIndex]] = attributes[itemAttributes[attributeIndex]];
                    }

                    item.Update(); //Trigger an item update so it gets inserted

                    //Update attachments
                    attachments = await GetSourceItemAttachments(list.Title, itemIndex);
                    for (attachmentIndex = 0; attachmentIndex < attachments.Count(); attachmentIndex++)
                    {
                        item.AttachmentFiles.Add(new AttachmentCreationInformation() //Queue a query to write the stream as an attachment
                        {
                            FileName = attachments.ElementAt(attachmentIndex).Key, //Set attachment name
                            ContentStream = attachments.ElementAt(attachmentIndex).Value //Set attachment content
                        });
                    }

                    item.Update(); //Trigger an item update so attachments get inserted
                    await list.Context.ExecuteQueryAsync(); //execute queued queries
                }
                catch (Exception ex)
                {
                    Console.Write(ex);
                    throw ex;
                }
            }
        }
        #endregion
        #region DocumentLibrary
        private async Task InitDocumentLibrary(List list)
        {
#if DEBUG
            Console.WriteLine(string.Format("InitDocumentLibrary:'{0}'", list.Title));
#endif
            Folder folder = list.RootFolder;
            //Get server pat
            list.Context.Load(list, x => x.ParentWebUrl);
            list.Context.Load(folder, x => x.ServerRelativeUrl);
            await list.Context.ExecuteQueryAsync();

            int index = folder.ServerRelativeUrl.IndexOf(list.ParentWebUrl);
            string cleanPath = (index < 0)
                ? folder.ServerRelativeUrl
                : folder.ServerRelativeUrl.Remove(index, list.ParentWebUrl.Length);
            await PopulateFolder(cleanPath);
        }
        private async Task PopulateFolder(string url)
        {
            ClientContext c;
            Folder folder;
            int fileIndex;
            int folderIndex;
            string sourceFileRelativeUrl;
            IEnumerable<string> folderNames = await GetSourceFolderNames(url);
            IEnumerable<string> fileNames = await GetSourceFileNames(url);
            using (c = context)
            {
                folder = c.Web.GetFolderByServerRelativeUrl(await CleanUrl(url));
                for (fileIndex = 0; fileIndex < fileNames.Count(); fileIndex++)
                {
                    sourceFileRelativeUrl = Path.Combine(url, fileNames.ElementAt(fileIndex));
                    folder.Files.Add(new FileCreationInformation()
                    {
                        Url = fileNames.ElementAt(fileIndex),
                        ContentStream = await GetSourceFileStream(sourceFileRelativeUrl),
                        Overwrite = false
                    });
                    await c.ExecuteQueryAsync(); //Add file
                }
                for (folderIndex = 0; folderIndex < folderNames.Count(); folderIndex++)
                {
                    folder.AddSubFolder(folderNames.ElementAt(folderIndex));
                    await c.ExecuteQueryAsync(); //Add folder
                    await PopulateFolder(Path.Combine(url, folderNames.ElementAt(folderIndex))); //Populate newly added folder
                }
            }
        }
        #endregion
        #endregion
    }
}
