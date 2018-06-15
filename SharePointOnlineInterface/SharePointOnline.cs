using Interfaces;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Threading.Tasks;

namespace SharePointOnlineInterface
{
    public class SharePointOnline : IDestination
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
        public SharePointOnline(string url, string username, string password)
        {
            if(username == null || password == null) //username and passwords are required for this class
            {
                throw new ArgumentException(message: "Arguments 'username' and 'password' are required for SharePoint Online URLs");
            }

            this.url = url; //Assign the passed URL to a private variable

            //Convert the password into a secure string
            SecureString pass = new SecureString();
            int characterIndex;
            for (characterIndex = 0; characterIndex < password.Length; characterIndex++)
            {
                pass.AppendChar(password[characterIndex]);
            }
            
            this.credentials = new SharePointOnlineCredentials(username, pass); //Create credentials to be used
        }
        #region MethodsNeededToUseThisAsADestination
        public void InjectDependencies(Func<string, int, IDictionary<string, string>> GetSourceItemAttributes, Func<string, int, Task<IDictionary<string, Stream>>> GetSourceItemAttachments, Func<string, Task<IEnumerable<string>>> GetSourceFolderNames, Func<string, Task<IEnumerable<string>>> GetSourceFileNames, Func<string, Task<Stream>> GetSourceFileStream)
        {
            //Save the injected methods as private ones to be used later
            this.GetSourceItemAttributes = GetSourceItemAttributes;
            this.GetSourceItemAttachments = GetSourceItemAttachments;
            this.GetSourceFolderNames = GetSourceFolderNames;
            this.GetSourceFileNames = GetSourceFileNames;
            this.GetSourceFileStream = GetSourceFileStream;
        }
        public async Task AddList(string title, int type, int itemCount)
        {
            List list = await GetOrCreateList(title, type); //Returns the expected list

            if (Enum.TryParse(type.ToString(), out ListTemplateType listType)) //parse the list type into the expected types
            {
                switch (listType) //Execute different methods depending on the type
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
        { //Returns a URL with the relative URL appended to it
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                await c.ExecuteQueryAsync();
                return c.Web.ServerRelativeUrl + url;
            }
        }
        private async Task<List> GetOrCreateList(string title, int type)
        { //Creates a list if it does not already exist and returns it
            List output;
            ClientContext c;
            using (c = context)
            {
                //Look for any list with that title already
                c.Load(c.Web.Lists, x => x.Where(y => y.Title == title));
                await c.ExecuteQueryAsync();

                if (c.Web.Lists.Any())
                {
#if DEBUG
                    Console.WriteLine(string.Format("Existing List:'{0}'", title));
#endif
                    output = c.Web.Lists.First(); //Get the existing list
                }
                else
                {
#if DEBUG
                    Console.WriteLine(string.Format("Adding List:'{0}'", title));
#endif
                    //Add new list
                    output = c.Web.Lists.Add(new ListCreationInformation()
                    {
                        Title = title, //New list title
                        TemplateType = type //New list type
                    });
                    await c.ExecuteQueryAsync();
                }
                return output;
            }
        }
        private async Task DeleteList(List list)
        { //This is used to delete large lists that can not be deleted otherwise. It removes the items in chunks before attempting to remote the list
            ListItemCollection items;
            int itemIndex;
            do
            {
                items = list.GetItems(getItemsToDelete); //Get the next batch of items
                list.Context.Load(items, x => x.ListItemCollectionPosition); //Used to determine if it has all of the items or not
                list.Context.Load(items, x => x.Include(y => y.Id)); //Get the list Ids
                await list.Context.ExecuteQueryAsync();
                for (itemIndex = 0; itemIndex < items.Count; itemIndex++) //Iterate the items
                {
                    items[itemIndex].DeleteObject(); //Deleat items one at a time
                }
            }
            while (items.ListItemCollectionPosition != null); //Check if that was the last group of items

        }
        private async Task<int> GetLastItemId(List list)
        { //This is used to get the last item id from a list
            ListItemCollection items;
            items = list.GetItems(getLastItemIdQuery); //Get the last Item added
            list.Context.Load(items, x => x.Include(y => y.Id)); //Queue a query to get the last Item Id
            await list.Context.ExecuteQueryAsync(); //Execute query to get the last item Id
            if (items.Any())
            {
                return items.First().Id; //return the id
            }
            return 0; //return 0 because there are no items
        }
        #region GenericList
        private async Task InitGenericList(List list, int itemCount)
        { //Used to populate Generic lists
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
                    attachments.Select(x => { x.Value.Dispose(); return x; }); //Dispose sreams
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }
        #endregion
        #region DocumentLibrary
        private async Task InitDocumentLibrary(List list)
        { //Used to populate Document libraries
#if DEBUG
            Console.WriteLine(string.Format("InitDocumentLibrary:'{0}'", list.Title));
#endif
            //Get the root folder to the list to start
            Folder folder = list.RootFolder;
            list.Context.Load(list, x => x.ParentWebUrl);
            list.Context.Load(folder, x => x.ServerRelativeUrl);
            await list.Context.ExecuteQueryAsync();

            //Convert the path into a relative one
            int index = folder.ServerRelativeUrl.IndexOf(list.ParentWebUrl);
            string cleanPath = (index < 0) ? folder.ServerRelativeUrl : folder.ServerRelativeUrl.Remove(index, list.ParentWebUrl.Length);
            await PopulateFolder(cleanPath);
        }
        private async Task PopulateFolder(string url)
        {
            ClientContext c;
            Folder folder;
            int fileIndex;
            int folderIndex;
            Stream fileStream;
            string sourceFileRelativeUrl;
            IEnumerable<string> folderNames = await GetSourceFolderNames(url);
            IEnumerable<string> fileNames = await GetSourceFileNames(url);
            using (c = context)
            {
                folder = c.Web.GetFolderByServerRelativeUrl(await CleanUrl(url)); //Get the folder
                //Iterate the files
                for (fileIndex = 0; fileIndex < fileNames.Count(); fileIndex++)
                {
                    sourceFileRelativeUrl = Path.Combine(url, fileNames.ElementAt(fileIndex));

                    using (fileStream = await GetSourceFileStream(sourceFileRelativeUrl))
                    {
                        //Add file
                        folder.Files.Add(new FileCreationInformation()
                        {
                            Url = fileNames.ElementAt(fileIndex), //Set filename
                            ContentStream = fileStream, //Get and set file stream
                            Overwrite = false //Do not overwrite files to save time
                        });
                        await c.ExecuteQueryAsync();
                    }
                }
                //Iterate the folders
                for (folderIndex = 0; folderIndex < folderNames.Count(); folderIndex++)
                {
                    folder.AddSubFolder(folderNames.ElementAt(folderIndex)); //Add folder
                    await c.ExecuteQueryAsync();
                    await PopulateFolder(Path.Combine(url, folderNames.ElementAt(folderIndex))); //Populate newly added folder
                }
            }
        }
        #endregion
        #endregion
    }
}
