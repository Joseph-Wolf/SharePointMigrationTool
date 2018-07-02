using Interfaces;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

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
        private string getItemByIdCamlView { get; } = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
        private int MaxInsertCount { get; } = 10;
        private string[] itemAttributes { get; } = new string[] { "Title", "Modified", "Created" };
        private SharePointOnlineCredentials credentials { get; set; }
        private string url { get; set; }
        #region Dependencies
        private Func<string, int, IDictionary<string, string>> GetSourceItemAttributes { get; set; }
        private Func<string, int, IEnumerable<string>> GetSourceItemAttachmentPaths { get; set; }
        private Func<string, IEnumerable<string>> GetSourceFolderNames { get; set; }
        private Func<string, IEnumerable<string>> GetSourceFileNames { get; set; }
        private Func<string, Stream> GetSourceFileStream { get; set; }
        #endregion
        #endregion
        #region Public
        public SharePointOnline(string url, string username, string password)
        {
            if (username == null || password == null) //username and passwords are required for this class
            {
                throw new ArgumentException(message: "Arguments 'username' and 'password' are required for SharePoint Online URLs");
            }

            this.url = url; //Assign the passed URL to a private variable

            //Convert the password into a secure string
            SecureString pass = new SecureString();
            foreach (var character in password)
            {
                pass.AppendChar(character);
            }

            this.credentials = new SharePointOnlineCredentials(username, pass); //Create credentials to be used
        }
        #region MethodsNeededToUseThisAsADestination
        public void InjectDependencies(Func<string, int, IDictionary<string, string>> GetSourceItemAttributes, Func<string, int, IEnumerable<string>> GetSourceItemAttachmentPaths, Func<string, IEnumerable<string>> GetSourceFolderNames, Func<string, IEnumerable<string>> GetSourceFileNames, Func<string, Stream> GetSourceFileStream)
        {
            //Save the injected methods as private ones to be used later
            this.GetSourceItemAttributes = GetSourceItemAttributes;
            this.GetSourceItemAttachmentPaths = GetSourceItemAttachmentPaths;
            this.GetSourceFolderNames = GetSourceFolderNames;
            this.GetSourceFileNames = GetSourceFileNames;
            this.GetSourceFileStream = GetSourceFileStream;
        }
        public void AddList(string title, int type, int count)
        {
            List list = GetOrCreateList(title, type); //Returns the expected list

            if (Enum.TryParse(type.ToString(), out ListTemplateType listType)) //parse the list type into the expected types
            {
                switch (listType) //Execute different methods depending on the type
                {
                    case ListTemplateType.GenericList:
#if DEBUG
                        Console.WriteLine(string.Format("InitGenericList:'{0}'", title));
#endif
                        InitGenericList(title, count);
                        break;
                    case ListTemplateType.DocumentLibrary:
#if DEBUG
                        Console.WriteLine(string.Format("InitDocumentLibrary:'{0}'", title));
#endif
                        InitDocumentLibrary(title);
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
        private string CleanUrl(string url)
        { //Returns a URL with the relative URL appended to it
            ClientContext c;
            using (c = context)
            {
                c.Load(c.Web, x => x.ServerRelativeUrl);
                c.ExecuteQuery();
                return c.Web.ServerRelativeUrl + url;
            }
        }
        private List GetOrCreateList(string title, int type = 100)
        { //Creates a list if it does not already exist and returns it
            List output;
            using (var c = context)
            {
                //Look for any list with that title already
                c.Load(c.Web.Lists, x => x.Where(y => y.Title == title));
                c.ExecuteQuery();

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
                    c.ExecuteQuery();
                }
                return output;
            }
        }
        private void DeleteList(List list)
        { //This is used to delete large lists that can not be deleted otherwise. It removes the items in chunks before attempting to remote the list
            ListItemCollection items;
            do
            {
                items = list.GetItems(getItemsToDelete); //Get the next batch of items
                list.Context.Load(items, x => x.ListItemCollectionPosition); //Used to determine if it has all of the items or not
                list.Context.Load(items, x => x.Include(y => y.Id)); //Get the list Ids
                list.Context.ExecuteQuery();
                foreach (var item in items) //Iterate the items
                {
                    item.DeleteObject(); //Deleat items one at a time
                }
            }
            while (items.ListItemCollectionPosition != null); //Check if that was the last group of items

        }
        private int GetLastItemId(string listTitle)
        { //This is used to get the last item id from a list
            using (var c = context)
            {
                c.Load(c.Web, x => x.Lists.Where(y => y.Title == listTitle));
                c.ExecuteQuery();
                if (c.Web.Lists.Any())
                {
                    var items = c.Web.Lists.First().GetItems(getLastItemIdQuery); //Get the last Item added
                    c.Load(items, x => x.Include(y => y.Id)); //Queue a query to get the last Item Id
                    c.ExecuteQuery(); //Execute query to get the last item Id
                    if (items.Any())
                    {
                        return items.First().Id; //return the id
                    }
                }
            }
            return 0; //return 0 because there are no items
        }
        #region GenericList
        private void InitGenericList(string title, int count)
        { //Used to populate Generic lists
            var moreRecords = true;
            do
            {
                moreRecords = AddListItem(title, count);
            }
            while (moreRecords == true);
        }
        private void UpdateAttachments(string title, int id)
        { //Update attachments
            using (var c = context)
            {
                c.Load(c.Web, x => x.Lists.Where(y => y.Title == title));
                c.ExecuteQuery();
                if (c.Web.Lists.Any())
                {
                    var list = c.Web.Lists.First();
                    var items = list.GetItems(new CamlQuery()
                    {
                        ViewXml = string.Format(getItemByIdCamlView, id)
                    });
                    c.Load(items);
                    c.ExecuteQuery();
                    if (items.Any())
                    {
                        var item = items.First();
                        var attachmentPaths = GetSourceItemAttachmentPaths(title, id);
                        foreach (var path in attachmentPaths)
                        {
                            using(var stream = GetSourceFileStream(path))
                            {
                                item.AttachmentFiles.Add(new AttachmentCreationInformation() //Queue a query to write the stream as an attachment
                                {
                                    FileName = Path.GetFileName(path), //Set attachment name
                                    ContentStream = stream //Set attachment content
                                });
                                c.ExecuteQuery(); //execute queued queries
                            }
                        }
                    }
                }
            }
        }
        private void UpdateItemProperties(string title, int id)
        { //Update properties of an item
            using (var c = context)
            {
                c.Load(c.Web, x => x.Lists.Where(y => y.Title == title));
                c.ExecuteQuery();
                if (c.Web.Lists.Any())
                {
                    var list = c.Web.Lists.First();
                    var items = list.GetItems(new CamlQuery()
                    {
                        ViewXml = string.Format(getItemByIdCamlView, id)
                    });
                    c.Load(items);
                    c.ExecuteQuery();
                    if (items.Any())
                    {
                        var item = items.First();
                        var attributes = GetSourceItemAttributes(title, id); //Get properties from source
                        foreach (var attribute in itemAttributes)
                        {
                            item[attribute] = attributes[attribute];
                        }

                        item.Update(); //Trigger an item update so it gets inserted
                        c.ExecuteQuery();
                    }
                }
            }
            UpdateAttachments(title, id);
        }
        private bool AddListItem(string title, int count)
        {
            var itemId = -1;
            using (var c = context)
            {
                c.Load(c.Web, x => x.Lists.Where(y => y.Title == title));
                c.ExecuteQuery();
                if (c.Web.Lists.Any())
                {
                    var list = c.Web.Lists.First();
                    //Add blank item
                    var item = list.AddItem(new ListItemCreationInformation());
                    item.Update(); //Trigger an item update so it gets inserted
                    c.ExecuteQuery();
                    itemId = item.Id;
                }
            }
            if(itemId > -1)
            {
                UpdateItemProperties(title, itemId);
                return itemId >= count ? false : true; //Set boolean to indicate if there are more records to process
            }
            return false;
        }
        #endregion
        #region DocumentLibrary
        private void InitDocumentLibrary(string listTitle)
        { //Used to populate Document libraries
            using(var c = context)
            {
                c.Load(c.Web, x => x.Lists.Where(y => y.Title == listTitle));
                c.ExecuteQuery();
                if (c.Web.Lists.Any())
                {
                    List list = c.Web.Lists.First();
                    //Get the root folder to the list to start
                    Folder folder = list.RootFolder;
                    list.Context.Load(list, x => x.ParentWebUrl);
                    list.Context.Load(folder, x => x.ServerRelativeUrl);
                    list.Context.ExecuteQuery();

                    //Convert the path into a relative one
                    var index = folder.ServerRelativeUrl.IndexOf(list.ParentWebUrl);
                    var cleanPath = (index < 0) ? folder.ServerRelativeUrl : folder.ServerRelativeUrl.Remove(index, list.ParentWebUrl.Length);
                    PopulateFolder(cleanPath);
                }
            }
        }
        private void PopulateFolder(string url)
        {
            var folderNames = GetSourceFolderNames(url);
            var fileNames = GetSourceFileNames(url);
            using (var c = context)
            {
                var folder = c.Web.GetFolderByServerRelativeUrl(CleanUrl(url)); //Get the folder
                c.Load(folder, x => x.Files.Include(y => y.Name), x => x.Folders.Include(y => y.Name)); //Get folder and file names to avoid inserting duplicates
                c.ExecuteQuery();
                //Iterate the files
                foreach (var fileName in fileNames)
                {
                    if (!folder.Files.Any(x => x.Name == fileName)) //Make sure the file doesn't exist already
                    {
                        var sourceFileRelativeUrl = Path.Combine(url, fileName);

                        using (var fileStream = GetSourceFileStream(sourceFileRelativeUrl))
                        {
                            //Add file
                            folder.Files.Add(new FileCreationInformation()
                            {
                                Url = fileName, //Set filename
                                ContentStream = fileStream, //Get and set file stream
                                Overwrite = false //Do not overwrite files to save time
                            });
                            c.ExecuteQuery();
                        }
                    }
                }
                //Iterate the folders
                foreach (var folderName in folderNames)
                {
                    if (!folder.Folders.Any(x => x.Name == folderName))
                    {
                        folder.AddSubFolder(folderName); //Add folder
                        c.ExecuteQuery();
                    }
                    PopulateFolder(Path.Combine(url, folderName)); //Populate the folder
                }
            }
        }
        #endregion
        #endregion
    }
}
