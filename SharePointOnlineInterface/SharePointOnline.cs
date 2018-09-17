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
        private int startIndex { get; set; } = -1;
        private int endIndex { get; set; } = -1;
        private ISource source { get; set; }
        private string getItemByIdCamlView { get; } = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Number'>{0}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>";
        private SharePointOnlineCredentials credentials { get; set; }
        private string url { get; set; }
        #endregion
        #region Public
        public SharePointOnline(string url, string username, string password, int startIndex = -1, int endIndex = -1)
        {
            if (username == null || password == null) //username and passwords are required for this class
            {
                throw new ArgumentException(message: "Arguments 'username' and 'password' are required for SharePoint Online URLs");
            }

            this.url = url; //Assign the passed URL to a private variable
            this.startIndex = startIndex;
            this.endIndex = endIndex;

            //Convert the password into a secure string
            SecureString pass = new SecureString();
            foreach (var character in password)
            {
                pass.AppendChar(character);
            }

            this.credentials = new SharePointOnlineCredentials(username, pass); //Create credentials to be used
        }
        #region MethodsNeededToUseThisAsADestination
        public void InjectDependencies(ISource _source)
        {
            this.source = _source;
        }
        public void AddList(string title, int type)
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
                        InitGenericList(title);
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
        #region GenericList
        private void InitGenericList(string title)
        { //Used to populate Generic lists
            for (var i = startIndex; i < endIndex; i++){
                UpdateAttachments(title, i);
            }

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
                        var attachmentPaths = source.GetItemAttachmentPaths(title, id);
                        foreach (var path in attachmentPaths)
                        {
                            using(var stream = source.GetFileStream(path))
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
      
#endregion
#endregion
    }
}
