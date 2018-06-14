using SharePoint2010Interface;
using SharePointOnlineInterface;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MigrateToO365Async
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Task> tasks = new List<Task>();
            //TODO: Find a better way to pass in arguments and potentially credentials for SP2010
            var source = new SharePoint2010(args[0]);
            var destination = new SharePointOnline(args[1], args[2], args[3], source.GetItemAttributes, source.GetItemAttachments, source.GetFolderNames, source.GetFileNames, source.GetFileStream);

            var sourceLists = source.GetLists();
            //sourceLists = sourceLists.Where(x => x.Title == "AnnuityMet"); //Debugging
            foreach (var list in sourceLists)
            {
                tasks.Add(destination.AddList(list.Title, list.Type, list.ItemCount));
            }
            Task.WaitAll(tasks.ToArray());
            tasks.Clear();
        }
    }
}
