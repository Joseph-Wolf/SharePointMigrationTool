using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface ISource
    {
        IEnumerable<SourceList> GetLists();
        IDictionary<string, string> GetItemAttributes(string listTitle, int itemId);
        IEnumerable<string> GetFolderNames(string url);
        IEnumerable<string> GetFileNames(string url);
        Stream GetFileStream(string url);
        IEnumerable<string> GetItemAttachmentPaths(string listTitle, int itemId);
    }
}
