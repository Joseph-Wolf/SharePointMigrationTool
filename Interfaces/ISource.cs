using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface ISource
    {
        IEnumerable<SourceList> GetLists();
        IDictionary<string, string> GetItemAttributes(string listTitle, int itemId);
        Task<IEnumerable<string>> GetFolderNames(string url);
        Task<IEnumerable<string>> GetFileNames(string url);
        Task<Stream> GetFileStream(string url);
        Task<IEnumerable<string>> GetItemAttachmentPaths(string listTitle, int itemId);
    }
}
