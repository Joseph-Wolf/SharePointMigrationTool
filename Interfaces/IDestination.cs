using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface IDestination
    {
        void InjectDependencies(Func<string, int, IDictionary<string, string>> GetSourceItemAttributes, Func<string, int, IEnumerable<string>> GetSourceItemAttachmentPaths, Func<string, IEnumerable<string>> GetSourceFolderNames, Func<string, IEnumerable<string>> GetSourceFileNames, Func<string, Stream> GetSourceFileStream);
        void AddList(string title, int type, int itemCount);
    }
}
