using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface IDestination
    {
        void InjectDependencies(Func<string, int, IDictionary<string, string>> GetSourceItemAttributes, Func<string, int, Task<IDictionary<string, Stream>>> GetSourceItemAttachments, Func<string, Task<IEnumerable<string>>> GetSourceFolderNames, Func<string, Task<IEnumerable<string>>> GetSourceFileNames, Func<string, Task<Stream>> GetSourceFileStream);
        Task AddList(string title, int type, int itemCount);
    }
}
