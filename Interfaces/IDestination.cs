using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface IDestination
    {
        void InjectDependencies(ISource source);
        void AddList(string title, int type);
    }
}
