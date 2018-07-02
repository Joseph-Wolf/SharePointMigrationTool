using Interfaces;
using SharePoint2010Interface;
using SharePointOnlineInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MigrateToO365Async
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceUrl = null;
            string sourceUsername = null;
            string sourcePassword = null;
            string destinationUrl = null;
            string destinationUsername = null;
            string destinationPassword = null;
            string key;
            string value;
            ISource source;
            IDestination destination;
            IEnumerable<SourceList> sourceLists;
            char commandSplitCharacter = '=';

            //Parse argument array into variables
            foreach(var argument in args)
            {
                if (argument.IndexOf(commandSplitCharacter) >= 0) //Check that argument has expected splitting characters
                {
                    key = argument.Split(commandSplitCharacter).First();
                    value = argument.Split(commandSplitCharacter).Last();
                    switch (key.ToUpper())
                    {
                        case "SOURCEURL":
                            sourceUrl = value;
                            break;
                        case "SOURCEUSERNAME":
                            sourceUsername = value;
                            break;
                        case "SOURCEPASSWORD":
                            sourcePassword = value;
                            break;
                        case "DESTINATIONURL":
                            destinationUrl = value;
                            break;
                        case "DESTINATIONUSERNAME":
                            destinationUsername = value;
                            break;
                        case "DESTINATIONPASSWORD":
                            destinationPassword = value;
                            break;
                        default:
                            break;
                    }
                }
                else
                { //Invalid arguments passed so post the expected ones
                    Console.WriteLine("Arguments should be in the format of 'name=value'.");
                    Console.WriteLine("Available Arguments include:");
                    Console.WriteLine();
                    Console.WriteLine("sourceURL");
                    Console.WriteLine("   The source SharePoint URL that you want to migrate the files from");
                    Console.WriteLine("sourceUsername");
                    Console.WriteLine("   The destination SharePoint Username to use for connecting");
                    Console.WriteLine("sourcePassword");
                    Console.WriteLine("   The destination SharePoint Password to use for connecting");
                    Console.WriteLine("destinationURL");
                    Console.WriteLine("   The destination SharePoint URL that you want to migrate the files to");
                    Console.WriteLine("destinationUsername");
                    Console.WriteLine("   The destination SharePoint Username to use for connecting");
                    Console.WriteLine("destinationPassword");
                    Console.WriteLine("   The destination SharePoint Password to use for connecting");
                    return;
                }
            }
            //Require specific variables
            if (sourceUrl == null || destinationUrl == null)
            {
                Console.WriteLine("Arguments 'SourceUrl' and 'DestinationUrl' are required.");
                return;
            }
            try
            {
                //TODO: Determine which SharePoint class source/destination should be used automatically
                source = new SharePoint2010(sourceUrl, sourceUsername, sourcePassword);
                destination = new SharePointOnline(destinationUrl, destinationUsername, destinationPassword);
            }
            catch (ArgumentException ex) //Catch exceptions thrown by invalid arguments (if username/password is required)
            {
                Console.WriteLine(ex.Message);
                return;
            }
            //Inject required methods from the source to the destination class
            destination.InjectDependencies(source.GetItemAttributes, source.GetItemAttachmentPaths, source.GetFolderNames, source.GetFileNames, source.GetFileStream);
            //Get lists present in the source
            sourceLists = source.GetLists();
            //Iterate and add lists
            //sourceLists = sourceLists.Where(x => x.Title == "ActivityAttachment"); //Debugging
            foreach (var list in sourceLists)
            {
                destination.AddList(list.Title, list.Type, list.ItemCount);
            }
        }
    }
}
