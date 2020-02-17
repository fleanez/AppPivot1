using System;

namespace AppStitch
{
    class StitchTool
    {
        static void Main(string[] args)
        {

            if (args.Length == 2)
            {
                int nFiles = Utils.StitchAllFilesInFolders(args[0], args[1]);
                Console.WriteLine($"Finished stitching {nFiles} '{args[2]}' files to folder { args[1]}");
            }
            else if (args.Length == 3)
            {
                int nFiles = Utils.StitchAllFilesInFolders(args[0], args[1], args[2]);
                Console.WriteLine($"Finished stitching {nFiles} '{args[2]}' files to folder { args[1]}");
            }
            else
            {
                Console.WriteLine("Incorrect argument number! 2 first arguments are not optional:");
                Console.WriteLine("Usage:");
                Console.WriteLine("   STITCHTOOL [SOURCE_FOLDER] [DESTINATION_FOLDER] [FILE_EXTENSION]");
                Console.WriteLine("Arguments:");
                Console.WriteLine("   [SOURCE_FOLDER]         Source folder (eg. c:/temp/src)");
                Console.WriteLine("   [DESTINATION_FOLDER]    Destination folder (eg. c:/temp/dest)");
                Console.WriteLine("   [FILE_EXTENSION]        (optional) Filter file extension (eg. .csv). Include dot (.) is necessary");
            }
            Console.ReadKey();
        }
    }
}
