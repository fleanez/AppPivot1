using System;
using System.IO;
using System.Text;

namespace AppStitch
{
    class Utils
    {
        /// <summary>
        /// Stitch (append) data from source file to destination file
        /// </summary>
        /// <param name="strSrcFile">complete path to source file</param>
        /// <param name="strDestFile">complete path to destination file</param>
        /// <returns></returns>
        public static int StitchTextFiles (String strSrcFile, String strDestFile)
        {
            String[] srcFileLines = File.ReadAllLines(strSrcFile);
            var csv = new StringBuilder();
            int nCont = 0;

            if (File.Exists(strDestFile))
            {
                foreach (String srcLine in File.ReadLines(strSrcFile))
                {
                    if (nCont > 0)
                    {
                        csv.AppendLine(srcLine);
                    }
                    nCont++;
                }
                File.AppendAllText(strDestFile, csv.ToString());
                csv.Clear();
            }

            return nCont;
        }

        /// <summary>
        /// Stitch (append) data from files in sorce folder to files with the same name in destination folder
        /// </summary>
        /// <param name="strSrcFolder">sorce folder</param>
        /// <param name="strDestFolder">destination folder</param>
        /// <param name="strExtension">file extension</param>
        /// <returns></returns>
        public static int StitchAllFilesInFolders(String strSrcFolder, String strDestFolder, String strExtension = "")
        {
            String strSrcExtension;
            String strSrcFileName;
            String strDestFileName;

            if (!Directory.Exists(strSrcFolder))
            {
                throw new Exception("Destination folder not found: " + strDestFolder);
            }

            int nCont = 0;
            foreach (String strSrcFilePath in Directory.GetFiles(strSrcFolder))
            {
                strSrcExtension = Path.GetExtension(strSrcFilePath);
                if ((strSrcExtension.ToLower() == strExtension.ToLower()) || strExtension == "")
                {
                    strSrcFileName = Path.GetFileName(strSrcFilePath);
                    foreach (String strDestFilePath in Directory.GetFiles(strDestFolder))
                    {
                        strDestFileName = Path.GetFileName(strDestFilePath);
                        if (strSrcFileName == strDestFileName)
                        {
                            StitchTextFiles(strSrcFilePath, strDestFilePath);
                            nCont++;
                            break;
                        }
                    }
                }
            }
            return nCont;
        }

    }
}
