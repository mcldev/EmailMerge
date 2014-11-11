using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EmailMerge;

namespace wtf
{
    class Program
    {
        /// <summary>
        /// Runs the Email Merge program to either merge pst2 into pst1, or into a new merged pst file.
        /// </summary>
        /// <param name="args">Inputs in order:
        /// 1- pst1 filename
        /// 2- pst2 filename
        /// 3- [optional] merged filename
        /// 4- [optional] duplicates filename from pst1
        /// 5- [optional] duplicates filename from pst2
        /// 6- [optional] boolean save duplicates
        /// 7- [optional] comma delim string of folders to ignore</param>
        static void Main(string[] args)
        {
            try
            {
                if (!args.Any() || args.Count() < 2) throw new ArgumentException("You must provide at least 2 pst filenames to merge");

                Console.WriteLine("Loading and Merging PST Files");

                string pst1 = args[0];
                string pst2 = args[1];

                string pstMerged = null;
                string duplicates1 = null;
                string duplicates2 = null;
                bool saveDuplicates = true;
                string[] foldersToIgnore = null;

                for (int i = 2; i < args.Count(); i++)
                {
                    switch (i)
                    {
                        case 2:
                            pstMerged = args[i];
                            break;
                        case 3:
                            duplicates1 = args[i];
                            break;
                        case 4:
                            duplicates2 = args[i];
                            break;
                        case 5:
                            if (!bool.TryParse(args[i], out saveDuplicates))
                                throw new ArgumentException(String.Format("Error converting string to boolean: '{0}'", args[i]));
                            break;
                        case 6:
                            foldersToIgnore = args[i].Split(',');
                            break;
                    }
                }

                //Merge PST Files
                PSTFile.MergePSTFiles(pstFilename1: pst1,
                                        pstFilename2: pst2,
                                        mergedFilename: pstMerged,
                                        duplicatesFilename1: duplicates1,
                                        duplicatesFilename2: duplicates2,
                                        saveDuplicates: saveDuplicates,
                                        foldersToIgnore: foldersToIgnore);
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            Console.WriteLine("Merge Finished... press any key to exit");
            Console.ReadKey(true);
        }

        
    }
}
