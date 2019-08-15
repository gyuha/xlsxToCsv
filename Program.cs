using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsxToCsv
{
    class Program
    {
        static ExcelConvert excelConvert = new ExcelConvert();

        static int fileList()
        {
            var allfiles = System.IO.Directory.GetFiles(".", "*.*", System.IO.SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));
            if ( allfiles.Count() == 0)
            {
                Console.WriteLine("File not found..");
                return 1;
            }
            foreach (string src in allfiles)
            {
                var tar = Path.ChangeExtension(src, "csv");
                Console.WriteLine(src + " => " + tar);
                try {
                    if(excelConvert.Convert(src, tar) == false)
                    {
                        Console.WriteLine("ERROR : ["+src+"] can't read excel file.");
                    }
                } catch( InvalidCastException e)
                {
                    Console.WriteLine(e.ToString());
                }
            }
            return 0;
        }

        static void printHelp()
        {
            string filename = Process.GetCurrentProcess().ProcessName;
            Console.WriteLine("Excel to CSV(UTF-8).  \n\r");
            Console.WriteLine("USING 1.");
            Console.WriteLine("\t"+ filename);
            Console.WriteLine("\n\rUSING 2.");
            Console.WriteLine("\t"+ filename +" [Source FilePath] [Target FilePath] [-h]\n\r");
            Console.WriteLine("\t  -h\t Display help");
        }

        static int Main(string[] args)
        {
            foreach(string a in args )
            {
                if( String.Compare(a, "-h", true) == 0)
                {
                    printHelp();
                    return 0;
                }
            }

            if (args.Length != 0 && args.Length != 2)
            {
                printHelp();
                return 1;
            }

            // 폴더를 검색해서 변경
            if (args.Length == 0)
            {
                return fileList();
            }
            
            // cli file change
            string srcFile = args[0];
            string tarFile = args[1];

            if(!File.Exists(srcFile))
            {
                Console.WriteLine("File not found. [" + srcFile + "]");
                return 1;
            }

            string tarPath = Path.GetDirectoryName(tarFile);
            if (!Directory.Exists(tarPath) && tarPath.Length > 0)
            {
                Console.WriteLine("Target Folder not exist.");
                return 1;
            }

            Console.WriteLine(srcFile + " => " + tarFile);
            if (excelConvert.Convert(srcFile, tarFile) == false)
            {
                Console.WriteLine("ERROR : [" + srcFile + "] can't read excel file.");
                return 1;
            }

            return 0;
        }
    }
}
