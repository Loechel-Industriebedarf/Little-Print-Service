using Spire.Pdf;
using Spire.Pdf.Print;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PrintService
{
    public class Program
    { 
        public static string ToPrintFolder { get; set; }
        public static string PrintedFolder { get; set; }
        public static string PrinterName { get; set; }
        public static string PrintKeyword { get; set; } //Only print, if this keyword is in the filename
        public static int SleepTime { get; set; }

        /// <summary>
        /// The main method calls a function to read the programs settings.
        /// After that, it enters the core loop.
        /// The core loop looks every second if a specific folder contains new files.
        /// If there are new files, they get printed and moved to a different directory.
        /// </summary>
        /// <param name="args">Console args. They are not used here at the moment.</param>
        static void Main(string[] args)
        {
            ReadSettings();

            while (true)
            {
                try
                {
                    CoreLoop(ToPrintFolder, PrintedFolder, PrinterName); //Print all files in directory
                    System.Threading.Thread.Sleep(SleepTime); //Wait x seconds, then check again
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }      
        }

        /// <summary>
        /// This method scans a folder for files. 
        /// If it finds any, it prints them and moves them to a diffent directory after.
        /// </summary>
        /// <param name="toPrintFolder">Folder with files that should be printed.</param>
        /// <param name="printedFolder">After a file was printed, move it to this folder.</param>
        /// <param name="printerName">Name of the printer, that should be used.</param>
        private static void CoreLoop(string toPrintFolder, string printedFolder, string printerName)
        {
            string[] PrintKeywords = PrintKeyword.Split('$');

            DirectoryInfo d = new DirectoryInfo(toPrintFolder);
            //Only print, if the keyword is in the filename
            List<FileInfo> Files = new List<FileInfo>();
            foreach(string keyword in PrintKeywords)
            {
                Files.AddRange(d.GetFiles("*" + keyword + "*")); //Get all files from directory
            }
            

            //Loop throught files
            foreach (FileInfo file in Files)
            {
                //Print file using freespire.pdf
                PdfDocument pdfdocument = new PdfDocument();
                pdfdocument.LoadFromFile(file.FullName);
                pdfdocument.PrintSettings.PrinterName = printerName;
                pdfdocument.PrintSettings.Copies = 1;
                pdfdocument.Print();
                pdfdocument.Dispose();


                string destFileName = printedFolder + "\\" + file.Name;
                //If file already exists: Delete it
                if (File.Exists(destFileName)) { File.Delete(destFileName); }
                //Move file to new folder
                File.Move(file.FullName, printedFolder + "\\" + file.Name);

                //Console log
                Console.WriteLine(DateTimeOffset.Now.ToString("dd.MM.yyyy HH:mm:ss") + ": " + file.Name + " wurde erfolgreich gedruckt.");
            }
        }

        /// <summary>
        /// Reads settings from xml file.
        /// </summary>
        public static void ReadSettings()
        {
            XDocument doc = XDocument.Load("settings.xml"); //Settings file

            var toPrintFolderTemp = doc.Descendants("toPrintFolder");
            var printedFolderTemp = doc.Descendants("printedFolder");
            var printerNameTemp = doc.Descendants("printerName");
            var printKeywordTemp = doc.Descendants("printKeyword");
            var sleepTimeTemp = doc.Descendants("sleepTime");

            foreach (var foo in toPrintFolderTemp) { ToPrintFolder = foo.Value; }
            foreach (var foo in printedFolderTemp) { PrintedFolder = foo.Value; }
            foreach (var foo in printerNameTemp) { PrinterName = foo.Value; }
            foreach (var foo in printKeywordTemp) { PrintKeyword = foo.Value; }
            foreach (var foo in sleepTimeTemp) { SleepTime = Convert.ToInt16(foo.Value); }
        }
    }
}
