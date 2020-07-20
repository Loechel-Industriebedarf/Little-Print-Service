using Spire.Pdf;
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
        public static string toPrintFolder { get; set; }
        public static string printedFolder { get; set; }
        public static string printerName { get; set; }

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
                        CoreLoop(toPrintFolder, printedFolder, printerName); //Print all files in directory
                        System.Threading.Thread.Sleep(1000); //Wait 1 second, then check again
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
            DirectoryInfo d = new DirectoryInfo(toPrintFolder);
            FileInfo[] Files = d.GetFiles("*"); //Get all files from directory

            //Loop throught files
            foreach (FileInfo file in Files)
            {
                //Print file
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

            foreach (var foo in toPrintFolderTemp) { toPrintFolder = foo.Value; }
            foreach (var foo in printedFolderTemp) { printedFolder = foo.Value; }
            foreach (var foo in printerNameTemp) { printerName = foo.Value; }
        }
    }
}
