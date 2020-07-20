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
        /// 
        /// </summary>
        /// <param name="toPrintFolder"></param>
        /// <param name="printedFolder"></param>
        /// <param name="printerName"></param>
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
        /// 
        /// </summary>
        public static void ReadSettings()
        {
            XDocument doc = XDocument.Load("settings.xml");

            var toPrintFolderTemp = doc.Descendants("toPrintFolder");
            var printedFolderTemp = doc.Descendants("printedFolder");
            var printerNameTemp = doc.Descendants("printerName");

            foreach (var foo in toPrintFolderTemp) { toPrintFolder = foo.Value; }
            foreach (var foo in printedFolderTemp) { printedFolder = foo.Value; }
            foreach (var foo in printerNameTemp) { printerName = foo.Value; }
        }
    }
}
