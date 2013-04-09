using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Communication;
using System.Diagnostics;
using System.Threading;
using System.IO;

using POILibCommunication;

namespace POI_Uploader_Web
{
    class POIPDFProcessor
    {
        static POISlideSaver saver;
        public static void Process(String fn,string nameFromTextBox, string presentorFromTextBox)
        {

            saver = new POISlideSaver(nameFromTextBox, presentorFromTextBox);
            //Determine the page count
            StreamReader sr = new StreamReader(File.OpenRead(fn));
            Regex regex = new Regex(@"/Type\s*/Page[^s]");
            MatchCollection matches = regex.Matches(sr.ReadToEnd());
            int numPages = matches.Count;
            sr.Close();

            
            
            //Start a PDF process and set to full screen mode
            ManipulateProcess.StartProcess(fn);

            //Get the index of the file name without the path
            int fnStartIndex = fn.LastIndexOf('\\') + 1; 

            //Process pdfProcess = ManipulateProcess.GetProcess(@"PDF");
            Process pdfProcess = ManipulateProcess.GetPdfProcess(fn.Substring(fnStartIndex));
            ManipulateProcess.SetForeGround(pdfProcess);
            EmulateIO.KeyStroke(@"^l");

            MemoryStream myStream;
            FileStream fs;
            String path = Directory.GetCurrentDirectory();

            for (int i = 0; i < numPages; i++)
            {
                ManipulateProcess.SetForeGround(pdfProcess);
                EmulateIO.KeyStroke(@"{PGDN}");

                Thread.Sleep(100);

                //Take the screen shot  q
                myStream = ScreenShot.TakeScreenShot(pdfProcess);
                String savedFileName = saver.FolderPath + @"/" + i + @".png";
                fs = File.OpenWrite(savedFileName);
                myStream.WriteTo(fs);
                fs.Close();
                saver.saveSlideImageToPresentation(i);

                
            }
            pdfProcess.CloseMainWindow();
            saver.saveToPOIFile();
        }
    }
}
