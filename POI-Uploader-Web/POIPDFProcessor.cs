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
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace POI_Uploader_Web
{
    class POIPDFProcessor
    {
        static POISlideSaver saver;
        static string inputPdf;
        static string folderPath;

        public static void Process(String fn,string name, string description, int presId)
        {
            int numPages = 0;
            saver = new POISlideSaver(name, description, presId);

            folderPath = saver.FolderPath;
            inputPdf = fn;

            //Use pdf reader to extract information
            try
            {
                PdfReader reader = new PdfReader(fn);
                numPages = reader.NumberOfPages;

                PdfReaderContentParser parser = new PdfReaderContentParser(reader);
                
                for (int i = 0; i < numPages; i++)
                {
                    string pageText = parser.ProcessContent(
                        i + 1, 
                        new SimpleTextExtractionStrategy()
                    ).GetResultantText();

                    Console.WriteLine("temp");

                    //Convert the slide into png
                    startPdfToPngConversion(i);
                    saver.saveSlideImageToPresentation(i);
                }

                saver.saveToPOIFile();
            }
            catch (Exception e)
            {
                POIGlobalVar.POIDebugLog(e);
            }
            
        }

        public static void startPdfToPngConversion(int slideIndex)
        {
            string outputFN = Path.Combine(folderPath, slideIndex + ".png");

            //Start a cmd process which trigger ffmpeg
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";
            startInfo.Verb = "runas";

            int pageIndex = slideIndex + 1;
            startInfo.Arguments = "/C gswin64c -dNOPAUSE -dBATCH -dNOPROMPT"
                + " -dFirstPage=" + pageIndex
                + " -dLastPage=" + pageIndex
                + " -sDEVICE=pngalpha -r96 -sOutputFile=" + outputFN
                + " " + inputPdf;

            POIGlobalVar.POIDebugLog(startInfo.Arguments);

            process.StartInfo = startInfo;
            process.Start();
            process.WaitForExit();
        }
    }
}
