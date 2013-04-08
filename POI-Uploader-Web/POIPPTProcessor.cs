using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

using System.Threading;
using System.IO;
using POILibCommunication;
using System.Diagnostics;


namespace POI_Uploader_Web
{
    class POIPPTProcessor
    {
        
        static string folderPath;
        static POISlideSaver saver;
        
        public static void Process(String fn,string nameFromTextBox, string presentorFromTextBox)
        {

            saver = new POISlideSaver(nameFromTextBox, presentorFromTextBox);

            folderPath = saver.FolderPath;
            
            PowerPoint.Application myApp = new PowerPoint.Application();
            PowerPoint.Presentations myPres = myApp.Presentations;
            

            //Open a certain PPT
            PowerPoint.Presentation sourcePre = myPres.Open(fn, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, Office.MsoTriState.msoFalse);
            PowerPoint.Presentation container;


            
            String fileName;

            DateTime startTime = DateTime.Now;

            sourcePre.SaveAs(folderPath, PowerPoint.PpSaveAsFileType.ppSaveAsPNG);

            foreach (PowerPoint.Slide curSlide in sourcePre.Slides)
            {
                List<int> durationList = new List<int>();
                fileName = folderPath + "/" + (curSlide.SlideIndex-1)+".PNG";;
                string savedFileName = folderPath + "/Slide" + (curSlide.SlideIndex)+".PNG";
                FileStream ins = new FileStream(savedFileName, FileMode.Open, FileAccess.Read);
                FileStream os = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write);
                ins.CopyTo(os);
                os.Flush();

                int animationCount = curSlide.TimeLine.MainSequence.Count;
                if (animationCount > 0)
                {
                    PowerPoint.Sequence curAnimationSequence = curSlide.TimeLine.MainSequence;
                    float totalTime = 0;
                    foreach (PowerPoint.Effect effect in curAnimationSequence)
                    {
                        Console.WriteLine(effect.Timing.Duration);

                        totalTime += effect.Timing.Duration;

                        durationList.Add((int)effect.Timing.Duration);
                        break;
                    }

                    
                    container = myPres.Add(Office.MsoTriState.msoTrue);
                    
                    curSlide.Copy();
                    container.Windows[1].Activate();
                    myApp.CommandBars.ExecuteMso(@"PasteSourceFormatting");

                    container.CreateVideo(folderPath + "/" + (curSlide.SlideIndex-1)+".wmv", true, (int) totalTime);
                    while (container.CreateVideoStatus != PowerPoint.PpMediaTaskStatus.ppMediaTaskStatusDone)
                    {
                        Thread.Sleep(1000);
                    }

                    container.Close();

                    //Once the video creation is done, convert it to wmv
                    Thread convThread = new Thread(StartVideoConversion);
                    String [] param = new String[2];
                    param[0] = folderPath;
                    param[1] = (curSlide.SlideIndex-1).ToString();
                    convThread.Start(param);
                }
            
                
                if (animationCount > 0)
                {
                    saver.saveSlideAnimationToPresentation(curSlide.SlideIndex - 1, durationList);
                }
                else
                {
                    saver.saveSlideImageToPresentation(curSlide.SlideIndex - 1);
                }

                
            }

            DateTime endTime = DateTime.Now;

            Console.WriteLine(@"Time consumed in seconds: " + (endTime - startTime).TotalSeconds);

            saver.saveToPOIFile();
            myApp.Quit();
        }

        private static void StartVideoConversion(object data)
        {
            String[] param = data as String[];
            String folderPath = param[0];
            String fileName = param[1];

            String inputFN = Path.Combine(folderPath, fileName + ".wmv");
            String outputFN = Path.Combine(folderPath, fileName + ".mp4");

            //Start a cmd process which trigger ffmpeg
            Process process = new Process();

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";

            //Important: specify the options for ffmpeg
            //The PIPE_NAME is already defined so can be used here as input file
            //"/C" let the console close after operation completes
            startInfo.Arguments = "/C ffmpeg -i " + inputFN + " -f mp4 -acodec libfaac -ab 128k -ar 48000 -ac 2 -vcodec libx264 " + outputFN;
            process.StartInfo = startInfo;
            process.Start();
        }

        
        private static void StartScreenShot(object data)
        {
            IntPtr handle = (IntPtr)data;
            int counter = 0;
            MemoryStream myStream;
            FileStream fs;

            String path = Directory.GetCurrentDirectory();
            while (true)
            {
                myStream = Communication.ScreenShot.TakeScreenShot(handle);

                fs = File.OpenWrite(path + @"/img" + @"_" + counter + @".jpg");
                myStream.WriteTo(fs);
                fs.Close();

                counter++;
                Console.WriteLine(counter);
                Thread.Sleep(1);
            }
        }
    }
}
