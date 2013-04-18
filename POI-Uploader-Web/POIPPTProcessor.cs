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
using System.Collections;
using System.Text.RegularExpressions;


namespace POI_Uploader_Web
{
    class POIPPTProcessor
    {
        
        static string folderPath;
        static POISlideSaver saver;
        static int presID;
        
        public static void Process(String fn,string name, string description, int presId)
        {
            presID = presID;

            saver = new POISlideSaver(name, description, presId);
            
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

                //Copy to the local disk
                ins.CopyTo(os);
                os.Flush();

                ins.Close();
                os.Close();


                GetTextCommentsOnEachSlideAndStoreToFile(curSlide);

                int animationCount = curSlide.TimeLine.MainSequence.Count;
                if (animationCount > 0)
                {
                    PowerPoint.Sequence curAnimationSequence = curSlide.TimeLine.MainSequence;
                    float totalTime = 0;
                    float curDuration = 0;
                    float allButLastDuration = 0;
                    System.Collections.IEnumerator enumerator = curAnimationSequence.GetEnumerator();

                    foreach (PowerPoint.Effect effect in curAnimationSequence)
                    {
                        if (curDuration == 0)
                        {
                            curDuration += effect.Timing.Duration;
                        }
                        else
                        {
                            if (effect.Timing.TriggerType == PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                            {
                                durationList.Add((int)curDuration);
                                allButLastDuration += curDuration;
                                curDuration = 0;
                            }

                            curDuration += effect.Timing.Duration;
                        }

                        totalTime += effect.Timing.Duration;
                    }
                    durationList.Add((int)(totalTime - allButLastDuration));


                    
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

                    //Once the video creation is done, convert it to wmv, wait until completion
                    String [] param = new String[2];
                    param[0] = folderPath;
                    param[1] = (curSlide.SlideIndex-1).ToString();
                    StartVideoConversion(param);
                    GetMouseClickImageFromSlideAccordingToTime(curSlide.SlideIndex, durationList,(int)totalTime);
                 
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

            GC.Collect();
            GC.WaitForPendingFinalizers();

            

            sourcePre.Close();
           

            saver.saveToPOIFile();

            foreach (Process process in System.Diagnostics.Process.GetProcessesByName("POWERPNT.EXE"))
            {
                process.Kill();
            }
            myApp.Quit();
            
        }

        private static void GetTextCommentsOnEachSlideAndStoreToFile(PowerPoint.Slide slide)
        {
            PowerPoint.Shapes shapes = slide.Shapes;
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    String text = shape.TextFrame.TextRange.Text;
                    StoreStringToFileWithPresIDAndIndex(slide.SlideIndex, text);
                     
                }
            }
        }
        private static void StoreStringToFileWithPresIDAndIndex(int index,string text)
        {
            

            text = ReplaceNoAlphanumericWithSpace(text);

            if (!String.IsNullOrWhiteSpace(text))
            {
                WriteProcessedTextToFile(index, text);
            }
        }

        private static void WriteProcessedTextToFile(int index, string text)
        {
            FileStream keywordFileStream = new FileStream(Path.Combine(POIArchive.ArchiveHome,
                POIGlobalVar.KeywordsFileName), FileMode.Append);
            TextWriter stringTextWriter = new StreamWriter(keywordFileStream);
            String stringWithPresIDAndIndex = presID + " " + index + " " + text;
            stringTextWriter.WriteLine(stringWithPresIDAndIndex);
            stringTextWriter.Close();
            keywordFileStream.Close();
        }
        private static string ReplaceNoAlphanumericWithSpace(string text)
        {
            text = text.Replace(System.Environment.NewLine, " ").Replace("\r", " ");
            Regex rgx = new Regex("[^a-zA-Z0-9 -]");
            text = rgx.Replace(text, " ");

            return text;
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

            process.WaitForExit();
        }

        private static void GetMouseClickImageFromSlideAccordingToTime(int slideIndex, List<int> mouseClickTimeList,int totalTime)
        {
            String movieName = folderPath +"/"+ (slideIndex - 1) + ".wmv";
            int mouseClickTime = 0;
            int mouseClickIndex = 0;
            for (mouseClickIndex=0;mouseClickIndex<mouseClickTimeList.Count; mouseClickIndex++)
            {
                mouseClickTime += mouseClickTimeList[mouseClickIndex];
                ExtractImageFromMovieUseFFmpeg(mouseClickIndex, slideIndex, mouseClickTime, movieName);
            }
            GetLastImageOfTheMovie(mouseClickIndex, slideIndex, totalTime, movieName);

        }

        private static void GetLastImageOfTheMovie(int lastImageIndex,int slideIndex, int totalTime, String movieName)
        {
            ExtractImageFromMovieUseFFmpeg(lastImageIndex, slideIndex, totalTime, movieName);
        }
        private static void ExtractImageFromMovieUseFFmpeg(int imageIndex, int slideIndex, int mouseClickTime, String movieName)
        {
            String imageName = folderPath + "/" + slideIndex + "_" + imageIndex + ".png";
            String ffmpegImageExtractionCommand = "/C ffmpeg -i " + movieName + " -ss " + TimeFormatForFFmpeg(mouseClickTime) + " -f image2 -vframes 1 " + imageName;

            ExecuteFFmpegImageCommand(ffmpegImageExtractionCommand);
        }
        private static String TimeFormatForFFmpeg(int totalSeconds)
        {
            int hours = totalSeconds / 3600;
            int minutes = (totalSeconds - 3600 * hours) / 60;
            int seconds = (totalSeconds - 3600 * hours - 60 * minutes);

            return String.Format("{0:00}:{1:00}:{2:00}", hours, minutes, seconds);
        }
        private static void ExecuteFFmpegImageCommand(String command)
        {
            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "cmd.exe";

            //"/C" let the console close after operation completes
            startInfo.Arguments = command;
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
