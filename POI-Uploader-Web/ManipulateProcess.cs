using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Diagnostics;
using System.Threading;

//For using win32 API
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace Communication
{
    public static class ManipulateProcess
    {
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern bool SetForegroundWindow(IntPtr hwnd);

        //Start a new process according to the fileName
        public static void StartProcess(string fileName)
        {
            ProcessStartInfo pdfInfo = new ProcessStartInfo(fileName);
            pdfInfo.WindowStyle = ProcessWindowStyle.Normal;
            Process pdfProcess = Process.Start(pdfInfo);

        }

        public static Process GetPdfProcess(string fn)
        {
            string procName = @"AcroRd32";

            Process myProcess = null;
            while (true)
            {
                Thread.Sleep(100);
                Process[] myProcessArray = Process.GetProcessesByName(procName);
                int size = myProcessArray.Length;

                //Remember to enumerate over all the possible processes to find proper window handle
                for (int i = 0; i < size; i++)
                {
                    myProcessArray[i].Refresh();
                    if (myProcessArray[i].MainWindowHandle != IntPtr.Zero &&
                        myProcessArray[i].MainWindowTitle.StartsWith(fn) )
                    {
                        myProcess = myProcessArray[i];
                        break;
                    }
                }

                if (myProcess != null) break;
            }

            return myProcess;
        }

        //Get a certain process with proper MainWindowHandle
        public static Process GetProcess(string procType)
        {
            string procName = null;
            switch(procType)
            {
                case @"PDF":
                    procName = @"AcroRd32";
                    break;
            }

            if(procName == null)
            {
                return null;
            }

            Process myProcess = null;
            while (true)
            {
                Thread.Sleep(100);
                Process[] myProcessArray = Process.GetProcessesByName(procName);
                int size = myProcessArray.Length;

                //Remember to enumerate over all the possible processes to find proper window handle
                for (int i = 0; i < size; i++)
                {
                    myProcessArray[i].Refresh();
                    if (myProcessArray[i].MainWindowHandle != IntPtr.Zero)
                    {
                        myProcess = myProcessArray[i];
                        break;
                    }
                }

                if (myProcess != null) break;
            }

            Console.WriteLine("Return process: " + myProcess.Id);
            Console.WriteLine("Return process: " + myProcess.MainWindowTitle);

            return myProcess;
        }

        public static void SetForeGround(Process process)
        {
            SetForegroundWindow(process.MainWindowHandle);
        }

        [DllImport("User32")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);

        const int SW_HIDE = 0;

        public static void HideWindow(int hwnd)
        {
            ShowWindow(hwnd, SW_HIDE);
        }
    }
}
