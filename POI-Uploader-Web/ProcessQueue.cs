using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Web;

using System.Threading;


namespace POI_Uploader_Web
{
    //Singleton class for queuing the incoming uploading request
    public class ProcessQueue : ConcurrentQueue<object>
    {
        static ProcessQueue instance = null;
        public delegate void RequestHandler(object arg);

        public WaitHandle queueCV = null;

        private ProcessQueue()
        {
            //Initialize the conditional variable
            queueCV = new AutoResetEvent(false);
        }

        public static ProcessQueue SharedInstance
        {
            get 
            {
                if (instance == null)
                {
                    instance = new ProcessQueue();
                }

                return instance;
            }
        }

        public static void EnqueueRequest(object arg)
        {
            SharedInstance.Enqueue(arg);

            //Signal the new request
            (SharedInstance.queueCV as AutoResetEvent).Set();
        }

        public static object DequeueRequest()
        {
            object arg = null;
            SharedInstance.TryDequeue(out arg);

            return arg;
        }

        public static void Run(RequestHandler handler)
        {
            //Force the queue to initialize
            ProcessQueue myQueue = ProcessQueue.SharedInstance;

            Thread handlerThread = new Thread(() =>
                {
                    while (true)
                    {
                        if (ProcessQueue.SharedInstance.Count > 0)
                        {
                            handler(ProcessQueue.DequeueRequest());
                        }
                        else
                        {
                            //Wait for the new request to come
                            ProcessQueue.SharedInstance.queueCV.WaitOne();
                        }
                    }
                }
            );

            handlerThread.Start();
        }
    }
}