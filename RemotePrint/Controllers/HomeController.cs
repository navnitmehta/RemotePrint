using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Configuration;
using System.Web;
using System.Web.Mvc;
using System.Printing;
using System.Printing.IndexedProperties;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RemotePrint.Controllers
{
    public class HomeController : Controller
    {
        private string printServerAddress = @"\\ISTS-NAVNIT-PC\NPIEA8082"; //"RemotePrintingTest";
        string printQueue = "HP LaserJet 400 MFP M425 PCL 6"; //"RemotePrintingTest"

        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            TrackPrinterStatus();

            return View();
        }

        public ActionResult Print()
        {
            ViewBag.Title = "Print";

            //--We split by page

            FileStream fs = new FileStream(@"C:\LBMSystems\test\testGrid_14112015_5-30.pcl", FileMode.Open);

            PlcEnumerator pclEnumerator = new PlcEnumerator(fs);

            RawPrinterHelper.SendBytesToPrinter(printServerAddress, "Pay Stub " + DateTime.Now.ToFileTime(), s =>
            {
                byte[] pageToPrint;

                if (pclEnumerator.GetPage(out pageToPrint))
                {
                    return pageToPrint;
                }
                return null;
            });

            //RawPrinterHelper.SendFileToPrinter(printServerAddress, @"C:\LBMSystems\test\testGrid_9112015_4-10.pcl");

            TrackPrinterStatus();

            return View("Index");

            /*************************************************************************************************************

            Below is printing sample without using Win32 API, this class does not have way to print page by page 

            ************************************************************************************************************/
            var printServer = new PrintServer(printServerAddress);

            var queues =
                printServer.GetPrintQueues()
                    .Select(q => new {q.Name, q.Location, q.Comment, q.FullName, q.IsDirect, q.QueueAttributes});

            var queue = printServer.GetPrintQueues().First();
            //(q => q.Name == printQueue); //printServer.GetPrintQueues().First(q => q.Name == "RemotePrintingTest");

            if (queue.IsInitializing)
            {
                Thread.Sleep(10);
            }

            var anotherPrintJob = queue.AddJob(Guid.NewGuid().ToString());

            Stream destinationStream = anotherPrintJob.JobStream;

            for (int j = 0; j < 10; j++)
            {
                StreamReader sourceStreamReader = new StreamReader(@"C:\LBMSystems\test\testGrid_14112015_5-30.pcl");
                sourceStreamReader.BaseStream.CopyTo(destinationStream);
                sourceStreamReader.Close();
            }

            destinationStream.Close();
            
            queue.Dispose();
            printServer.Dispose();

            TrackPrinterStatus();

            return View("Index");
        }

        private void TrackPrinterStatus()
        {
            var printServerToMonitor = new PrintServer(printServerAddress);

            var i = 0;
            while (i < 1000)
            {
                foreach (var queueToMonitor in printServerToMonitor.GetPrintQueues())
                {
                    queueToMonitor.Refresh();

                    foreach (var jobToMonitor in queueToMonitor.GetPrintJobInfoCollection())
                    {
                        queueToMonitor.Refresh();
                        jobToMonitor.Refresh();

                        string queueStatus = String.Empty;

                        var s =
                            string.Format("Queue {0}:", queueToMonitor.FullName) +
                            "Printing " + jobToMonitor.NumberOfPagesPrinted + " of " + jobToMonitor.NumberOfPages +
                            ", Status:" + jobToMonitor.JobStatus +
                            ", IsPrinting:" + jobToMonitor.IsPrinting +
                            ", IsCompleted:" + jobToMonitor.IsCompleted +
                            ", IsSpooling:" + jobToMonitor.IsSpooling +
                            ", JobIdentifier:" + jobToMonitor.JobIdentifier +
                            ", JobSize:" + jobToMonitor.JobSize +
                            ", TimeSinceStartedPrinting:" + jobToMonitor.TimeSinceStartedPrinting +
                            ", IsPaperOut:" + jobToMonitor.IsPaperOut +
                            ", Queue.IsPrinting:" + queueToMonitor.IsPrinting +
                            ", Queue.NumberOfJobs:" + queueToMonitor.NumberOfJobs +
                            Environment.NewLine;
                        Debug.WriteLine(s);

                        SpotTroubleUsingJobAttributes(jobToMonitor);
                        SpotTroubleUsingQueueAttributes(ref queueStatus, queueToMonitor);
                    }
                }

                i++;
                Thread.Sleep(1500);
            }

            printServerToMonitor.Dispose();
        }

        // Check for possible trouble states of a printer using the flags of the QueueStatus property 
        internal static void SpotTroubleUsingQueueAttributes(ref String statusReport, PrintQueue pq)
        {
            if ((pq.QueueStatus & PrintQueueStatus.PaperProblem) == PrintQueueStatus.PaperProblem)
            {
                statusReport = statusReport + "Has a paper problem. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NoToner) == PrintQueueStatus.NoToner)
            {
                statusReport = statusReport + "Is out of toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.DoorOpen) == PrintQueueStatus.DoorOpen)
            {
                statusReport = statusReport + "Has an open door. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Error) == PrintQueueStatus.Error)
            {
                statusReport = statusReport + "Is in an error state. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.NotAvailable) == PrintQueueStatus.NotAvailable)
            {
                statusReport = statusReport + "Is not available. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Offline) == PrintQueueStatus.Offline)
            {
                statusReport = statusReport + "Is off line. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutOfMemory) == PrintQueueStatus.OutOfMemory)
            {
                statusReport = statusReport + "Is out of memory. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperOut) == PrintQueueStatus.PaperOut)
            {
                statusReport = statusReport + "Is out of paper. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.OutputBinFull) == PrintQueueStatus.OutputBinFull)
            {
                statusReport = statusReport + "Has a full output bin. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.PaperJam) == PrintQueueStatus.PaperJam)
            {
                statusReport = statusReport + "Has a paper jam. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused)
            {
                statusReport = statusReport + "Is paused. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.TonerLow) == PrintQueueStatus.TonerLow)
            {
                statusReport = statusReport + "Is low on toner. ";
            }
            if ((pq.QueueStatus & PrintQueueStatus.UserIntervention) == PrintQueueStatus.UserIntervention)
            {
                statusReport = statusReport + "Needs user intervention. ";
            }

            Debug.WriteLine(statusReport);

            // Check if queue is even available at this time of day 
            // The method below is defined in the complete example.
            //ReportAvailabilityAtThisTime(ref statusReport, pq);
        }

        private static Boolean ReportAvailabilityAtThisTime(PrintQueue pq)
        {
            Boolean available = true;
            if (pq.StartTimeOfDay != pq.UntilTimeOfDay) // If the printer is not available 24 hours a day
            {
                DateTime utcNow = DateTime.UtcNow;
                Int32 utcNowAsMinutesAfterMidnight = (utcNow.TimeOfDay.Hours * 60) + utcNow.TimeOfDay.Minutes;

                // If now is not within the range of available times . . .
                if (!((pq.StartTimeOfDay < utcNowAsMinutesAfterMidnight)
                   &&
                   (utcNowAsMinutesAfterMidnight < pq.UntilTimeOfDay)))
                {
                    available = false;
                }
            }
            return available;
        }//end ReportAvailabilityAtThisTime

        internal static void SpotTroubleUsingJobAttributes(PrintSystemJobInfo theJob)
        {
            if ((theJob.JobStatus & PrintJobStatus.Blocked) == PrintJobStatus.Blocked)
            {
                Debug.WriteLine("The job is blocked.");
            }
            if (((theJob.JobStatus & PrintJobStatus.Completed) == PrintJobStatus.Completed)
                ||
                ((theJob.JobStatus & PrintJobStatus.Printed) == PrintJobStatus.Printed))
            {
                Debug.WriteLine(
                    "The job has finished. Have user recheck all output bins and be sure the correct printer is being checked.");
            }
            if (((theJob.JobStatus & PrintJobStatus.Deleted) == PrintJobStatus.Deleted)
                ||
                ((theJob.JobStatus & PrintJobStatus.Deleting) == PrintJobStatus.Deleting))
            {
                Debug.WriteLine(
                    "The user or someone with administration rights to the queue has deleted the job. It must be resubmitted.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Error) == PrintJobStatus.Error)
            {
                Debug.WriteLine("The job has errored.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Offline) == PrintJobStatus.Offline)
            {
                Debug.WriteLine("The printer is offline. Have user put it online with printer front panel.");
            }
            if ((theJob.JobStatus & PrintJobStatus.PaperOut) == PrintJobStatus.PaperOut)
            {
                Debug.WriteLine("The printer is out of paper of the size required by the job. Have user add paper.");
            }

            if (((theJob.JobStatus & PrintJobStatus.Paused) == PrintJobStatus.Paused)
                ||
                ((theJob.HostingPrintQueue.QueueStatus & PrintQueueStatus.Paused) == PrintQueueStatus.Paused))
            {
                HandlePausedJob(theJob);
                //HandlePausedJob is defined in the complete example.
            }

            if ((theJob.JobStatus & PrintJobStatus.Printing) == PrintJobStatus.Printing)
            {
                Debug.WriteLine("The job is printing now.");
            }
            if ((theJob.JobStatus & PrintJobStatus.Spooling) == PrintJobStatus.Spooling)
            {
                Debug.WriteLine("The job is spooling now.");
            }
            if ((theJob.JobStatus & PrintJobStatus.UserIntervention) == PrintJobStatus.UserIntervention)
            {
                Debug.WriteLine("The printer needs human intervention.");
            }

        } //end SpotTroubleUsingJobAttributes

        internal static void HandlePausedJob(PrintSystemJobInfo theJob)
        {
            // If there's no good reason for the queue to be paused, resume it and 
            // give user choice to resume or cancel the job.
            Debug.WriteLine("The user or someone with administrative rights to the queue" +
                              "\nhas paused the job or queue." +
                              "\nResume the queue? (Has no effect if queue is not paused.)" +
                              "\nEnter \"Y\" to resume, otherwise press return: ");
            String resume = Console.ReadLine();
            if (resume == "Y")
            {
                theJob.HostingPrintQueue.Resume();

                // It is possible the job is also paused. Find out how the user wants to handle that.
                Debug.WriteLine("Does user want to resume print job or cancel it?" +
                                  "\nEnter \"Y\" to resume (any other key cancels the print job): ");
                String userDecision = Console.ReadLine();
                if (userDecision == "Y")
                {
                    theJob.Resume();
                }
                else
                {
                    theJob.Cancel();
                }
            } //end if the queue should be resumed

        } //end HandlePausedJob
    }
}
