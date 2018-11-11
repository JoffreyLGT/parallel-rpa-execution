using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace ParallelBotsExecution.FormFilling
{
    class Program
    {
        // Configs
        const int nbChrome = 2; // Number of Chrome instances running in parallel
        const bool chromeHeadless = false; // Set true to activate the headless mode
        // Configs

        static Random rnd = new Random(); // Used to generate random ids

        static void Main(string[] args)
        {
            DateTime start = DateTime.Now;
            Console.WriteLine(start.ToString() + " Script started");

            List<ExcelLine> linesToWrite = new List<ExcelLine>();
            Dictionary<string, bool> processesStatus = new Dictionary<string, bool>();

            string dataFile = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\FormFilling\Test-cases.xlsx";
            ExcelManager manager = new ExcelManager(dataFile);

            for (int i = 0; i < nbChrome; i++)
            {
                Task.Factory.StartNew(() => RunRobot(ref manager, ref linesToWrite, ref processesStatus, chromeHeadless));
            }

            var excelTask = new Task(() => FillExcel(ref manager, ref linesToWrite, ref processesStatus));
            excelTask.Start();
            // We have to indicate to the program that we want the ExcelTask to finish before continuing.
            // Otherwise, the bot could finish it's execution without completing any task.
            excelTask.Wait(); 

            DateTime end = DateTime.Now;
            Console.WriteLine(end.ToString() + " Script finished");

            Console.WriteLine("Total time:" + end.Subtract(start).ToString());
            Console.ReadKey();
        }

        /// <summary>
        /// Run one bot that will do the process.
        /// </summary>
        /// <param name="manager"></param>
        /// <param name="linesToWrite"></param>
        /// <param name="processesStatus"></param>
        /// <param name="headless"></param>
        static void RunRobot(ref ExcelManager manager, ref List<ExcelLine> linesToWrite, ref Dictionary<string, bool> processesStatus, bool headless = false)
        {
            // Generate a new ID and add it in the processes status so we know when it's finished
            string id = rnd.Next().ToString();
            processesStatus.Add(id, false);

            // Create a new bot
            Robot robot = new Robot(headless: headless);

            // Loop until we find an empty line
            ExcelLine line = null;
            do
            {
                // Read the next line from the manager
                line = manager.ReadNextLine();
                if (line.Content == null) continue;

                // Fill the form
                bool result = robot.FillForm(line.Content.FirstName, line.Content.LastName, line.Content.UserName, line.Content.Address, line.Content.Country, line.Content.State, line.Content.Zip, line.Content.NameOnCard, line.Content.CreditCardNumber, line.Content.Expirationdate, line.Content.Cvv);
                
                // Set the status regarding the result
                line.Content.BotStatus = result ? "ok" : "ko";
                // Add the line in the list of lines to write
                linesToWrite.Add(line);
            } while (line.Content != null);

            // Destroy the bot and set the process status to finished
            robot.Quit();
            processesStatus[id] = true;
        }

        /// <summary>
        /// Process that extract the lines from the lineToWrite list and write them in Excel.
        /// It is the only process writing in the Excel file to avoid the concurrency.
        /// It stops when no more bots are running and when there is no more line to write.
        /// </summary>
        /// <param name="manager"></param>
        /// <param name="linesToWrite"></param>
        /// <param name="processesStatus"></param>
        private static void FillExcel(ref ExcelManager manager, ref List<ExcelLine> linesToWrite, ref Dictionary<string, bool> processesStatus)
        {
            do
            {
                if (linesToWrite.Count > 0)
                {
                    manager.WriteBotStatus(linesToWrite.ElementAt(0));
                    linesToWrite.RemoveAt(0);
                }
                else
                {
                    Thread.Sleep(1000);
                }
            } while (processesStatus.Count == 0 ||
                     processesStatus.Values.Contains(false) ||
                     linesToWrite.Count > 0);
        }
    }
}
