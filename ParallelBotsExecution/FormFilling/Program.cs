using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections.Concurrent;

namespace ParallelBotsExecution.FormFilling
{
    class Program
    {
        // Configs
        const int nbChrome = 2; // Number of Chrome instances running in parallel
        const bool chromeHeadless = false; // Set true to activate the headless mode
        static readonly string dataFile = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\FormFilling\Test-cases.xlsx";
        // Configs
        static readonly object myLock = new object();

        static async Task Main()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            Console.WriteLine("Script started");
            ExcelManager manager = new ExcelManager(dataFile);

            List<Task> tasks = new List<Task>();
            for (int i = 0; i < nbChrome; i++)
            {
                tasks.Add(Task.Run(() => RunRobot(manager, chromeHeadless)));
            }
            await Task.WhenAll(tasks);
            Console.WriteLine("Script finished");

            stopwatch.Stop();
            Console.WriteLine($"Execution time: {stopwatch.Elapsed.ToString(@"hh\:mm\:ss")}");
            Console.ReadKey();
        }

        /// <summary>
        /// Run one bot that will do the process.
        /// </summary>
        /// <param name="manager"></param>
        /// <param name="headless"></param>
        static void RunRobot(ExcelManager manager, bool headless = false)
        {
            // Create a new bot
            Robot robot = new Robot(headless: headless);

            foreach (ExcelLine line in manager)
            {
                // Fill the form
                bool result = robot.FillForm(line.Content.FirstName, line.Content.LastName, line.Content.UserName, line.Content.Address, line.Content.Country, line.Content.State, line.Content.Zip, line.Content.NameOnCard, line.Content.CreditCardNumber, line.Content.Expirationdate, line.Content.Cvv);

                // Set the status regarding the result
                line.Content.BotStatus = result ? "ok" : "ko";
                lock (myLock)
                {
                    manager.WriteBotStatus(line);
                }
            }

            // Destroy the bot and set the process status to finished
            robot.Quit();
        }
    }
}