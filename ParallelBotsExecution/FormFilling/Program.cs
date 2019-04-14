using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ParallelBotsExecution.FormFilling
{
    class Program
    {
        // Configs
        const int nbChrome = 2; // Number of Chrome instances running in parallel
        const bool chromeHeadless = false; // Set true to activate the headless mode
        static readonly string dataFile = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\FormFilling\Test-cases.xlsx";
        // Configs

        static async Task Main()
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            Console.WriteLine("Script started");
            ExcelManager manager = new ExcelManager(dataFile);

            double tickPerLineDone = 100 / (manager.LastLineToProcess - manager.FirstLineToProcess + 1);
            int nbLineDone = 0;

            var progression = new Progress<ExcelLine>();
            progression.ProgressChanged += (_, line) =>
            {
                nbLineDone++; 
                Console.WriteLine($"{nbLineDone * tickPerLineDone}%| Line {line.LineNumber} processed with status {line.Content.BotStatus}.");
            };

            List<Task> tasks = new List<Task>();
            for (int i = 0; i < nbChrome; i++)
            {
                tasks.Add(RunRobot(manager, chromeHeadless, progression));
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
        static async Task RunRobot(ExcelManager manager, bool headless = false, IProgress<ExcelLine> progress = null)
        {
            await Task.Run(() =>
            {
                // Create a new bot
                Robot robot = new Robot(headless: headless);
                ExcelLine line;
                while ((line = manager.ReadNextLine()) != null)
                {
                    // Fill the form
                    bool result = robot.FillForm(line.Content.FirstName, line.Content.LastName, line.Content.UserName, line.Content.Address, line.Content.Country, line.Content.State, line.Content.Zip, line.Content.NameOnCard, line.Content.CreditCardNumber, line.Content.Expirationdate, line.Content.Cvv);

                    // Set the status regarding the result
                    line.Content.BotStatus = result ? "ok" : "ko";
                    manager.WriteBotStatus(line);
                    progress?.Report(line);
                }
                // Destroy the bot and set the process status to finished
                robot.Quit();
            });
        }
    }
}