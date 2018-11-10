using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ParallelBotsExecution
{
    class RedditFetching
    {
        static List<string> subjects = new List<string>();
        static Dictionary<string, bool> hasFinished = new Dictionary<string, bool>();

        static void Main(string[] args)
        {

            Application excel = new Application();
            excel.Visible = true;
            Workbook wb = excel.Workbooks.Add();
            Worksheet ws = wb.Worksheets.Add();

            for (int i = 0; i < 3; i++)
            {
                var nb = i;
                var t = new Task(() => RedditFetchingProcess("Instance " + nb.ToString()));
                t.Start();
            }

            int line = 0;
            var excelTask = new Task(() => FillExcel(ws, ref line));
            excelTask.Start();

            Console.ReadKey();
        }

        private static void FillExcel(Worksheet ws, ref int line)
        {
            do
            {
                if (subjects.Count > 0)
                {
                    line++;
                    ws.Cells[line, 1] = subjects.ElementAt(0);
                    subjects.RemoveAt(0);
                }
                else
                {
                    Thread.Sleep(1000);
                }
            } while (hasFinished.Values.Contains(false));
        }

        /// <summary>
        /// Open a new ChromeDriver, go on the old version of reddit and fetch the post titles.
        /// </summary>
        /// <param name="name"></param>
        static void RedditFetchingProcess(string name)
        {
            hasFinished.Add(name, false);
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            driver.Navigate().GoToUrl(@"https://old.reddit.com/");

            IReadOnlyCollection<IWebElement> titles = driver.FindElements(By.XPath("//p[@class=\"title\"]"));
            foreach (IWebElement title in titles)
            {
                subjects.Add(title.Text);
                Console.WriteLine(name + " " + title.Text);
                Thread.Sleep(500);
            }
            hasFinished[name] = true;
        }
    }
}
