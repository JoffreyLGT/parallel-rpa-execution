using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ParallelBotsExecution
{
    class Program
    {
        static List<string> subjects = new List<string>();

        static void Main(string[] args)
        {
            for (int i = 0; i < 3; i++)
            {
                var nb = i;
                var t = new Task(() => RedditFetching("Instance " + nb.ToString()));
                t.Start();
            }
            Console.ReadKey();
        }

        static void RedditFetching(string name)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            driver.Navigate().GoToUrl(@"https://old.reddit.com/");

            bool endloop = false;
            do
            {
                IReadOnlyCollection<IWebElement> titles = driver.FindElements(By.XPath("//p[@class=\"title\"]"));
                foreach (IWebElement title in titles)
                {
                    subjects.Add(title.Text);
                    Console.WriteLine(name + " " + title.Text);
                    Thread.Sleep(500);
                }
                try
                {
                    driver.FindElement(By.XPath("//span[@class=\"next-button\"]/a")).Click();
                }
                catch (Exception)
                {
                    endloop = true;
                }
            } while (!endloop);

        }
    }
}
