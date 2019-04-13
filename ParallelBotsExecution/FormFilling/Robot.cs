using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using OpenQA.Selenium.Support.UI;

namespace ParallelBotsExecution.FormFilling
{
    class Robot:IDisposable
    {
        private readonly IWebDriver driver;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="implicitWait">in seconds</param>
        internal Robot(int implicitWait = 3, bool headless = false)
        {
            ChromeOptions options = new ChromeOptions();
            if (headless)
            {
                options.AddArgument("--headless");
                options.AddArgument("--window-size=1920,1080");
            }
            this.driver = new ChromeDriver(options);

            this.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(implicitWait);
        }

        public void Dispose()
        {
            driver.Dispose();
        }

        /// <summary>
        /// Open the Form and fill it with the information.
        /// </summary>
        /// <param name="number"></param>
        /// <param name="firstName"></param>
        /// <param name="lastName"></param>
        /// <param name="address"></param>
        /// <param name="country"></param>
        /// <param name="state"></param>
        /// <param name="zip"></param>
        /// <param name="nameOnCard"></param>
        /// <param name="creditCardNumber"></param>
        /// <param name="expirationDate"></param>
        /// <param name="cvv"></param>
        /// <returns>true if the form is submitted</returns>
        internal bool FillForm(
            string firstName,
            string lastName,
            string userName,
            string address,
            string country,
            string state,
            string zip,
            string nameOnCard,
            string creditCardNumber,
            string expirationDate,
            string cvv)
        {

            driver.Navigate().GoToUrl(@"https://getbootstrap.com/docs/4.1/examples/checkout/");

            // User info
            driver.FindElement(By.Id("firstName")).SendKeys(firstName);
            driver.FindElement(By.Id("lastName")).SendKeys(lastName);
            driver.FindElement(By.Id("username")).SendKeys(userName);
            driver.FindElement(By.Id("address")).SendKeys(address);
            new SelectElement(driver.FindElement(By.Id("country"))).SelectByText(country);
            new SelectElement(driver.FindElement(By.Id("state"))).SelectByText(state);
            driver.FindElement(By.Id("zip")).SendKeys(zip);

            // Credit card info
            driver.FindElement(By.Id("cc-name")).SendKeys(nameOnCard);
            driver.FindElement(By.Id("cc-number")).SendKeys(creditCardNumber);
            driver.FindElement(By.Id("cc-expiration")).SendKeys(expirationDate);
            driver.FindElement(By.Id("cc-cvv")).SendKeys(cvv);


            // Submit the form
            driver.FindElement(By.CssSelector("button.btn.btn-primary.btn-lg.btn-block")).Click();

            // When the form is submitted, all the fields are emptied.
            // So if firstName is empty, it was a success.
            return driver.FindElement(By.Id("firstName")).GetAttribute("value").Length == 0;
        }

        internal void Quit()
        {
            driver.Quit();
            Dispose();
        }
    }
}
