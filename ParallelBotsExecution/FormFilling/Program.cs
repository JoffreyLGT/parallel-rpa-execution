using System.IO;
using System.Reflection;

namespace ParallelBotsExecution.FormFilling
{
    class Program
    {
        static void Main(string[] args)
        {
            string dataFile = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location) + @"\FormFilling\Test-cases.xlsx";
            ExcelManager manager = new ExcelManager(dataFile);
            RunRobot(ref manager);
        }

        static void RunRobot(ref ExcelManager manager)
        {
            Robot robot = new Robot();

            ExcelLine line = null;
            do
            {
                line = manager.ReadNextLine();
                if (line == null) continue;

                bool result = robot.FillForm(line.Content.FirstName, line.Content.LastName, line.Content.UserName, line.Content.Address, line.Content.Country, line.Content.State, line.Content.Zip, line.Content.NameOnCard, line.Content.CreditCardNumber, line.Content.Expirationdate, line.Content.Cvv);
                line.Content.BotStatus = result ? "ok" : "ko";
                manager.WriteBotStatus(line);
            } while (line != null);
        }
    }
}
