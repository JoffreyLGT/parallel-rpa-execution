using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ParallelBotsExecution.FormFilling
{
    class ExcelManager
    {
        private Application app;
        private Workbook wb;
        private Worksheet ws;
        private int lastReturnedLine;

        internal ExcelManager(string file)
        {
            // Use the existing instance if there is one
            try
            {
                app = (Application)Marshal.GetActiveObject("Excel.Application");
                wb = app.ActiveWorkbook;
            }
            catch (COMException)
            {
                app = new Application();
                app.Visible = true;
                wb = app.Workbooks.Open(file);
            }

            ws = wb.ActiveSheet;
            lastReturnedLine = 1; // We have a header in the file so we start at 1, not 0.

        }

        internal ExcelLine ReadNextLine()
        {
            lastReturnedLine++;
            ExcelLine line = new ExcelLine();
            line.LineNumber = lastReturnedLine;
            line.Content = ReadLine(lastReturnedLine);
            return line;
        }

        private ExcelLineContent ReadLine(int number)
        {
            ExcelLineContent line = new ExcelLineContent();
            line.Number = GetCellValue(number, (int)ExcelLineContent.Columns.number);
            if (string.IsNullOrEmpty(line.Number)) return null;
            line.FirstName = GetCellValue(number, (int)ExcelLineContent.Columns.firstName);
            line.LastName = GetCellValue(number, (int)ExcelLineContent.Columns.lastName);
            line.UserName = GetCellValue(number, (int)ExcelLineContent.Columns.userName);
            line.Address = GetCellValue(number, (int)ExcelLineContent.Columns.address);
            line.Country = GetCellValue(number, (int)ExcelLineContent.Columns.country);
            line.State = GetCellValue(number, (int)ExcelLineContent.Columns.state);
            line.Zip = GetCellValue(number, (int)ExcelLineContent.Columns.zip);
            line.NameOnCard = GetCellValue(number, (int)ExcelLineContent.Columns.nameOnCard);
            line.CreditCardNumber = GetCellValue(number, (int)ExcelLineContent.Columns.creditCardNumber);
            line.Expirationdate = GetCellValue(number, (int)ExcelLineContent.Columns.expirationdate);
            line.Cvv = GetCellValue(number, (int)ExcelLineContent.Columns.cvv);
            line.BotStatus = GetCellValue(number, (int)ExcelLineContent.Columns.botStatus);
            return line;
        }

        internal void WriteBotStatus(ExcelLine line)
        {
            ws.Cells[line.LineNumber, (int)ExcelLineContent.Columns.botStatus] = line.Content.BotStatus;
        }

        private string GetCellValue(int row, int col)
        {
            string content = string.Empty;
            try
            {
                content = ((Range)ws.Cells[row, col]).Value.ToString();
            }
            catch (Exception)
            {
                // Do nothing. For example, an exception is thrown when the cell is empty.
            }
            return content;
        }
    }
}
