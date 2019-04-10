using Microsoft.Office.Interop.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ParallelBotsExecution.FormFilling
{
    class ExcelManager:IEnumerable<ExcelLine>
    {
        private readonly Application app;
        private readonly Workbook wb;
        private readonly Worksheet ws;
        public int LastReturnedLine { get; set; }

        /// <summary>
        /// Default constructor.
        /// If Excel is already opened, connects to the opened worksheet.
        /// Otherwise, open the file at the path in parameter.
        /// </summary>
        /// <param name="file"></param>
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
                app = new Application
                {
                    Visible = true
                };
                wb = app.Workbooks.Open(file);
            }

            ws = wb.ActiveSheet;
            LastReturnedLine = 1; // We have a header in the file so we start at 1, not 0.

        }

        /// <summary>
        /// Read the next line in the Excel file.
        /// </summary>
        /// <returns></returns>
        internal ExcelLine ReadNextLine()
        {
            LastReturnedLine++;
            ExcelLine line = new ExcelLine
            {
                LineNumber = LastReturnedLine,
                Content = ReadLine(LastReturnedLine)
            };
            return line;
        }

        /// <summary>
        /// Read the content of the Excel line.
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        private ExcelLineContent ReadLine(int number)
        {
            ExcelLineContent line = new ExcelLineContent
            {
                Number = GetCellValue(number, (int)ExcelLineContent.Columns.number),
                FirstName = GetCellValue(number, (int)ExcelLineContent.Columns.firstName),
                LastName = GetCellValue(number, (int)ExcelLineContent.Columns.lastName),
                UserName = GetCellValue(number, (int)ExcelLineContent.Columns.userName),
                Address = GetCellValue(number, (int)ExcelLineContent.Columns.address),
                Country = GetCellValue(number, (int)ExcelLineContent.Columns.country),
                State = GetCellValue(number, (int)ExcelLineContent.Columns.state),
                Zip = GetCellValue(number, (int)ExcelLineContent.Columns.zip),
                NameOnCard = GetCellValue(number, (int)ExcelLineContent.Columns.nameOnCard),
                CreditCardNumber = GetCellValue(number, (int)ExcelLineContent.Columns.creditCardNumber),
                Expirationdate = GetCellValue(number, (int)ExcelLineContent.Columns.expirationdate),
                Cvv = GetCellValue(number, (int)ExcelLineContent.Columns.cvv),
                BotStatus = GetCellValue(number, (int)ExcelLineContent.Columns.botStatus)
            };
            if (string.IsNullOrEmpty(line.Number)) return null;

            return line;
        }

        /// <summary>
        /// Write the bot status in its column in Excel.
        /// </summary>
        /// <param name="line"></param>
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

        public IEnumerator<ExcelLine> GetEnumerator()
        {
            ExcelLine line;
            while((line = ReadNextLine()).Content != null)
            {
                yield return line;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
