namespace ParallelBotsExecution.FormFilling
{
    class ExcelLineContent
    {
        public string Number { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string UserName { get; set; }
        public string Address { get; set; }
        public string Country { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string NameOnCard { get; set; }
        public string CreditCardNumber { get; set; }
        public string Expirationdate { get; set; }
        public string Cvv { get; set; }
        public string BotStatus { get; set; }

        /// <summary>
        /// Define the columns order in the Excel file.
        /// </summary>
        internal enum Columns
        {
            number = 1,
            firstName,
            lastName,
            userName,
            address,
            country,
            state,
            zip,
            nameOnCard,
            creditCardNumber,
            expirationdate,
            cvv,
            botStatus
        }
    }
}
