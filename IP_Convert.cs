using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;

namespace IP_Convert
{
    /// <summary>
    /// Takes in R365 Payment Distribution File in CSV format, splits up multiple invoices items
    /// on a single line item, remove extraneous data and formatting, produces Truist compatible CSV file
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public class CsvRecord
        {
            public string GLAccountName { get; set; }
            public string BankRoutingNumber { get; set; }
            public string BankAccountNumber { get; set; }
            public DateOnly FileDate { get; set; }
            public string Amount { get; set; }
            public string InvoiceNumbers { get; set; }
            public string InvoiceDates { get; set; }
            public string InvoiceAmounts { get; set; }
            public string InvoiceTotals { get; set; }
            public string InvoiceDescriptions { get; set; }
            public string CheckNumber { get; set; }
            public string VendorID { get; set; }
            public string VendorName { get; set; }
            public string VendorAddressLine1 { get; set; }
            public string VendorAddressLine2 { get; set; }
            public string VendorCity { get; set; }
            public string VendorState { get; set; }
            public string VendorZip { get; set; }
            public string VendorEmail { get; set; }
        }

        public class OriginalRecord
        {
            public string Payee_Country { get; set; }
            public string PayerAddress1 { get; set; }
            public string PayerAddress2 { get; set; }
            public string PayerCity { get; set; }
            public string PayerState { get; set; }
            public string PayerZip { get; set; }
            public string PayerCountry { get; set; }
        }

        public class UpdatedRecord
        {
            public string GLAccountName { get; set; }
            public string BankRoutingNumber { get; set; }
            public string BankAccountNumber { get; set; }
            public DateOnly FileDate { get; set; }
            public string Amount { get; set; }
            public string InvoiceNumbers { get; set; }
            public string InvoiceDates { get; set; }
            public string InvoiceAmounts { get; set; }
            public string InvoiceTotals { get; set; }
            public string InvoiceDescriptions { get; set; }
            public string CheckNumber { get; set; }
            public string VendorID { get; set; }
            public string VendorName { get; set; }
            public string VendorAddressLine1 { get; set; }
            public string VendorAddressLine2 { get; set; }
            public string VendorCity { get; set; }
            public string VendorState { get; set; }
            public string VendorZip { get; set; }
            public string Payee_Country { get; set; }
            public string VendorEmail { get; set; }
            public string PayerAddress1 { get; set; }
            public string PayerAddress2 { get; set; }
            public string PayerCity { get; set; }
            public string PayerState { get; set; }
            public string PayerZip { get; set; }
            public string PayerCountry { get; set; }
        }


        private void SelectFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                FilePathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void RunButton_Click(object sender, RoutedEventArgs e)
        {

            if (!string.IsNullOrEmpty(FilePathTextBox.Text))
            {
                try
                {

                    List<OriginalRecord> list;

                    var config = new CsvConfiguration(CultureInfo.InvariantCulture) { HeaderValidated = null, MissingFieldFound = null };

                    using (var reader = new StreamReader("data.csv"))
                    using (var csv = new CsvReader(reader, config))
                    {
                        list = csv.GetRecords<OriginalRecord>().ToList();
                    }

                    var newList = list.Select(x => new UpdatedRecord()
                    {
                        Payee_Country = x.Payee_Country,
                        PayerAddress1 = x.PayerAddress1,
                        PayerAddress2 = x.PayerAddress2,
                        PayerCity = x.PayerCity,
                        PayerState = x.PayerState,
                        PayerZip = x.PayerZip,
                        PayerCountry = x.PayerCountry
                    }).ToList();

                    // Read the CSV file
                    List<CsvRecord> records;

                    using (var reader = new StreamReader(FilePathTextBox.Text))
                    using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                    {
                        records = csv.GetRecords<CsvRecord>().ToList();
                    }

                    // Create a new list to store the modified records
                    List<CsvRecord> modifiedRecords = new List<CsvRecord>();

                    // Split records containing pipes

                    foreach (var record in records)
                    {

                        // Split the InvoiceAmounts and InvoiceTotals based on the pipe character
                        var invoiceAmountsArray = record.InvoiceAmounts.Split('|');
                        var invoiceTotalsArray = record.InvoiceTotals.Split('|');
                        var invoiceNumbersArray = record.InvoiceNumbers.Split('|');
                        var invoiceDatesArray = record.InvoiceDates.Split('|');
                        var invoiceDescriptionsArray = record.InvoiceDescriptions.Split('|');

                        // Ensure the arrays have the same length
                        int itemCount = Math.Min(invoiceAmountsArray.Length, invoiceTotalsArray.Length);

                        // Create separate line items for each amount
                        for (int i = 0; i < itemCount; i++)
                        {
                            var newRecord = new UpdatedRecord
                            {
                                GLAccountName = record.GLAccountName.Trim(' ', '"'),
                                PayerAddress1 = "xxx",
                                PayerAddress2 = "xxx",
                                PayerCity = "xxx",
                                PayerState = "xxx",
                                PayerZip = "xxx",
                                PayerCountry = "US",
                                BankRoutingNumber = record.BankRoutingNumber.Trim(' ', '"'),
                                BankAccountNumber = record.BankAccountNumber.Trim(' ', '"'),
                                FileDate = record.FileDate,
                                Amount = record.Amount,
                                InvoiceNumbers = invoiceNumbersArray[i].Trim(' ', '"'),
                                InvoiceDates = invoiceDatesArray[i].Trim(' ', '"'),
                                InvoiceAmounts = invoiceAmountsArray[i].Trim(' ', '"'),
                                InvoiceTotals = invoiceTotalsArray[i].Trim(' ', '"'),
                                InvoiceDescriptions = invoiceDescriptionsArray[i].Trim(' ', '"'),
                                CheckNumber = record.CheckNumber.Trim(' ', '"'),
                                VendorID = record.VendorID.Trim(' ', '"'),
                                VendorName = record.VendorName.Trim(' ', '"'),
                                VendorAddressLine1 = record.VendorAddressLine1.Trim(' ', '"'),
                                VendorAddressLine2 = record.VendorAddressLine2.Trim(' ', '"'),
                                VendorCity = record.VendorCity.Trim(' ', '"'),
                                VendorState = record.VendorState.Trim(' ', '"'),
                                VendorZip = record.VendorZip.Trim(' ', '"'),
                                Payee_Country = "US",
                                VendorEmail = record.VendorEmail.Trim(' ', '"')
                            };

                            // Add the new record to the list
                            modifiedRecords.Add(newRecord);
                        }
                    }

                    // Perform your edits on the 'records' list here

                    foreach (var record in modifiedRecords)
                    {
                        if (!string.IsNullOrWhiteSpace(record.CheckNumber) && !record.CheckNumber.All(char.IsDigit))
                        {
                            record.CheckNumber = "";
                        }

                        record.BankRoutingNumber = record.BankRoutingNumber.PadLeft(9, '0');
                        record.BankAccountNumber = record.BankAccountNumber.PadLeft(13, '0');
                        record.CheckNumber = record.CheckNumber.PadLeft(10, '0');

                        if (!string.IsNullOrEmpty(record.VendorID) && record.VendorID.Length > 17)
                        {

                            record.VendorID = record.VendorID.Substring(0, 17);
                        }

                        decimal tempValue = decimal.Parse(record.Amount);
                        record.Amount = tempValue.ToString("F2");

                        tempValue = decimal.Parse(record.InvoiceTotals);
                        record.InvoiceTotals = tempValue.ToString("F2");

                        tempValue = decimal.Parse(record.InvoiceAmounts);
                        record.InvoiceAmounts = tempValue.ToString("F2");

                        if (!string.IsNullOrEmpty(record.InvoiceNumbers) && record.InvoiceNumbers.Length > 30)
                        {

                            record.InvoiceNumbers = record.InvoiceNumbers.Substring(0, 30);
                        }

                        record.FileDate = DateOnly.ParseExact(record.FileDate.ToString("MM/dd/yyyy"), "MM/dd/yyyy", CultureInfo.InvariantCulture);

                        DateTime tempDate = DateTime.Parse(record.InvoiceDates);
                        record.InvoiceDates = tempDate.ToString("MM/dd/yyyy");
                        record.InvoiceDescriptions = record.InvoiceDescriptions.Trim(' ', '"');

                        if (!string.IsNullOrEmpty(record.InvoiceDescriptions) && record.InvoiceDescriptions.Length > 50)
                        {

                            record.InvoiceDescriptions = record.InvoiceDescriptions.Substring(0, 50);
                        }

                        if (!string.IsNullOrEmpty(record.GLAccountName) && record.GLAccountName.Length > 35)
                        {

                            record.GLAccountName = record.GLAccountName.Substring(0, 35);
                        }
                    }

                    // Save the modified CSV file to the desktop
                    var desktopPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
                    var outputPath = System.IO.Path.Combine(desktopPath, "Fixed_IP.csv");

                    using (var writer = new StreamWriter(outputPath))
                    using (var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture)
                    { HasHeaderRecord = false }))
                    {
                        csv.WriteRecords(modifiedRecords);
                    }

                    MessageBox.Show($"File converted and saved to: {outputPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}