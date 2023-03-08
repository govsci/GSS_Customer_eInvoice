using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using Send_Electronic_Invoice.Classes;
using Send_Electronic_Invoice.Objects;

namespace Send_Electronic_Invoice
{
    public class Program
    {

        public Program()
        {
            List<Customer> customers = Database.GetCustomers();
            foreach (Customer customer in customers)
            {
                if (Constants.PublishProfile == 0 || Constants.PublishProfile == 2 || Constants.PublishProfile == 3 || Constants.PublishProfile == 4 || Constants.PublishProfile == 6)
                {
                    Database.GetInvoices(customer);
                    Database.GetCreditMemos(customer);
                    Database.GetDummyInvoices(customer);
                }

                BuildInvoices(customer);

                GetManualInvoices(customer);

                if (customer.CsvString.Length > 0)
                    new Send_Invoice(null, customer);
                else
                    SendInvoices(customer);
            }

            new SendReport(customers);
        }

        private void BuildInvoices(Customer customer)
        {
            if (customer.SellToCustomerNo == "C14533")
            {
                try
                {
                    BuildCsv csv = new NRELcsv();
                    customer.CsvFileName = $"NRELGSSINV{DateTime.Now.ToString("MMddyy")}.csv";
                    customer.CsvString = csv.BuildTheCsv(customer);
                }
                catch(Exception ex)
                {
                    customer.Errors.Add(new CodeError(ex, "Program", "BuildInvoices", null));
                }
            }
            else
            {
                foreach (InvoiceHeader invoice in customer.Invoices)
                {
                    try
                    {
                        BuildXML xml = new BuildXML();
                        invoice.cXMLInvoice = xml.CreateCxmlInvoice(customer, invoice);
                    }
                    catch (Exception ex)
                    {
                        invoice.Errors.Add(new CodeError(ex, "Program", "BuildInvoices(Customer customer)", null));
                    }
                }
            }
        }

        private void GetManualInvoices(Customer customer)
        {
            string folder = $@"{Constants.InvoicePortFolder}{customer.SellToCustomerNo}\";
            if (Directory.Exists(folder))
                Manual_Invoices.Get_Manual_Invoices(customer, folder);
        }

        private void SendInvoices(Customer customer)
        {
            foreach (InvoiceHeader inv in customer.Invoices)
            {
                if (inv.cXMLInvoice == null)
                    inv.NeedsReview = true;
                else if (!inv.NeedsReview)
                    new Send_Invoice(inv, customer);

                if (inv.NeedsReview)
                {
                    string path =  $@"{Constants.InvoicePortFolder}Needs Review\";
                    if (!Directory.Exists(path)) Directory.CreateDirectory(path);
                    Functions.CreateDump(path, $"{Functions.GetDumpname(DateTime.Now)}.{inv.No}.{inv.Your_Reference}.xml", inv.cXMLInvoice.InnerXml, false);
                    Database.UpdateInvoice(inv.No, inv.Your_Reference, 0);
                }
            }
        }

        public static void Main(string[] args)
        {
            try
            {
                if (args.Length > 0)
                {
                    Constants.ApplicationErrors = new List<CodeError>();
                    Constants.PublishProfile = int.Parse(args[0]);
                    Functions.SetupConnections();
                    new Program();
                }
            }
            catch(Exception ex)
            {
                if (Constants.ApplicationErrors == null)
                    Constants.ApplicationErrors = new List<CodeError>();
                Constants.ApplicationErrors.Add(new CodeError(ex, "Program", "Main", null));
            }
            finally
            {
                if (Constants.ApplicationErrors != null && Constants.ApplicationErrors.Count > 0)
                    Email.SendErrorMessage(Constants.ApplicationErrors);
            }
        }
    }
}
