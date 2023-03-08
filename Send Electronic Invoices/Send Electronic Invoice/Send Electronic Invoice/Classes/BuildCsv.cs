using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Send_Electronic_Invoice.Objects;

namespace Send_Electronic_Invoice.Classes
{
    public abstract class BuildCsv
    {
        public abstract string BuildTheCsv(Customer customer);
    }

    public class NRELcsv : BuildCsv
    {
        private List<int> LineNumbersUsed = new List<int>();
        public override string BuildTheCsv(Customer customer)
        {
            if (customer.Invoices.Count > 0)
            {
                StringBuilder csv = new StringBuilder();
                csv.AppendLine($"File Name{customer.TextDelimiter}" +
                    $"Source{customer.TextDelimiter}" +
                    $"PO Number{customer.TextDelimiter}" +
                    $"Supplier Name{customer.TextDelimiter}" +
                    $"Invoice Date{customer.TextDelimiter}" +
                    $"Invoice Number{customer.TextDelimiter}" +
                    $"Invoice Amount{customer.TextDelimiter}" +
                    $"PO Line Number{customer.TextDelimiter}" +
                    $"PO Line Amount{customer.TextDelimiter}" +
                    $"Quantity Shipped{customer.TextDelimiter}" +
                    $"Unit Price");

                foreach (InvoiceHeader invoice in customer.Invoices)
                {
                    foreach (InvoiceLine line in invoice.SalesInvLines)
                    {
                        if (line.Quantity > 0.00M)
                        {
                            string invoiceTotal = (invoice.InvoiceLineTotal + invoice.ShippingAmount).ToString("C").Replace("$", "").Replace(",", "");
                            string orgLineLineNo = line.OrderLine == null ? line.Line_No.ToString() : line.OrderLine.LineNo.ToString();
                            string lineTotal = line.LineTotal.ToString("C").Replace("$", "").Replace(",", "");
                            string quantity = line.Quantity.ToString("G29");
                            string unitPrice = line.Unit_Price.ToString("C").Replace("$", "").Replace(",", "");
                            string invoiceNo = $"G{invoice.No.Replace("-", "")}";
                            if (invoiceNo.EndsWith("R")) invoiceNo = invoiceNo.Remove(invoiceNo.LastIndexOf("R"));

                            if (invoice.Your_Reference == "226528" && line.Line_No == 3)
                                orgLineLineNo = "2";

                            csv.AppendLine($"{customer.CsvFileName}{customer.TextDelimiter}" +
                                $"{customer.FromDomain}{customer.TextDelimiter}" +
                                $"{invoice.Your_Reference}{customer.TextDelimiter}" +
                                $"{customer.FromIdentity}{customer.TextDelimiter}" +
                                $"{invoice.Posting_Date.ToString("MM/dd/yyyy")}{customer.TextDelimiter}" +
                                $"{invoiceNo}{customer.TextDelimiter}" +
                                $"{invoiceTotal}{customer.TextDelimiter}" +
                                $"{orgLineLineNo}{customer.TextDelimiter}" +
                                $"{lineTotal}{customer.TextDelimiter}" +
                                $"{quantity}{customer.TextDelimiter}" +
                                $"{unitPrice}");
                        }
                    }
                    Database.UpdateInvoice(invoice.No, invoice.Your_Reference, 1);
                }

                return csv.ToString().Remove(csv.ToString().LastIndexOf(Environment.NewLine));
            }
            else return "";
        }
    }
}
