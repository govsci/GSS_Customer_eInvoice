using Org.BouncyCastle.Asn1.IsisMtt.X509;
using Send_Electronic_Invoice.Objects;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Send_Electronic_Invoice.Classes
{
    public static class Manual_Invoices
    {
        public static void Get_Manual_Invoices(Customer customer, string folder)
        {
            GetFiles(customer, folder);
        }

        private static void GetFiles(Customer customer, string folder)
        {
            string[] files = Directory.GetFiles(folder);
            foreach (string file in files)
            {
                try
                {
                    if (file.EndsWith(".xml"))
                    {
                        XmlDocument cXML = new XmlDocument();
                        cXML.Load(file);
                        if (cXML != null)
                        {
                            FileInfo info = new FileInfo(file);

                            ReadXmlFile(customer, cXML, info.Name);
                            File.Delete(file);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Constants.ApplicationErrors.Add(new CodeError(ex, "Manual_Invoices", "GetFiles(Customer customer)", new SqlCommand(file)));
                }
            }
        }

        private static void ReadXmlFile(Customer customer, XmlDocument cXML, string fileName)
        {
            //Functions.CreateDump($@"{Constants.InvoiceFolder}ManualCopy\", $"{Functions.GetDumpname(DateTime.Now)}.xml", cXML.InnerXml, true);
            string poNumber = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderReference/@orderID"));
            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(poNumber, customer.SQL_Condition.Replace(".", @"\.").Replace("%", ".+?"));
            if (match.Success)
            {
                string[] street = new string[1]; street[0] = "";
                if (Functions.SetStringValue(cXML.SelectSingleNode("//InvoicePartner/Contact[@role='billTo']/PostalAddress/Street")).Length > 0)
                {
                    street = new string[cXML.SelectNodes("//InvoicePartner/Contact[@role='billTo']/PostalAddress/Street").Count];
                    for (int i = 0; i < street.Length; i++)
                        street[i] = cXML.SelectNodes("//InvoicePartner/Contact[@role='billTo']/PostalAddress/Street")[i].InnerXml;
                }

                string shipToCode = "", shipToName = "", shipToCity = "", shipToState = "", shipToPostCode = "", shipToCountry = "", shipToCountryCode = "";
                string[] shipToStreet = new string[1]; street[0] = "";
                if (cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']") != null)
                {
                    if (cXML.SelectNodes("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/Street") != null && cXML.SelectNodes("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/Street").Count > 0)
                    {
                        shipToStreet = new string[cXML.SelectNodes("//InvoiceDetailShipping/Contact[@role='shipTo']").Count];
                        for (int i = 0; i < shipToStreet.Length; i++)
                            shipToStreet[i] = cXML.SelectNodes("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/Street")[i].InnerXml;
                    }

                    shipToCode = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/@addressID"));
                    shipToName = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/Name"));
                    shipToCity = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/City"));
                    shipToState = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/State"));
                    shipToPostCode = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/PostalCode"));
                    shipToCountry = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/Country"));
                    shipToCountryCode = Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailShipping/Contact[@role='shipTo']/PostalAddress/Country/@isoCountryCode"));
                }

                InvoiceHeader invoice = new InvoiceHeader(
                    Functions.SetStringValue(cXML.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID"))
                    , new Address(""
                        , Functions.SetStringValue(cXML.SelectSingleNode("//InvoicePartner/Contact[@role='billTo']/Name"))
                        , street.Length > 0 ? street[0] : ""
                        , street.Length > 1 ? street[1] : ""
                        , Functions.SetStringValue(cXML.SelectSingleNode("//InvoicePartner/Contact[@role='billTo']/PostalAddress/City"))
                        , Functions.SetStringValue(cXML.SelectSingleNode("//InvoicePartner/Contact[@role='billTo']/PostalAddress/State"))
                        , Functions.SetStringValue(cXML.SelectSingleNode("//InvoicePartner/Contact[@role='billTo']/PostalAddress/PostalCode")))
                    , new Address(shipToCode
                        , shipToName
                        , shipToStreet.Length > 0 ? shipToStreet[0] : ""
                        , shipToStreet.Length > 1 ? shipToStreet[1] : ""
                        , shipToCity
                        , shipToState
                        , shipToPostCode)
                    , poNumber
                    , DateTime.Parse(Functions.SetDateTimeValue(cXML.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceDate")))
                    , fileName);

                ReadXmlLines(customer, cXML, ref invoice);

                invoice.cXMLInvoice = cXML;
                customer.Invoices.Add(invoice);
            }
        }

        private static void ReadXmlLines(Customer customer, XmlDocument cXML, ref InvoiceHeader invoice)
        {
            foreach (XmlNode node in cXML.SelectNodes("//InvoiceDetailItem"))
            {
                try
                {
                    InvoiceLine line = new InvoiceLine(
                        customer
                        , Functions.SetIntegerValue(node.SelectSingleNode("@invoiceLineNumber"))
                        , Functions.SetStringValue(node.SelectSingleNode("InvoiceDetailItemReference/ItemID/SupplierPartID"))
                        , Functions.SetStringValue(node.SelectSingleNode("InvoiceDetailItemReference/ItemID/SupplierPartID"))
                        , Functions.SetStringValue(node.SelectSingleNode("UnitOfMeasure"))
                        , Functions.SetStringValue(node.SelectSingleNode("InvoiceDetailItemReference/Description"))
                        , Functions.SetDecimalValue(node.SelectSingleNode("@quantity"))
                        , Functions.SetDecimalValue(node.SelectSingleNode("UnitPrice/Money"))
                        , 2
                        , invoice.No.StartsWith("CMP")
                        );
                    line.Refr_Line_No = Functions.SetIntegerValue(node.SelectSingleNode("InvoiceDetailItemReference/@lineNumber"));

                    InboundOrder order = null;
                    line.OrderLine = Database.GetInboundOrderLine(ref order, invoice.Your_Reference, line.Line_No, line.No, line.Supplier_Part_ID, line.Unit_of_Measure_Code);
                    invoice.OriginalOrder = order;
                    invoice.SalesInvLines.Add(line);
                }
                catch (Exception ex)
                {
                    invoice.Errors.Add(new CodeError(ex, "Manual_Invoices", "ReadXmlLines", null));
                }
            }
        }
    }
}
