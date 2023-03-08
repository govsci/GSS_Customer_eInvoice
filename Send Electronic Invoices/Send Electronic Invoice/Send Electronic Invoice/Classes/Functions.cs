using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Send_Electronic_Invoice.Objects;
using System.IO;
using System.Xml;
using System.Data.SqlClient;

namespace Send_Electronic_Invoice.Classes
{
    public static class Functions
    {
        public static string TimeStamp(DateTime value) { return value.ToString("yyyyMMddHHmmssfffffff"); }
        public static string cXMLTimeStamp(DateTime value) { return value.ToString("yyyy-MM-ddTHH:mm:ss"); }
        public static string OrderDate(DateTime value) { return value.ToString("yyyy-MM-dd"); }
        public static string GetDumpname(DateTime value) { return value.ToString("yyyyMMddTHHmmssffff"); }

        public static string RemoveSpecialChar(string input)
        {
            input = input.Replace("®", "&#174;");
            input = input.Replace("©", "&#169;");
            input = input.Replace("™", "&#8482;");
            return input;
        }

        public static void SetupConnections()
        {
            switch (Constants.PublishProfile)
            {
                case 3:
                    Constants.connectionGssNav = System.Configuration.ConfigurationManager.ConnectionStrings["GssNav01"].ConnectionString;
                    Constants.connectionEcommerce = System.Configuration.ConfigurationManager.ConnectionStrings["TstEcomDb"].ConnectionString;
                    Constants.DeploymentMode = "test";
                    break;
                case 4:
                    Constants.connectionGssNav = System.Configuration.ConfigurationManager.ConnectionStrings["GssNav"].ConnectionString;
                    Constants.connectionEcommerce = System.Configuration.ConfigurationManager.ConnectionStrings["TstEcomDb"].ConnectionString;
                    Constants.DeploymentMode = "test";
                    break;
                case 6:
                    Constants.connectionGssNav = System.Configuration.ConfigurationManager.ConnectionStrings["GssNav01"].ConnectionString;
                    Constants.connectionEcommerce = System.Configuration.ConfigurationManager.ConnectionStrings["TstEcomDb"].ConnectionString;
                    Constants.DeploymentMode = "test";
                    break;
                default:
                    Constants.connectionGssNav = System.Configuration.ConfigurationManager.ConnectionStrings["GssNav"].ConnectionString;
                    Constants.connectionEcommerce = System.Configuration.ConfigurationManager.ConnectionStrings["PrdEcomDb"].ConnectionString;
                    Constants.DeploymentMode = "production";
                    break;
            }
        }

        public static string CreateDump(string root, string name, string content, bool folderCreate)
        {
            string year = DateTime.Now.ToString("yyyy");
            string month = DateTime.Now.ToString("MM");
            string day = DateTime.Now.ToString("dd");
            string path = folderCreate ? root + "\\" + year + "\\" + month + "\\" + day : root;

            if (name.Length == 0)
                name = DateTime.Now.ToString("yyyyMMddThhmmssffff");

            try
            {
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);

                using (FileStream fs = File.Create($"{path}\\{name}"))
                {
                    Byte[] contents = new UTF8Encoding(true).GetBytes(content);
                    fs.Write(contents, 0, contents.Length);
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Functions", "CreateDump(string root, string name, string content, bool folderCreate)", new SqlCommand($"{path}\\{name}")));
            }

            return $"{path}\\{name}";
        }

        public static string FindFile(string invoiceNo, string SellTo)
        {
            string path = $@"{Constants.InvoiceSentFolder}{DateTime.Now.ToString("yyyy")}\{DateTime.Now.ToString("MM")}\{DateTime.Now.ToString("dd")}";

            if (!Directory.Exists(path))
                return "No";

            foreach (string file in Directory.GetFiles(path))
            {
                FileInfo finfo = new FileInfo(file);
                if (finfo.Extension.ToUpper().EndsWith("XML"))
                {
                    XmlDocument xml = new XmlDocument();
                    xml.Load(file);
                    string invNumber = xml.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID").InnerXml;
                    if (invoiceNo == invNumber)
                        return "Yes";
                }
            }

            return "No";
        }

        public static InvoiceHeader CalculateInvoiceTotals(Customer customer, InvoiceHeader invoice)
        {
            foreach (InvoiceLine line in invoice.SalesInvLines)
            {
                if (customer.ShipHeader == 1 && line.Type == 1 && Constants.GLAccounts.Contains(line.No))
                {
                    if (invoice.No.StartsWith("CMP") && line.No == "45200")
                    {
                        decimal quantity = line.Quantity < 0.00M ? line.Quantity * -1 : line.Quantity;
                        decimal unitPrice = line.Unit_Price < 0.00M ? line.Unit_Price * -1 : line.Unit_Price;
                        invoice.ShippingAmount += (quantity * unitPrice);
                    }
                    else
                        invoice.ShippingAmount += (line.Quantity * line.Unit_Price);
                }
                else if (line.Type == 2 || (line.Type == 1 && Constants.GLAccounts.Contains(line.No)))
                    invoice.InvoiceLineTotal += line.LineTotal;
            }

            return invoice;
        }

        public static string SetStringValue(XmlNode xml)
        {
            try
            {
                return xml.InnerText;
            }
            catch
            {
                return "";
            }
        }

        public static string SetDateTimeValue(XmlNode xml)
        {
            try
            {
                return xml.InnerText;
            }
            catch
            {
                return DateTime.Now.ToString("yyyy-MM-dd");
            }
        }

        public static int SetIntegerValue(XmlNode xml)
        {
            try
            {
                return int.Parse(xml.InnerText);
            }
            catch
            {
                return 0;
            }
        }

        public static decimal SetDecimalValue(XmlNode xml)
        {
            try
            {
                return decimal.Parse(xml.InnerText);
            }
            catch
            {
                return 0.00M;
            }
        }
    }
}
