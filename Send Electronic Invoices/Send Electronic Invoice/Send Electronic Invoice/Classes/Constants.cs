using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Xml;
using System.Net.Security;
using Tamir.SharpSsh;
using Renci.SshNet;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography.X509Certificates;
using Send_Electronic_Invoice.Objects;

namespace Send_Electronic_Invoice.Classes
{
    public static class Constants
    {
        public static List<CodeError> ApplicationErrors { get; set; }
        public static int PublishProfile = 2;
        /* PublishProfile Options
         * 0 - Capture all invoices from Navision, manual invoices, and send to customers
         * 1 - Capture only manual invoices and send to customers
         * 2 - Capture all invoices from Navision and save to local C drive (does not send to customers), for testing
         * 3 - Capture all invoices from Nav01
         * 4 - Capture all invoices from Navision using EcommerceDB and save to local C drive
         * 5 - Only capture manual invoices and save to local C drive
         * 6 - Capture all invoices from Nav01 and save to local C drive
        */

        public static string connectionGssNav = ConfigurationManager.ConnectionStrings["GssNav01"].ConnectionString;
        public static string connectionEcommerce = ConfigurationManager.ConnectionStrings["TstEcomDb"].ConnectionString;
        public static string DeploymentMode = "test";
        
        public static string InvoiceConfirmedFolder = ConfigurationManager.AppSettings["InvoiceConfirmedFolder"];
        public static string InvoiceCreatedFolder = ConfigurationManager.AppSettings["InvoiceCreatedFolder"];
        public static string InvoiceFailedFolder = ConfigurationManager.AppSettings["InvoiceFailedFolder"];
        public static string InvoiceOutgoingFolder = ConfigurationManager.AppSettings["InvoiceOutgoingFolder"];
        public static string InvoiceSentFolder = ConfigurationManager.AppSettings["InvoiceSentFolder"];
        public static string InvoicePortFolder = ConfigurationManager.AppSettings["InvoicePortFolder"];
        public static string InvoiceReportFolder = ConfigurationManager.AppSettings["InvoiceReportFolder"];
        public static string InvoiceEncryptedFolder = ConfigurationManager.AppSettings["InvoiceEncryptedFolder"];

        public static decimal minThreshold = -0.02M;
        public static decimal maxThreshold = 0.02M;

        public static string[] GLAccounts = ConfigurationManager.AppSettings["GlAccounts"].Split(';');
    }
    public static class GSSContact
    {
        public static string Name { get { return "Government Scientific Source, Inc."; } }
        public static string Street { get { return "12355 Sunrise Valley Dr. Suite 400"; } }
        public static string City { get { return "Reston"; } }
        public static string State { get { return "VA"; } }
        public static string PostalCode { get { return "20191"; } }
        public static string CountryCode { get { return "US"; } }
        public static string Country { get { return "United States"; } }

    }
}
