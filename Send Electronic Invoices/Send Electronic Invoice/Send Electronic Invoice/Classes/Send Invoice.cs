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
using WinSCP;
using System.Security.Policy;

namespace Send_Electronic_Invoice.Classes
{
    public class Send_Invoice
    {
        public List<int> PublishProfiles = new List<int>() { 0, 1, 3 };
        public Send_Invoice(InvoiceHeader invoice, Customer customer)
        {
            string fileName = ""
               , fileContents = ""
               , docNo = ""
               , custPO = "";
            bool csv = false;
            try
            {
                if (invoice != null)
                {
                    docNo = invoice.cXMLInvoice.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID").InnerXml;
                    custPO = invoice.cXMLInvoice.SelectSingleNode("//InvoiceDetailOrderInfo/OrderReference/@orderID").InnerXml;

                    fileName = $"{docNo}.{custPO}";
                    fileContents = invoice.cXMLInvoice.InnerXml;
                }
                else
                {
                    fileName = customer.CsvFileName;
                    fileContents = customer.CsvString;
                    csv = true;
                }

                Functions.CreateDump(Constants.InvoiceCreatedFolder, $"{Functions.GetDumpname(DateTime.Now)}.{(docNo.Length > 0 ? fileName + ".xml" : fileName)}", fileContents, true);
                if (PublishProfiles.Contains(Constants.PublishProfile))
                {
                    switch (customer.ConnectionMethod.ToUpper())
                    {
                        case "HTTP":
                            SendHttp(fileContents, customer.URL, fileName, customer);
                            break;
                        case "HTTPS":
                            SendHttp(fileContents, customer.URL, fileName, customer);
                            break;
                        case "FTP":
                            if (customer.EncryptionKey.Length > 0)
                                SendEncryptionFtp(fileContents, customer, fileName);
                            else
                                SendFtp(fileContents, customer, fileName);
                            break;
                        case "SFTP":
                            if (customer.EncryptionKey.Length > 0)
                                SendEncryptionSftp(fileContents, customer, fileName);
                            else
                                SendSftp(fileContents, customer, fileName);
                            break;
                        case "SCP":
                            if (customer.EncryptionKey.Length > 0)
                                SendEncryptedScp(fileContents, customer, fileName, csv);
                            else
                                SendScp(fileContents, customer, fileName, csv);
                            break;
                    }
                    if (docNo.Length > 0)
                    {
                        Database.UpdateInvoice(docNo, custPO, 1);
                        invoice.InvoiceSent = DateTime.Now;
                    }
                }
            }
            catch (Exception ex)
            {
                Functions.CreateDump(Constants.InvoiceFailedFolder, $"{Functions.GetDumpname(DateTime.Now)}.{fileName}.xml", fileContents, true);

                if (invoice != null)
                    invoice.Errors.Add(new CodeError(ex, "Send_Invoice", "Send_Invoice", null));
                else
                    customer.Errors.Add(new CodeError(ex, "Send_Invoice", "Send_Invoice", null));
            }
        }

        public static void SendHttp(string contents, string url, string fileName, Customer customer)
        {
            // Force bypassing SSL handshake
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                    | SecurityProtocolType.Tls11
                    | SecurityProtocolType.Tls12
                    | SecurityProtocolType.Ssl3;
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };

            ServicePointManager.ServerCertificateValidationCallback +=
            delegate (
                object sender,
                System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                System.Security.Cryptography.X509Certificates.X509Chain chain,
                System.Net.Security.SslPolicyErrors sslPolicyErrors)
            {
                return true;
            };

            if (url.Contains("ariba"))
            {
                MSXML2.ServerXMLHTTP xmlHttp = new MSXML2.ServerXMLHTTP();
                xmlHttp.open("POST", url, false, null, null);
                xmlHttp.setRequestHeader("Content-type", "text/xml");
                xmlHttp.send(contents);
                string result = xmlHttp.responseText;
                Functions.CreateDump(Constants.InvoiceConfirmedFolder, $"{Functions.GetDumpname(DateTime.Now)}.{fileName}.xml", result, true);
                Functions.CreateDump(Constants.InvoiceSentFolder, $"{Functions.GetDumpname(DateTime.Now)}.{fileName}.xml", contents, true);
            }
            else
            {
                byte[] buffer = Encoding.UTF8.GetBytes(contents);
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Timeout = 60000;
                req.Method = "POST";
                req.ContentType = "text/xml;charset=utf-8";
                req.ContentLength = buffer.Length;
                Stream PostData = req.GetRequestStream();
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();
                HttpWebResponse WebResp = (HttpWebResponse)req.GetResponse();
                if (WebResp.StatusCode == HttpStatusCode.OK)
                {
                    Stream s = WebResp.GetResponseStream();
                    Encoding enc = Encoding.GetEncoding("utf-8");
                    StreamReader readStream = new StreamReader(s, enc);
                    Functions.CreateDump(Constants.InvoiceConfirmedFolder, $"{Functions.GetDumpname(DateTime.Now)}.{fileName}.xml", readStream.ReadToEnd(), true);
                    Functions.CreateDump(Constants.InvoiceSentFolder, $"{Functions.GetDumpname(DateTime.Now)}.{fileName}.xml", contents, true);
                }
                WebResp.Close();
            }
        }
        public static void SendEncryptionSftp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            PGP pgp = new PGP();
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);

            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            pgp.Encrypt(file, Constants.InvoiceEncryptedFolder, customer.EncryptionKey, fileName, customer.FileExtension);

            System.Threading.Thread.Sleep(300);

            connectSFTP(customer, file = pgp.GetFilePath());
        }
        public static void SendSftp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);
            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            System.Threading.Thread.Sleep(300);

            connectSFTP(customer, file);
        }
        public static void SendEncryptionFtp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            PGP pgp = new PGP();
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);
            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            pgp.Encrypt(file, Constants.InvoiceEncryptedFolder, customer.EncryptionKey, fileName, customer.FileExtension);

            System.Threading.Thread.Sleep(300);

            ListFilesOnServerSsl(customer, new Uri(@"ftp://" + customer.FTP_Username + ":dummy@" + customer.FTP_Server + ":" + customer.FTP_Port + customer.FTP_Folder + fileName + customer.FileExtension), pgp.GetFilePath());
        }
        public static void SendFtp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);
            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            System.Threading.Thread.Sleep(300);

            ListFilesOnServerSsl(customer, new Uri(@"ftp://" + customer.FTP_Username + ":dummy@" + customer.FTP_Server + ":" + customer.FTP_Port + customer.FTP_Folder + fileName + customer.FileExtension), file);
        }

        public static void SendEncryptedScp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            PGP pgp = new PGP();
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);
            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            pgp.Encrypt(file, Constants.InvoiceEncryptedFolder, customer.EncryptionKey, fileName, customer.FileExtension);

            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = customer.FTP_Server,
                UserName = customer.FTP_Username,
                Password = customer.FTP_Password,
                SshHostKeyFingerprint = customer.FTP_HostKey
            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                session.Open(sessionOptions);
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                //session.RemoveFiles($"{customer.FTP_Folder}/*");

                TransferEventArgs args = session.PutFileToDirectory(pgp.GetFilePath(), customer.FTP_Folder, false, transferOptions);

                if (args.Error != null)
                    throw args.Error;
            }
        }

        public static void SendScp(string contents, Customer customer, string sftpFileName, bool csv = false)
        {
            string file = "";
            string fileName = csv ? sftpFileName : Functions.GetDumpname(DateTime.Now);
            file = Functions.CreateDump(Constants.InvoiceSentFolder, fileName, contents, true);

            SessionOptions sessionOptions = new SessionOptions
            {
                Protocol = Protocol.Sftp,
                HostName = customer.FTP_Server,
                UserName = customer.FTP_Username,
                Password = customer.FTP_Password,
                SshHostKeyFingerprint = customer.FTP_HostKey
            };

            using (WinSCP.Session session = new WinSCP.Session())
            {
                session.Open(sessionOptions);
                TransferOptions transferOptions = new TransferOptions();
                transferOptions.TransferMode = TransferMode.Binary;

                TransferEventArgs args = session.PutFileToDirectory(file, customer.FTP_Folder, true, transferOptions);

                if (args.Error != null)
                    throw args.Error;
            }
        }

        public static void connectSFTP(Customer cus, string filePath)
        {
            using (Renci.SshNet.SftpClient client = new SftpClient(cus.FTP_Server, Convert.ToInt16(cus.FTP_Port), cus.FTP_Username, cus.FTP_Password))
            {
                client.Connect();
                string fileName = filePath.Split('\\')[filePath.Split('\\').Length - 1];
                using (var fileStream = System.IO.File.OpenRead(filePath))
                    client.UploadFile(fileStream, fileName, true);
            }

            /*Sftp s = new Sftp(cus.FTP_Server, cus.FTP_Username, cus.FTP_Password);
            s.Connect(Convert.ToInt16(cus.FTP_Port));//.Connect();
            s.Put(filePath);
            s.Close();*/
        }

        public static bool ListFilesOnServerSsl(Customer cus, Uri serverUri, string filePath)
        {
            StreamReader tmp = new StreamReader(filePath);
            string fileContent = tmp.ReadToEnd();
            tmp.Close();

            // The serverUri should start with the ftp:// scheme.
            if (serverUri.Scheme != Uri.UriSchemeFtp)
            {
                return false;
            }
            // Get the object used to communicate with the server.
            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(serverUri);
            request.Method = WebRequestMethods.Ftp.UploadFile; //  ("LIST"; //WebRequestMethods.Ftp...ListDirectory;
            request.Credentials = new NetworkCredential(cus.FTP_Username, cus.FTP_Password);
            request.EnableSsl = true;
            request.KeepAlive = false;



            Stream requestStream = request.GetRequestStream();
            // Copy the contents of the file to the request stream.
            //     StreamReader sourceStream = new StreamReader(localFilePath);
            //     byte[] fileContents = Encoding.UTF8.GetBytes((sourceStream.ReadToEnd());
            byte[] fileContents = Encoding.UTF8.GetBytes(fileContent);


            //   sourceStream.Close();
            request.ContentLength = fileContents.Length;


            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();




            // Get the ServicePoint object used for this request, and limit it to one connection.
            // In a real-world application you might use the default number of connections (2),
            // or select a value that works best for your application.

            //         ServicePoint sp = request.ServicePoint;
            //       Console.WriteLine("ServicePoint connections = {0}.", sp.ConnectionLimit);
            //     sp.ConnectionLimit = 1;

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            Console.WriteLine("The content length is {0}", response.ContentLength);
            // The following streams are used to read the data returned from the server.
            Stream responseStream = null;
            StreamReader readStream = null;
            responseStream = response.GetResponseStream();
            readStream = new StreamReader(responseStream, System.Text.Encoding.UTF8);

            if (readStream != null)
            {
                // Display the data received from the server.
                Console.WriteLine(readStream.ReadToEnd());
            }
            Console.WriteLine("List status: {0}", response.StatusDescription);
            if (readStream != null) readStream.Close();
            if (response != null) response.Close();

            Console.WriteLine("Banner message: {0}",
                response.BannerMessage);

            Console.WriteLine("Welcome message: {0}",
                response.WelcomeMessage);

            Console.WriteLine("Exit message: {0}",
                response.ExitMessage);
            return true;
        }
    }
}
