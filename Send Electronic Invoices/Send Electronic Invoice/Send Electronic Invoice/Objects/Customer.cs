using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Send_Electronic_Invoice.Objects
{
    public class Customer
    {
        public Customer(string _sellToCustomerNo, string _fromDomain, string _fromIdentity, string _toDomain, string _toIdentity, string _senderDomain
            , string _senderIdentity, string _sharedSecret, string _userAgent, string _url, string _encryptionKey, string _ftpServer, string _ftpPort
            , string _ftpUsername, string _ftpPassword, string _connectionMethod, string _sqlCondition, string _schedule, string _fileExtension
            , string _ftpFolder, int _shipHeader, int _negativeQty, string cxmlDocType, string cmOrgInvRef, string masterAgreeRefr, string textDelimiter
            , string ftpHostKey)
        {
            SellToCustomerNo = _sellToCustomerNo;
            FromDomain = _fromDomain;
            FromIdentity = _fromIdentity;
            ToDomain = _toDomain;
            ToIdentity = _toIdentity;
            SenderDomain = _senderDomain;
            SenderIdentity = _senderIdentity;
            SharedSecret = _sharedSecret;
            UserAgent = _userAgent;
            URL = _url;
            EncryptionKey = _encryptionKey;
            FTP_Server = _ftpServer;
            FTP_Port = _ftpPort;
            FTP_Username = _ftpUsername;
            FTP_Password = _ftpPassword;
            ConnectionMethod = _connectionMethod;
            SQL_Condition = _sqlCondition;
            Schedule = _schedule;
            FileExtension = _fileExtension;
            FTP_Folder = _ftpFolder;
            ShipHeader = _shipHeader;
            NegativeQty = _negativeQty;
            CxmlDocumentType = cxmlDocType;
            Credit_Memo_Org_Inv_Reference = cmOrgInvRef;
            MasterAgreementReference = masterAgreeRefr;
            TextDelimiter = textDelimiter;
            CsvString = CsvFileName = "";
            FTP_HostKey = ftpHostKey;

            Invoices = new List<InvoiceHeader>();
            Errors = new List<CodeError>();
        }
        public string SellToCustomerNo { get; }
        public string FromDomain { get; }
        public string FromIdentity { get; }
        public string ToDomain { get; }
        public string ToIdentity { get; }
        public string SenderDomain { get; }
        public string SenderIdentity { get; }
        public string SharedSecret { get; }
        public string UserAgent { get; }
        public string URL { get; }
        public string EncryptionKey { get; }
        public string FTP_Server { get; }
        public string FTP_Port { get; }
        public string FTP_Username { get; }
        public string FTP_Password { get; }
        public string ConnectionMethod { get; }
        public string SQL_Condition { get; }
        public string Schedule { get; }
        public string FileExtension { get; }
        public string FTP_Folder { get; }
        public int ShipHeader { get; }
        public int NegativeQty { get; }
        public string CxmlDocumentType { get; }
        public string Credit_Memo_Org_Inv_Reference { get; }
        public List<InvoiceHeader> Invoices { get; }
        public string MasterAgreementReference { get; }
        public string TextDelimiter { get; }
        public string CsvString { get; set; }
        public string CsvFileName { get; set; }
        public string FTP_HostKey { get; set; }
        public List<CodeError> Errors { get; }
    }
}
