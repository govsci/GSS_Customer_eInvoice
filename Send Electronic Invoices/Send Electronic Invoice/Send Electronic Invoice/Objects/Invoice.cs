using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Send_Electronic_Invoice.Objects
{
    public class InvoiceHeader
    {
        public InvoiceHeader(string _no, Address billTo, Address shipTo, string _yourReference, DateTime _invoiceDate, string file)
        {
            No = _no;
            Bill_To = billTo;
            Ship_To = shipTo;
            Your_Reference = _yourReference.Replace("\\","_").Replace("/","_");
            Posting_Date = _invoiceDate;

            SalesInvLines = new List<InvoiceLine>();
            cXMLInvoice = new XmlDocument();
            Errors = new List<CodeError>();
            FileName = file;

            InvoiceLineTotal = 0.0M;
            ShippingAmount = 0.0M;
            Subtotal = 0.0M;
        }
        public InvoiceHeader(string _no, Address billTo, Address shipTo, string _yourReference, DateTime _invoiceDate, string orgInvNo, DateTime orgInvDate, string file)
        {
            No = _no;
            Bill_To = billTo;
            Ship_To = shipTo;
            Your_Reference = _yourReference.Replace("\\", "_").Replace("/", "_");
            Posting_Date = _invoiceDate;
            OriginalInvoiceNo = orgInvNo;
            OriginalInvoiceDate = orgInvDate;

            SalesInvLines = new List<InvoiceLine>();
            cXMLInvoice = new XmlDocument();
            Errors = new List<CodeError>();
            FileName = file;

            InvoiceLineTotal = 0.0M;
            ShippingAmount = 0.0M;
            Subtotal = 0.0M;
        }

        public string No { get; }
        public Address Bill_To { get; }
        public Address Ship_To { get; }
        public string Your_Reference { get; }
        public List<InvoiceLine> SalesInvLines { get; }
        public XmlDocument cXMLInvoice { get; set; }
        public InboundOrder OriginalOrder { get; set; }
        public DateTime InvoiceSent { get; set; }
        public DateTime Posting_Date { get; }
        public bool NeedsReview { get; set; }
        public string OriginalInvoiceNo { get; }
        public DateTime OriginalInvoiceDate { get; }
        public decimal InvoiceLineTotal { get; set; }
        public decimal ShippingAmount { get; set; }
        public decimal Subtotal { get; set; }
        public List<CodeError> Errors { get; }
        public string FileName { get; }
    }

    public class InvoiceLine
    {
        public InvoiceLine(Customer customer, int _lineNo, string _no, string _vendorItemNo, string _unitOfMeasure, string _description, decimal _quantity, decimal _unitPrice, int _type, bool creditMemo)
        {
            Line_No = _lineNo;
            No = _no;
            Vendor_Item_No = _vendorItemNo;
            Unit_of_Measure_Code = _unitOfMeasure;
            Description = _description;
            Quantity = customer.NegativeQty == 1 && creditMemo ? _quantity * -1 : _quantity;
            Unit_Price = customer.NegativeQty != 1 && creditMemo ? _unitPrice * -1 : _unitPrice;
            Type = _type;
            Invoice_Line_No = 0;
            Refr_Line_No = 0;
            Supplier_Part_ID = "";
            ValidationError = "";

            LineTotal = (Unit_Price * Quantity);
            if (customer.SellToCustomerNo == "BERKELEY")
                LineTotal = Math.Round(LineTotal, 2, MidpointRounding.AwayFromZero);
            else
                LineTotal = Math.Round(LineTotal, 2);
        }

        public int Line_No { get; }
        public int Invoice_Line_No { get; set; }
        public int Refr_Line_No { get; set; }
        public string No { get; }
        public string Supplier_Part_ID { get; set; }
        public decimal Quantity { get; }
        public int Type { get; }
        public decimal Unit_Price { get; }
        public string Vendor_Item_No { get; }
        public string Unit_of_Measure_Code { get; }
        public string Description { get; }
        public decimal LineTotal { get; }
        public InboundOrderLine OrderLine { get; set; }
        public string ValidationError { get; set; }
    }

    public class Address
    {
        public Address(string shipToCode, string name, string street1, string street2, string city, string state,string postCode, string country = "United States", string countryCode = "US")
        {
            ShipToCode = shipToCode;
            Name = name;
            Street1 = street1;
            Street2 = street2;
            City = city;
            State = state;
            PostCode = postCode;
            Country = country;
            CountryCode = countryCode;
        }

        public string ShipToCode { get; }
        public string Name { get; }
        public string Street1 { get; }
        public string Street2 { get; }
        public string City { get; }
        public string State { get; }
        public string PostCode { get; }
        public string Country { get; }
        public string CountryCode { get; }
    }
}
