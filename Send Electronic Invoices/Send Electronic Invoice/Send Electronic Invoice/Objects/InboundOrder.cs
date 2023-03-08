using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Send_Electronic_Invoice.Objects
{
    public class InboundOrder
    {
        public InboundOrder(string _orderId, decimal total, string orderDate, string payloadId, string paymentTerms, Address billto)
        {
            OrderID = _orderId;
            POTotal = total;
            OrderDate = orderDate;
            PayloadID = payloadId;
            PaymentTerms = paymentTerms;
            BillToAddress = billto;
        }

        public string OrderID { get; }
        public decimal POTotal { get; }
        public string OrderDate { get; }
        public string PayloadID { get; }
        public string PaymentTerms { get; }
        public Address BillToAddress { get; }
    }
    public class InboundOrderLine
    {
        public InboundOrderLine(string gss, string partNo, string desc, string uom, int line, decimal qty, decimal price)
        {
            GSSPartNo = gss;
            SupplierPartID = partNo;
            Description = desc;
            UnitOfMeasure = uom;
            LineNo = line;
            Quantity = qty;
            Unit_Price = price > 0.00M ? price : 1;
        }
        public string GSSPartNo { get; }
        public string SupplierPartID { get; }
        public string Description { get; }
        public string UnitOfMeasure { get; }
        public int LineNo { get; }
        public decimal Quantity { get; }
        public decimal Unit_Price { get; }
    }
}
