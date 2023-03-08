using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Send_Electronic_Invoice.Objects;

namespace Send_Electronic_Invoice.Classes
{
    public class BuildXML
    {
        private Customer customer;
        private InvoiceHeader salesInvoice;
        private int GlLineNumber = 0, TotalNumLines = 0;

        private string Purpose = ""
            , Payload_ID = "";

        public XmlDocument CreateCxmlInvoice(Customer _customer, InvoiceHeader _salesInvoice)
        {
            XmlDocument cXML = new XmlDocument();
            customer = _customer;
            salesInvoice = _salesInvoice;

            Purpose = salesInvoice.No.StartsWith("CMP") ? "creditMemo" : "standard";
            Payload_ID = Functions.TimeStamp(DateTime.Now) + "." + salesInvoice.No + "@govsci.com";

            XmlDeclaration xml = cXML.CreateXmlDeclaration("1.0", "UTF-8", null);
            cXML.AppendChild(xml);

            if (customer.CxmlDocumentType.Length > 0)
            {
                XmlDocumentType docType = cXML.CreateDocumentType("cXML", null, customer.CxmlDocumentType, null);
                cXML.AppendChild(docType);
            }

            XmlElement cxmlele = cXML.CreateElement("cXML");
            cxmlele.SetAttribute("version", "1.0");
            cxmlele.SetAttribute("payloadID", Payload_ID);
            cxmlele.SetAttribute("xml:lang", "en-US");
            cxmlele.SetAttribute("timestamp", Functions.cXMLTimeStamp(DateTime.Now));
            cXML.AppendChild(cxmlele);

            cXML.SelectSingleNode("cXML").AppendChild(CreateHeader(cXML));
            cXML.SelectSingleNode("cXML").AppendChild(CreateRequest(cXML));

            if (TotalNumLines == 0)
                salesInvoice.NeedsReview = true;

            return cXML;
        }

        private XmlElement CreateHeader(XmlDocument cXML)
        {
            XmlElement header = cXML.CreateElement("Header");

            //From
            XmlElement from = cXML.CreateElement("From");
            XmlElement fcred = cXML.CreateElement("Credential");
            fcred.SetAttribute("domain", customer.FromDomain);
            XmlElement fid = cXML.CreateElement("Identity");
            fid.AppendChild(cXML.CreateTextNode(customer.FromIdentity));
            fcred.AppendChild(fid);
            from.AppendChild(fcred);
            header.AppendChild(from);

            //To
            XmlElement to = cXML.CreateElement("To");
            XmlElement tcred = cXML.CreateElement("Credential");
            tcred.SetAttribute("domain", customer.ToDomain);
            XmlElement tid = cXML.CreateElement("Identity");
            tid.AppendChild(cXML.CreateTextNode(customer.ToIdentity));
            tcred.AppendChild(tid);
            to.AppendChild(tcred);
            header.AppendChild(to);

            //Sender
            XmlElement sender = cXML.CreateElement("Sender");
            XmlElement scred = cXML.CreateElement("Credential");
            scred.SetAttribute("domain", customer.SenderDomain);
            XmlElement sid = cXML.CreateElement("Identity");
            sid.AppendChild(cXML.CreateTextNode(customer.SenderIdentity));
            scred.AppendChild(sid);
            XmlElement ss = cXML.CreateElement("SharedSecret");
            ss.AppendChild(cXML.CreateTextNode(customer.SharedSecret));
            scred.AppendChild(ss);
            sender.AppendChild(scred);
            XmlElement ua = cXML.CreateElement("UserAgent");
            ua.AppendChild(cXML.CreateTextNode(customer.UserAgent));
            sender.AppendChild(ua);
            header.AppendChild(sender);

            return header;
        }

        private XmlElement CreateRequest(XmlDocument cXML)
        {
            XmlElement request = cXML.CreateElement("Request");
            request.SetAttribute("deploymentMode", Constants.DeploymentMode);

            XmlElement invoiceDetailRequest = cXML.CreateElement("InvoiceDetailRequest");

            XmlElement invoiceDetailRequestHeader = cXML.CreateElement("InvoiceDetailRequestHeader");
            invoiceDetailRequestHeader.SetAttribute("invoiceID", salesInvoice.No);
            invoiceDetailRequestHeader.SetAttribute("operation", "new");

            invoiceDetailRequestHeader.SetAttribute("purpose", Purpose);
            if (customer.URL.Contains("ariba"))
            {
                invoiceDetailRequestHeader.SetAttribute("invoiceDate", DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ssK"));
                invoiceDetailRequestHeader.SetAttribute("invoiceOrigin", "supplier");
            }
            else
                invoiceDetailRequestHeader.SetAttribute("invoiceDate", salesInvoice.Posting_Date.ToString("yyyy-MM-dd"));

            XmlElement invoiceDetailHeaderIndicator = cXML.CreateElement("InvoiceDetailHeaderIndicator");
            invoiceDetailRequestHeader.AppendChild(invoiceDetailHeaderIndicator);

            XmlElement invoiceDetailLineIndicator = cXML.CreateElement("InvoiceDetailLineIndicator");
            if (customer.ConnectionMethod == "HTTPS") invoiceDetailLineIndicator.SetAttribute("isAccountingInLine", "yes");
            invoiceDetailRequestHeader.AppendChild(invoiceDetailLineIndicator);

            invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML,"", "remitTo", GSSContact.Name, new string[] { GSSContact.Street }, GSSContact.City, GSSContact.State, GSSContact.PostalCode, GSSContact.CountryCode, GSSContact.Country));

            if (customer.URL.Contains("ariba"))
            {
                if (salesInvoice.OriginalOrder != null)
                    invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, salesInvoice.OriginalOrder.BillToAddress.ShipToCode, "billTo", salesInvoice.OriginalOrder.BillToAddress.Name, new string[] { salesInvoice.OriginalOrder.BillToAddress.Street1, salesInvoice.OriginalOrder.BillToAddress.Street2 }, salesInvoice.OriginalOrder.BillToAddress.City, salesInvoice.OriginalOrder.BillToAddress.State, salesInvoice.OriginalOrder.BillToAddress.PostCode, salesInvoice.OriginalOrder.BillToAddress.CountryCode, salesInvoice.OriginalOrder.BillToAddress.Country));
                else
                    invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "billTo", salesInvoice.Bill_To.Name, new string[] { salesInvoice.Bill_To.Street1, salesInvoice.Bill_To.Street2 }, salesInvoice.Bill_To.City, salesInvoice.Bill_To.State, salesInvoice.Bill_To.PostCode, salesInvoice.Bill_To.CountryCode, salesInvoice.Bill_To.Country));
                invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "from", GSSContact.Name, new string[] { GSSContact.Street }, GSSContact.City, GSSContact.State, GSSContact.PostalCode, GSSContact.CountryCode, GSSContact.Country));
                invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "billFrom", GSSContact.Name, new string[] { GSSContact.Street }, GSSContact.City, GSSContact.State, GSSContact.PostalCode, GSSContact.CountryCode, GSSContact.Country));
                invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "soldTo", salesInvoice.Bill_To.Name, new string[] { salesInvoice.Bill_To.Street1, salesInvoice.Bill_To.Street2 }, salesInvoice.Bill_To.City, salesInvoice.Bill_To.State, salesInvoice.Bill_To.PostCode, salesInvoice.Bill_To.CountryCode, salesInvoice.Bill_To.Country));

                XmlElement invoiceDetailShipping = cXML.CreateElement("InvoiceDetailShipping");
                invoiceDetailShipping.AppendChild(CreateInvoiceContact(cXML, "", "shipFrom", GSSContact.Name, new string[] { GSSContact.Street }, GSSContact.City, GSSContact.State, GSSContact.PostalCode, GSSContact.CountryCode, GSSContact.Country));
                invoiceDetailShipping.AppendChild(CreateInvoiceContact(cXML, salesInvoice.Ship_To.ShipToCode, "shipTo", salesInvoice.Ship_To.Name, new string[] { salesInvoice.Ship_To.Street1, salesInvoice.Ship_To.Street2 }, salesInvoice.Ship_To.City, salesInvoice.Ship_To.State, salesInvoice.Ship_To.PostCode, salesInvoice.Ship_To.CountryCode, salesInvoice.Ship_To.Country));
                invoiceDetailRequestHeader.AppendChild(invoiceDetailShipping);

                XmlElement paymentTerms = cXML.CreateElement("PaymentTerm");
                paymentTerms.SetAttribute("payInNumberOfDays", salesInvoice.OriginalOrder != null ? salesInvoice.OriginalOrder.PaymentTerms : "30");
                invoiceDetailRequestHeader.AppendChild(paymentTerms);

                invoiceDetailRequestHeader.AppendChild(CreateExtrinsic(cXML, "invoiceSourceDocument", "", "PurchaseOrder"));
            }
            else
            {
                invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "billTo", salesInvoice.Bill_To.Name, new string[] { salesInvoice.Bill_To.Street1, salesInvoice.Bill_To.Street2 }, salesInvoice.Bill_To.City, salesInvoice.Bill_To.State, salesInvoice.Bill_To.PostCode, salesInvoice.Bill_To.CountryCode, salesInvoice.Bill_To.Country));
                invoiceDetailRequestHeader.AppendChild(CreateInvoicePartner(cXML, "", "shipFrom", GSSContact.Name, new string[] { GSSContact.Street }, GSSContact.City, GSSContact.State, GSSContact.PostalCode, GSSContact.CountryCode, GSSContact.Country));
            }

            if (customer.Credit_Memo_Org_Inv_Reference.Length > 0)
            {
                switch (customer.Credit_Memo_Org_Inv_Reference)
                {
                    case "Coupa":
                        invoiceDetailRequestHeader.AppendChild(CreateExtrinsic(cXML, "CustomFields", "original_invoice_date", salesInvoice.OriginalInvoiceDate.ToString("yyyy-MM-dd")));
                        invoiceDetailRequestHeader.AppendChild(CreateExtrinsic(cXML, "CustomFields", "original_invoice", salesInvoice.OriginalInvoiceNo));
                        break;
                }
            }

            invoiceDetailRequest.AppendChild(invoiceDetailRequestHeader);

            XmlElement invoiceDetailOrder = cXML.CreateElement("InvoiceDetailOrder");

            XmlElement invoiceDetailOrderInfo = cXML.CreateElement("InvoiceDetailOrderInfo");
            XmlElement orderReference = cXML.CreateElement("OrderReference");
            orderReference.SetAttribute("orderID", salesInvoice.Your_Reference);
            XmlElement documentReference = cXML.CreateElement("DocumentReference");
            documentReference.SetAttribute("payloadID", salesInvoice.OriginalOrder != null ? salesInvoice.OriginalOrder.PayloadID : "");
            orderReference.AppendChild(documentReference);
            invoiceDetailOrderInfo.AppendChild(orderReference);

            if (customer.MasterAgreementReference.Length > 0)
            {
                string masterAgree = "";
                if (customer.MasterAgreementReference.StartsWith("["))
                    masterAgree = GetPropValue(salesInvoice, customer.MasterAgreementReference.Replace("[", "").Replace("]", "")).ToString();
                else
                    masterAgree = customer.MasterAgreementReference;

                XmlElement masterAgreementRefr = cXML.CreateElement("MasterAgreementReference");
                XmlElement masterDocumentRefr = cXML.CreateElement("DocumentReference");
                masterDocumentRefr.SetAttribute("payloadID", masterAgree);
                masterAgreementRefr.AppendChild(masterDocumentRefr);
                invoiceDetailOrderInfo.AppendChild(masterAgreementRefr);
            }

            invoiceDetailOrder.AppendChild(invoiceDetailOrderInfo);

            //Append Items
            if (salesInvoice != null)
                AddInvoiceDetailItemLines(cXML, ref invoiceDetailOrder);

            invoiceDetailRequest.AppendChild(invoiceDetailOrder);

            string SubTotalDisplay = salesInvoice.InvoiceLineTotal.ToString("G29");
            if (!SubTotalDisplay.Contains(".")) SubTotalDisplay += ".00";

            string ShippingAmountDisplay = salesInvoice.ShippingAmount.ToString("G29");
            if (!ShippingAmountDisplay.Contains(".")) ShippingAmountDisplay += ".00";

            string GrossAmountDisplay = (salesInvoice.InvoiceLineTotal + salesInvoice.ShippingAmount).ToString("G29");
            if (!GrossAmountDisplay.Contains(".")) GrossAmountDisplay += ".00";

            XmlElement invoiceDetailSummary = cXML.CreateElement("InvoiceDetailSummary");

            XmlElement subtotalAmount = cXML.CreateElement("SubtotalAmount");
            XmlElement subtotalAmountMoney = cXML.CreateElement("Money");
            subtotalAmountMoney.SetAttribute("currency", "USD");
            subtotalAmountMoney.AppendChild(cXML.CreateTextNode(SubTotalDisplay));
            subtotalAmount.AppendChild(subtotalAmountMoney);
            invoiceDetailSummary.AppendChild(subtotalAmount);

            XmlElement tax = cXML.CreateElement("Tax");
            XmlElement taxMoney = cXML.CreateElement("Money");
            taxMoney.SetAttribute("currency", "USD");
            tax.AppendChild(taxMoney);
            XmlElement taxDescription = cXML.CreateElement("Description");
            taxDescription.SetAttribute("xml:lang", "en");
            taxDescription.AppendChild(cXML.CreateTextNode("Total Tax"));
            tax.AppendChild(taxDescription);
            invoiceDetailSummary.AppendChild(tax);

            if (!customer.URL.Contains("ariba"))
            {
                XmlElement specialHandlingAmount = cXML.CreateElement("SpecialHandlingAmount");
                XmlElement specialHandlingAmountMoney = cXML.CreateElement("Money");
                specialHandlingAmountMoney.SetAttribute("currency", "USD");
                specialHandlingAmountMoney.AppendChild(cXML.CreateTextNode("0"));
                specialHandlingAmount.AppendChild(specialHandlingAmountMoney);
                XmlElement specialHandlingDescription = cXML.CreateElement("Description");
                specialHandlingDescription.SetAttribute("xml:lang", "en");
                specialHandlingDescription.AppendChild(cXML.CreateTextNode(""));
                specialHandlingAmount.AppendChild(specialHandlingDescription);
                invoiceDetailSummary.AppendChild(specialHandlingAmount);
            }

            XmlElement shippingAmount = cXML.CreateElement("ShippingAmount");
            XmlElement shippingAmountMoney = cXML.CreateElement("Money");
            shippingAmountMoney.SetAttribute("currency", "USD");
            shippingAmountMoney.AppendChild(cXML.CreateTextNode(ShippingAmountDisplay));
            shippingAmount.AppendChild(shippingAmountMoney);
            invoiceDetailSummary.AppendChild(shippingAmount);

            XmlElement grossAmount = cXML.CreateElement("GrossAmount");
            XmlElement grossAmountMoney = cXML.CreateElement("Money");
            grossAmountMoney.SetAttribute("currency", "USD");
            grossAmountMoney.AppendChild(cXML.CreateTextNode(GrossAmountDisplay));
            grossAmount.AppendChild(grossAmountMoney);
            invoiceDetailSummary.AppendChild(grossAmount);

            if (!customer.URL.Contains("ariba"))
            {
                XmlElement invoiceDetailDiscount = cXML.CreateElement("InvoiceDetailDiscount");
                XmlElement invoiceDetailDiscountMoney = cXML.CreateElement("Money");
                invoiceDetailDiscountMoney.SetAttribute("currency", "USD");
                invoiceDetailDiscountMoney.AppendChild(cXML.CreateTextNode("0.00"));
                invoiceDetailDiscount.AppendChild(invoiceDetailDiscountMoney);
                invoiceDetailSummary.AppendChild(invoiceDetailDiscount);
            }

            XmlElement netAmount = cXML.CreateElement("NetAmount");
            XmlElement netAmountMoney = cXML.CreateElement("Money");
            netAmountMoney.SetAttribute("currency", "USD");
            netAmountMoney.AppendChild(cXML.CreateTextNode(GrossAmountDisplay));
            netAmount.AppendChild(netAmountMoney);
            invoiceDetailSummary.AppendChild(netAmount);

            if (!customer.URL.Contains("ariba"))
            {
                XmlElement depositAmount = cXML.CreateElement("DepositAmount");
                XmlElement depositAmountMoney = cXML.CreateElement("Money");
                depositAmountMoney.SetAttribute("currency", "USD");
                depositAmountMoney.AppendChild(cXML.CreateTextNode("0.00"));
                depositAmount.AppendChild(depositAmountMoney);
                invoiceDetailSummary.AppendChild(depositAmount);
            }

            XmlElement dueAmount = cXML.CreateElement("DueAmount");
            XmlElement dueAmountMoney = cXML.CreateElement("Money");
            dueAmountMoney.SetAttribute("currency", "USD");
            dueAmountMoney.AppendChild(cXML.CreateTextNode(GrossAmountDisplay));
            dueAmount.AppendChild(dueAmountMoney);
            invoiceDetailSummary.AppendChild(dueAmount);

            invoiceDetailRequest.AppendChild(invoiceDetailSummary);
            request.AppendChild(invoiceDetailRequest);

            return request;
        }

        private void AddInvoiceDetailItemLines(XmlDocument cXML, ref XmlElement invoiceDetailOrder)
        {
            foreach (InvoiceLine line in salesInvoice.SalesInvLines)
            {
                if (line.Quantity != 0)
                {
                    if (line.Type == 2 || (customer.ShipHeader != 1 && line.Type == 1 && Constants.GLAccounts.Contains(line.No)))
                    {
                        if (line.OrderLine == null)
                            salesInvoice.NeedsReview = true;

                        line.ValidationError = ValidateInvoiceLine(line);
                        if (line.ValidationError.Length > 0)
                            salesInvoice.NeedsReview = true;

                        invoiceDetailOrder.AppendChild(CreateInvoiceDetailItem(cXML, line, TotalNumLines + 1));
                        TotalNumLines++;
                    }
                }
                else if (line.Type == 1 && line.Description.StartsWith("MAP:"))
                    GlLineNumber = GlLineNumber + 1;
            }

            
        }

        private string ValidateInvoiceLine(InvoiceLine line)
        {
            if (line.OrderLine == null)
                return "Original PO could not be found";

            decimal difference = 1 - (line.Unit_Price / line.OrderLine.Unit_Price);
            bool priceDiff = difference >= Constants.minThreshold && difference <= Constants.maxThreshold;

            if (!priceDiff)
                return "Sales Invoice price does not match Original PO Price";
            else
                return "";
        }

        private XmlElement CreateInvoicePartner(XmlDocument cXML, string addressID, string role, string name, string[] streets, string city, string state, string postalCode, string countryCode, string country)
        {
            XmlElement invoicePartner = cXML.CreateElement("InvoicePartner");
            invoicePartner.AppendChild(CreateInvoiceContact(cXML, addressID, role, name, streets, city, state, postalCode, countryCode, country));
            return invoicePartner;
        }

        private XmlElement CreateInvoiceContact(XmlDocument cXML, string addressID, string role, string name, string[] streets, string city, string state, string postalCode, string countryCode, string country)
        {
            XmlElement contact = cXML.CreateElement("Contact");
            contact.SetAttribute("role", role);
            if (addressID.Length > 0)
                contact.SetAttribute("addressID", addressID);

            XmlElement nameElement = cXML.CreateElement("Name");
            nameElement.SetAttribute("xml:lang", "en");
            nameElement.AppendChild(cXML.CreateTextNode(name));
            contact.AppendChild(nameElement);

            XmlElement postalAddress = cXML.CreateElement("PostalAddress");

            foreach (string street in streets)
            {
                if (street != null && street.Length > 0 && street.ToUpper() != "ACCOUNTS PAYABLE")
                {
                    XmlElement streetElement = cXML.CreateElement("Street");
                    streetElement.AppendChild(cXML.CreateTextNode(street));
                    postalAddress.AppendChild(streetElement);
                }
            }

            XmlElement cityElement = cXML.CreateElement("City");
            cityElement.AppendChild(cXML.CreateTextNode(city));
            postalAddress.AppendChild(cityElement);

            XmlElement stateElement = cXML.CreateElement("State");
            stateElement.AppendChild(cXML.CreateTextNode(state));
            postalAddress.AppendChild(stateElement);

            XmlElement postalCodeElement = cXML.CreateElement("PostalCode");
            postalCodeElement.AppendChild(cXML.CreateTextNode(postalCode));
            postalAddress.AppendChild(postalCodeElement);

            XmlElement countryElement = cXML.CreateElement("Country");
            countryElement.SetAttribute("isoCountryCode", countryCode);
            countryElement.AppendChild(cXML.CreateTextNode(country));
            postalAddress.AppendChild(countryElement);

            contact.AppendChild(postalAddress);

            return contact;
        }

        private XmlElement CreateInvoiceDetailItem(XmlDocument cXML, InvoiceLine line, int lineNo)
        {
            line.Invoice_Line_No = lineNo;
            
            string Unit_Of_Measure = "", Description = "";

            if (line.OrderLine != null)
            {
                line.Supplier_Part_ID = line.OrderLine.SupplierPartID;
                Description = line.OrderLine.Description;
                line.Refr_Line_No = line.OrderLine.LineNo;
                Unit_Of_Measure = line.OrderLine.UnitOfMeasure;
            }
            else
            {
                if (line.Vendor_Item_No.Length > 0)
                    line.Supplier_Part_ID = line.Vendor_Item_No;
                else
                    line.Supplier_Part_ID = line.No;
                line.Refr_Line_No = line.Line_No - GlLineNumber;
                Description = line.Description;
            }

            string UnitPriceDisplay = line.Unit_Price.ToString("G29");
            if (!UnitPriceDisplay.Contains(".")) UnitPriceDisplay += ".00";

            XmlElement invoiceDetailItem = cXML.CreateElement("InvoiceDetailItem");
            invoiceDetailItem.SetAttribute("invoiceLineNumber", lineNo.ToString());
            invoiceDetailItem.SetAttribute("quantity", line.Quantity.ToString("G29"));

            XmlElement unitOfMeasure = cXML.CreateElement("UnitOfMeasure");
            unitOfMeasure.AppendChild(cXML.CreateTextNode(Unit_Of_Measure));
            invoiceDetailItem.AppendChild(unitOfMeasure);

            XmlElement unitPrice = cXML.CreateElement("UnitPrice");
            XmlElement unitPriceMoney = cXML.CreateElement("Money");
            unitPriceMoney.SetAttribute("currency", "USD");
            unitPriceMoney.AppendChild(cXML.CreateTextNode(UnitPriceDisplay));
            unitPrice.AppendChild(unitPriceMoney);
            invoiceDetailItem.AppendChild(unitPrice);

            XmlElement invoiceDetailItemReference = cXML.CreateElement("InvoiceDetailItemReference");
            invoiceDetailItemReference.SetAttribute("lineNumber", line.Refr_Line_No.ToString());

            XmlElement itemId = cXML.CreateElement("ItemID");
            XmlElement supplierPartId = cXML.CreateElement("SupplierPartID");
            supplierPartId.AppendChild(cXML.CreateTextNode(line.Supplier_Part_ID));
            itemId.AppendChild(supplierPartId);
            invoiceDetailItemReference.AppendChild(itemId);

            XmlElement description = cXML.CreateElement("Description");
            description.SetAttribute("xml:lang", "en");
            description.AppendChild(cXML.CreateTextNode(Description));
            invoiceDetailItemReference.AppendChild(description);

            invoiceDetailItem.AppendChild(invoiceDetailItemReference);

            string LineTotalDisplay = line.LineTotal.ToString("G29");
            if (!LineTotalDisplay.Contains(".")) LineTotalDisplay += ".00";

            XmlElement subtotalAmount = cXML.CreateElement("SubtotalAmount");
            XmlElement subtotalAmountMoney = cXML.CreateElement("Money");
            subtotalAmountMoney.SetAttribute("currency", "USD");
            subtotalAmountMoney.AppendChild(cXML.CreateTextNode(LineTotalDisplay));
            subtotalAmount.AppendChild(subtotalAmountMoney);
            invoiceDetailItem.AppendChild(subtotalAmount);

            if (customer.URL.Contains("ariba"))
            {
                XmlElement grossAmount = cXML.CreateElement("GrossAmount");
                XmlElement grossAmountMoney = cXML.CreateElement("Money");
                grossAmountMoney.SetAttribute("currency", "USD");
                grossAmountMoney.AppendChild(cXML.CreateTextNode(LineTotalDisplay));
                grossAmount.AppendChild(grossAmountMoney);
                invoiceDetailItem.AppendChild(grossAmount);

                XmlElement netAmount = cXML.CreateElement("NetAmount");
                XmlElement netAmountMoney = cXML.CreateElement("Money");
                netAmountMoney.SetAttribute("currency", "USD");
                netAmountMoney.AppendChild(cXML.CreateTextNode(LineTotalDisplay));
                netAmount.AppendChild(netAmountMoney);
                invoiceDetailItem.AppendChild(netAmount);

                invoiceDetailItem.AppendChild(CreateExtrinsic(cXML, "punchinItemFromCatalog", "", "no"));
                invoiceDetailItem.AppendChild(CreateExtrinsic(cXML, "productType", "", "Material"));
            }
            else
            {
                XmlElement tax = cXML.CreateElement("Tax");
                XmlElement taxMoney = cXML.CreateElement("Money");
                taxMoney.SetAttribute("currency", "USD");
                tax.AppendChild(taxMoney);
                XmlElement taxDescription = cXML.CreateElement("Description");
                taxDescription.SetAttribute("xml:lang", "en");
                taxDescription.AppendChild(cXML.CreateTextNode("Total Tax"));
                tax.AppendChild(taxDescription);
                invoiceDetailItem.AppendChild(tax);
            }

            return invoiceDetailItem;
        }

        private XmlElement CreateExtrinsic(XmlDocument cXML, string name, string identifier, string description)
        {
            XmlElement extrinsic = cXML.CreateElement("Extrinsic");
            extrinsic.SetAttribute("name", name);

            if (identifier.Length > 0)
            {
                XmlElement idReference = cXML.CreateElement("IdReference");
                idReference.SetAttribute("identifier", identifier);
                idReference.SetAttribute("domain", name);

                XmlElement descElement = cXML.CreateElement("Description");
                descElement.SetAttribute("xml:lang", "en");
                descElement.AppendChild(cXML.CreateTextNode(description));

                idReference.AppendChild(descElement);
                extrinsic.AppendChild(idReference);
            }
            else
            {
                extrinsic.AppendChild(cXML.CreateTextNode(description));
            }

            return extrinsic;
        }

        public object GetPropValue(object src, string propName)
        {
            try
            {
                return src.GetType().GetProperty(propName).GetValue(src, null);
            }
            catch
            {
                return "";
            }
        }
    }
}