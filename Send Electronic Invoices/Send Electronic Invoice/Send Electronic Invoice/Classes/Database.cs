using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using Org.BouncyCastle.Tls;
using Send_Electronic_Invoice.Objects;

namespace Send_Electronic_Invoice.Classes
{
    public static class Database
    {
        public static List<Customer> GetCustomers()
        {
            List<Customer> customers = new List<Customer>();
            SqlCommand cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]");
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionEcommerce))
                {
                    dbcon.Open();
                    cmd.Connection = dbcon;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET CUSTOMERS"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            bool timepass = false;
                            string dispatchTimes = rs["dispatchTimes"].ToString();
                            if (dispatchTimes.Length == 0)
                                timepass = true;
                            else
                            {
                                string[] dispatchTimesArray = dispatchTimes.Split(',');
                                foreach (string dispatchTime in dispatchTimesArray)
                                {
                                    try
                                    {
                                        TimeSpan time = TimeSpan.Parse(dispatchTime);
                                        if (DateTime.Now.TimeOfDay >= time && DateTime.Now.TimeOfDay <= time.Add(new TimeSpan(0, 5, 0)))
                                            timepass = true;
                                    }
                                    catch
                                    {
                                    }
                                }
                            }

                            if (timepass && (rs["schedule"].ToString().Length == 0 || rs["schedule"].ToString().Contains(DateTime.Now.ToString("dd")))
                                && (rs["blockSchedule"].ToString().Length == 0 || !rs["blockSchedule"].ToString().Contains(DateTime.Now.ToString("dd"))))
                                customers.Add(new Customer(
                                    rs["Sell-to Customer No_"].ToString(),
                                    rs["fromDomain"].ToString(),
                                    rs["fromIdentity"].ToString(),
                                    rs["toDomain"].ToString(),
                                    rs["toIdentity"].ToString(),
                                    rs["senderDomain"].ToString(),
                                    rs["senderIdentity"].ToString(),
                                    rs["sharedSecret"].ToString(),
                                    rs["userAgent"].ToString(),
                                    rs["URL"].ToString(),
                                    rs["encryptionKey"].ToString(),
                                    rs["ftpServer"].ToString(),
                                    rs["ftpPort"].ToString(),
                                    rs["ftpUsername"].ToString(),
                                    rs["ftpPassword"].ToString(),
                                    rs["connectionMethod"].ToString(),
                                    rs["sqlCondition"].ToString(),
                                    rs["schedule"].ToString(),
                                    rs["fileExtension"].ToString(),
                                    rs["ftpFolder"].ToString(),
                                    int.Parse(rs["shipHeader"].ToString()),
                                    int.Parse(rs["negativeQty"].ToString()),
                                    rs["cxmlDocType"].ToString(),
                                    rs["cmOrgInvRef"].ToString(),
                                    rs["masterAgreeRefr"].ToString(),
                                    rs["textDelimiter"].ToString(),
                                    rs["ftpHostKey"].ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetCustomers", cmd));
            }
            return customers;
        }

        public static void GetInvoices(Customer customer)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@sellToCust", customer.SellToCustomerNo));
                    cmd.Parameters.Add(new SqlParameter("@filter", customer.SQL_Condition));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET POSTED INVOICES"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            InvoiceHeader invoice = new InvoiceHeader(
                                rs["No_"].ToString()
                                , new Address("", rs["Bill-to Name"].ToString(), rs["Bill-to Address"].ToString(), rs["Bill-to Address 2"].ToString(), rs["Bill-to City"].ToString(), rs["Bill-to County"].ToString(), rs["Bill-to Post Code"].ToString())
                                , new Address(rs["Ship-to Code"].ToString(), rs["Ship-to Name"].ToString(), rs["Ship-to Address"].ToString(), rs["Ship-to Address 2"].ToString(), rs["Ship-to City"].ToString(), rs["Ship-to County"].ToString(), rs["Ship-to Post Code"].ToString())
                                , rs["Your Reference"].ToString()
                                , DateTime.Parse(rs["Posting Date"].ToString())
                                , "");

                            GetInvoiceLine(customer, ref invoice);

                            if (invoice.SalesInvLines.FindAll(s =>s.Type == 2 && s.Quantity > 0.00M).Count == 0 && customer.SellToCustomerNo == "BERKELEY")
                            {
                                InvoiceLine line = new InvoiceLine(customer, 1, "GSSTRANS", "GSSTRANS", "EA", "Expedite Shipping", 1.00M, 0.01M, 2, false);
                                invoice.SalesInvLines.Add(line);
                            }

                            invoice = Functions.CalculateInvoiceTotals(customer, invoice);

                            customer.Invoices.Add(invoice);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetInvoices", cmd));
            }
        }
        private static void GetInvoiceLine(Customer customer, ref InvoiceHeader invoice)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@invoiceNo", invoice.No));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET POSTED INVOICE LINES"));
                    using (SqlDataReader lineRs = cmd.ExecuteReader())
                    {
                        while (lineRs.Read())
                        {
                            InboundOrder order = null;

                            InvoiceLine line = new InvoiceLine(
                                customer, 
                                (int.Parse(lineRs["Line No_"].ToString()) / 10000),
                                lineRs["No_"].ToString(),
                                lineRs["Vendor Item No_"].ToString(),
                                lineRs["Unit of Measure Code"].ToString(),
                                lineRs["Description"].ToString(),
                                decimal.Parse(lineRs["Quantity"].ToString()),
                                decimal.Parse(lineRs["Unit Price"].ToString()),
                                int.Parse(lineRs["Type"].ToString()),
                                false);

                            line.OrderLine = GetInboundOrderLine(ref order, invoice.Your_Reference, line.Line_No, line.No, line.Vendor_Item_No, line.Unit_of_Measure_Code);
                            if(invoice.OriginalOrder == null && order !=null)
                                invoice.OriginalOrder = order;

                            invoice.SalesInvLines.Add(line);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetInvoiceLine", cmd));
            }
        }

        public static void GetCreditMemos(Customer customer)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@sellToCust", customer.SellToCustomerNo));
                    cmd.Parameters.Add(new SqlParameter("@filter", customer.SQL_Condition));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET POSTED CREDITS"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            InvoiceHeader invoice = new InvoiceHeader(
                                rs["No_"].ToString()
                                , new Address("", rs["Bill-to Name"].ToString(), rs["Bill-to Address"].ToString(), rs["Bill-to Address 2"].ToString(), rs["Bill-to City"].ToString(), rs["Bill-to County"].ToString(), rs["Bill-to Post Code"].ToString())
                                , new Address(rs["Ship-to Code"].ToString(), rs["Ship-to Name"].ToString(), rs["Ship-to Address"].ToString(), rs["Ship-to Address 2"].ToString(), rs["Ship-to City"].ToString(), rs["Ship-to County"].ToString(), rs["Ship-to Post Code"].ToString())
                                , rs["Your Reference"].ToString()
                                , DateTime.Parse(rs["Posting Date"].ToString())
                                , rs["Invoice No_"].ToString()
                                , DateTime.Parse(rs["Invoice Date"].ToString())
                                , "");

                            GetCreditMemoLines(customer, ref invoice);

                            if (invoice.SalesInvLines.FindAll(s => s.Type == 2 && s.Quantity > 0.00M).Count == 0 && customer.SellToCustomerNo == "BERKELEY")
                            {
                                InvoiceLine line = new InvoiceLine(customer, 1, "GSSTRANS", "GSSTRANS", "EA", "Expedite Shipping", 1.00M, 0.01M, 2, true);
                                invoice.SalesInvLines.Add(line);
                            }

                            invoice = Functions.CalculateInvoiceTotals(customer, invoice);

                            customer.Invoices.Add(invoice);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetCreditMemos", cmd));
            }


        }
        private static void GetCreditMemoLines(Customer customer, ref InvoiceHeader invoice)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@invoiceNo", invoice.No));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET POSTED CREDIT LINES"));
                    using (SqlDataReader lineRs = cmd.ExecuteReader())
                    {
                        while (lineRs.Read())
                        {
                            InboundOrder order = null;

                            InvoiceLine line = new InvoiceLine(
                                customer,
                                (int.Parse(lineRs["Line No_"].ToString()) / 10000),
                                lineRs["No_"].ToString(),
                                lineRs["Vendor Item No_"].ToString(),
                                lineRs["Unit of Measure Code"].ToString(),
                                lineRs["Description"].ToString(),
                                decimal.Parse(lineRs["Quantity"].ToString()),
                                decimal.Parse(lineRs["Unit Price"].ToString()),
                                int.Parse(lineRs["Type"].ToString()),
                                true);

                            line.OrderLine = GetInboundOrderLine(ref order, invoice.Your_Reference, line.Line_No, line.No, line.Vendor_Item_No, line.Unit_of_Measure_Code);
                            if (invoice.OriginalOrder == null && order != null)
                                invoice.OriginalOrder = order;

                            invoice.SalesInvLines.Add(line);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetCreditMemoLine", cmd));
            }          
        }

        public static void GetDummyInvoices(Customer customer)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@sellToCust", customer.SellToCustomerNo));
                    cmd.Parameters.Add(new SqlParameter("@filter", customer.SQL_Condition));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET SALES INVOICES"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            InvoiceHeader invoice = new InvoiceHeader(
                                rs["No_"].ToString()
                                , new Address("", rs["Bill-to Name"].ToString(), rs["Bill-to Address"].ToString(), rs["Bill-to Address 2"].ToString(), rs["Bill-to City"].ToString(), rs["Bill-to County"].ToString(), rs["Bill-to Post Code"].ToString())
                                , new Address(rs["Ship-to Code"].ToString(), rs["Ship-to Name"].ToString(), rs["Ship-to Address"].ToString(), rs["Ship-to Address 2"].ToString(), rs["Ship-to City"].ToString(), rs["Ship-to County"].ToString(), rs["Ship-to Post Code"].ToString())
                                , rs["Your Reference"].ToString()
                                , DateTime.Parse(rs["Posting Date"].ToString())
                                , "");

                            GetDummyInvoiceLines(customer, ref invoice);

                            if (invoice.SalesInvLines.FindAll(s =>s.Type == 2 && s.Quantity > 0.00M).Count == 0 && customer.SellToCustomerNo == "BERKELEY")
                            {
                                InvoiceLine line = new InvoiceLine(customer, 1, "GSSTRANS", "GSSTRANS", "EA", "Expedite Shipping", 1.00M, 0.01M, 2, invoice.No.StartsWith("CMP"));
                                invoice.SalesInvLines.Add(line);
                            }

                            invoice = Functions.CalculateInvoiceTotals(customer, invoice);

                            customer.Invoices.Add(invoice);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetDummyInvoices", cmd));
            }


        }
        private static void GetDummyInvoiceLines(Customer customer, ref InvoiceHeader invoice)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@invoiceNo", invoice.No));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET SALES INVOICE LINES"));
                    using (SqlDataReader lineRs = cmd.ExecuteReader())
                    {
                        while (lineRs.Read())
                        {
                            InboundOrder order = null;

                            InvoiceLine line = new InvoiceLine(
                                customer,
                                (int.Parse(lineRs["Line No_"].ToString()) / 10000),
                                lineRs["No_"].ToString(),
                                lineRs["Vendor Item No_"].ToString(),
                                lineRs["Unit of Measure Code"].ToString(),
                                lineRs["Description"].ToString(),
                                decimal.Parse(lineRs["Quantity"].ToString()),
                                decimal.Parse(lineRs["Unit Price"].ToString()),
                                int.Parse(lineRs["Type"].ToString()),
                                invoice.No.StartsWith("CMP"));

                            line.OrderLine = GetInboundOrderLine(ref order, invoice.Your_Reference, line.Line_No, line.No, line.Vendor_Item_No, line.Unit_of_Measure_Code);
                            if (invoice.OriginalOrder == null && order != null)
                                invoice.OriginalOrder = order;

                            invoice.SalesInvLines.Add(line);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetDummyInvoiceLine", cmd));
            }
        }
        public static void UpdateInvoice(string no, string custPONo, int sent)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "UPDATE INVOICE"));
                    cmd.Parameters.Add(new SqlParameter("@invoiceNo", no));
                    cmd.Parameters.Add(new SqlParameter("@customerPoNo", custPONo));
                    cmd.Parameters.Add(new SqlParameter("@wasSent", sent));
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "UpdateInvoice", cmd));
            }
        }

        //public static InboundOrder GetInboundOrder(string poNum)
        //{
        //    SqlCommand cmd = null;

        //    List<InboundOrderLines> lines = new List<InboundOrderLines>();
        //    decimal total = 0.0M;
        //    string orderDate = "", payloadId = "", paymentTerms = "", addressID = "", name = "", street1 = "", street2 = "", city = "", state = "", postCode = "";

        //    try
        //    {
        //        using (SqlConnection dbcon = new SqlConnection(Constants.connectionEcommerce))
        //        {
        //            dbcon.Open();
        //            cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
        //            cmd.CommandType = CommandType.StoredProcedure;
        //            cmd.Parameters.Add(new SqlParameter("@poNo", poNum));
        //            cmd.Parameters.Add(new SqlParameter("@method", "GET ORDER"));
        //            using (SqlDataReader rs = cmd.ExecuteReader())
        //            {
        //                int lineNumber = 1;
        //                while (rs.Read())
        //                {
        //                    int tmpLineNo = 0;
        //                    try { tmpLineNo = int.Parse(rs["lineNumber"].ToString()); }
        //                    catch { tmpLineNo = 0; }

        //                    InboundOrderLines line = new InboundOrderLines(
        //                        rs["gssPartNumber"].ToString().StartsWith("GSSTRANS") ? "GSSTRANS" : rs["gssPartNumber"].ToString()
        //                        , rs["supplierPartID"].ToString()
        //                        , rs["description"].ToString()
        //                        , rs["originalUOM"].ToString()
        //                        , tmpLineNo > 0 ? tmpLineNo : lineNumber
        //                        , decimal.Parse(rs["qty"].ToString())
        //                        , decimal.Parse(rs["unitPrice"].ToString()));

        //                    orderDate = rs["Order Date"].ToString();
        //                    payloadId = rs["payloadId"].ToString();
        //                    paymentTerms = rs["paymentTerms"].ToString();
        //                    addressID = rs["billToAddressID"].ToString();
        //                    name = rs["billToName"].ToString();
        //                    street1 = rs["billToStreet1"].ToString();
        //                    street2 = rs["billToStreet2"].ToString();
        //                    city = rs["billToCity"].ToString();
        //                    state = rs["billToState"].ToString();
        //                    postCode = rs["billToPostalCode"].ToString();

        //                    lines.Add(line);

        //                    lineNumber++;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Email.SendErrorMessage(ex, "Send_Electronic_Invoice", "Constants.GetInboundOrder", cmd);
        //    }

        //    foreach (InboundOrderLines l in lines)
        //        total += (l.Unit_Price * l.Quantity);

        //    InboundOrder order = new InboundOrder(poNum, lines, total, orderDate, payloadId, paymentTerms, new Address(addressID, name, street1, street2, city, state, postCode));

        //    return order;
        //}

        public static InboundOrderLine GetInboundOrderLine(ref InboundOrder order, string poNum, int lineNumber, string gssPartNo, string supplierPartNo, string unitOfMeasure)
        {
            SqlCommand cmd = null;
            InboundOrderLine line = null;
            string lineNumbersUsed = ";";

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@poNo", poNum));
                    cmd.Parameters.Add(new SqlParameter("@lineNumber", lineNumber));
                    cmd.Parameters.Add(new SqlParameter("@gssPartNo", gssPartNo));
                    cmd.Parameters.Add(new SqlParameter("@supplierPartId", supplierPartNo));
                    cmd.Parameters.Add(new SqlParameter("@unitOfMeasure", unitOfMeasure));
                    cmd.Parameters.Add(new SqlParameter("@lineNumbersUsed", lineNumbersUsed));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET ORDER LINE"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                        {
                            line = new InboundOrderLine(
                                rs["gssPartNumber"].ToString().StartsWith("GSSTRANS") ? "GSSTRANS" : rs["gssPartNumber"].ToString()
                                , rs["supplierPartID"].ToString()
                                , rs["description"].ToString()
                                , rs["originalUOM"].ToString()
                                , int.Parse(rs["lineNumber"].ToString())
                                , decimal.Parse(rs["qty"].ToString())
                                , decimal.Parse(rs["unitPrice"].ToString()));

                            lineNumbersUsed += $"{rs["lineNumber"].ToString()};";

                            order = new InboundOrder(poNum
                                , 0.00M
                                , rs["Order Date"].ToString()
                                , rs["payloadId"].ToString()
                                , rs["paymentTerms"].ToString()
                                , new Address(rs["billToAddressID"].ToString()
                                    , rs["billToName"].ToString()
                                    , rs["billToStreet1"].ToString()
                                    , rs["billToStreet2"].ToString()
                                    , rs["billToCity"].ToString()
                                    , rs["billToState"].ToString()
                                    , rs["billToPostalCode"].ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "Database", "GetInboundOrderLine", cmd));
            }

            return line;
        }

        public static int GetInvoiceHistory(string invNumber)
        {
            int count = 1;
            SqlCommand cmd = new SqlCommand("[dbo].[Ecommerce.BizTalk.InternalOrders.Control]");

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.connectionGssNav))
                {
                    dbcon.Open();
                    cmd.Connection = dbcon;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@invoiceNo", invNumber));
                    cmd.Parameters.Add(new SqlParameter("@method", "GET INVOICE HISTORY"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                        if (rs.Read())
                            count = int.Parse(rs[0].ToString());
                }
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "SendReport", "GetInvoiceHistory", cmd));
            }

            return count;
        }
    }
}
