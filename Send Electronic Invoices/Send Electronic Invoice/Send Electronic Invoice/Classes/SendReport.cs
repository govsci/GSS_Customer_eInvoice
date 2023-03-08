using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using Send_Electronic_Invoice.Objects;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;

namespace Send_Electronic_Invoice.Classes
{
    public class SendReport
    {
        private HSSFWorkbook workbook;
        private List<int> LineNumbersUsed = new List<int>();
        public delegate string F1(XmlNode x);

        public SendReport(List<Customer> customers)
        {
            try
            {
                int workSheetCount = 0;

                InitializeWorkbook();

                foreach (Customer c in customers)
                {
                    if (c.Invoices.Count > 0)
                    {
                        AddWorksheet(c);
                        workSheetCount++;
                    }
                }

                string excelPath = "";
                if (workSheetCount > 0)
                {
                    excelPath = $@"{Constants.InvoiceReportFolder}\{DateTime.Now.ToString(@"yyyy\\MM\\dd\\")}";
                    if (!Directory.Exists(excelPath)) Directory.CreateDirectory(excelPath);

                    excelPath += Functions.TimeStamp(DateTime.Now) + ".InvoiceReport.xls";

                    WriteToFile(excelPath);
                    workbook.Close();

                    //excelPath = Convert(excelPath);
                }
                else
                    workbook.Close();

                SendEmail(excelPath, customers);
            }
            catch (Exception ex)
            {
                Constants.ApplicationErrors.Add(new CodeError(ex, "SendReport", "SendReport", null));
            }
        }

        private void InitializeWorkbook()
        {
            workbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Government Scientific Source, Inc.";
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Invoices Sent Report";
            workbook.SummaryInformation = si;
        }
        private void AddWorksheet(Customer customer)
        {
            F1 setStringValue = x => x == null ? "" : x.InnerXml;

            string sheetName = $"{customer.SellToCustomerNo} {customer.SQL_Condition} Invoices";
            if (sheetName.Length > 30) sheetName = sheetName.Remove(30);

            ISheet worksheet = workbook.CreateSheet(sheetName);
            WorksheetAddHeaderRow(ref worksheet, customer);

            int rowNo = 2;
            foreach (InvoiceHeader inv in customer.Invoices)
            {
                foreach (InvoiceLine line in inv.SalesInvLines)
                {
                    if (line.Quantity != 0)
                    {
                        try
                        {
                            decimal percent = 0.0M, unitPrice = line.Unit_Price < 0.00M ? line.Unit_Price * -1 : line.Unit_Price;
                            if (line.OrderLine != null && line.OrderLine.Unit_Price > 0.00M)
                                percent = (line.Unit_Price - line.OrderLine.Unit_Price) / (line.OrderLine.Unit_Price);
                            else
                                percent = 1.00M;

                            StringBuilder inverrors = new StringBuilder();
                            foreach (CodeError err in inv.Errors)
                                inverrors.Append($"{(inverrors.Length > 0 ? ";" : "")}{err.Error.Message}");

                            StringBuilder custerrors = new StringBuilder();
                            foreach (CodeError err in inv.Errors)
                                custerrors.Append($"{(custerrors.Length > 0 ? ";" : "")}{err.Error.Message}");

                            IRow row = worksheet.CreateRow(rowNo);

                            row.CreateCell(0).SetCellValue(inv.No);
                            row.CreateCell(1).SetCellValue(inv.Posting_Date.ToString("yyyy-MM-dd"));
                            row.CreateCell(2).SetCellValue(inv.Your_Reference);
                            row.CreateCell(3).SetCellValue(inv.ShippingAmount.ToString("G29"));
                            row.CreateCell(4).SetCellValue((inv.InvoiceLineTotal + inv.ShippingAmount).ToString("G29"));
                            row.CreateCell(5).SetCellValue(inv.OriginalOrder != null ? inv.OriginalOrder.POTotal.ToString("G29") : "0.00");
                            row.CreateCell(6).SetCellValue(line.Invoice_Line_No.ToString());
                            row.CreateCell(7).SetCellValue(line.Refr_Line_No.ToString());
                            row.CreateCell(8).SetCellValue(line.Supplier_Part_ID);
                            row.CreateCell(9).SetCellValue(line.OrderLine != null ? line.OrderLine.Description : line.Description);
                            row.CreateCell(10).SetCellValue(line.OrderLine != null ? line.OrderLine.UnitOfMeasure : line.Unit_of_Measure_Code);
                            row.CreateCell(11).SetCellValue(line.Quantity.ToString("G29"));
                            row.CreateCell(12).SetCellValue(line.Unit_Price.ToString("G29"));
                            row.CreateCell(13).SetCellValue(line.LineTotal.ToString("G29"));
                            row.CreateCell(14).SetCellValue(line.OrderLine != null ? line.OrderLine.LineNo.ToString() : "0");
                            row.CreateCell(15).SetCellValue(line.OrderLine != null ? line.OrderLine.GSSPartNo.ToString() : "");
                            row.CreateCell(16).SetCellValue(line.OrderLine != null ? line.OrderLine.SupplierPartID.ToString() : "");
                            row.CreateCell(17).SetCellValue(line.OrderLine != null ? line.OrderLine.Description.ToString() : "");
                            row.CreateCell(18).SetCellValue(line.OrderLine != null ? line.OrderLine.UnitOfMeasure.ToString() : "");
                            row.CreateCell(19).SetCellValue(line.OrderLine != null ? line.OrderLine.Quantity.ToString("G29") : "0");
                            row.CreateCell(20).SetCellValue(line.OrderLine != null ? line.OrderLine.Unit_Price.ToString("G29") : "0");
                            row.CreateCell(21).SetCellValue(line.OrderLine != null ? (line.OrderLine.Quantity * line.OrderLine.Unit_Price).ToString("G29") : "0");
                            row.CreateCell(22).SetCellValue((percent * 100).ToString("G29") + "%");
                            row.CreateCell(23).SetCellValue(Database.GetInvoiceHistory(inv.No).ToString());
                            row.CreateCell(24).SetCellValue(Functions.FindFile(inv.No, customer.SellToCustomerNo));
                            row.CreateCell(25).SetCellValue(inv.InvoiceSent != null ? inv.InvoiceSent.ToString("MM/dd/yyyy hh:mm tt") : "");
                            row.CreateCell(26).SetCellValue(inv.OriginalOrder != null ? inv.OriginalOrder.OrderDate : "");
                            row.CreateCell(27).SetCellValue(inv.FileName);
                            row.CreateCell(28).SetCellValue(line.ValidationError);
                            row.CreateCell(29).SetCellValue(inverrors.ToString());
                            row.CreateCell(30).SetCellValue(custerrors.ToString());

                            HSSFFont hssfFont = (HSSFFont)workbook.CreateFont();
                            hssfFont.FontHeightInPoints = 12;
                            hssfFont.FontName = "Arial";

                            HSSFCellStyle hssfStyle = (HSSFCellStyle)workbook.CreateCellStyle();
                            hssfStyle.SetFont(hssfFont);

                            hssfStyle.Alignment = HorizontalAlignment.Left;
                            hssfStyle.BorderBottom = BorderStyle.Thin;
                            hssfStyle.BorderLeft = BorderStyle.Thin;
                            hssfStyle.BorderRight = BorderStyle.Thin;
                            hssfStyle.BorderTop = BorderStyle.Thin;

                            if ((percent > 0.0M && percent < 0.05M) || (percent > -0.05M && percent < -0.0M))
                            {
                                hssfStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
                                hssfStyle.FillPattern = FillPattern.SolidForeground;
                            }
                            else if (percent > 0.05M || percent < -0.05M)
                            {
                                hssfStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Red.Index;
                                hssfStyle.FillPattern = FillPattern.SolidForeground;
                            }

                            foreach (ICell cell in row.Cells)
                                cell.CellStyle = hssfStyle;

                            rowNo++;
                        }
                        catch (Exception ex)
                        {
                            Constants.ApplicationErrors.Add(new CodeError(ex, "SendReport", "AddWorksheet", new SqlCommand(inv.No)));
                        }
                    }
                }
                LineNumbersUsed = new List<int>();
            }

            for (int i = 0; i < 31; i++)
            {
                worksheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        private void WorksheetAddHeaderRow(ref ISheet worksheet, Customer customer)
        {
            IRow headerRow = worksheet.CreateRow(0);
            ICell headerCell = headerRow.CreateCell(1);
            headerCell.SetCellValue(customer.SellToCustomerNo + " Invoices");

            HSSFFont hssfFont = (HSSFFont)workbook.CreateFont();
            hssfFont.FontHeightInPoints = 18;
            hssfFont.FontName = "Times New Roman";
            hssfFont.IsBold = true;

            HSSFCellStyle hssfStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            hssfStyle.SetFont(hssfFont);

            headerCell.CellStyle = hssfStyle;

            worksheet.Header.Left = HSSFHeader.Page;
            worksheet.Header.Center = customer.SellToCustomerNo + " Invoices";

            //header
            IRow hRow = worksheet.CreateRow(1);
            hRow.CreateCell(0).SetCellValue("Invoice No.");
            hRow.CreateCell(1).SetCellValue("Invoice Date");
            hRow.CreateCell(2).SetCellValue("Cust. PO No.");
            hRow.CreateCell(3).SetCellValue("Inv. Ship Amount");
            hRow.CreateCell(4).SetCellValue("Inv. Total");
            hRow.CreateCell(5).SetCellValue("PO Total");
            hRow.CreateCell(6).SetCellValue("Inv. Line Number");
            hRow.CreateCell(7).SetCellValue("Inv. Ref. Line Number");
            hRow.CreateCell(8).SetCellValue("Inv. Part No.");
            hRow.CreateCell(9).SetCellValue("Inv. Description");
            hRow.CreateCell(10).SetCellValue("Inv. Unit Of Measure");
            hRow.CreateCell(11).SetCellValue("Inv. Quantity");
            hRow.CreateCell(12).SetCellValue("Inv. Unit Price");
            hRow.CreateCell(13).SetCellValue("Inv. Line Amount");
            hRow.CreateCell(14).SetCellValue("PO Line Number");
            hRow.CreateCell(15).SetCellValue("PO GSS Part No.");
            hRow.CreateCell(16).SetCellValue("PO Vendor Part No.");
            hRow.CreateCell(17).SetCellValue("PO Description");
            hRow.CreateCell(18).SetCellValue("PO Unit Of Measure");
            hRow.CreateCell(19).SetCellValue("PO Quantity");
            hRow.CreateCell(20).SetCellValue("PO Unit Price");
            hRow.CreateCell(21).SetCellValue("PO Line Amount");
            hRow.CreateCell(22).SetCellValue("Inv and PO Line Diff. (%)");
            hRow.CreateCell(23).SetCellValue("No. of Attempts");
            hRow.CreateCell(24).SetCellValue("Invoice Sent");
            hRow.CreateCell(25).SetCellValue("Date/Time Sent");
            hRow.CreateCell(26).SetCellValue("PO Date");
            hRow.CreateCell(27).SetCellValue("NR File Name");
            hRow.CreateCell(28).SetCellValue("Validation Errors");
            hRow.CreateCell(29).SetCellValue("Invoice App Errors");
            hRow.CreateCell(30).SetCellValue("Customer Error");

            HSSFFont hssfColFont = (HSSFFont)workbook.CreateFont();
            hssfColFont.FontHeightInPoints = 12;
            hssfColFont.FontName = "Arial";
            hssfColFont.IsBold = true;

            HSSFCellStyle hssfColStyle = (HSSFCellStyle)workbook.CreateCellStyle();
            hssfColStyle.SetFont(hssfColFont);

            var palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex(57, 188, 214, 238);

            hssfColStyle.FillForegroundColor = palette.GetColor(57).Indexed;
            hssfColStyle.FillPattern = FillPattern.SolidForeground;

            hssfColStyle.Alignment = HorizontalAlignment.Center;
            hssfColStyle.WrapText = true;
            hssfColStyle.BorderBottom = BorderStyle.Thin;
            hssfColStyle.BorderLeft = BorderStyle.Thin;
            hssfColStyle.BorderRight = BorderStyle.Thin;
            hssfColStyle.BorderTop = BorderStyle.Thin;

            foreach (ICell cell in hRow.Cells)
                cell.CellStyle = hssfColStyle;
        }

        private void WriteToFile(string excelPath)
        {
            FileStream file = new FileStream(excelPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }

        private string Convert(string path)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(path);
            wb.SaveAs(Filename: path + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();

            return path + "x";
        }

        private void SendEmail(string excelPath, List<Customer> customers)
        {
            StringBuilder body = new StringBuilder();
            body.Append("Please review the attached Invoice Report for invoices that were sent today, " + DateTime.Now.ToString(@"MM/dd/yyyy a\t hh:mm tt") + "."
                + "<br /><br />1.  \"Inv. Ship Amount\" is the shipping amount in the Invoice Summary (Header) of the invoice."
                + "<br />2.  \"Inv Ref. Line Number\" is the PO line number that the invoice line is referencing."
                + "<br />3.  \"Inv. and PO Line Diff (%)\" is the difference, in percent, between the Unit Price on the Invoice and the Unit Price on the PO."
                + "<br />4.  If a row is highlighted Yellow, the difference between the Invoice Line Unit Price and the PO Line Unit Price is between 0% and 5%."
                + "<br />5.  If a row is highlighted Red, the difference between the Invoice Line Unit Price and the PO Line Unit Price is over 5%.");

            StringBuilder invoiceMsg = new StringBuilder();
            StringBuilder customerMsg = new StringBuilder();

            foreach (Customer customer in customers)
            {
                foreach (InvoiceHeader invoice in customer.Invoices.FindAll(i => i.NeedsReview || i.Errors.Count > 0))
                {
                    StringBuilder validationErrors = new StringBuilder();
                    Database.UpdateInvoice(invoice.No, invoice.Your_Reference, 0);

                    decimal invLineTotal = 0.0M;
                    decimal shipTotal = 0.0M;
                    foreach (InvoiceLine line in invoice.SalesInvLines)
                    {
                        if (line.Type == 2)
                            invLineTotal += (line.Quantity * line.Unit_Price);
                        else
                            shipTotal += (line.Quantity * line.Unit_Price);
                        if (line.ValidationError.Length > 0)
                            validationErrors.Append($"{(validationErrors.Length>0 ? "<br>" : "")}Line {line.Line_No}: {line.ValidationError}");
                    }

                    string partial = "";
                    if (invoice.SalesInvLines.Where(i => i.Type == 2 && i.Quantity > 0).Count() < invoice.SalesInvLines.Where(v => v.Type == 2).Count())
                        partial = "Partial";
                    else
                        partial = "Full";

                    StringBuilder inverrors = new StringBuilder();
                    foreach (CodeError err in invoice.Errors)
                        inverrors.Append($"{(inverrors.Length > 0 ? "<br>" : "")}{err.Error.Message}");

                    invoiceMsg.Append($"<tr><td>{invoice.No}</td>" +
                        $"<td>{invoice.Your_Reference}</td>" +
                        $"<td>{partial}</td>" +
                        $"<td>{shipTotal.ToString("G29")}</td>" +
                        $"<td>{invLineTotal.ToString("G29")}</td>" +
                        $"<td>{(invoice.OriginalOrder != null ? invoice.OriginalOrder.POTotal.ToString("G29") : "N/A")}</td>" +
                        $"<td>{(invoice.cXMLInvoice == null ? "No" : "Yes")}</td>" +
                        $"<td>{validationErrors}</td>" +
                        $"<td>{inverrors.ToString()}</td>" +
                        $"</tr>");
                }

                StringBuilder custerrors = new StringBuilder();
                foreach (CodeError err in customer.Errors)
                    custerrors.Append($"{(custerrors.Length > 0 ? ";" : "")}{err.Error.Message}");

                if (custerrors.ToString().Length > 0)
                    customerMsg.Append($"<tr><td>{customer.SellToCustomerNo}</td><td>{custerrors}</td></tr>");
            }
            
            if (invoiceMsg.ToString().Length > 0)
            {
                string tablebody = "<br><br>The following invoice(s) were either marked for manual review or errored out:<br><table border='1'><tbody>"
                    + "<tr><th>Invoice No.</th><th>Customer PO #</th><th>Full/Partial</th><th>Shipping Total</th><th>Invoice Lines Total</th><th>Original PO Total</th><th>e-Document Created</th><th>Validation Error(s)</th><th>App Error</th></tr>" 
                    + $"{invoiceMsg.ToString()}</tbody></table>";
                body.Append(tablebody);
            }

            if (customerMsg.ToString().Length > 0)
            {
                string tablebody = "<br<br>Error(s) occurred for the following customer(s):<br><table border='1'><tbody><tr><th>Customer No.</th><th>Error</th></tr>" +
                    $"{customerMsg.ToString()}</tbody></table>";
                body.Append(tablebody);
            }

            if (Constants.PublishProfile == 0 || Constants.PublishProfile == 1)
                Email.SendEmail(body.ToString(), "Electronic Invoices Sent", "", "kpratt@govsci.com;jfrancis@govsci.com;narshad@govsci.com;tdonovan@govsci.com", "gss-it-development@govsci.com;gpfaffe@govsci.com", "", excelPath, true);
            else
                Email.SendEmail(body.ToString(), "Electronic Invoices Sent", "", "zlingelbach@govsci.com", "", "", excelPath, true);
        }
    }
}
