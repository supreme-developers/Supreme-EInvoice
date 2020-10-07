using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.Sql;
using System.Configuration;
using System.Diagnostics;
using System.Net;
using System.Text.RegularExpressions;
namespace EInvoice_Chevron
{
    public class EInvoice
    {

        //private int _DTID;
        //private String _FileName;
        string methodName = "";

    //public EInvoice(int DTID, String FileName)
    //{
    //    this._DTID = DTID;
    //    this._FileName = FileName;
        
    //}

    //#region Public Properties
    //    public int EmployeeID { get { return _DTID; } set { _DTID = value; } }
    //    public String EmployeeNumber { get { return _FileName; } set { _FileName = value; } }
    //#endregion

        public void Main(string FileName, int DTID)
        {
            
            //The following gets the method to call specific to the current customer as given in the field 'EDIMethod_Name' in the tblCustomers table.
            string cmdtext = "";
            SqlConnection Dbconn = new SqlConnection();
            Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString; 
            //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
            SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

            cmdtext = "select C.[EDIMethod_Name]" +
                      " from tblDelHeader H" +
                      " left join tblCustomers C on H.CustomerID = C.CustomerId" +
                      " where [Delivery Ticket ID] = @DelTicket";
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = cmdtext;
            cmd.Parameters.Add(new SqlParameter("@DelTicket", DTID));
            Dbconn.Open();
            // SqlDataReader dr = cmd.ExecuteReader();
            try
            {
                methodName = cmd.ExecuteScalar().ToString();
                Type thisType = this.GetType();
                object[] arguments = {FileName, DTID};
                MethodInfo theMethod = thisType.GetMethod(methodName);
                theMethod.Invoke(this, arguments);
            }
            catch (Exception excep)
            {
                string exception = excep.Message;
            }
            finally
            {
                Dbconn.Close();
                Dbconn.Dispose();
                cmd.Dispose();
                GC.Collect();

            }
        }
        public void Chevron_EInvoice(string FileName, int DTID)
        {
            Excel.Application xl = null;
            Excel._Workbook wb = null;
            Excel._Worksheet sheet = null;

            //VBIDE.VBComponent module = null;
            bool SaveChanges = false;
            try
            {
                if (File.Exists(FileName)) { File.Delete(FileName); }
                GC.Collect(); //System Garbage Collector

                // Create a new instance of Excel from scratch
                xl = new Excel.Application();
                xl.Application.DisplayAlerts = false;
                xl.Visible = false;

                // Add one workbook to the instance of Excel
                wb = (Excel._Workbook)(xl.Workbooks.Add(Missing.Value));

                // Get a reference to worksheet in the workbook
                sheet = (Excel._Worksheet)(wb.Sheets[1]);

                sheet.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                string cmdtext = "";

                SqlConnection Dbconn = new SqlConnection();
                Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
                SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

                cmdtext = "sp_AR_EDI_CreateEDIInvoice_Chevron";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = cmdtext;
                cmd.Parameters.Add(new SqlParameter("@DTID", DTID));
                Dbconn.Open();
                SqlDataReader dr = cmd.ExecuteReader();

                try
                {
                    int i = 0;
                    decimal invoiceTotal = new decimal();
                    while (dr.Read()) //sift through records and create Excel Sheet.
                    {
                        int detailrow = 45 + i;//changed from 43 -> 45 EG 9.13.2018
                        if (i == 0)
                        {
                            ((Excel.Range)sheet.Columns["D", Type.Missing]).ColumnWidth = 30;
                            ((Excel.Range)sheet.Columns["C", Type.Missing]).ColumnWidth = 30;
                            ((Excel.Range)sheet.Columns["F", Type.Missing]).ColumnWidth = 30;
                            sheet.Cells[1,1] = "SAP single Invoice Template - v3";
                          
                            sheet.Cells[7, 4] = dr["Invoice Number"].ToString();
                            sheet.Cells[8, 4] = dr["Delivery Ticket Number"].ToString();
                            sheet.Cells[9, 4] = dr["Inv Date"].ToString();
                            sheet.Cells[10, 4] = "USD";
                            sheet.Cells[13, 4] = dr["Buyer ID"].ToString();
                            sheet.Cells[14, 4] = dr["Representative"].ToString();
                            sheet.Cells[17, 4] = dr["Supreme EDI Number"].ToString();
                            sheet.Cells[18, 4] = "Elene Matherne";
                            sheet.Cells[19, 4] = dr["ContactEmail"].ToString();
                            sheet.Cells[22, 4] = dr["PO Number"].ToString();
                            sheet.Cells[23, 4] = dr["AFE"].ToString();
                            sheet.Cells[27, 4] = dr["EDI ID"].ToString();
                            sheet.Cells[29, 4] = dr["Rental Period Start Date"].ToString();
                            sheet.Cells[30, 4] = dr["Rental Period End Date"].ToString();
                            decimal num3 = Convert.ToDecimal(dr["State Tax"]) + Convert.ToDecimal(dr["Parish Tax"]);
                            string str1 = num3.ToString();
                            sheet.Cells[39, 4] = str1;
                        }

                        Excel.Range myRange = sheet.Application.get_Range("B43", "O43"); //<-----------not needed
                        sheet.Cells[detailrow, 2] = dr["Catalog"];
                        sheet.Cells[detailrow, 3] = dr["Item Description"];
                        sheet.Cells[detailrow, 5] = dr["Qty"];
                        sheet.Cells[detailrow, 6] = dr["UOM"];
                        sheet.Cells[detailrow, 7] = dr["Unit Price"];
                        sheet.Cells[detailrow, 8] = dr["Line Total"];
                        sheet.Cells[detailrow, 10] = dr["ItemizedDiscountDollars"];

                        invoiceTotal = invoiceTotal + (Convert.ToDecimal(dr["Qty"]) * Convert.ToDecimal(dr["Unit Price"]));
                        sheet.Cells[detailrow, 11] = Convert.ToDecimal(dr["Qty"]) * Convert.ToDecimal(dr["Unit Price"]);
                        sheet.Cells[detailrow, 13] = Convert.ToDecimal(dr["LineItemStateTax"]) + Convert.ToDecimal(dr["LineItemParishTax"]);
                        //sheet.Cells[detailrow, 11] = Convert.ToInt32(dr["Line Total"]) + Convert.ToInt32(dr["ItemizedDiscountDollars"]);
                        //sheet.Cells[detailrow, 13] = Convert.ToInt32(dr["LineItemStateTax"]) + Convert.ToInt32(dr["LineItemParishTax"]);
                        sheet.Cells[detailrow, 14] = "";
                        sheet.Cells[detailrow, 15] = "";
                        i++;
                    }
                    sheet.Cells[37, 4] = invoiceTotal;
                }
                catch (Exception exc)
                {
                    String msg;
                    msg = "Error: ";
                    msg = String.Concat(msg, exc.Message);
                    msg = String.Concat(msg, " Line: ");
                    msg = String.Concat(msg, exc.Source);
                    Console.WriteLine(msg);
                }

                finally
                {
                    Dbconn.Close();
                    Dbconn.Dispose();
                }

                xl.Visible = false;
                xl.UserControl = false;
                // Set a flag saying that all is well and it is ok to save our changes to a file.
                SaveChanges = true;
                //  Save the file to disk
                wb.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookNormal,
                          null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                          false, false, null, null, null);
            }
            catch (Exception err)
            {
                String msg;
                msg = "Error: ";
                msg = String.Concat(msg, err.Message);
                msg = String.Concat(msg, " Line: ");
                msg = String.Concat(msg, err.Source);
                Console.WriteLine(msg);
            }
            finally
            {

                try
                {
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.
                    xl.Visible = false;
                    xl.UserControl = false;
                    // Closes the document and does not give option to save.
                    wb.Close(SaveChanges, null, null);
                    xl.Workbooks.Close();
                }
                catch { }
                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.
                xl.Quit();
                if (sheet != null) { Marshal.ReleaseComObject(sheet); }
                if (wb != null) { Marshal.ReleaseComObject(wb); }
                if (xl != null) { Marshal.ReleaseComObject(xl); }

                //module = null;
                sheet = null;
                wb = null;
                xl = null;
                GC.Collect();
            }



            //((Excel.Range)sheet.Columns["D", Type.Missing]).ColumnWidth = 30;
            //((Excel.Range)sheet.Columns["C", Type.Missing]).ColumnWidth = 30;
            //((Excel.Range)sheet.Columns["F", Type.Missing]).ColumnWidth = 30;
            //sheet.Cells[35, 6] = "b2een SAP Template- Version 1.5";
            ////Header set up.
            //sheet.Cells[6, 4] = dr["Invoice Number"].ToString();
            //sheet.Cells[7, 4] = dr["Delivery Ticket Number"].ToString();
            //sheet.Cells[8, 4] = dr["Inv Date"].ToString();
            //sheet.Cells[9, 4] = "USD";
            //sheet.Cells[12, 4] = dr["Buyer ID"].ToString();
            //sheet.Cells[13, 4] = dr["Representative"].ToString();
            //sheet.Cells[16, 4] = dr["Supreme EDI Number"].ToString();
            //sheet.Cells[17, 4] = "Eric Gautreaux";
            //sheet.Cells[18, 4] = dr["ContactEmail"].ToString();
            //sheet.Cells[21, 4] = dr["PO Number"].ToString();
            //sheet.Cells[22, 4] = dr["AFE"].ToString();
            //sheet.Cells[25, 4] = dr["EDI ID"].ToString();
            //sheet.Cells[26, 4] = dr["WellNo"].ToString();
            //sheet.Cells[27, 4] = dr["Rental Period Start Date"].ToString();
            //sheet.Cells[28, 4] = dr["Rental Period End Date"].ToString();
            //sheet.Cells[31, 3] = "";
            //sheet.Cells[35, 4] = dr["Sub Total"].ToString();
            //sheet.Cells[37, 4] = Convert.ToInt32(dr["State Tax"]) + Convert.ToInt32(dr["Parish Tax"]);

            //------------------------------removed 9.13.2016 After update---------------------------------------//
            //sheet.Cells[6, 3] = "";
            //sheet.Cells[6, 4] = dr["Invoice Number"].ToString();
            //sheet.Cells[7, 4] = dr["Delivery Ticket Number"].ToString();
            //sheet.Cells[8, 4] = dr["Inv Date"].ToString();
            //sheet.Cells[9, 4] = "USD";
            //sheet.Cells[13, 4] = dr["Buyer ID"].ToString();
            //sheet.Cells[14, 4] = dr["Representative"].ToString();
            //sheet.Cells[17, 4] = dr["Supreme EDI Number"].ToString();
            //sheet.Cells[18, 4] = "Elene Matherne";
            //sheet.Cells[19, 4] = dr["ContactEmail"].ToString();
            //sheet.Cells[22, 4] = dr["PO Number"].ToString();
            //sheet.Cells[23, 4] = dr["AFE"].ToString();

            //sheet.Cells[27, 4] = dr["EDI ID"].ToString();

            //sheet.Cells[29, 4] = dr["Rental Period Start Date"].ToString();
            //sheet.Cells[30, 4] = dr["Rental Period End Date"].ToString();
            //------------------------------removed END 9.13.2016 After update---------------------------------------//
        }

        public void openInvoice(string FileName, int DTID)
        {
            //-----------------------------Get CustomerCodes-=----------------------------------------------------
            string CustomerNumber = "";
            string commandtext = "";
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
            //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
            SqlCommand command = new SqlCommand(commandtext, conn);

            commandtext = "select C.[Customer Number]" +
                      " from tblDelHeader H" +
                      " left join tblCustomers C on H.CustomerID = C.CustomerId" +
                      " where [Delivery Ticket ID] = @DelTicket";
            command.CommandType = System.Data.CommandType.Text;
            command.CommandText = commandtext;
            command.Parameters.Add(new SqlParameter("@DelTicket", DTID));
            conn.Open();

            try
            {
                CustomerNumber = command.ExecuteScalar().ToString();


            }
            catch (Exception excep)
            {
                string exception = excep.Message;
            }
            finally
            {
                conn.Close();
                conn.Dispose();
                command.Dispose();
                GC.Collect();

            }
            //---------------------------------------------------------------------------------------------------

            Excel.Application xl = null;
            Excel._Workbook wb = null;
            Excel._Worksheet sheet = null;

            //VBIDE.VBComponent module = null;
            bool SaveChanges = false;
            try
            {
                if (File.Exists(FileName)) { File.Delete(FileName); }
                GC.Collect(); //System Garbage Collector

                // Create a new instance of Excel from scratch
                xl = new Excel.Application();
                xl.Application.DisplayAlerts = false;
                xl.Visible = false;

                // Add one workbook to the instance of Excel
                wb = (Excel._Workbook)(xl.Workbooks.Add(Missing.Value));

                // Get a reference to worksheet in the workbook
                sheet = (Excel._Worksheet)(wb.Sheets[1]);

                sheet.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                string cmdtext = "";

                SqlConnection Dbconn = new SqlConnection();
                Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
                //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
                SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

                cmdtext = "sp_AR_EDI_CreateEDIInvoice_OpenInvoice";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = cmdtext;
                cmd.Parameters.Add(new SqlParameter("@DTID", DTID));
                Dbconn.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                try
                {
                    
                    int i = 1;
                    while (dr.Read()) //sift through records and create Excel Sheet.
                    {
                        sheet.Cells[i, 1] = dr["Invoice Number"].ToString();  //Max 25 Char

                        //((Excel.Range)sheet.Cells[i, 2]).EntireColumn.NumberFormat = "yyyy-MM-dd"; 
                        //sheet.Cells[i, 2] = String.Format("{0:yyyy-MM-dd}", Convert.ToDateTime(dr["Inv Date"].ToString())); //yyyy-mm-dd

                        

                       // ((Excel.Range)sheet.Cells[i, 2]).EntireColumn.NumberFormat = "yyyy-MM-dd";
                       
                        //sheet.Cells[i, 2] = Convert.ToDateTime(dr["Inv Date"].ToString()).ToOADate();
                        ((Excel.Range)sheet.Cells[i, 2]).EntireColumn.NumberFormat = "@";
                        sheet.Cells[i, 2].Value2 = Convert.ToDateTime(dr["Inv Date"].ToString())
                                                   .ToString("yyyy-MM-dd");

                        //sheet.Cells[i, 2] = dr["invdate"].ToString();

                        int invdesccount = Math.Min(dr["Job type"].ToString().Length, 2000);

                        sheet.Cells[i, 3] = "";
                        sheet.Cells[i, 4] = "USD";
                        sheet.Cells[i, 5] = "";//Max 9 digits operator DUNS leave blank
                        sheet.Cells[i, 6] = dr["CustomerCode"];//Customer Code
                        sheet.Cells[i, 7] = "";//Operator Contact Email - Leave blank
                        sheet.Cells[i, 8] = "";//Postal or Zip code leave blank
                        sheet.Cells[i, 9] = "";//supplier contact email - leave blank
                        sheet.Cells[i, 10] = "";//supplier zip code - leave blank
                        sheet.Cells[i, 11] = dr["Qty"]; //six dec places
                        sheet.Cells[i, 12] = dr["measure"];
                        sheet.Cells[i, 13] = dr["Unit Price"];
                        sheet.Cells[i, 14] = "";//Discount Percent??
                        //Ensure item description doesn't exceed 2000 chars
                        int itmdesccount = Math.Min(dr["Item Description"].ToString().Length, 2000);
                        sheet.Cells[i, 15] = Regex.Replace(dr["Item Description"].ToString().Substring(0, itmdesccount), @"[^\w\.@-]", "");
                        //Product or Service Code
                        sheet.Cells[i, 16] = dr["Catalog"].ToString();

                        if (dr["Rental Period Start Date"].ToString() != "")
                        {
                            ((Excel.Range)sheet.Cells[i, 17]).EntireColumn.NumberFormat = "@";
                            sheet.Cells[i, 17].Value2 = Convert.ToDateTime(dr["Rental Period End Date"].ToString())
                                                       .ToString("yyyy-MM-dd");

                            //Service Date???
                            //((Excel.Range)sheet.Cells[i, 17]).EntireColumn.NumberFormat = "yyyy-MM-dd";
                        }

                        sheet.Cells[i, 18] = dr["PO Number"].ToString();
                        sheet.Cells[i, 19] = dr["Job Number"].ToString();
                        sheet.Cells[i, 20] = dr["AFE"].ToString();

                        //The following are labeled optional
                        sheet.Cells[i, 21] = ""; //cost center number not required.
                        sheet.Cells[i, 22] = dr["WellNo"];//LocationorWell
                        sheet.Cells[i, 23] = "";//FieldorLEase
                        if (Convert.ToDecimal(dr["LineItemStateTax"].ToString()) > 0 && Convert.ToDecimal(dr["Line Total"].ToString()) > 0)
                        {
                            //sheet.Cells[i, 24] = Decimal.Divide(Convert.ToDecimal(dr["LineItemStateTax"].ToString()), Convert.ToDecimal(dr["Line Total"].ToString())) * 100;  //Tax1 Perc
                            sheet.Cells[i, 24] = String.Format("{0:0.0000}", GetStateTaxRate(DTID));
                            sheet.Cells[i, 25] = "State";//Tax1 Type
                        }
                        else
                        {
                            sheet.Cells[i, 24] = ""; //Tax1 Perc
                            sheet.Cells[i, 25] = "";//Tax1 Type
                        }
                        sheet.Cells[i, 26] = "";//Tax2Perc
                        sheet.Cells[i, 27] = "";//Tax2 Type
                        
                        if (Convert.ToDecimal(dr["LineItemParishTax"].ToString()) > 0 && Convert.ToDecimal(dr["Line Total"].ToString()) > 0)
                        {
                            //sheet.Cells[i, 28] = Decimal.Divide(Convert.ToDecimal(dr["LineItemParishTax"].ToString()), Convert.ToDecimal(dr["Line Total"].ToString())) * 100;//Tax3Per
                            sheet.Cells[i, 28] = String.Format("{0:0.0000}", GetParishTaxRate(DTID));
                            sheet.Cells[i, 29] = "Parish";//Tax3Type
                        }

                        else
                        {
                            sheet.Cells[i, 28] = "";
                            sheet.Cells[i, 29] = "";//Tax3Type
                        }
                        sheet.Cells[i, 30] = "";//Tax4per
                        sheet.Cells[i, 31] = "";//Tax4 type
                        sheet.Cells[i, 32] = "";//operational Cat
                        sheet.Cells[i, 33] = "";//Operator Cod
                        ((Excel.Range)sheet.Cells[i, 41]).EntireColumn.NumberFormat = "@";
                        sheet.Cells[i, 41] = Convert.ToDateTime(dr["Rental Period Start Date"].ToString())
                                                           .ToString("yyyy-MM-dd");
                        if (CustomerNumber == "MCMORAN" || CustomerNumber == "MURPHY")
                        {
                            sheet.Cells[i, 34] = "";//Asset Number
                            sheet.Cells[i, 35] = dr["OrderedBy"];
                            sheet.Cells[i, 36] = "";//Transaction number
                            sheet.Cells[i, 37] = "";//Requisitioner LastName~FirstName~MiddleInitial
                            sheet.Cells[i, 38] = "";//ContractNumber
                            sheet.Cells[i, 39] = dr["Delivery Ticket Number"];//Field Ticket Number
                        }
                        if (CustomerNumber == "MARATHONTULSAOK" || CustomerNumber == "MARLAF" || CustomerNumber == "MAROK")
                        {
                            sheet.Cells[i, 22] = "";
                            sheet.Cells[i, 40] = "10";//column AN
                            //column Q
                            if (dr["Rental Period End Date"].ToString() != "")
                            {
                                ((Excel.Range)sheet.Cells[i, 17]).EntireColumn.NumberFormat = "@";
                                sheet.Cells[i, 17].Value2 = Convert.ToDateTime(dr["Rental Period End Date"].ToString())
                                                           .ToString("yyyy/MM/dd").Replace('/', '-');
                            }

                            sheet.Cells[i, 38] = dr["PO Number"].ToString(); //Column AL


                            if (CustomerNumber == "SWIFT")//changes made 10.30.2014 EG per Brandy Theriot
                            {
                                sheet.Cells[i, 40] = "1.1.1";
                                sheet.Cells[i, 14] = "20";
                                //column Q
                                if (dr["Rental Period End Date"].ToString() != "")
                                {
                                    ((Excel.Range)sheet.Cells[i, 17]).EntireColumn.NumberFormat = "@";
                                    sheet.Cells[i, 17].Value2 = Convert.ToDateTime(dr["Rental Period End Date"].ToString())
                                                               .ToString("yyyy/MM/dd").Replace('/', '-');
                                }

                            }
                            //sheet.Cells[i, 40] = dr["PO Number"].ToString();
                        }
                        i++;
                    }
                }
                catch (Exception exc)
                {
                    updateErrorTable("File Error after sp fired ------" + exc.Message);
                    updateErrorTable(exc.InnerException.Message);
                    String msg;
                    msg = "Error: ";
                    msg = String.Concat(msg, exc.Message);
                    msg = String.Concat(msg, " Line: ");
                    msg = String.Concat(msg, exc.Source);
                    Console.WriteLine(msg);
                }
                finally
                {
                    Dbconn.Close();
                    Dbconn.Dispose();
                }
                xl.Visible = false;
                xl.UserControl = false;
                // Set a flag saying that all is well and it is ok to save our changes to a file.
                //SaveChanges = true;
                //  Save the file to disk
                //DLLTest("before file create");
                //DLLTest(FileName);
          
                wb.SaveAs(FileName, Excel.XlFileFormat.xlCSV,
                          null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                          false, false, null, null, null);
                //DLLTest("after file create");
            }
            catch (Exception err)
            {
                updateErrorTable("File Error First sp------" + err.Message);
                updateErrorTable(err.InnerException.Message);
            }
            finally
            {
                try
                {
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.
                    xl.Visible = false;
                    xl.UserControl = false;
                    // Closes the document and does not give option to save.
                    // wb.Close(SaveChanges, null, null);
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(wb);
                    // if (xl != null) { Marshal.ReleaseComObject(xl); }

                    //module = null;
                    sheet = null;
                    wb = null;
                    xl = null;
                }
                catch (Exception exec)
                {

                }
                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.

                GC.Collect();
            }
        }

        public void OildexCustomers_EInvoice(string FileName, int DTID)
        {
            
            Excel.Application xl = null;
            Excel._Workbook wb = null;
            Excel._Worksheet sheet = null;

            //VBIDE.VBComponent module = null;
            bool SaveChanges = false;
            try
            {
                if (File.Exists(FileName)) { File.Delete(FileName); }
                GC.Collect(); //System Garbage Collector

                // Create a new instance of Excel from scratch
                xl = new Excel.Application();
                xl.Application.DisplayAlerts = false;
                xl.Visible = false;

                // Add one workbook to the instance of Excel
                wb = (Excel._Workbook)(xl.Workbooks.Add(Missing.Value));

                // Get a reference to worksheet in the workbook
                sheet = (Excel._Worksheet)(wb.Sheets[1]);

                sheet.Columns.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                sheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                string cmdtext = "";

                SqlConnection Dbconn = new SqlConnection();
                Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
                //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
                SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

                cmdtext = "sp_AR_EDI_CreateEDIInvoice_Oildex";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandText = cmdtext;
                cmd.Parameters.Add(new SqlParameter("@DTID", DTID));
                Dbconn.Open();
                SqlDataReader dr = cmd.ExecuteReader();
                try
                {
                    int i = 1;
                    sheet.Cells[i, 1] = "BuyerName";
                    sheet.Cells[i, 2] = "SellerId";
                    sheet.Cells[i, 3] = "SellerName";
                    sheet.Cells[i, 4] = "InvoiceNumber";
                    sheet.Cells[i, 5] = "PurchaseOrderNumber";
                    sheet.Cells[i, 6] = "InvoiceDescription";
                    sheet.Cells[i, 7] = "InvoiceDate";
                    sheet.Cells[i, 8] = "InvoiceAmountDue";
                    sheet.Cells[i, 9] = "DueDate";
                    sheet.Cells[i, 10] = "RemittanceAddressee";
                    sheet.Cells[i, 11] = "RemittanceStreet";
                    sheet.Cells[i, 12] = "RemittanceSuite";
                    sheet.Cells[i, 13] = "RemittanceCity";
                    sheet.Cells[i, 14] = "RemittanceState";
                    sheet.Cells[i, 15] = "RemittancePostalCode";
                    sheet.Cells[i, 16] = "ShipFromAddressee";
                    sheet.Cells[i, 17] = "ShipFromStreet";
                    sheet.Cells[i, 18] = "ShipFromSuite";
                    sheet.Cells[i, 19] = "ShipFromCity";
                    sheet.Cells[i, 20] = "ShipFromState";
                    sheet.Cells[i, 21] = "ShipFromPostalCode";
                    sheet.Cells[i, 22] = "ShipToAddressee";
                    sheet.Cells[i, 23] = "ShipToStreet";
                    sheet.Cells[i, 24] = "ShipToSuite";
                    sheet.Cells[i, 25] = "ShipToCity";
                    sheet.Cells[i, 26] = "ShipToState";
                    sheet.Cells[i, 27] = "ShipToPostalCode";
                    sheet.Cells[i, 28] = "AttentionTo";
                    sheet.Cells[i, 29] = "Quantity";
                    sheet.Cells[i, 30] = "UnitPrice";
                    sheet.Cells[i, 31] = "UnitsOfMeasure";
                    sheet.Cells[i, 32] = "AmountDue";
                    sheet.Cells[i, 33] = "Description";
                    sheet.Cells[i, 34] = "ProductServiceID";
                    sheet.Cells[i, 35] = "AFENo";
                    sheet.Cells[i, 36] = "WellName";
                    sheet.Cells[i, 37] = "TicketNo";
                    sheet.Cells[i, 38] = "ServiceDate1";
                    sheet.Cells[i, 39] = "ServiceDate2";
                    sheet.Cells[i, 40] = "InvoiceTaxExempt1";
                    sheet.Cells[i, 41] = "InvoiceTaxAmount1";
                    sheet.Cells[i, 42] = "InvoiceTaxType1";

                    i++;
                    
                    while (dr.Read()) //sift through records and create Excel Sheet.
                    {
                        if (i == 1)
                        {
                            
                        }
                            sheet.Cells[i, 1] = dr["BuyerName"].ToString();
                            sheet.Cells[i, 2] = dr["SellerID"].ToString(); 
                            
                            //((Excel.Range)sheet.Cells[i, 2]).EntireColumn.NumberFormat = "@";
                            //int invdesccount = Math.Min(dr["Job type"].ToString().Length, 2000);

                            sheet.Cells[i, 3] = dr["Company Name"].ToString().Replace(Convert.ToChar(34).ToString(),"");

                            if (!dr["Invoice Number"].Equals(DBNull.Value))
                                sheet.Cells[i, 4] = dr["Invoice Number"];
                            else
                                sheet.Cells[i, 4] = "NA";

                            if (!dr["PO Number"].Equals(DBNull.Value))
                                sheet.Cells[i, 5] = dr["PO Number"].ToString();
                            else
                                sheet.Cells[i, 5] = "NA";


                            if (!dr["Job Type"].Equals(DBNull.Value))
                                sheet.Cells[i, 6] = dr["Job Type"].ToString();
                            else
                                sheet.Cells[i, 6] = "No Description";
                           
                            if (dr["Inv Date"].Equals(DBNull.Value)) //Works like IsNull
                                sheet.Cells[i, 7] = "NA";
                            else
                                sheet.Cells[i, 7] = dr["Inv Date"];

                            sheet.Cells[i, 8] = dr["Total Cost"].ToString();//Invoice Total

                           
                            if (!dr["DueDate"].Equals(DBNull.Value))
                                sheet.Cells[i, 9] = dr["DueDate"];
                            else
                                sheet.Cells[i, 9] = "NA";

                            sheet.Cells[i, 10] = dr["RemOffice"].ToString();
                            sheet.Cells[i, 11] = dr["RemAdd"].ToString(); //REmittance STreet
                            sheet.Cells[i, 12] = "";
                            sheet.Cells[i, 13] = dr["RemCity"].ToString();
                            sheet.Cells[i, 14] = dr["RemState"].ToString();
                            sheet.Cells[i, 15] = dr["RemPostalCode"].ToString();

                            sheet.Cells[i, 16] = dr["Company Name"].ToString().Replace(Convert.ToChar(34).ToString(),"");
                            sheet.Cells[i, 17] = dr["Address1"].ToString();
                            sheet.Cells[i, 18] = "";
                            sheet.Cells[i, 19] = dr["City"].ToString();
                            sheet.Cells[i, 20] = dr["State"].ToString();
                            sheet.Cells[i, 21] = dr["PostalCode"].ToString();
                            sheet.Cells[i, 22] = dr["Invoice To"].ToString();
                            sheet.Cells[i, 23] = dr["InvoiceToAddress2"].ToString();
                            sheet.Cells[i, 24] = "";
                            sheet.Cells[i, 25] = dr["InvoiceToCity"].ToString();
                            sheet.Cells[i, 26] = dr["InvoiceToState"].ToString();
                            sheet.Cells[i, 27] = dr["InvoiceToZipCode"].ToString();
                            sheet.Cells[i, 28] = dr["AttentionTo"].ToString();
                            sheet.Cells[i, 29] = dr["Qty"].ToString();
                            sheet.Cells[i, 30] = dr["Unit Price"].ToString();
                            sheet.Cells[i, 31] = dr["UOM"].ToString();
                            sheet.Cells[i, 32] = dr["Line Total"].ToString();

                            int itmdesccount = Math.Min(dr["Item Description"].ToString().Length, 2000);
                            sheet.Cells[i, 33] = Regex.Replace(dr["Item Description"].ToString().Substring(0, itmdesccount), @"[^\w\.@-]", "");  

                            //sheet.Cells[i, 33] = dr["Item Description"].ToString();

                            sheet.Cells[i, 34] = dr["Item Number"].ToString(); //Product Service ID?

                            if (dr["AFE"].Equals(DBNull.Value)) //Works like IsNull
                                sheet.Cells[i, 35] = "999";
                            else
                                sheet.Cells[i, 35] = dr["AFE"].ToString();;

                            sheet.Cells[i, 36] = dr["WellNo"].ToString();

                            sheet.Cells[i, 37] = dr["Delivery Ticket Number"].ToString();//Ticket Number??

                            sheet.Cells[i, 38] = dr["Rental Period Start Date"].ToString();
                            sheet.Cells[i, 39] = dr["Rental Period End Date"].ToString();

                            sheet.Cells[i, 40] = dr["InvoiceTaxExempt"].ToString();
                            sheet.Cells[i, 41] = Convert.ToDecimal(dr["State Tax"].ToString()) + Convert.ToDecimal(dr["Parish Tax"].ToString());
                             sheet.Cells[i, 42] = "StateAndLocal";

                            //Ensure item description doesn't exceed 2000 chars
                           // int itmdesccount = Math.Min(dr["Item Description"].ToString().Length, 2000);    
                        i++;
                    }
                }
                catch (Exception exc)
                {
                    updateErrorTable(exc.Message);
                }
                finally
                {
                    Dbconn.Close();
                    Dbconn.Dispose();
                }
                xl.Visible = false;
                xl.UserControl = false;
                // Set a flag saying that all is well and it is ok to save our changes to a file.
                //SaveChanges = true;
                //  Save the file to disk
                wb.SaveAs(FileName, Excel.XlFileFormat.xlCSV,
                          null, null, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                          false, false, null, null, null);

                updateErrorTable(FileName);
                
            }
            catch (Exception err)
            {
                updateErrorTable(err.Message);
            }
            finally
            {
                try
                {
                    
                    // Repeat xl.Visible and xl.UserControl releases just to be sure
                    // we didn't error out ahead of time.
                    xl.Visible = false;
                    xl.UserControl = false;
                    // Closes the document and does not give option to save.
                    // wb.Close(SaveChanges, null, null);
                    xl.Quit();
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(wb);
                    

                    Marshal.FinalReleaseComObject(sheet);
                    Marshal.FinalReleaseComObject(wb);
                    if (xl != null) 
                    { 
                        Marshal.ReleaseComObject(xl);
                        Marshal.FinalReleaseComObject(xl);
                    }

                    //module = null;
                    sheet = null;
                    wb = null;
                    xl = null;
                   // SendFile(FileName);
                }
                catch (Exception exec)
                {

                }
                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.

                GC.Collect();
            }
        }

        //public static void SendFile(string filename)
        //{
        //    FileInfo fileInfo = new FileInfo(filename);
        //    FtpWebRequest reqFTP = (FtpWebRequest)FtpWebRequest.Create(new Uri("http://sftp.oildex.com" + "/" + Path.GetFileName(filename)));

        //    reqFTP.KeepAlive = false;

        //    // Specify the command to be executed.
        //    reqFTP.Method = WebRequestMethods.Ftp.UploadFile;

        //    // use binary 
        //    reqFTP.UseBinary = true;

        //    reqFTP.ContentLength = fileInfo.Length;

        //    // Buffer size set to 2kb
        //    const int buffLength = 2048;
        //    byte[] buff = new byte[buffLength];

        //    // Stream to which the file to be upload is written
        //    Stream strm = reqFTP.GetRequestStream();

        //    FileStream fs = fileInfo.OpenRead();

        //    // Read from the file stream 2kb at a time
        //    int cLen = fs.Read(buff, 0, buffLength);

        //    // Do a while till the stream ends
        //    while (cLen != 0)
        //    {
        //        // FTP Upload Stream
        //        strm.Write(buff, 0, cLen);
        //        cLen = fs.Read(buff, 0, buffLength);
        //    }

        //    // Close 
        //    strm.Close();
        //    fs.Close();
        //}

        //private static void SendFile(string FileName)
        //{
        //    FileStream rdr = new FileStream(FileName + ".csv", FileMode.Open);
        //    HttpWebRequest req = (HttpWebRequest)WebRequest.Create("http://sftp.oildex.com");
        //    HttpWebResponse resp;
        //    req.Method = "Post";
        //    req.Credentials = new NetworkCredential("x0179040nv1", "Supreme!", "PROD");

        //    req.ContentLength = rdr.Length;
        //    req.ContentType = "application/Excel";
        //    req.AllowWriteStreamBuffering = true;
        //    Stream reqStream = req.GetRequestStream();
        //    byte[] inData = new byte[rdr.Length];
        //    int bytesRead = rdr.Read(inData, 0, Convert.ToInt32(rdr.Length));


        //    reqStream.Write(inData, 0, Convert.ToInt32(rdr.Length));
        //    rdr.Close();


        //}

        //private static void SendFile(string FileName)
        //{
        //    string local_filename = Path.GetFileName(FileName);
        //    FtpWebRequest request = (FtpWebRequest)WebRequest.Create("http://sftp.oildex.com" + "/" + Path.GetFileName(FileName));

        //    //hold ftp credentials in DB somewhere.

        //    request.Method = WebRequestMethods.Ftp.UploadFile;
        //    request.Credentials = new NetworkCredential("x0179040nv1", "Supreme!", "PROD");
        //    request.UsePassive = true;
        //    request.UseBinary = true;
        //    request.KeepAlive = false;

        //    request.Timeout = 20000;


        //    FileStream stream = File.OpenRead(FileName);
        //    byte[] buffer = new byte[stream.Length];
        //    stream.Read(buffer, 0, buffer.Length);
        //    stream.Close();
        //    try
        //    {

        //        Stream reqStream = request.GetRequestStream();
        //        reqStream.Write(buffer, 0, buffer.Length);
        //        reqStream.Close();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine(ex);
        //    }
        //}



        //public decimal GetStateTaxRate(string State)
        //{
        //    string cmdtext = "";
        //    decimal rate = 0;
        //    SqlConnection Dbconn = new SqlConnection();
        //    Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
        //    SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

        //    cmdtext = "select [State Tax Rate]" +
        //              " from tblStateRate" +
        //              " where State = @State";
        //    cmd.CommandType = System.Data.CommandType.Text;
        //    cmd.CommandText = cmdtext;
        //    cmd.Parameters.Add(new SqlParameter("@State", State));
        //    Dbconn.Open();
        //    // SqlDataReader dr = cmd.ExecuteReader();
        //    try
        //    {
        //        rate = Convert.ToDecimal(cmd.ExecuteScalar().ToString()) * 100;

        //    }
        //    catch (Exception ex)
        //    {
        //        String msg;
        //        msg = "Error: ";
        //        msg = String.Concat(msg, ex.Message);
        //        msg = String.Concat(msg, " Line: ");
        //        msg = String.Concat(msg, ex.Source);
        //        Console.WriteLine(msg);
        //    }
        //    finally
        //    {                
        //        Dbconn.Close();
        //        Dbconn.Dispose();
        //        cmd.Dispose();
        //        GC.Collect();
        //    }

        //    return rate;
        //}

        public decimal GetStateTaxRate(int DTID)
        {
            
            string cmdtext = "";
            decimal rate = 0;
            SqlConnection Dbconn = new SqlConnection();
            Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
            SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

            cmdtext = "select [State Tax Rate]" +
                      " from tbldelheader " +
                      " where [Delivery Ticket ID] = @DTID";
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = cmdtext;
            cmd.Parameters.Add(new SqlParameter("@DTID", DTID));
            Dbconn.Open();
            // SqlDataReader dr = cmd.ExecuteReader();
            try
            {
                rate = Convert.ToDecimal(cmd.ExecuteScalar().ToString()) * 100;
                
            }
            catch (Exception ex)
            {
                updateErrorTable(ex.Message);
                updateErrorTable(ex.InnerException.Message);
                String msg;
                msg = "Error: ";
                msg = String.Concat(msg, ex.Message);
                msg = String.Concat(msg, " Line: ");
                msg = String.Concat(msg, ex.Source);
                Console.WriteLine(msg);
            }
            finally
            {
                Dbconn.Close();
                Dbconn.Dispose();
                cmd.Dispose();
                GC.Collect();
            }

            return rate;
        }

        public decimal GetParishTaxRate(int DTID)
        {
            string cmdtext = "";
            decimal rate = 0;
            SqlConnection Dbconn = new SqlConnection();
            Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
            SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

            cmdtext = "select [Parish Tax Rate]" +
                      " from tbldelheader " +
                      " where [Delivery Ticket ID] = @DTID";
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = cmdtext;
            cmd.Parameters.Add(new SqlParameter("@DTID", DTID));
            Dbconn.Open();
            // SqlDataReader dr = cmd.ExecuteReader();
            try
            {
                rate = Convert.ToDecimal(cmd.ExecuteScalar().ToString()) * 100;

            }
            catch (Exception ex)
            {
                updateErrorTable(ex.Message);
                updateErrorTable(ex.InnerException.Message);
                String msg;
                msg = "Error: ";
                msg = String.Concat(msg, ex.Message);
                msg = String.Concat(msg, " Line: ");
                msg = String.Concat(msg, ex.Source);
                Console.WriteLine(msg);
            }
            finally
            {
                Dbconn.Close();
                Dbconn.Dispose();
                cmd.Dispose();
                GC.Collect();
            }

            return rate;
        }

        //public decimal GetParishTaxRate(string Parish)
        //{
        //    string cmdtext = "";
        //    decimal rate = 0;
        //    SqlConnection Dbconn = new SqlConnection();
        //    Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
        //    SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

        //    cmdtext = "select [Parish Tax Rate]" +
        //              " from tblParishRates" +
        //              " where Parish = @Parish";
        //    cmd.CommandType = System.Data.CommandType.Text;
        //    cmd.CommandText = cmdtext;
        //    cmd.Parameters.Add(new SqlParameter("@Parish", Parish));
        //    Dbconn.Open();
        //    // SqlDataReader dr = cmd.ExecuteReader();
        //    try
        //    {
        //        rate = Convert.ToDecimal(cmd.ExecuteScalar().ToString()) * 100;

        //    }
        //    catch (Exception ex)
        //    {
        //        String msg;
        //        msg = "Error: ";
        //        msg = String.Concat(msg, ex.Message);
        //        msg = String.Concat(msg, " Line: ");
        //        msg = String.Concat(msg, ex.Source);
        //        Console.WriteLine(msg);
        //    }
        //    finally
        //    {
        //        Dbconn.Close();
        //        Dbconn.Dispose();
        //        cmd.Dispose();
        //        GC.Collect();
        //    }

        //    return rate;
        //}

        public void updateErrorTable(string error)
        {

            string cmdtext = "";

            SqlConnection Dbconn = new SqlConnection();
            Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
            //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
            SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

            cmdtext = "Insert into tblWriteError(Error) Values(@Error)";
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = cmdtext;
            cmd.Parameters.Add(new SqlParameter("@Error", error));
            Dbconn.Open();
            cmd.ExecuteNonQuery();

            Dbconn.Close();
            Dbconn.Dispose();
            cmd.Dispose();
        }
        public void DLLTest(string data)
        {

            string cmdtext = "";

            SqlConnection Dbconn = new SqlConnection();
            Dbconn.ConnectionString = Properties.Settings.Default.SSIRentConnectionString;
            //Dbconn.ConnectionString = Properties.Settings.Default.RentTestConnectionString;
            SqlCommand cmd = new SqlCommand(cmdtext, Dbconn);

            cmdtext = "Insert into aaaEric(SQL, FromProcedure) Values(@SQL, @FromProcedure)";
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.CommandText = cmdtext;
            cmd.Parameters.Add(new SqlParameter("@SQL", "EDI Entry ->" + data));
            cmd.Parameters.Add(new SqlParameter("@FromProcedure", DateTime.Now.ToShortDateString()));
            Dbconn.Open();
            cmd.ExecuteNonQuery();

            Dbconn.Close();
            Dbconn.Dispose();
            cmd.Dispose();


        }
    }


}