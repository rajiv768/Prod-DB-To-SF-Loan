using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;
using System.Xml;
using System.Net;
using System.IO;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;

//using PrecisePrediction.Salesforce;

namespace SalesForceAPI
{
    
    class Program
    {
        private readonly static string connectionString = ConfigurationManager.AppSettings["sqlConn"];
        private readonly static SqlConnection myConnection = new SqlConnection(connectionString);
        private static SqlDataReader myReader = null;
        private static StreamWriter sw = null;

        //private readonly static string uname = ConfigurationManager.AppSettings["SfdcUser"];
        //private readonly static string sfdcPassword = ConfigurationManager.AppSettings["SfdcPassword"];
        //private readonly static string sfdcToken = ConfigurationManager.AppSettings["SfdcToken"];

        private readonly static string uname = "rajiv.arora@finolytics.com";
        private readonly static string sfdcPassword = "Donald0068";
        private readonly static string sfdcToken = "IRrdBOFEmkCmo55CeTq5TSJRm";

        private readonly static string logPath = ConfigurationManager.AppSettings["logPath"];

        private readonly static string last_LAS_Loan_Id__c = ConfigurationManager.AppSettings["last_LAS_Loan_Id__c"];

        //private readonly static string pw = sfdcPassword + sfdcToken;
        private readonly static string pw = sfdcPassword;

        private readonly static EnterpriseWsdl.SessionHeader header = new EnterpriseWsdl.SessionHeader();            

        private static string sessionId = string.Empty;
        private static string serverUrl = string.Empty;
        private static EndpointAddress apiAddr;

        private static List<Applicant> applc;
        private static List<Applicant2> applc2;
        private static int LasLoanId;

        //private static List<string> ids;

        //private static PrecisePrediction.Salesforce.EnterpriseWsdl.QueryResult apiResult;

        static void Main(string[] args)
        {
            //string query = "SELECT Id, LAS_Loan_Id__c, Applicant__r.First_Name__c, Applicant__r.Last_Name__c, Applicant__r.Blacklist__c, Applicant__r.Date_Of_Birth__c, " +
            //    "Sort_Code__c, Bank_Account__c, Loan_Amount__c, Email__c, Mobile__c, Outstanding_Balance__c, Our_Status__c, Original_Due_Date__c FROM Loan__c WHERE " +
            //    "Our_Status__c in ( 'Not Handled', 'Manual Check Approved') and Payout_Date__c = null";

            //apiAddr = new Salesforce().Login();
            //apiResult = sf.ReadObject(query);
            //if (apiResult.records != null)
            //{
            //    for (int i = 0; i < apiResult.records.Length; i++)
            //    {
            //        PrecisePrediction.Salesforce.EnterpriseWsdl.Loan__c loan = (PrecisePrediction.Salesforce.EnterpriseWsdl.Loan__c)apiResult.records[i];
            //    }
            //}

            //******************************************************************************************************************

            OpenLogFile();

            LogMessage(sw, "Started...");

            //apiAddr = LogIntoSalesforce();

            //using (EnterpriseWsdl.SoapClient apiClient = new EnterpriseWsdl.SoapClient("Soap", apiAddr))
            //{
            //string query = "SELECT LAS_Loan_Id__c  FROM Loan__c order by CreatedDate desc,LAS_Loan_Id__c desc limit 1";
            //EnterpriseWsdl.QueryResult apiResult;

            //apiClient.query(header, null, null, null, query, out apiResult);

            //if (apiResult.records != null)
            //{
            //    EnterpriseWsdl.Loan__c l = (EnterpriseWsdl.Loan__c)apiResult.records[0];
            //    LasLoanId = Convert.ToInt32(l.LAS_Loan_Id__c);
            //    //LasLoanId = 200008074;

            // READ DATA FROM EXCEL FILE & INSERT INTO SALESFORCE, USE THESE 3 METHODS
            OpenSQLConnection();
            GetCustomerInfoFromSQL();
            CloseSQLConnection();
            //}
            //}

            //// READ DATA FROM EXCEL FILE & INSERT INTO SALESFORCE
            //if (!string.IsNullOrEmpty(MyExcel.DB_PATH))
            //{
            //    MyExcel.InitializeExcel();
            //    applc2 = MyExcel.ReadMyExcel2();
            //    MyExcel.CloseExcel();

            //    // To SELECT, UPDATE & DELETE records from SALESFORCE, run this method
            //    if (applc2 != null)
            //    {
            //        CallEnterpriseApi();
            //    }
            //}

            LogMessage(sw, "Finished...");
            //Console.WriteLine("Please press enter key to exit...");
            //Console.ReadLine();

            sw.Close();
        }

        private static void OpenLogFile()
        {
            string path = @logPath + String.Format("{0:dd.MM.yyyy}", DateTime.Today) + ".txt";
            
            if (!File.Exists(path))
            {
                sw = File.CreateText(path);
                //// Create a file to write to. 
                //using (sw = File.CreateText(path))
                //{
                //    LogInformation(sw, applicant2);
                //}
            }
            else
            {
                sw = File.AppendText(path);
                //// Open a file to write to. 
                //using (StreamWriter sw = File.AppendText(path))
                //{
                //    LogInformation(sw, applicant2);
                //}
            }            
        }

        private static void LogApplicantInformation(StreamWriter sw, Applicant2 applicant2)
        {
            sw.WriteLine("Log time " + String.Format("{0:dd.MM.yyyy hh:mm:ss}", DateTime.Now));
            sw.WriteLine("App Id : " + applicant2.LAS_Loan_Id__c + " DOB : " + applicant2.Date_Of_Birth__c + " Email : " + applicant2.Email__c + " Mobile : " + applicant2.Mobile__c + " added successfully.");
            sw.WriteLine("");
        }

        private static void LogLoanInformation(StreamWriter sw, Applicant2 applicant2)
        {
            sw.WriteLine("Log time " + String.Format("{0:dd.MM.yyyy hh:mm:ss}", DateTime.Now));
            sw.WriteLine("Loan Amount : " + applicant2.Loan_Amount__c + " Loan Days : " + applicant2.Loan_Days__c + " added successfully.");
            sw.WriteLine("");
        }

        private static void LogErrorMessage(StreamWriter sw, string errorMessage)
        {
            sw.WriteLine("Log time " + String.Format("{0:dd.MM.yyyy hh:mm:ss}", DateTime.Now));
            sw.WriteLine("Error Message : " + errorMessage);
            sw.WriteLine("");
        }

        private static void LogMessage(StreamWriter sw, string message)
        {
            sw.WriteLine("Log time " + String.Format("{0:dd.MM.yyyy hh:mm:ss}", DateTime.Now));
            sw.WriteLine("Message : " + message);
            sw.WriteLine("");
        }

        private static string PhoneFormat(string phone)
        {
            if (!string.IsNullOrEmpty(phone) && phone.Length > 1)
            {
                if (phone.Substring(0, 2) == "44")
                    phone = "+44" + phone.Substring(2);
                else if (phone.Substring(0, 2) != "44")
                    phone = "+44" + phone.Substring(0);
            }
            //Console.WriteLine(phone);
            return phone;
        }

        private static string DateFormat(string date)
        {
            if (!string.IsNullOrEmpty(date))
            {
                DateTime dt = Convert.ToDateTime(date);
                date = String.Format("{0:dd/MM/yyyy}", dt);
            }
            //Console.WriteLine(date);
            return date;
        }

        private static void CallEnterpriseApi()
        {
            apiAddr = LogIntoSalesforce(); 

            using (EnterpriseWsdl.SoapClient apiClient = new EnterpriseWsdl.SoapClient("Soap", apiAddr))
            {
                InsertRecordsToSalesForce(apiClient);
            }
        }

        private static EndpointAddress LogIntoSalesforce()
        {
            using (EnterpriseWsdl.SoapClient loginClient = new EnterpriseWsdl.SoapClient())
            {
                EnterpriseWsdl.LoginResult result = loginClient.login(null, uname, pw);

                sessionId = result.sessionId;
                serverUrl = result.serverUrl;

                Console.WriteLine("Session ID: " + sessionId);
                Console.WriteLine("Server URL: " + serverUrl);
            }

            EndpointAddress apiAddr = new EndpointAddress(serverUrl);

            //EnterpriseWsdl.SessionHeader header = new EnterpriseWsdl.SessionHeader();
            header.sessionId = sessionId;
            return apiAddr;
        }

        private static void InsertRecordsToSalesForce(EnterpriseWsdl.SoapClient apiClient)
        {
            //ids = new List<string>();
            int i = 0;

            foreach (var applicant2 in applc2)
            {
                i += 1;
                Console.WriteLine("Processing ..." + i);

                ////build up soql query
                //string query = "SELECT Applicant__c FROM Loan__c where LAS_Loan_Id__c = '" + applicant2.LAS_Loan_Id__c + "'";
                //EnterpriseWsdl.QueryResult apiResult;

                //apiClient.query(header, null, null, null, query, out apiResult);

                //if (apiResult.records == null)
                //{
                try
                {
                    string query = "SELECT Id FROM Applicant__c where LAS_Customer_Id__c = '" + applicant2.Applicant__r + "'";
                    EnterpriseWsdl.QueryResult apiResult;

                    apiClient.query(header, null, null, null, query, out apiResult);

                    if (apiResult.records == null)
                    {
                        InsertApplicantToSalesForce(apiClient, applicant2);
                    }
                    else
                    {
                        EnterpriseWsdl.Applicant__c applicant = (EnterpriseWsdl.Applicant__c)apiResult.records[0];
                        InsertLoanToSalesforce(apiClient, applicant2, applicant.Id);
                    }

                    InsertStatusToSQL(applicant2);
                }
                catch (Exception e)
                {
                    LogErrorMessage(sw, applicant2.Applicant__r + " have some invalid value!");
                }
                    
                //}
            }
        }

        // Manually deleting one or many records from SALESFORCE
        private static void DeleteMassRecordsFromSalesForce(EnterpriseWsdl.SoapClient apiClient, string[] ids)
        {
            Console.WriteLine("Deleting mass records ...");

            EnterpriseWsdl.DeleteResult[] deleteResults;

            EnterpriseWsdl.LimitInfo[] limitInfo;

            apiClient.delete(header, null, null, null, null, null, null, null, null, null, ids, out limitInfo, out deleteResults);
            EnterpriseWsdl.DeleteResult deleteResult = deleteResults[0];

            if (deleteResult.success)
            {
                Console.WriteLine("Mass records candidate " + deleteResult.id + " deleted succesfully.");
            }
            else
            {
                Console.WriteLine("Delete failed");
            }
        }

        //private static void InsertApplicantsTosalesForce(EnterpriseWsdl.SoapClient apiClient)
        private static void InsertApplicantToSalesForce(EnterpriseWsdl.SoapClient apiClient, Applicant2 applicant2)
        {
            EnterpriseWsdl.Applicant__c applic2 = new EnterpriseWsdl.Applicant__c();
                
            if (!String.IsNullOrEmpty(applicant2.LAS_Customer_Id__c))
                applic2.LAS_Customer_Id__c = applicant2.LAS_Customer_Id__c;

            if (!String.IsNullOrEmpty(applicant2.LAS_User_Name__c))
                applic2.LAS_User_Name__c = applicant2.LAS_User_Name__c;

            if (!String.IsNullOrEmpty(applicant2.Title__c))
                applic2.Title__c = applicant2.Title__c;

            if (!String.IsNullOrEmpty(applicant2.First_Name__c))
                applic2.First_Name__c = applicant2.First_Name__c.Replace("'", " ").Replace("-", "");

            if (!String.IsNullOrEmpty(applicant2.Middle_Name__c))
                applic2.Middle_Name__c = applicant2.Middle_Name__c.Replace("'", " ").Replace("-", "");

            if (!String.IsNullOrEmpty(applicant2.Last_Name__c))
                applic2.Last_Name__c = applicant2.Last_Name__c.Replace("'", " ").Replace("-", "");

            if (!String.IsNullOrEmpty(applicant2.Email__c))
                applic2.Email__c = applicant2.Email__c;

            if (!String.IsNullOrEmpty(applicant2.Mobile__c))
                applic2.Mobile__c = PhoneFormat(applicant2.Mobile__c);

            if (!String.IsNullOrEmpty(applicant2.Date_Of_Birth__c))
            {
                string dateofbirth = DateFormat(applicant2.Date_Of_Birth__c);
                //DateTime dt = DateTime.Parse(dateofbirth);
                applic2.Date_Of_Birth__c = DateTime.ParseExact(dateofbirth, "d/M/yyyy", CultureInfo.InvariantCulture);
                applic2.Date_Of_Birth__cSpecified = true;
            }

            if (!String.IsNullOrEmpty(applicant2.Residential_Status__c))
                applic2.Residential_Status__c = applicant2.Residential_Status__c;

            if (!String.IsNullOrEmpty(applicant2.Marital_Status__c))
                applic2.Marital_Status__c = applicant2.Marital_Status__c;

            if (!String.IsNullOrEmpty(applicant2.Postal_Code__c))
                applic2.Postal_Code__c = applicant2.Postal_Code__c;

            if (!String.IsNullOrEmpty(applicant2.Mobile__c))
                applic2.Street__c = applicant2.Street__c;

            if (!String.IsNullOrEmpty(applicant2.County__c))
                applic2.County__c = applicant2.County__c;

            if (!String.IsNullOrEmpty(applicant2.Landline__c))
                applic2.Landline__c = PhoneFormat(applicant2.Landline__c);

            if (!String.IsNullOrEmpty(applicant2.Flat_Number__c))
                applic2.Flat_Number__c = applicant2.Flat_Number__c;

            if (!String.IsNullOrEmpty(applicant2.Town__c))
                applic2.Town__c = applicant2.Town__c;

            if (!String.IsNullOrEmpty(applicant2.Password__c))
                applic2.Password__c = applicant2.Password__c;

            EnterpriseWsdl.sObject[] outreachArray = new EnterpriseWsdl.sObject[] { applic2 };

            EnterpriseWsdl.SaveResult[] createResult;
            EnterpriseWsdl.LimitInfo[] limitInfo;

            apiClient.create(header, null, null, null, null, null, null, null, null, null, null, outreachArray, out limitInfo, out createResult);

            if (createResult[0].success)
            {
                LogApplicantInformation(sw, applicant2);

                InsertLoanToSalesforce(apiClient, applicant2, createResult[0].id);
            }
            else
            {
                LogErrorMessage(sw, "Applicant insertion failed due to : " + createResult[0].errors[0].statusCode);
            }
        }

        private static void InsertLoanToSalesforce(EnterpriseWsdl.SoapClient apiClient, Applicant2 applicant2, string applicantId)
        {
            EnterpriseWsdl.Loan__c loan2 = new EnterpriseWsdl.Loan__c();

            loan2.Applicant__c = applicantId;

            loan2.LAS_Loan_Id__c = applicant2.LAS_Loan_Id__c;

            loan2.Bank_Account__c = applicant2.Bank_Account__c.Replace("-", "").Replace(" ", "").Trim();
            loan2.Sort_Code__c = applicant2.Sort_Code__c.Replace("-", "").Replace(" ", "").Trim();

            if (!String.IsNullOrEmpty(applicant2.Loan_Amount__c))
            {
                loan2.Loan_Amount__c = double.Parse(applicant2.Loan_Amount__c);
                loan2.Loan_Amount__cSpecified = true;
            }

            if (!String.IsNullOrEmpty(applicant2.Loan_Days__c))
            {
                loan2.Loan_Days__c = double.Parse(applicant2.Loan_Days__c);
                loan2.Loan_Days__cSpecified = true;
            }

            if (!String.IsNullOrEmpty(applicant2.Monthly_Net_Income__c))
            {
                loan2.Monthly_Net_Income__c = double.Parse(applicant2.Monthly_Net_Income__c);
                loan2.Monthly_Net_Income__cSpecified = true;
            }

            if (!String.IsNullOrEmpty(applicant2.AppliedOn__c))
            {
                string appliedOn = DateFormat(applicant2.AppliedOn__c);
                //DateTime dt2 = Convert.ToDateTime(appliedOn);
                loan2.AppliedOn__c = DateTime.ParseExact(appliedOn, "d/M/yyyy", CultureInfo.InvariantCulture);
                loan2.AppliedOn__cSpecified = true;
            }

            EnterpriseWsdl.sObject[] outreachArray2 = new EnterpriseWsdl.sObject[] { loan2 };

            EnterpriseWsdl.SaveResult[] createResult2;
            EnterpriseWsdl.LimitInfo[] limitInfo2;

            apiClient.create(header, null, null, null, null, null, null, null, null, null, null, outreachArray2, out limitInfo2, out createResult2);

            if (createResult2[0].success)
            {
                LogLoanInformation(sw, applicant2);
            }
            else
            {
                LogErrorMessage(sw, "Loan insertion failed due to : " + createResult2[0].errors[0].statusCode);
            }
        }

        private static void InsertRecordsTosalesForce2(EnterpriseWsdl.SoapClient apiClient)
        {
            foreach (var applicant in applc)
            {
                //Console.WriteLine("Adding new applicant ...");
                LogErrorMessage(sw, "Adding new applicant ...");

                EnterpriseWsdl.Applicant__c applic = new EnterpriseWsdl.Applicant__c();
                EnterpriseWsdl.Loan__c loan = new EnterpriseWsdl.Loan__c();

                string dateofbirth = applicant.DateOfBirth;
                DateTime dt = Convert.ToDateTime(dateofbirth);
                applic.Date_Of_Birth__c = dt;
                applic.Date_Of_Birth__cSpecified = true;

                string email = applicant.Email;
                applic.Email__c = email;

                string mobile = PhoneFormat(applicant.MobilePhone);
                applic.Mobile__c = mobile;

                string street = applicant.Street;
                applic.Street__c = street;

                string postalcode = applicant.PostalCode;
                applic.Postal_Code__c = postalcode;

                string county = applicant.County;
                applic.County__c = county;

                string landline = PhoneFormat(applicant.Landline);
                applic.Landline__c = landline;

                string workphone = PhoneFormat(applicant.WorkPhone);
                applic.Work_Phone__c = workphone;

                EnterpriseWsdl.sObject[] outreachArray = new EnterpriseWsdl.sObject[] { applic };

                EnterpriseWsdl.SaveResult[] createResult;
                EnterpriseWsdl.LimitInfo[] limitInfo;

                apiClient.create(header, null, null, null, null, null, null, null, null, null, null, outreachArray, out limitInfo, out createResult);

                if (createResult[0].success)
                {
                    //Console.WriteLine(createResult[0].id + " applicant added successfully.");
                    LogErrorMessage(sw, createResult[0].id + " applicant added successfully.");

                    loan.Applicant__c = createResult[0].id;

                    loan.LAS_Loan_Id__c = applicant.LASLoanId;
                    loan.Bank_Account__c = applicant.BankAccount.Replace("-", "").Replace(" ", "").Trim();
                    loan.Sort_Code__c = applicant.SortCode.Replace("-", "").Replace(" ", "").Trim();
                    
                    string loanamount = applicant.LoanAmount;
                    loan.Loan_Amount__c = double.Parse(loanamount);
                    loan.Loan_Amount__cSpecified = true;

                    string lastpaymentdate = applicant.LastPaymentDate;
                    DateTime lpd = Convert.ToDateTime(lastpaymentdate);
                    loan.Last_Payment_Date__c = lpd;
                    loan.Last_Payment_Date__cSpecified = true;

                    string amountpaid = applicant.AmountPaid;
                    loan.Amount_Paid__c = double.Parse(amountpaid);
                    loan.Amount_Paid__cSpecified = true;

                    EnterpriseWsdl.sObject[] outreachArray2 = new EnterpriseWsdl.sObject[] { loan };

                    EnterpriseWsdl.SaveResult[] createResult2;
                    EnterpriseWsdl.LimitInfo[] limitInfo2;

                    apiClient.create(header, null, null, null, null, null, null, null, null, null, null, outreachArray2, out limitInfo2, out createResult2);

                    if (createResult2[0].success)
                    {
                        //Console.WriteLine(createResult2[0].id + " Record added successfully.");
                        LogErrorMessage(sw, createResult2[0].id + " Record added successfully.");
                    }
                    else
                    {
                        //Console.WriteLine("Insertion failed.");
                        //Console.WriteLine(createResult2[0].errors[0].statusCode);
                        LogErrorMessage(sw, createResult2[0].errors[0].statusCode.ToString());
                    }
                    //}
                }
                else
                {
                    //Console.WriteLine("Insertion failed.");
                    //Console.WriteLine(createResult[0].errors[0].statusCode);
                    LogErrorMessage(sw, createResult[0].errors[0].statusCode.ToString());
                }
            }
        }

        private static void OpenSQLConnection()
        {
            try
            {
                myConnection.Open();
                LogMessage(sw, "Database opened");
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.ToString());
                LogErrorMessage(sw, e.ToString());
            }
        }

        private static void CloseSQLConnection()
        {
            try
            {
                myConnection.Close();
                LogMessage(sw, "Database closed");
            }
            catch (Exception e)
            {
                //Console.WriteLine(e.ToString());
                LogErrorMessage(sw, e.ToString());
            }
        }

        private static void GetCustomerInfoFromSQL()
        {
            try
            {
                string sqlCommand = "select * from (  " +
                "select id as LAS_Loan_Id__c, customer_id as Applicant__r, loan_amount as Loan_Amount__c, loan_time as Loan_Days__c, monthly_net_income as Monthly_Net_Income__c, title as Title__c, first_name as First_Name__c,  " +
                "middle_name as Middle_Name__c, last_name as Last_Name__c, email_address as Email__c,mobile as Mobile__c,date_of_birth as Date_Of_Birth__c, residential_status as Residential_Status__c,  " +
                "maritial_status as Marital_Status__c, postcode as Postal_Code__c, street_number as Street__c, county as County__c, land_line as Landline__c,appliedon as appliedon,null as Payout_Date__c,  " +
                "bank_account_number as Bank_Account__c, bank_sort_code as Sort_Code__c, flat_number as Flat_Number__c, town as Town__c,  " +
                "(select top 1 decision from  scorecard.ApplicationFormCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  ApplicationFormCheck,  " +
                "(select top 1 decision from  scorecard.CallReportCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallReportCheck,  " +
                "(select top 1 decision from  scorecard.CallValidateBankAndIDCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallValidateBankAndIDCheck,  " +
                "(select top 1 decision from  scorecard.CallValidateCardCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallValidateCardCheck,  " +
                "(select count(*) from  scorecard.ApplicationFormCheck_logg where applyloans_id =  appl.id) as  ApplicationFormCheck_count,  " +
                "(select count(*) from  scorecard.CallReportCheck_logg where applyloans_id =  appl.id) as  CallReportCheck_count,  " +
                "(select count(*) from  scorecard.CallValidateBankAndIDCheck_logg where applyloans_id =  appl.id) as  CallValidateBankAndIDCheck_count,  " +
                "(select count(*) from  scorecard.CallValidateCardCheck_logg where applyloans_id =  appl.id) as  CallValidateCardCheck_count,  " +
                "customer_id as LAS_Customer_Id__c, " +
                "(select max(username) from dbo.customer_logins cl2 where cl2.customer_id = appl.customer_id)  as LAS_User_Name__c, " +
                "(select max(auth_access) from dbo.customer_logins cl2 where cl2.customer_id = appl.customer_id)  as Password__c " +
                "from [dbo].[applyloans]  appl " +
                "where id in (select loan_id from  communications where activity_description = 'loan_issued'  )  " +
                "and not id in (select loan_id from  communications where activity_description = 'loan_to_SF'  )  " +
                "and not  email_address like '%softprodigy%'  " +
                "and not  email_address in ('gauravsoft@gmail.com','%finolytics%') " +
                "and not  mobile in ('07777777777') " +
                ") foo ";

                //string sqlCommand = "select * from ( " +
                //"select id as LAS_Loan_Id__c, customer_id as Applicant__r, loan_amount as Loan_Amount__c, loan_time as Loan_Days__c, monthly_net_income as Monthly_Net_Income__c, title as Title__c, first_name as First_Name__c, " +
                //"middle_name as Middle_Name__c, last_name as Last_Name__c, email_address as Email__c,mobile as Mobile__c,date_of_birth as Date_Of_Birth__c, residential_status as Residential_Status__c, " +
                //"maritial_status as Marital_Status__c, postcode as Postal_Code__c, street_number as Street__c, county as County__c, land_line as Landline__c,appliedon as appliedon,null as Payout_Date__c, " +
                //"bank_account_number as Bank_Account__c, bank_sort_code as Sort_Code__c, flat_number as Flat_Number__c, town as Town__c, " +
                //"(select top 1 decision from  scorecard.ApplicationFormCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  ApplicationFormCheck, " +
                //"(select top 1 decision from  scorecard.CallReportCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallReportCheck, " +
                //"(select top 1 decision from  scorecard.CallValidateBankAndIDCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallValidateBankAndIDCheck, " +
                //"(select top 1 decision from  scorecard.CallValidateCardCheck_logg where applyloans_id =  appl.id order by loggDate desc) as  CallValidateCardCheck, " +
                //"(select count(*) from  scorecard.ApplicationFormCheck_logg where applyloans_id =  appl.id) as  ApplicationFormCheck_count, " +
                //"(select count(*) from  scorecard.CallReportCheck_logg where applyloans_id =  appl.id) as  CallReportCheck_count, " +
                //"(select count(*) from  scorecard.CallValidateBankAndIDCheck_logg where applyloans_id =  appl.id) as  CallValidateBankAndIDCheck_count, " +
                //"(select count(*) from  scorecard.CallValidateCardCheck_logg where applyloans_id =  appl.id) as  CallValidateCardCheck_count, " +
                //"customer_id as LAS_Customer_Id__c," +
                //"(select max(username) from dbo.customer_logins cl2 where cl2.customer_id = appl.customer_id)  as LAS_User_Name__c " +
                //"from [dbo].[applyloans]  appl " +
                //") foo where (select top 1 activity_description from communications where activity_category  = 2 and loan_id = foo.LAS_Loan_Id__c ) = 'loan_issued' " +
                //"and not  email__c like '%softprodigy%' " +
                //"and not  email__C in ('gauravsoft@gmail.com') " +
                //"and LAS_Loan_Id__c > " + LasLoanId +
                //"and cast(appliedon as datetime) > dateadd(hour, -24, (select max(cast(appliedon as datetime))  FROM   dbo.applyloans))";


                SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
                myReader = myCommand.ExecuteReader();

                applc2 = MyExcel.ReadSQL(myReader);

                CloseSQLConnection();

                //if (applc2 != null)
                if (applc2.Count > 0)
                {
                    LogMessage(sw, "Calling API");
                    CallEnterpriseApi();
                }
            }
            catch (Exception e)
            {
                LogErrorMessage(sw, e.ToString());
            }
        }

        private static void InsertStatusToSQL(Applicant2 applicant2)
        {
            OpenSQLConnection();
            try
            {
                string sqlCommand = "INSERT INTO [dbo].[communications] " +
                "([customer_id], [account_id], [loan_id], [type_activity], [date_activity],  " +
                "[activity_category], [activity_description], [date_activity_backup]) " +
                "VALUES (" + applicant2.LAS_Customer_Id__c + ", null," + applicant2.LAS_Loan_Id__c + ",'2', CURRENT_TIMESTAMP, '12', 'loan_to_SF', null)";

                SqlCommand myCommand = new SqlCommand(sqlCommand, myConnection);
                //myCommand.Connection.Open();
                myCommand.ExecuteNonQuery();

                LogMessage(sw, applicant2.LAS_Loan_Id__c + " inserted in communication");
            }
            catch (Exception e)
            {
                LogErrorMessage(sw, e.ToString());
            }
        }
    }
}
