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

namespace SalesForceAPI
{
    
    class Program
    {
        private readonly static string connectionString = ConfigurationManager.AppSettings["sqlConn"];
        private readonly static SqlConnection myConnection = new SqlConnection(connectionString);
        private static SqlDataReader myReader = null;

        private readonly static string uname = ConfigurationManager.AppSettings["SfdcUser"];
        private readonly static string sfdcPassword = ConfigurationManager.AppSettings["SfdcPassword"];
        private readonly static string sfdcToken = ConfigurationManager.AppSettings["SfdcToken"];

        private readonly static string pw = sfdcPassword + sfdcToken;

        private static string sessionId = string.Empty;
        private static string serverUrl = string.Empty;


        static void Main(string[] args)
        {
            //string connectionString = ConfigurationManager.AppSettings["sqlConn"];
            //SqlConnection myConnection = new SqlConnection(connectionString);
            //SqlDataReader myReader = null;

            //OpenSQLConnection(myConnection);
            //GetCustomerInfoFromSQL(myConnection, myReader);
            //CloseSQLConnection(myConnection);

            OpenSQLConnection();
            GetCustomerInfoFromSQL();
            CloseSQLConnection();

            Console.ReadLine();

            //CallEnterpriseApi(myReader);
        }

        //private static void CallEnterpriseApi(SqlDataReader myReader)
        private static void CallEnterpriseApi()
        {           

            //string sessionId = string.Empty;
            //string serverUrl = string.Empty;

            using (EnterpriseWsdl.SoapClient loginClient = new EnterpriseWsdl.SoapClient())
            {
                //string uname = ConfigurationManager.AppSettings["SfdcUser"];
                //string sfdcPassword = ConfigurationManager.AppSettings["SfdcPassword"];
                //string sfdcToken = ConfigurationManager.AppSettings["SfdcToken"];

                //string pw = sfdcPassword + sfdcToken;

                EnterpriseWsdl.LoginResult result = loginClient.login(null, uname, pw);

                sessionId = result.sessionId;
                serverUrl = result.serverUrl;

                Console.WriteLine("Session ID: " + sessionId);
                Console.WriteLine("Server URL: " + serverUrl);
            }

            EndpointAddress apiAddr = new EndpointAddress(serverUrl);
            
            EnterpriseWsdl.SessionHeader header = new EnterpriseWsdl.SessionHeader();
            header.sessionId = sessionId;

            using (EnterpriseWsdl.SoapClient apiClient = new EnterpriseWsdl.SoapClient("Soap", apiAddr))
            {
                //SelectRecordsFromSalesForce(header, apiClient);

                //InsertRecordsTosalesForce(myReader, header, apiClient);

                //UpdateRecordsInSalesForce(header, apiClient);

                DeleteRecordsFromSalesForce(header, apiClient);
            }
        }

        private static void UpdateRecordsInSalesForce()
        {
                        //Console.ReadLine();

                        //Console.WriteLine("Updating record ...");

                        //EnterpriseWsdl.customer_infos__c updateCustomer = new EnterpriseWsdl.customer_infos__c();
                        //updateCustomer.first_name__c = "Rajiv";
                        //addCustomer.first_name__c = "Rajiv";


                        //Console.ReadLine();

                        //String[] ids = new String[] { Id };

                        //EnterpriseWsdl.DeleteResult[] deleteResults = 
        }

        private static void InsertRecordsTosalesForce(EnterpriseWsdl.SessionHeader header, EnterpriseWsdl.SoapClient apiClient)
        {
            Console.WriteLine("Adding new records ...");

            EnterpriseWsdl.customer_infos__c addCustomer = new EnterpriseWsdl.customer_infos__c();

            while (myReader.Read())
            {
                addCustomer.customerid__c = myReader["id"].ToString();
                addCustomer.account_number__c = myReader["account_number"].ToString();
                addCustomer.title__c = myReader["title"].ToString();

                addCustomer.first_name__c = myReader["first_name"].ToString();
                addCustomer.middle_name__c = myReader["middle_name"].ToString();
                addCustomer.last_name__c = myReader["last_name"].ToString();

                addCustomer.address_line_1__c = myReader["address_line_1"].ToString();
                addCustomer.address_line_2__c = myReader["address_line_2"].ToString();
                addCustomer.address_city__c = myReader["address_city"].ToString();
                addCustomer.adress_county__c = myReader["adress_county"].ToString();
                addCustomer.adress_postcode__c = myReader["adress_postcode"].ToString();
                addCustomer.adress_country__c = myReader["adress_country"].ToString();

                addCustomer.email__c = myReader["email"].ToString();
                addCustomer.new_email_address__c = myReader["new_email_address"].ToString();

                addCustomer.phone_landline__c = myReader["phone_landline"].ToString();
                addCustomer.new_phone_landline__c = myReader["new_phone_landline"].ToString();

                addCustomer.mobile__c = myReader["mobile"].ToString() != "" ? myReader["mobile"].ToString() : "70000000";
                addCustomer.new_mobile__c = myReader["new_mobile"].ToString() != "" ? myReader["new_mobile"].ToString() : "0";

                addCustomer.name_of_bank__c = myReader["name_of_bank"].ToString();
                addCustomer.name_of_new_bank__c = myReader["name_of_new_bank"].ToString();

                addCustomer.bank_sortcode__c = myReader["bank_sortcode"].ToString();
                addCustomer.new_bank_sortcode__c = myReader["new_bank_sortcode"].ToString();

                addCustomer.bank_account_number__c = myReader["bank_account_number"].ToString();
                addCustomer.new_account_number__c = myReader["new_account_number"].ToString();

                addCustomer.name_on_account__c = myReader["name_on_account"].ToString();
                addCustomer.name_on_new_account__c = myReader["name_on_new_account"].ToString();

                addCustomer.monthly_net_income__c = myReader["monthly_net_income"].ToString() != "" ? myReader["monthly_net_income"].ToString() : "0";

                EnterpriseWsdl.sObject[] outreachArray = new EnterpriseWsdl.sObject[] { addCustomer };
                //outreachArray[0].Id = myReader["id"].ToString();

                EnterpriseWsdl.SaveResult[] createResult;
                EnterpriseWsdl.LimitInfo[] limitInfo;

                apiClient.create(header, null, null, null, null, null, null, null, null, null, outreachArray, out limitInfo, out createResult);

                if (createResult[0].success)
                {
                    Console.WriteLine("Record added successfully.");
                }
                else
                {
                    Console.WriteLine("Insertion failed.");
                }
            }
        }

        private static void SelectRecordsFromSalesForce(EnterpriseWsdl.SessionHeader header, EnterpriseWsdl.SoapClient apiClient)
        {
            Console.WriteLine("Querying ...");

            //build up soql query
            string query = "SELECT first_name__c FROM customer_infos__c";
            EnterpriseWsdl.QueryResult apiResult;
            apiClient.query(header, null, null, null, query, out apiResult);

            //EnterpriseWsdl.DeleteResult[] deleteResults;
            //string[] ids = new string[]{};
            //EnterpriseWsdl.LimitInfo[] limitInfo;

            for (int i = 0; i < apiResult.records.Length; i++)
            {
                EnterpriseWsdl.customer_infos__c c = (EnterpriseWsdl.customer_infos__c)apiResult.records[i];

                Console.WriteLine("Candidate Id : {0}, Name : {1} & AccountNumber : {2} ", c.customerid__c, c.first_name__c, c.account_number__c);

                //ids[i] = c.first_name__c;

                //apiClient.delete(header, null, null, null, null, null, null, null, null, ids, out limitInfo, out deleteResults);
                //EnterpriseWsdl.DeleteResult deleteResult = deleteResults[i];

                //if (deleteResult.success)
                //{
                //    Console.WriteLine("Record ID " + deleteResult.id + " deleted succesfully.");
                //}
                //else
                //{
                //    Console.WriteLine("Delete failed");
                //}
            }

            //foreach (var id in ids)
            //{
            //    Console.WriteLine(id);
            //}

            Console.ReadLine();
        }

        private static void DeleteRecordsFromSalesForce(EnterpriseWsdl.SessionHeader header, EnterpriseWsdl.SoapClient apiClient)
        {
            Console.WriteLine("Deleting ...");

            EnterpriseWsdl.DeleteResult[] deleteResults;
            string[] ids = new string[] { "a012000001f7eLM" };
            EnterpriseWsdl.LimitInfo[] limitInfo;

            apiClient.delete(header, null, null, null, null, null, null, null, null, ids, out limitInfo, out deleteResults);
            EnterpriseWsdl.DeleteResult deleteResult = deleteResults[0];

            if (deleteResult.success)
            {
                Console.WriteLine("Record ID " + deleteResult.id + " deleted succesfully.");
            }
            else
            {
                Console.WriteLine("Delete failed");
            }

            //foreach (var id in ids)
            //{
            //    Console.WriteLine(id);
            //}

            Console.ReadLine();
        }

        //private static void OpenSQLConnection(SqlConnection myConnection)
        private static void OpenSQLConnection()
        {
            try
            {
                myConnection.Open();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        //private static void CloseSQLConnection(SqlConnection myConnection)
        private static void CloseSQLConnection()
        {
            try
            {
                myConnection.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        //private static void GetCustomerInfoFromSQL(SqlConnection myConnection, SqlDataReader myReader)
        private static void GetCustomerInfoFromSQL()
        {
            try
            {
                SqlCommand myCommand = new SqlCommand("select * from customer_infos order by id", myConnection);
                myReader = myCommand.ExecuteReader();

                //while (myReader.Read())
                //{

                //    Console.WriteLine("Candidate Id : {0}, Name : {1}, AccountNumber : {2} ", myReader["id"].ToString(), myReader["first_name"].ToString(), myReader["account_number"].ToString());
                //}
                
                //CallEnterpriseApi(myReader);
                CallEnterpriseApi();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
