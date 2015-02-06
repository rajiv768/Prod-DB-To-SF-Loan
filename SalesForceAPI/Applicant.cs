using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesForceAPI
{
    class Applicant
    {
        public string DateOfBirth { get; set; }
        public string LASLoanId { get; set; } 
        public string Email { get; set; }
        public string MobilePhone { get; set; }
        public string Street { get; set; }
        public string County { get; set; }
        public string PostalCode { get; set; }
        public string Landline { get; set; }
        public string WorkPhone { get; set; }
        public string LoanAmount { get; set; }
        public string LastPaymentDate { get; set; }
        public string AmountPaid { get; set; }
        public string BankAccount { get; set; }
        public string SortCode { get; set; }
    }

    class Applicant2
    {
        public string LAS_Loan_Id__c { get; set; }
        public string LAS_Customer_Id__c { get; set; }
        public string LAS_User_Name__c { get; set; }         
        public string Applicant__r { get; set; }
        public string Loan_Amount__c { get; set; }
        public string Loan_Days__c { get; set; }
        public string Monthly_Net_Income__c { get; set; }
        public string Title__c { get; set; }
        public string First_Name__c { get; set; }
        public string Middle_Name__c { get; set; }
        public string Last_Name__c { get; set; }
        public string Email__c { get; set; }
        public string Mobile__c { get; set; }
        public string Date_Of_Birth__c { get; set; }
        public string Bank_Account__c { get; set; }
        public string Sort_Code__c { get; set; }
        public string Password__c { get; set; }

        public string Residential_Status__c { get; set; }
        public string Marital_Status__c { get; set; }
        public string Postal_Code__c { get; set; }
        public string Street__c { get; set; }
        public string Flat_Number__c { get; set; }
        public string Town__c { get; set; }
        public string County__c { get; set; }
        public string Landline__c { get; set; }
        public string AppliedOn__c { get; set; }
        public string Payout_Date__c { get; set; }
        public string ApplicationFormCheck { get; set; }
        public string CallReportCheck { get; set; }
        public string CallValidateBankAndIDCheck { get; set; }


        public string CallValidateCardCheck { get; set; }
        public string ApplicationFormCheck_count { get; set; }
        public string CallReportCheck_count { get; set; }
        public string CallValidateBankAndIDCheck_count { get; set; }
        public string CallValidateCardCheck_count { get; set; }
    }
    
            //Configuration config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
            //config.AppSettings.Settings.Add("last_LAS_Loan_Id__c", "1243");
            //config.Save(ConfigurationSaveMode.Minimal);

            //XmlDocument XmlDoc = new XmlDocument();
            //XmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            //foreach (XmlElement xElement in XmlDoc.DocumentElement)
            //{
            //    if (xElement.Name == "appSettings")
            //    {
            //        foreach (XmlNode xNode in xElement.ChildNodes)
            //        {
            //            if (xNode.Attributes[0].Value == "last_LAS_Loan_Id__c")
            //            {
            //                xNode.Attributes[1].Value = "1243";
            //            }
            //        }
            //    }
            //}
            //XmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
            
}
