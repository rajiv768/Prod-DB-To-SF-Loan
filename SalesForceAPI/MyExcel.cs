using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace SalesForceAPI
{
    class MyExcel
    {
        public static string DB_PATH = @"D:\Finolytics\5records.csv";
        public static List<Applicant> applicantList = new List<Applicant>();
        public static List<Applicant2> applicantList2 = new List<Applicant2>();
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static int lastRow = 0;

        public static void InitializeExcel()
        {
            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(DB_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets[1]; // Explict cast is not required here
            lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
        }

        public static List<Applicant> ReadMyExcel()
        {
            applicantList.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "K" + index.ToString()).Cells.Value;
                applicantList.Add(new Applicant
                {
                    DateOfBirth = MyValues.GetValue(1, 1)!= null ? MyValues.GetValue(1, 1).ToString() : "",
                    Email =  MyValues.GetValue(1, 2)!= null ? MyValues.GetValue(1, 2).ToString() : "",
                    MobilePhone =  MyValues.GetValue(1, 3)!= null ? MyValues.GetValue(1, 3).ToString() : "",
                    Street =  MyValues.GetValue(1, 4)!= null ? MyValues.GetValue(1, 4).ToString() : "",
                    County =  MyValues.GetValue(1, 5)!= null ? MyValues.GetValue(1, 5).ToString() : "",
                    PostalCode =  MyValues.GetValue(1, 6)!= null ? MyValues.GetValue(1, 6).ToString() : "",
                    Landline =  MyValues.GetValue(1, 7)!= null ? MyValues.GetValue(1, 7).ToString() : "",
                    WorkPhone =  MyValues.GetValue(1, 8)!= null ? MyValues.GetValue(1, 8).ToString() : "",
                    LoanAmount =  MyValues.GetValue(1, 9)!= null ? MyValues.GetValue(1, 9).ToString() : "",
                    LastPaymentDate =  MyValues.GetValue(1, 10)!= null ? MyValues.GetValue(1, 10).ToString() : "",
                    AmountPaid = MyValues.GetValue(1, 11) != null ? MyValues.GetValue(1, 11).ToString() : ""
                });
            }
            return applicantList;
        }

        public static List<Applicant2> ReadMyExcel2()
        {
            applicantList2.Clear();
            for (int index = 2; index <= lastRow; index++)
            {
                System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "AB" + index.ToString()).Cells.Value;
                applicantList2.Add(new Applicant2
                {
                    Applicant__r  = MyValues.GetValue(1, 1)!= null ? MyValues.GetValue(1, 1).ToString() : "",
                    Loan_Amount__c  = MyValues.GetValue(1, 2)!= null ? MyValues.GetValue(1, 2).ToString() : "",
                    Loan_Days__c  = MyValues.GetValue(1, 3)!= null ? MyValues.GetValue(1, 3).ToString() : "",
                    Monthly_Net_Income__c  = MyValues.GetValue(1, 4)!= null ? MyValues.GetValue(1, 4).ToString() : "",
                    Title__c  = MyValues.GetValue(1, 5)!= null ? MyValues.GetValue(1, 5).ToString() : "",
                    First_Name__c  = MyValues.GetValue(1, 6)!= null ? MyValues.GetValue(1, 6).ToString() : "",
                    Middle_Name__c  = MyValues.GetValue(1, 7)!= null ? MyValues.GetValue(1, 7).ToString() : "",
                    Last_Name__c  = MyValues.GetValue(1, 8)!= null ? MyValues.GetValue(1, 8).ToString() : "",
                    Email__c  = MyValues.GetValue(1, 9)!= null ? MyValues.GetValue(1, 9).ToString() : "",
                    Mobile__c  = MyValues.GetValue(1, 10)!= null ? MyValues.GetValue(1, 10).ToString() : "",
                    Date_Of_Birth__c  = MyValues.GetValue(1, 11)!= null ? MyValues.GetValue(1, 11).ToString() : "",
                    
                    Residential_Status__c  = MyValues.GetValue(1, 12)!= null ? MyValues.GetValue(1, 12).ToString() : "",
                    Marital_Status__c  = MyValues.GetValue(1, 13)!= null ? MyValues.GetValue(1, 13).ToString() : "",
                    Postal_Code__c  = MyValues.GetValue(1, 14)!= null ? MyValues.GetValue(1, 14).ToString() : "",
                    Street__c  = MyValues.GetValue(1, 15)!= null ? MyValues.GetValue(1, 15).ToString() : "",
                    County__c  = MyValues.GetValue(1, 16)!= null ? MyValues.GetValue(1, 16).ToString() : "",
                    Landline__c  = MyValues.GetValue(1, 17)!= null ? MyValues.GetValue(1, 17).ToString() : "",
                    AppliedOn__c = MyValues.GetValue(1, 18) != null ? MyValues.GetValue(1, 18).ToString() : "",
                    Payout_Date__c  = MyValues.GetValue(1, 19)!= null ? MyValues.GetValue(1, 19).ToString() : "",
                    ApplicationFormCheck  = MyValues.GetValue(1, 20)!= null ? MyValues.GetValue(1, 20).ToString() : "",
                    CallReportCheck  = MyValues.GetValue(1, 21)!= null ? MyValues.GetValue(1, 21).ToString() : "",
                    CallValidateBankAndIDCheck  = MyValues.GetValue(1, 22)!= null ? MyValues.GetValue(1, 22).ToString() : "",
                    
                    
                    CallValidateCardCheck  = MyValues.GetValue(1, 23)!= null ? MyValues.GetValue(1, 23).ToString() : "",
                    ApplicationFormCheck_count  = MyValues.GetValue(1, 24)!= null ? MyValues.GetValue(1, 24).ToString() : "",
                    CallReportCheck_count  = MyValues.GetValue(1, 25)!= null ? MyValues.GetValue(1, 25).ToString() : "",
                    CallValidateBankAndIDCheck_count  = MyValues.GetValue(1, 26)!= null ? MyValues.GetValue(1, 26).ToString() : "",
                    CallValidateCardCheck_count  = MyValues.GetValue(1, 27)!= null ? MyValues.GetValue(1, 27).ToString() : ""
                });
            }
            return applicantList2;
        }

        public static List<Applicant2> ReadSQL(SqlDataReader myReader)
        {
            applicantList2.Clear();
            while (myReader.Read())
            {
                //System.Array MyValues = (System.Array)MySheet.get_Range("A" + index.ToString(), "K" + index.ToString()).Cells.Value;
                applicantList2.Add(new Applicant2
                {
                    LAS_Loan_Id__c = myReader["LAS_Loan_Id__c"] != null ? myReader["LAS_Loan_Id__c"].ToString() : "",
                    LAS_Customer_Id__c = myReader["LAS_Customer_Id__c"] != null ? myReader["LAS_Customer_Id__c"].ToString() : "",
                    LAS_User_Name__c = myReader["LAS_User_Name__c"] != null ? myReader["LAS_User_Name__c"].ToString() : "",
                    Applicant__r = myReader["Applicant__r"] != null ? myReader["Applicant__r"].ToString() : "",
                    Loan_Amount__c = myReader["Loan_Amount__c"] != null ? myReader["Loan_Amount__c"].ToString() : "",
                    Loan_Days__c = myReader["Loan_Days__c"] != null ? myReader["Loan_Days__c"].ToString() : "",
                    Monthly_Net_Income__c = myReader["Monthly_Net_Income__c"] != null ? myReader["Monthly_Net_Income__c"].ToString() : "",
                    Title__c = myReader["Title__c"] != null ? myReader["Title__c"].ToString() : "",
                    First_Name__c = myReader["First_Name__c"] != null ? myReader["First_Name__c"].ToString() : "",
                    Middle_Name__c = myReader["Middle_Name__c"] != null ? myReader["Middle_Name__c"].ToString() : "",
                    Last_Name__c = myReader["Last_Name__c"] != null ? myReader["Last_Name__c"].ToString() : "",
                    Email__c = myReader["Email__c"] != null ? myReader["Email__c"].ToString() : "",
                    Mobile__c = myReader["Mobile__c"] != null ? myReader["Mobile__c"].ToString() : "",
                    Date_Of_Birth__c = myReader["Date_Of_Birth__c"] != null ? myReader["Date_Of_Birth__c"].ToString() : "",
                    Bank_Account__c = myReader["Bank_Account__c"] != null ? myReader["Bank_Account__c"].ToString() : "",
                    Sort_Code__c = myReader["Sort_Code__c"] != null ? myReader["Sort_Code__c"].ToString() : "",
                    Password__c = myReader["Password__c"] != null ? myReader["Password__c"].ToString() : "",

                    Residential_Status__c = myReader["Residential_Status__c"] != null ? myReader["Residential_Status__c"].ToString() : "",
                    Marital_Status__c = myReader["Marital_Status__c"] != null ? myReader["Marital_Status__c"].ToString() : "",
                    Postal_Code__c = myReader["Postal_Code__c"] != null ? myReader["Postal_Code__c"].ToString() : "",
                    Street__c = myReader["Street__c"] != null ? myReader["Street__c"].ToString() : "",
                    Flat_Number__c = myReader["Flat_Number__c"] != null ? myReader["Flat_Number__c"].ToString() : "",
                    Town__c = myReader["Town__c"] != null ? myReader["Town__c"].ToString() : "",
                    County__c = myReader["County__c"] != null ? myReader["County__c"].ToString() : "",
                    Landline__c = myReader["Landline__c"] != null ? myReader["Landline__c"].ToString() : "",
                    AppliedOn__c = myReader["AppliedOn"] != null ? myReader["AppliedOn"].ToString() : "",
                    Payout_Date__c = myReader["Payout_Date__c"] != null ? myReader["Payout_Date__c"].ToString() : "",
                    ApplicationFormCheck = myReader["ApplicationFormCheck"] != null ? myReader["ApplicationFormCheck"].ToString() : "",
                    CallReportCheck = myReader["CallReportCheck"] != null ? myReader["CallReportCheck"].ToString() : "",
                    CallValidateBankAndIDCheck = myReader["CallValidateBankAndIDCheck"] != null ? myReader["CallValidateBankAndIDCheck"].ToString() : "",

                    CallValidateCardCheck = myReader["CallValidateCardCheck"] != null ? myReader["CallValidateCardCheck"].ToString() : "",
                    ApplicationFormCheck_count = myReader["ApplicationFormCheck_count"] != null ? myReader["ApplicationFormCheck_count"].ToString() : "",
                    CallReportCheck_count = myReader["CallReportCheck_count"] != null ? myReader["CallReportCheck_count"].ToString() : "",
                    CallValidateBankAndIDCheck_count = myReader["CallValidateBankAndIDCheck_count"] != null ? myReader["CallValidateBankAndIDCheck_count"].ToString() : "",
                    CallValidateCardCheck_count = myReader["CallValidateCardCheck_count"] != null ? myReader["CallValidateCardCheck_count"].ToString() : ""
                });
            }
            return applicantList2;
        }

        public static void CloseExcel()
        {
            MyBook.Saved = true;
            MyApp.Quit();

        }
    }
}
