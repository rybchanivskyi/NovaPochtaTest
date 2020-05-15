using System;
using System.Linq;
using Microsoft.Xrm.Tooling.Connector;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace NovaPochtaTest
{
    class Program
    {
        static void Main(string[] args)
        {
            const string url = "https://leadscrm.crm4.dynamics.com";
            const string userName = "admin@leadscrm.onmicrosoft.com";
            const string password = "Toropilka@02";
            const string fileName = "output.csv";

            string conn = $@"
            Url = {url};
            AuthType = Office365;
            UserName = {userName};
            Password = {password};
            RequireNewInstance = True";
            try
            {
                Executor(conn, fileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: \n" + ex.Message);
            }
            Console.WriteLine("Press any key to exit.");
            Console.ReadLine();
        }

        public static void WriteRecord(string companyId, string companyName, string contactId, string contactFullName, int contactType, string contactPhone, string contactEmail, StreamWriter file)
        {
            file.WriteLine(companyId + "," + companyName + "," + contactId + "," + contactFullName + "," + contactType + "," + (contactType == 1 ? contactEmail : contactPhone));
        }

        public static void Executor(string connectionStr, string fileName)
        {
            using (var svc = new CrmServiceClient(connectionStr))
            {
                Console.WriteLine("Try to receive data!");
                var res = (from contact in new OrganizationServiceContext(svc).CreateQuery("contact")
                           join account in new OrganizationServiceContext(svc).CreateQuery("account")
                           on contact["accountid"] equals account["accountid"]
                           where ((OptionSetValue)contact["statecode"]).Value == 0 && ((OptionSetValue)account["statecode"]).Value == 0
                           select new
                           {
                               CompanyId = account.Contains("accountid") ? account["accountid"].ToString() : "",
                               CompanyName = account.Contains("name") ? account["name"].ToString() : "",
                               ContactId = contact.Contains("contactid") ? contact["contactid"].ToString() : "",
                               ContactFullName = contact.Contains("fullname") ? contact["fullname"].ToString() : "",
                               ContactType = (OptionSetValue)contact["preferredcontactmethodcode"],
                               ContactPhone = contact.Contains("new_phone1") ? contact["new_phone1"].ToString() : "",
                               ContactEmail = contact.Contains("emailaddress1") ? contact["emailaddress1"].ToString() : "",
                           }).ToList();
                using (StreamWriter file = new StreamWriter(@fileName, false))
                {
                    Console.WriteLine($@"Start creating file: {fileName}");
                    foreach (var element in res)
                    {
                        WriteRecord(element.CompanyId, element.CompanyName, element.ContactId, element.ContactFullName,
                            element.ContactType.Value, element.ContactPhone, element.ContactEmail, file);
                    }
                }
                Console.WriteLine($@"All data was written to file - {Environment.CurrentDirectory}\{fileName}");
            }
        }
    }
}
