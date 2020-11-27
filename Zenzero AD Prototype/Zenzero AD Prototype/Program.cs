using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Zenzero_AD_Prototype
{
    class Program
    {
        enum DataField
        {
            GivenName,
            Surname,
            DisplayName,
            MailNickname,
            Mail,
            OtherMails,
        }

        static string UserDataEntry(DataField dataField)
        {
            Console.Write(dataField + ": ");

            return Console.ReadLine();
        }

        static async Task Main(string[] args)
        {

            var dataFieldValues = new Dictionary<string, string>()
            {
                { "GivenName", "" },
                { "Surname", "" },
                { "DisplayName", "" },
                { "MailNickname", "" },
                { "Mail", "" },
                { "OtherMails", "" },
                { "Password", "" },
            };


            foreach (var field in Enum.GetValues(typeof(DataField)))
            {
                dataFieldValues[field.ToString()] = UserDataEntry((DataField)field);
                Console.WriteLine();
            }

            Console.WriteLine("Your randomly generated password: TestPassword99 (You will be prompted to change this the first time you log in to your account)");

            //Application and tenant GUIDs
            string clientId = "a56d8c3c-36bd-47a6-b236-7f7b3f7f2f23";
            string tenantId = "8963d1dc-8cb9-4ca8-9c4e-ed8464a768a9";
            //Secret string used by AAD to authenticate application, should not be shared
            string clientSecret = "V0Hs-KF-qysVrifm-JT-Uj_97ip55dms2N";

            //Access token generation
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .WithRedirectUri("http://localhost")
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            User user;

            if (dataFieldValues["OtherMails"] == "")
            {
                //User object containing end user's credentials
                user = new User
                {
                    AccountEnabled = true,
                    //forename
                    GivenName = dataFieldValues["GivenName"],
                    Surname = dataFieldValues["Surname"],
                    //'username' displayed in the 'users' section of the Zenzero Tenant in AAD
                    DisplayName = dataFieldValues["DisplayName"],
                    //Used as a backup for when 'UserPrincipalName' specified below is invalid, in which case the UPN is set to 'MailNickname'@zenzerotest123.onmicrosoft.com
                    MailNickname = dataFieldValues["MailNickname"],
                    //Email address used for logging the end user into AAD
                    UserPrincipalName = dataFieldValues["MailNickname"] + "@zenzerotest123.onmicrosoft.com",
                    //Primary email address
                    Mail = dataFieldValues["Mail"],
                    PasswordProfile = new PasswordProfile
                    {
                        ForceChangePasswordNextSignIn = true,
                        Password = "TestPassword99"
                    }
                };
            } else
            {
                //User object containing end user's credentials
                user = new User
                {
                    AccountEnabled = true,
                    //forename
                    GivenName = dataFieldValues["GivenName"],
                    Surname = dataFieldValues["Surname"],
                    //'username' displayed in the 'users' section of the Zenzero Tenant in AAD
                    DisplayName = dataFieldValues["DisplayName"],
                    //Used as a backup for when 'UserPrincipalName' specified below is invalid, in which case the UPN is set to 'MailNickname'@zenzerotest123.onmicrosoft.com
                    MailNickname = dataFieldValues["MailNickname"],
                    //Email address used for logging the end user into AAD
                    UserPrincipalName = dataFieldValues["MailNickname"] + "@zenzerotest123.onmicrosoft.com",
                    //Primary email address
                    Mail = dataFieldValues["Mail"],
                    OtherMails = dataFieldValues["OtherMails"].Split(" "),
                    PasswordProfile = new PasswordProfile
                    {
                        ForceChangePasswordNextSignIn = true,
                        Password = "TestPassword99"
                    }
                };
            }
            

            await graphClient.Users
                .Request()
                .AddAsync(user);

        }

    }
}
