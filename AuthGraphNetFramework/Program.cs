using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AuthGraphNetFramework
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                getUsersAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            Console.ReadLine();
        }

        public async static Task getUsersAsync()
        {
            var clientId = "3d05fbdd-713c-40d7-be36-3b2a7344d860";
            var tenantId = "629fd4e8-9d26-4da5-85ff-cc01ca1948c4";
            var clientSecret = "C-vI8s0VlB1TCTY~lq39y1dg5Q~tZ9kxX.";


            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);
            GraphServiceClient graphClient = new GraphServiceClient(authProvider);

            var groups = await graphClient.Groups.Request().Select(x => new { x.Id, x.DisplayName }).GetAsync();

            var message = new Message
            {
                Subject = "Testing Microsfot Graph",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "<h3>Test works</h3>"
                },
                ToRecipients = new List<Recipient>()
    {
        new Recipient
        {
            EmailAddress = new EmailAddress
            {
                Address = "lisandrorossi444@gmail.com"
            }
        }
    },
                CcRecipients = new List<Recipient>()
    {
        new Recipient
        {
            EmailAddress = new EmailAddress
            {
                Address = "admin@lisandrorossi444.onmicrosoft.com"
            }
        }
    }
            };

            var saveToSentItems = false;

            await graphClient.Users["admin@lisandrorossi444.onmicrosoft.com"]
                    .SendMail(message, saveToSentItems)
                    .Request()
                    .PostAsync();


            foreach (var group in groups)
            {
                Console.WriteLine($"{group.DisplayName}, {group.Id}");
            }
        }
    }
}
