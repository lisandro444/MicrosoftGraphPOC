// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using graph_tutorial.TokenStorage;
using Microsoft.Identity.Client;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security.Claims;
using System.Web;

namespace graph_tutorial.Helpers
{
    public static class GraphHelper
    {
        // Load configuration settings from PrivateSettings.config
        private static string appId = ConfigurationManager.AppSettings["ida:AppId"];
        private static string appSecret = ConfigurationManager.AppSettings["ida:AppSecret"];
        private static string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
        private static List<string> graphScopes = new List<string>(ConfigurationManager.AppSettings["ida:AppScopes"].Split(' '));

        public static async Task<CachedUser> GetUserDetailsAsync(string accessToken)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                        return Task.FromResult(0);
                    }));

            var user = await graphClient.Me.Request()
                .Select(u => new
                {
                    u.DisplayName,
                    u.Mail,
                    u.UserPrincipalName
                })
                .GetAsync();

            return new CachedUser
            {
                Avatar = string.Empty,
                DisplayName = user.DisplayName,
                Email = string.IsNullOrEmpty(user.Mail) ?
                    user.UserPrincipalName : user.Mail
            };
        }

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            var events = await graphClient.Me.Events.Request()
                .Select("subject,organizer,start,end")
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            var message = new Message
            {
                Subject = "testing de mail",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = "The new cafeteria is open."
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

            await graphClient.Me
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();

            return events.CurrentPage;
        }

        private static GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var idClient = ConfidentialClientApplicationBuilder.Create(appId)
                            .WithRedirectUri(redirectUri)
                            .WithClientSecret(appSecret)
                            .Build();

                        var tokenStore = new SessionTokenStore(idClient.UserTokenCache,
                            HttpContext.Current, ClaimsPrincipal.Current);

                        var accounts = await idClient.GetAccountsAsync();

                        // By calling this here, the token can be refreshed
                        // if it's expired right before the Graph call is made
                        var result = await idClient.AcquireTokenSilent(graphScopes, accounts.FirstOrDefault())
                            .ExecuteAsync();

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));
        }
    }
}