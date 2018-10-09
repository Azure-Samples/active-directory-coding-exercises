/*
  The MIT License (MIT)

Copyright (c) 2015 Microsoft Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Batch
{
    class BatchDemo
    {
        public async Task RunAsync()
        {
            var tenantId = Config.GetConfig("tenantId");
            ConfidentialClientApplication daemonClient = new ConfidentialClientApplication(
                Config.GetConfig("clientId"),
                String.Format(Config.GetConfig("authorityFormat"), tenantId),
                Config.GetConfig("replyUri"),
                new ClientCredential(Config.GetConfig("clientSecret")),
                null,
                new TokenCache());
            var authenticationResult = await daemonClient.AcquireTokenForClientAsync(new string[] { "https://graph.microsoft.com/.default" });
            var accessToken = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);

            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri("https://graph.microsoft.com/beta/");
                client.DefaultRequestHeaders.Authorization = accessToken;

                await DemoBatch(client);
            }
        }

        async Task DemoBatch(HttpClient client)
        {
            /// <summary>
            /// Creates a batch that requests a user and the trending insights around them
            /// </summary>
            /// <todo>
            /// The app uses client_credential flow - if it used a delegated flow, we could use /me alias
            /// or the signed in user's upn. However, you need to replace {upn} with an existing user's UPN
            /// or you will see an error
            /// </todo>
            var request = new HttpRequestMessage(HttpMethod.Post, "$batch");
            request.Content = new StringContent(@"{
                  'requests': [
                    {
                      'id': '1',
                      'method': 'GET',
                      'url': '/users/{upn}?$select=givenName,surName,jobTitle'
                    },
                    {
                      'id': '2',
                      'dependsOn': [ '1' ],
                      'method': 'GET',
                      'url': '/users/{upn}/insights/trending'
                    }
                  ]
                }", Encoding.UTF8, "application/json");
            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
        }
    }
}