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

using device_code_flow_console;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace add_custom_data
{
    public class MyInformation : PublicAppUsingDeviceCodeFlow
    {
        public MyInformation(PublicClientApplication app) : base(app)
        {
        }

        /// <summary>
        /// Scopes to request access to the protected Web API (here Microsoft Graph)
        /// </summary>
        public static string[] Scopes { get; set; } = new string[] { "User.Read", "User.ReadBasic.All", "User.ReadWrite", "Group.ReadWrite.All", "Directory.AccessAsUser.All" };

        /// <summary>
        /// URLs of the protected Web APIs to call (here Microsoft Graph)
        /// </summary>
        public static string WebApiUrlMe { get; set; } = "https://graph.microsoft.com/v1.0/me";
        public static string WebApiUrlMyManager { get; set; } = "https://graph.microsoft.com/v1.0/me/manager";

        /// <summary>
        /// Calls the Web API and displays its information
        /// </summary>
        /// <returns></returns>
        public async Task DisplayMeAndMyManagerAsync()
        {
            AuthenticationResult authenticationResult = await AcquireATokenFromCacheOrDeviceCodeFlow(Scopes);
            if (authenticationResult != null)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{authenticationResult.Account.Username} successfully signed-in");

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Me");
                Console.ResetColor();
                await CallWebApiAndDisplayResultASync(WebApiUrlMe, authenticationResult);
                Console.WriteLine();

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("My manager");
                Console.ResetColor();
                await CallWebApiAndDisplayResultASync(WebApiUrlMyManager, authenticationResult);
            }
        }

 
     
        /// <summary>
        /// Calls the protected Web API and displays the result
        /// </summary>
        /// <param name="authenticationResult"><see cref="AuthenticationResult"/> returned by successfull call to MSAL.NET</param>
        public static async Task CallWebApiAndDisplayResultASync(string webApiUrl, AuthenticationResult authenticationResult)
        {
            if (authenticationResult != null)
            {

                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);

                HttpResponseMessage response = await client.GetAsync(webApiUrl);
                if (response.IsSuccessStatusCode)
                {
                    string json = await response.Content.ReadAsStringAsync();
                    JObject me = JsonConvert.DeserializeObject(json) as JObject;
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Display(me);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                    string content = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Content: {content}");
                }
                Console.ResetColor();

            }
        }

        /// <summary>
        /// Display the result of the Web API call
        /// </summary>
        /// <param name="result">Object to display</param>
        private static void Display(JObject result)
        {
            foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith('@')))
            {
                Console.WriteLine($"{child.Name} = {child.Value}");
            }
        }


    }
}
