using device_code_flow_console;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace add_custom_data
{
    class OpenExtensionsDemo : PublicAppUsingDeviceCodeFlow
    {
        public OpenExtensionsDemo(PublicClientApplication app) : base(app)
        {
        }

        /// <summary>
        /// This class follows the steps in this tutorial: https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_open_users
        /// However it first starts by acquiring a token
        /// </summary>
        public static string[] Scopes { get; set; } = new string[] { "User.Read", "User.ReadBasic.All", "User.ReadWrite", "Group.ReadWrite.All", "Directory.AccessAsUser.All" };

        public async Task RunAsync()
        {
            AuthenticationResult authenticationResult = await AcquireATokenFromCacheOrDeviceCodeFlow(Scopes);
            if (authenticationResult != null)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{authenticationResult.Account.Username} successfully signed-in");
                var accessToken = authenticationResult.AccessToken;

                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    //Use open extensions
                    /// <todo>
                    /// You'll need to flesh out each of the methods below
                    /// </todo>
                    await AddRoamingProfileInformationAsync(client);
                    await RetrieveRoamingProfileInformationAsync(client);
                    await UpdateRoamingProfileInformationAsync(client);
                    await DeleteRoamingProfileInformationAsync(client);
                }
            }
        }
        async Task AddRoamingProfileInformationAsync(HttpClient client)
        {
            Console.WriteLine("Add roaming info to the signed-in user");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to add an open extension to the signed in user
            /// Use the HttpRequestMessage helper
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_open_users as a reference
            /// </exercise_hint>
            
            ///<todo>
            /// Add code here
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            */
        }

        async Task RetrieveRoamingProfileInformationAsync(HttpClient client)
        {
            Console.WriteLine("Get the signed-in user profile and roaming data");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to get the signed in user and their roaming data extension
            /// Use the HttpRequestMessage helper
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_open_users as a reference
            /// </exercise_hint>

            ///<todo>
            /// Add code here
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            */
        }

        async Task UpdateRoamingProfileInformationAsync(HttpClient client)
        {
            Console.WriteLine("Update user profile");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to patch the roaming extension data
            /// Use the HttpRequestMessage helper
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_open_users as a reference
            /// </exercise_hint>

            ///<todo>
            /// Add code here
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);
            
            response.WriteCodeAndReasonToConsole();
            */
        }

         async Task DeleteRoamingProfileInformationAsync(HttpClient client)
        {
            Console.WriteLine("Remove the roaming profile from the signed-in user");
            Console.WriteLine();


            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to delete the open extension from the signed in user
            /// Use the HttpRequestMessage helper
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_open_users as a reference
            /// </exercise_hint>

            ///<todo>
            /// Add code here
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);

            response.WriteCodeAndReasonToConsole();
            */
        }


    }
}
