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
        class SchemaExtensionsDemo:PublicAppUsingDeviceCodeFlow
        {
            public SchemaExtensionsDemo(PublicClientApplication app) : base(app)
            {
            }

        /// <summary>
        /// This class follows the steps in this tutorial: https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups
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

                        // Use schema extensions

                        /// <todo>
                        /// Uncomment the code below - once you do that it won't build until you
                        /// flesh out each of the methods below
                        /// </todo>
                        /*
                        await ViewAvailableExtensionsAsync(client);

                        var schemaId = await RegisterSchemaExtensionAsync(client);

                        // Need to wait for schema to finish creating before creating group, otherwise get error
                        System.Threading.Thread.Sleep(3000);

                        var groupId = await CreateGroupWithExtendedDataAsync(client, schemaId);

                        System.Threading.Thread.Sleep(3000);

                        await UpdateCustomDataInGroupAsync(client, groupId, schemaId);

                        System.Threading.Thread.Sleep(3000);

                        await GetGroupAndExtensionDataAsync(client, schemaId);

                        System.Threading.Thread.Sleep(3000);

                        await DeleteGroupAndExtensionAsync(client, schemaId, groupId);
                        */
                    }
                }
            }
        
        async Task ViewAvailableExtensionsAsync(HttpClient client)
        {
            Console.WriteLine("Get the available schema extensions");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to get all available schema extensions
            /// Use the HttpRequestMessage helper
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
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

        async Task<string> RegisterSchemaExtensionAsync(HttpClient client)
        {
            Console.WriteLine("Register a new schema extension for groups");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to register a new schema extension for a group to represent a class
            /// Use the HttpRequestMessage helper  - you might want to set the content-type here to application/json too
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
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
            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);
            Console.WriteLine(JValue.Parse(responseBody).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            return (string)o["id"];
            */

            // return dummy string for now to allow the app to build
            /// <todo> remove when you flesh out this method </todo>
            return ("id");
        }

        async Task<string> CreateGroupWithExtendedDataAsync(HttpClient client, string schemaId)
        {
            Console.WriteLine("Create a new group with extended data");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph create a new group, with extension data
            /// Use the HttpRequestMessage helper - you might want to set the content-type here to application/json too
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
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
            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);
            Console.WriteLine(JValue.Parse(responseBody).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            return (string)o["id"];
            */

            // return dummy string for now, to allow the app to build
            /// <todo> remove when you flesh out this method </todo>
            return ("id");
        }

        async Task UpdateCustomDataInGroupAsync(HttpClient client, string groupId, string schemaId)
        {
            Console.WriteLine("Update custom data for groups");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to update some of the custom data on the group
            /// Use the HttpRequestMessage helper - you might want to set the content-type here to application/json too
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
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
            Console.WriteLine();
            */
        }

        async Task GetGroupAndExtensionDataAsync(HttpClient client, string schemaId)
        {
            Console.WriteLine("Find group based on custom data values (filter)");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to find the group based on the value of its courseId (or a different ext property)
            /// Use the HttpRequestMessage helper 
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
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

            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);

            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            */
        }

        async Task DeleteGroupAndExtensionAsync(HttpClient client, string schemaId, string groupId)
        {
            Console.WriteLine("Delete registered schema extension and the group ");
            Console.WriteLine();

            /// <exercise_hint>
            /// Add a REST http request to Microsoft Graph to delete the schema extension and the group
            /// Use https://developer.microsoft.com/en-us/graph/docs/concepts/extensibility_schema_groups as a reference
            /// </exercise_hint>

            ///<todo>
            /// Add code here to remove schema extension
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();

            Console.WriteLine();
            */

            ///<todo>
            /// Add code here to remove the group
            ///</todo> 

            /// <todo>
            /// remove comments below once you've added code to create a request object
            /// </todo>
            /*
            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();

            Console.WriteLine();
            */
        }
    }
}
