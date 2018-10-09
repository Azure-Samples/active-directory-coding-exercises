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
                    }
                }
            }
        
        async Task ViewAvailableExtensionsAsync(HttpClient client)
        {
            Console.WriteLine("Get the available schema extensions");
            Console.WriteLine();
            var request = new HttpRequestMessage(HttpMethod.Get, "schemaextensions");

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
        }

        async Task<string> RegisterSchemaExtensionAsync(HttpClient client)
        {
            Console.WriteLine("Register a new schema extension for groups");
            Console.WriteLine();

            var request = new HttpRequestMessage(HttpMethod.Post, "schemaExtensions");
            request.Content = new StringContent(@"{
                  'id': 'courses',
                  'description': 'Graph Learn training courses extensions',
                  'targetTypes': [
                    'Group'
                  ],
                  'properties': [
                    {
                      'name': 'courseId',
                      'type': 'Integer'
                    },
                    {
                      'name': 'courseName',
                      'type': 'String'
                    },
                    {
                      'name': 'courseType',
                      'type': 'String'
                    }
                  ]
                }", Encoding.UTF8, "application/json");

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);
            Console.WriteLine(JValue.Parse(responseBody).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            return (string)o["id"];
        }

        async Task<string> CreateGroupWithExtendedDataAsync(HttpClient client, string schemaId)
        {
            Console.WriteLine("Create a new group with extended data");
            Console.WriteLine();

            var request = new HttpRequestMessage(HttpMethod.Post, "groups");
            string json = @"{
                  'displayName': 'New Managers 2018',
                  'description': 'New Managers training course 2018',
                  'groupTypes': [
                    'Unified'
                  ],
                  'mailEnabled': true,
                  'mailNickName': 'newManagers" + Guid.NewGuid().ToString().Substring(8) + @"',
                  'securityEnabled': false,
                  '" + schemaId + @"': {
                    'courseId': 123,
                    'courseName': 'New Managers',
                    'courseType': 'Online'
                  }
                }";

            request.Content = new StringContent(json,
                Encoding.UTF8,
                "application/json");

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);
            Console.WriteLine(JValue.Parse(responseBody).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
            return (string)o["id"];
        }

        async Task UpdateCustomDataInGroupAsync(HttpClient client, string groupId, string schemaId)
        {
            Console.WriteLine("Update custom data for groups");
            Console.WriteLine();

            var request = new HttpRequestMessage(new HttpMethod("PATCH"), "groups/" + groupId);
            string json = @"{
                  '" + schemaId + @"': {
                    'courseId': '123',
                    'courseName': 'New Managers',
                    'courseType': 'Online'
                  }
                }";
            request.Content = new StringContent(
                json,
                Encoding.UTF8,
                "application/json");

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();
            Console.WriteLine();
        }

        async Task GetGroupAndExtensionDataAsync(HttpClient client, string schemaId)
        {
            Console.WriteLine("Find group based on custom data values (filter)");
            Console.WriteLine();

            var request = new HttpRequestMessage(
                HttpMethod.Get,
                "groups?$filter=" + schemaId + "/courseId eq '123'&$select=displayName,id,description," + schemaId);

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();

            var responseBody = await response.Content.ReadAsStringAsync();

            JObject o = JObject.Parse(responseBody);

            Console.WriteLine(JValue.Parse(await response.Content.ReadAsStringAsync()).ToString(Newtonsoft.Json.Formatting.Indented));
            Console.WriteLine();
        }

        async Task DeleteGroupAndExtensionAsync(HttpClient client, string schemaId, string groupId)
        {
            Console.WriteLine("Delete registered schema extension and the group ");
            Console.WriteLine();

            var request = new HttpRequestMessage(HttpMethod.Delete, "schemaextensions/" + schemaId);

            var response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();

            Console.WriteLine();

            request = new HttpRequestMessage(HttpMethod.Delete, "groups/" + groupId);

            response = await client.SendAsync(request);
            response.WriteCodeAndReasonToConsole();

            Console.WriteLine();
        }
    }
}
