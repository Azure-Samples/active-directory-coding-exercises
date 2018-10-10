using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace InclusivityFeedbackLoop
{
    public static class ScreenEmailFunc
    {
        private const string idaMicrosoftGraphUrl = "https://graph.microsoft.com";

        /// <summary>
        /// This is the Azure Function entry point. We're using an HttpTrigger function, which 
        /// receives the Microsoft Graph webhook and handles responding according to the Microsoft Graph
        /// requirements.
        /// </summary>
        /// <param name="req">Incoming HTTP request</param>
        /// <param name="log">An Azure Function log we can write to for debugging purposes</param>
        /// <returns></returns>
        [FunctionName("ScreenEmailFunc")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"Webhook was triggered!");

            // Handle validation scenario for creating a new webhook subscription
            string validationToken;
            if (GetValidationToken(req, out validationToken))
            {
                return PlainTextResponse(validationToken);
            }

            // Process each notification
            var response = await ProcessWebhookNotificationsAsync(req, log, async hook =>
            {
                return await CheckForSubscriptionChangesAsync(hook.SubscriptionId, hook.Resource, log);
            });
            return response;
        }

        /// <summary>
        /// Helper function that contains the work of parsing the incoming HTTP request JSON 
        /// body into webhook subscription notifications, and then calling the processSubscriptionNotification
        /// function for each received notification.
        /// </summary>
        /// <param name="req">Incoming HTTP Request.</param>
        /// <param name="log">Log used for writing out tracing information.</param>
        /// <param name="processSubscriptionNotification">Async function that is called per-notification in the request.</param>
        /// <returns></returns>
        private static async Task<HttpResponseMessage> ProcessWebhookNotificationsAsync(HttpRequestMessage req, TraceWriter log, Func<SubscriptionNotification, Task<bool>> processSubscriptionNotification)
        {
            // Read the body of the request and parse the notification
            string content = await req.Content.ReadAsStringAsync();
            log.Verbose($"Raw request content: {content}");

            // In a production application you should queue the work to be done in an Azure Queue and _not_ do the heavy lifting 
            // in the webhook request handler.

            var webhooks = JsonConvert.DeserializeObject<WebhookNotification>(content);
            if (webhooks?.Notifications != null)
            {
                // Since webhooks can be batched together, loop over all the notifications we receive and process them separately.
                foreach (var hook in webhooks.Notifications)
                {
                    log.Info($"Hook received for subscription: '{hook.SubscriptionId}' Resource: '{hook.Resource}', changeType: '{hook.ChangeType}'");
                    try
                    {
                        await processSubscriptionNotification(hook);
                    }
                    catch (Exception ex)
                    {
                        log.Error($"Error processing subscription notification. Subscription {hook.SubscriptionId} was skipped. {ex.Message}", ex);
                    }
                }

                // After we process all the messages, return an empty response.
                return req.CreateResponse(HttpStatusCode.NoContent);
            }
            else
            {
                log.Info($"Request was incorrect. Returning bad request.");
                return req.CreateResponse(HttpStatusCode.BadRequest);
            }
        }

        /// <summary>
        /// Data structure for the request payload for an incoming webhook.
        /// </summary>
        private class WebhookNotification
        {
            [JsonProperty("value")]
            public SubscriptionNotification[] Notifications { get; set; }
        }

        /// <summary>
        /// Data structure for the notification payload inside the incoming webhook.
        /// </summary>
        private class SubscriptionNotification
        {
            [JsonProperty("clientState")]
            public string ClientState { get; set; }
            [JsonProperty("resource")]
            public string Resource { get; set; }
            [JsonProperty("subscriptionId")]
            public string SubscriptionId { get; set; }

            [JsonProperty("changeType")]
            public string ChangeType { get; set; }
        }

        /// <summary>
        /// Parse the request query string to see if there is a validationToken parameter, and return the value if found.
        /// </summary>
        /// <param name="req">Incoming HTTP request.</param>
        /// <param name="token">Out parameter containing the validationToken, if found.</param>
        /// <returns></returns>
        private static bool GetValidationToken(HttpRequestMessage req, out string token)
        {
            Dictionary<string, string> qs = req.GetQueryNameValuePairs()
                                    .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
            return qs.TryGetValue("validationToken", out token);
        }

        /// <summary>
        /// Get the resource from the subscription and read the email that was sent to process it.
        /// </summary>
        /// <param name="subscriptionId">The subscription ID that generated the notification</param>
        /// <param name="resource">The resource in the notification.</param>
        /// <param name="log">TraceWriter for debug output</param>
        /// <returns></returns>
        private static async Task<bool> CheckForSubscriptionChangesAsync(string subscriptionId, string resource, TraceWriter log)
        {
            // Obtain an access token
            string accessToken = await RetrieveAccessTokenAsync(log);

            HttpClient client = new HttpClient();
            
            // Send Graph request to fetch mail
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/" + resource);

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Send the request and get the response.
            HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(continueOnCapturedContext: false);

            log.Info(response.ToString());

            if (response.IsSuccessStatusCode)
            {
                var result = await response.Content.ReadAsStringAsync();

                JObject obj = (JObject)JsonConvert.DeserializeObject(result);

                string subject = (string)obj["subject"];

                if (!subject.Equals("Inclusivity tips", StringComparison.OrdinalIgnoreCase))
                {
                    string content = (string)obj["body"]["content"];

                    await ScanEmail(content, log);
                }
            }

            return true;
        }

        /// <summary>
        /// Generate a plain text response, used to return the validationToken during the subscription validation flow
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private static HttpResponseMessage PlainTextResponse(string text)
        {
            HttpResponseMessage response = new HttpResponseMessage()
            {
                StatusCode = HttpStatusCode.OK,
                Content = new StringContent(
                        text,
                        System.Text.Encoding.UTF8,
                        "text/plain"
                    )
            };
            return response;
        }

        /// <summary>
        /// Retrieve an access token for a particular user from our token cache using ADAL.
        /// </summary>
        private static async Task<string> RetrieveAccessTokenAsync(TraceWriter log)
        {
            log.Verbose($"Retriving new accessToken");

            string authorityUrl = System.Environment.GetEnvironmentVariable("AuthorityUrl", EnvironmentVariableTarget.Process);

            var authContext = new AuthenticationContext(authorityUrl);

            string clientId = System.Environment.GetEnvironmentVariable("ClientId", EnvironmentVariableTarget.Process);
            string clientSecret = System.Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process);

            try
            {
                var clientCredential = new ClientCredential(clientId, clientSecret);
                var authResult = await authContext.AcquireTokenAsync(idaMicrosoftGraphUrl, clientCredential);
                return authResult.AccessToken;
            }
            catch (AdalException ex)
            {
                log.Info($"ADAL Error: Unable to retrieve access token: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Uses Moderator API to check content of the email.
        /// </summary>
        private static async Task ScanEmail(string body, TraceWriter log)
        {
            var client = new HttpClient();
            string cognitiveAPIKey = Environment.GetEnvironmentVariable("CognitiveAPIKey", EnvironmentVariableTarget.Process);
            string cognitiveAPIUrl = Environment.GetEnvironmentVariable("CognitiveAPIUrl", EnvironmentVariableTarget.Process);

            // Request headers
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", cognitiveAPIKey);

            HttpResponseMessage response;

            // Request body
            byte[] byteData = Encoding.UTF8.GetBytes(body);

            using (var content = new ByteArrayContent(byteData))
            {
                content.Headers.ContentType = new MediaTypeHeaderValue("text/plain");
                response = await client.PostAsync(cognitiveAPIUrl, content);
            }

            var result = await response.Content.ReadAsStringAsync();

            JObject obj = (JObject)JsonConvert.DeserializeObject(result);

            if (!obj["Terms"].HasValues)
            {
                log.Info("No terms from the term list were detected in the text.");
            }
            else
            {
                JArray terms = (JArray)obj["Terms"];
                string termValue = "";

                foreach (var item in terms.Children())
                {
                    var itemProperties = item.Children<JProperty>();
                    var term = itemProperties.FirstOrDefault(x => x.Name == "Term");
                    termValue = term.Value.ToString();
                }

                await SendEmail(log, termValue);
            }
        }

        private static async Task SendEmail(TraceWriter log, string term)
        {
            var token = await RetrieveAccessTokenAsync(log);
            Dictionary<string, string> alternative = new Dictionary<string, string>();
            alternative.Add("guys", "team, all");
            alternative.Add("girls", "all, ladies");

            string userid = System.Environment.GetEnvironmentVariable("UserId", EnvironmentVariableTarget.Process);
            string content = String.Format("<p>Hi!</p><p>Inclusive e-mails are one of the top 3 ways to create a happier, more collaborative work environment. We noticed that you recently sent an e-mail with some non-inclusive language. :( </p><p>Next time, instead of the word \"{0}\", try \"{1}\" instead.</p><p>Thanks!</p>",
                term, alternative[term]);
            await EmailHelper.ComposeAndSendMailAsync("Inclusivity tips", content, userid, token, log);

            log.Info("Non inclusive words!");
        }
    }
}