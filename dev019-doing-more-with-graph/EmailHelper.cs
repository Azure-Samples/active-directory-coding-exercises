using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace InclusivityFeedbackLoop
{
    class EmailHelper
    {
        /// <summary>
        /// Compose and send a new email.
        /// </summary>
        /// <param name="subject">The subject line of the email.</param>
        /// <param name="bodyContent">The body of the email.</param>
        /// <param name="recipients">A semicolon-separated list of email addresses.</param>
        /// <returns></returns>
        public static async Task ComposeAndSendMailAsync(string subject,
                                                            string bodyContent,
                                                            string recipient,
                                                            string token, TraceWriter log)
        {
            List<Recipient> recipientList = new List<Recipient>();

            recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = recipient.Trim() } });

            try
            {
                var email = new Message
                {
                    Body = new ItemBody
                    {
                        Content = bodyContent,
                        ContentType = BodyType.Html,
                    },
                    Subject = subject,
                    ToRecipients = recipientList
                };

                try
                {
                    //initialize HttpClient for REST call
                    HttpClient client = new HttpClient();
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                    //setup the client post
                    string contentString = JsonConvert.SerializeObject(email);

                    HttpContent content = new StringContent("{\"message\":" + contentString + "}");
                    //Specify the content type. 
                    content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                    HttpResponseMessage result = await client.PostAsync(
                        "https://graph.microsoft.com/v1.0/users/" + recipient + "/sendMail", content);

                    log.Info(result.ToString());

                    if (result.IsSuccessStatusCode)
                    {
                        //email send successfully.
                        log.Info("Email sent successfully. ");
                    }
                }
                catch (ServiceException exception)
                {
                    throw new Exception("We could not send the message: " + exception.Error == null ? "No error message returned." : exception.Error.Message);
                }
            }

            catch (Exception e)
            {
                throw new Exception("We could not send the message: " + e.Message);
            }
        }
    }
}