/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    using System;
    using System.Collections.Concurrent;
    using System.Linq;
    using System.Threading;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Primitives;
    using MicrosoftGraphAspNetCoreConnectSample.Models;

    public class NotificationController : Controller
    {
        private static ReaderWriterLockSlim SessionLock = new ReaderWriterLockSlim(LockRecursionPolicy.NoRecursion);

        [Authorize]
        public ActionResult Index()
        {
            if (User.Identity.IsAuthenticated)
            {
                // Get user's id for token cache.
                ViewBag.CurrentUserId = User.FindFirst(Startup.ObjectIdentifierType)?.Value;

                //Store the notifications in session state. A production
                //application would likely queue for additional processing.
                //Store the notifications in application state. A production
                //application would likely queue for additional processing.
                ConcurrentBag<Notification> notificationArray = NotificationsStore.GetNotifications();
                if (notificationArray == null)
                {
                    notificationArray = new ConcurrentBag<Notification>();
                }

                NotificationsStore.SaveNotifications(notificationArray);

                return View(notificationArray);
            }
            return View();
        }

        // The `notificationUrl` endpoint that's registered with the webhook subscription.
        [HttpPost]
        public ActionResult Listen()
        {
            StringValues queryValues;
            // Validate the new subscription by sending the token back to Microsoft Graph.
            // This response is required for each subscription.
            if (Request.Query.TryGetValue("validationToken", out queryValues))
            {
                string token = queryValues.First();
                return Content(token, "plain/text");
            }

            // Parse the received notifications.
            else
            {
                try
                {
                    using (var inputStream = new System.IO.StreamReader(Request.Body))
                    {
                        JObject jsonObject = JObject.Parse(inputStream.ReadToEnd());
                        if (jsonObject != null)
                        {

                            // Notifications are sent in a 'value' array. The array might contain multiple notifications for events that are
                            // registered for the same notification endpoint, and that occur within a short time span.
                            JArray value = JArray.Parse(jsonObject["value"].ToString());
                            foreach (var notification in value)
                            {
                                Notification current = JsonConvert.DeserializeObject<Notification>(notification.ToString());

                                // Check client state to verify the message is from Microsoft Graph.
                                SubscriptionStore subscription = SubscriptionStore.GetSubscriptionInfo(current.SubscriptionId);

                                // This sample only works with subscriptions that are still cached.
                                if (subscription != null)
                                {
                                    if (current.ClientState == subscription.ClientState)
                                    {
                                        //Store the notifications in application state. A production
                                        //application would likely queue for additional processing.

                                        ConcurrentBag<Notification> notificationArray = NotificationsStore.GetNotifications();
                                        if (notificationArray == null)
                                        {
                                            notificationArray = new ConcurrentBag<Notification>();
                                        }

                                        notificationArray.Add(current);
                                        NotificationsStore.SaveNotifications(notificationArray);
                                        //TempData["notifications"] = JsonConvert.SerializeObject(notificationArray);
                                    }
                                }
                            }

                        }
                    }
                }
                catch (Exception)
                {

                    // TODO: Handle the exception.
                    // Still return a 202 so the service doesn't resend the notification.
                }

                return new StatusCodeResult(202);
            }
        }

    }
}