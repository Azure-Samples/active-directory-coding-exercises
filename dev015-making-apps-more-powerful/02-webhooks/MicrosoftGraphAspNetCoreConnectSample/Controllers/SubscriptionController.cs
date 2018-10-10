/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

namespace MicrosoftGraphAspNetCoreConnectSample.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Newtonsoft.Json;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using MicrosoftGraphAspNetCoreConnectSample.Helpers;
    using MicrosoftGraphAspNetCoreConnectSample.Models;
    using ServiceException = Microsoft.Graph.ServiceException;
    using Subscription = Microsoft.Graph.Subscription;

    public class SubscriptionController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly IHostingEnvironment _env;
        private readonly IGraphSdkHelper _graphSdkHelper;
        private static string subscriptionId;

        public SubscriptionController(IConfiguration configuration, IHostingEnvironment hostingEnvironment, IGraphSdkHelper graphSdkHelper)
        {
            _configuration = configuration;
            _env = hostingEnvironment;
            _graphSdkHelper = graphSdkHelper;
        }

        public ActionResult Index()
        {
            return View();
        }

        [Authorize]
        public async Task<ActionResult> CreateSubscription()
        {
            if (User.Identity.IsAuthenticated)
            {
                string identifier = User.FindFirst(Startup.ObjectIdentifierType)?.Value;
                var graphClient = _graphSdkHelper.GetAuthenticatedClient(identifier);
                Subscription createdSubscription = null;
                
                /// <exercise_hint>
                /// Some hints that can help you.
                /// Creating subscription https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/subscription_post_subscriptions
                /// </exercise_hint>

                /// ADD CODE HERE

                SubscriptionViewModel viewModel = new SubscriptionViewModel
                {
                    Subscription = JsonConvert.DeserializeObject<Models.Subscription>(JsonConvert.SerializeObject(createdSubscription))
                };

                SubscriptionStore.SaveSubscriptionInfo(viewModel.Subscription.Id,
                    viewModel.Subscription.ClientState,
                    User.FindFirst(Startup.ObjectIdentifierType)?.Value,
                    User.FindFirst(Startup.TenantIdType)?.Value);
                subscriptionId = viewModel.Subscription.Id;

                return View("Subscription", createdSubscription);
            }

            return View("Subscription", null);
        }

        // Delete the current webhooks subscription and sign out the user.
        [Authorize]
        public async Task<ActionResult> DeleteSubscription()
        {
            string subscriptionId = SubscriptionController.subscriptionId;

            string identifier = User.FindFirst(Startup.ObjectIdentifierType)?.Value;
            var graphClient = _graphSdkHelper.GetAuthenticatedClient(identifier);

            try
            {
                await graphClient.Subscriptions[subscriptionId].Request().DeleteAsync();
            }
            catch (ServiceException e)
            {

            }

            return RedirectToAction("SignOut", "Account");
        }
    }
}