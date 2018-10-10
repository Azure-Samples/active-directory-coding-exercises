using System;
using System.Web;
using Microsoft.Extensions.Caching.Memory;

namespace MicrosoftGraphAspNetCoreConnectSample.Models
{
    public class SubscriptionStore
    {
        public string SubscriptionId { get; set; }
        public string ClientState { get; set; }
        public string UserId { get; set; }
        public string TenantId { get; set; }

        private static MemoryCache cache = new MemoryCache(new MemoryCacheOptions());
        
        private SubscriptionStore(string subscriptionId, Tuple<string, string, string> parameters)
        {
            SubscriptionId = subscriptionId;
            ClientState = parameters.Item1;
            UserId = parameters.Item2;
            TenantId = parameters.Item3;
        }

        // This sample temporarily stores the current subscription ID, client state, user object ID, and tenant ID.
        // This info is required so the NotificationController can retrieve an access token from the cache and validate the subscription.
        // Production apps typically use some method of persistent storage.
        public static void SaveSubscriptionInfo(string subscriptionId, string clientState, string userId, string tenantId)
        {
            cache.Set(
                "subscriptionId_" + subscriptionId,
                Tuple.Create(clientState, userId, tenantId),
                new TimeSpan(24, 0, 0));
        }

        public static SubscriptionStore GetSubscriptionInfo(string subscriptionId)
        {
            Tuple<string, string, string> subscriptionParams =
                (Tuple<string, string, string>) cache.Get("subscriptionId_" + subscriptionId);
            return new SubscriptionStore(subscriptionId, subscriptionParams);
        }
    }
}