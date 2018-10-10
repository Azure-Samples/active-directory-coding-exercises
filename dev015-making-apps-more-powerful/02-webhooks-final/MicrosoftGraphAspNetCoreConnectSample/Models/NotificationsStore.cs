using System;
using System.Collections.Concurrent;
using System.Web;
using Microsoft.Extensions.Caching.Memory;

namespace MicrosoftGraphAspNetCoreConnectSample.Models
{
    public class NotificationsStore
    {
        private static MemoryCache cache = new MemoryCache(new MemoryCacheOptions());

        // This sample temporarily stores the current subscription ID, client state, user object ID, and tenant ID.
        // This info is required so the NotificationController can retrieve an access token from the cache and validate the subscription.
        // Production apps typically use some method of persistent storage.
        public static void SaveNotifications(ConcurrentBag<Notification> notifications)
        {
            cache.Set(
                "notificationArray",
                notifications,
                new TimeSpan(24, 0, 0));
        }

        public static ConcurrentBag<Notification> GetNotifications()
        {
            ConcurrentBag<Notification> notifications = (ConcurrentBag<Notification>) cache.Get("notificationArray");
            return notifications;
        }
    }
}