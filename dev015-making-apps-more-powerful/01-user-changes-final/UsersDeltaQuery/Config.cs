using System;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace UsersDeltaQuery
{
        public static class Config
        {
            /// <summary>
            /// Static dictionary to store config parameters
            /// Fill in your app's values here
            /// </summary>
            private static Dictionary<string, string> _config = new Dictionary<string, string>() {
            // Adding Items to the dictionary
                {"clientId", "{your-client-id}" },
                {"clientSecret", "{your-client-secret}" },
                {"tenantId", "{your-tenant-name}" },
                {"authorityFormat", "https://login.microsoftonline.com/{0}/v2.0" },
                {"replyUri", "{your-reply-uri}" }
            };

            /// <summary>
            /// Access the Dictionary from external sources
            /// </summary>
            public static String GetConfig(String key)
            {
                // Try to get the result in the static Dictionary
                String result;
                if (_config.TryGetValue(key, out result))
                {
                    return result;
                }
                else
                {
                    return null;
                }
            }
        }
}