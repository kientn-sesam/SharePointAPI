using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
namespace SharePointAPI.Middleware
{
    
    public class AuthHelper
    {
        private static readonly HttpClient client = new HttpClient();
        private static Dictionary<string, Token> tokenCache = new Dictionary<string, Token>();
        public static ClientContext GetClientContextOauth(string url, string tokenUrl, string clientId, string clientSecret, string resource)
        {
            ClientContext cc = new ClientContext(url);
            cc.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                Int32 timeNow = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
                if (!AuthHelper.tokenCache.ContainsKey(url) || AuthHelper.tokenCache[url].expires_on < timeNow)
                {
                    var accessToken = getAccessToken(tokenUrl, clientId, clientSecret, resource).Result;
                    tokenCache[url] = accessToken;
                }

                e.WebRequestExecutor.WebRequest.Headers.Add("Authorization", "Bearer " + tokenCache[url].access_token);
            };

            return cc;
        }

        public static ClientContext GetClientContextForUsernameAndPassword(string url, string username, string password)
        {
            var secure = new SecureString();
            foreach (char c in password)
            {
                secure.AppendChar(c);
            }
            ClientContext cc = new ClientContext(url)
            {
                AuthenticationMode = ClientAuthenticationMode.Default,
                Credentials = new SharePointOnlineCredentials(username, secure)
            };
            
            cc.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
            {
                e.WebRequestExecutor.WebRequest.UserAgent = "ISV|Villegder|GovernanceCheck/1.0";
            };

            return cc;
        }

        internal static ClientContext GetClientContextOauth(string url, string clientId, string clientSecret)
        {
            Uri siteUri = new Uri(url);
            //string realm = AuthHelper.GetRealmFromTargetUrl(siteUri);
            return null;
        }

        public static async Task<Token> getAccessToken(string tokenUrl, string clientId, string clientSercret, string resource)
        {
            var values = new Dictionary<string, string>
            {
               { "grant_type", "client_credentials" },
               { "client_id", clientId },
                { "client_secret", clientSercret},
                { "resource", resource}
            };

            var content = new FormUrlEncodedContent(values);

            var response = await client.PostAsync(tokenUrl, content);

            var responseString = await response.Content.ReadAsStringAsync();
            Token json = JsonConvert.DeserializeObject<Token>(responseString);

            return json;
        }

    }



    public class Token
    {
        string token_type { get; set; }
        int expires_in { get; set; }
        int intext_expires_in { get; set; }
        public int expires_on { get; set; }
        int not_before { get; set; }
        string resource { get; set; }
        public string access_token { get; set; }
    }
}