using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Text.Json;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Company.Function
{
    static class Authentication
    {
        //private static string _clientId = Environment.GetEnvironmentVariable("ClientIDApp");
        //private static string _tenantId = Environment.GetEnvironmentVariable("TenantIDApp");
        //private static string _keyApp = Environment.GetEnvironmentVariable("ClientSecretApp");

        public static GraphServiceClient _graphClient;
        private static string _clientId = "31e57cf9-4de6-4015-ad5b-26658e694de6";
        private static string _tenantId = "e2291a86-71d9-402e-aebd-4f7659a41d35";
        private static string _clientSecret = "g.PhDGsga884nanPd~-8r~y64~2BmclAgm";

        private static string username = "simone.ferrazzo@reteinformatica.com";
        private static string password = "FabioFilzi.32a";

        private static string tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/token";
        private static string resourceSP = "https://reteinformatica.sharepoint.com"; 
        private static string resource = "https://graph.microsoft.com/"; 

        public static string resourceAdmin = "https://reteinformatica-admin.sharepoint.com/";


        public static async Task<GraphServiceClient> auth()
        {
            try
            {

                var pwdA = new SecureString();
                foreach (var c in password){
                    pwdA.AppendChar(c);
                }
                
                var httpClient = new System.Net.Http.HttpClient();
                string token = "";
                
                var pwd = new System.Net.NetworkCredential(string.Empty, pwdA).Password;
                
                var body = $"resource={resource}&client_id={_clientId}&grant_type=password&username={username}&password={pwd}";
                
                using (var stringContent = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded"))
                {
                    var result = await httpClient.PostAsync(tokenEndpoint, stringContent).ContinueWith((response) =>
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    });
                    
                    var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
                    token = tokenResult.GetProperty("access_token").GetString();
                }
                // _graphClient =  new GraphServiceClient(
                //             new DelegateAuthenticationProvider((requestMessage) =>
                //             {
                //                 requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                //                 return Task.FromResult(0);
                //             }
                // ));
                var delegateAuthProvider = new DelegateAuthenticationProvider((requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                    return Task.FromResult(0);
                });

                return  new GraphServiceClient(delegateAuthProvider);
            }
            catch (Exception e)
            {
                Console.WriteLine("Errore Autenticazione: " + e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }

        public static async Task<string> authSP()
        {
            try
            {
                //var context = new ClientContext("https://reteinformatica.sharepoint.com");
                var pwdA = new SecureString();
                foreach (var c in password){
                    pwdA.AppendChar(c);
                }
                
                
                var httpClient = new System.Net.Http.HttpClient();
                string token = "";
                
                var pwd = new System.Net.NetworkCredential(string.Empty, pwdA).Password;
                
                var body = $"resource={resourceSP}&client_id={_clientId}&grant_type=password&username={username}&password={pwd}";
                
                using (var stringContent = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded"))
                {
                    var result = await httpClient.PostAsync(tokenEndpoint, stringContent).ContinueWith((response) =>
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    });
                    
                    var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
                    token = tokenResult.GetProperty("access_token").GetString();
                }
                
                
                
                return token;
                
            }
            catch (Exception e)
            {
                Console.WriteLine("Errore Autenticazione sharepoint online: " + e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }
        public static async Task<string> authHubSite()
        {
            try
            {
                //var context = new ClientContext("https://reteinformatica.sharepoint.com");
                var pwdA = new SecureString();
                foreach (var c in password){
                    pwdA.AppendChar(c);
                }
                
                
                var httpClient = new System.Net.Http.HttpClient();
                string token = "";
                
                var pwd = new System.Net.NetworkCredential(string.Empty, pwdA).Password;
                
                var body = $"resource={resourceAdmin}&client_id={_clientId}&grant_type=password&username={username}&password={pwd}";
                
                using (var stringContent = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded"))
                {
                    var result = await httpClient.PostAsync(tokenEndpoint, stringContent).ContinueWith((response) =>
                    {
                        return response.Result.Content.ReadAsStringAsync().Result;
                    });
                    
                    var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
                    token = tokenResult.GetProperty("access_token").GetString();
                }
                return token;
            }
            catch (Exception e)
            {
                Console.WriteLine("Errore Autenticazione sharepoint online: " + e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }
    }
}
