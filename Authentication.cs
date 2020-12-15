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
using Microsoft.Extensions.Logging;

namespace Company.Function
{
    static class Authentication
    {
        //private static string _clientId = Environment.GetEnvironmentVariable("ClientIDApp");
        //private static string _tenantId = Environment.GetEnvironmentVariable("TenantIDApp");
        //private static string _keyApp = Environment.GetEnvironmentVariable("ClientSecretApp");
        private static string resource = "https://graph.microsoft.com/"; 
        private static string tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/token";
        public static GraphServiceClient _graphClient;
        //private static string _clientId = Environment.GetEnvironmentVariable("clientId_auth",EnvironmentVariableTarget.Process);
        //private static string username = Environment.GetEnvironmentVariable("username_auth",EnvironmentVariableTarget.Process);
        //private static string password = Environment.GetEnvironmentVariable("password_auth",EnvironmentVariableTarget.Process);
        //private static string resourceSP = Environment.GetEnvironmentVariable("resourceSP",EnvironmentVariableTarget.Process);
        //public static string resourceAdmin = Environment.GetEnvironmentVariable("resourceAdminSP",EnvironmentVariableTarget.Process);
        private static string _clientId = "31e57cf9-4de6-4015-ad5b-26658e694de6";
        private static string username = "simone.ferrazzo@reteinformatica.com";
        private static string password = "FabioFilzi.32a";
        private static string resourceSP = "https://reteinformatica.sharepoint.com";
        public static string resourceAdmin = "https://reteinformatica-admin.sharepoint.com/";


        public static async Task<GraphServiceClient> auth(ILogger log)
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
                var delegateAuthProvider = new DelegateAuthenticationProvider((requestMessage) =>
                {
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

                    return Task.FromResult(0);
                });

                return  new GraphServiceClient(delegateAuthProvider);
            }
            catch (Exception e)
            {
                log.LogError("Errore autenticazione a graph: {0}",e.Message);
                //Console.WriteLine("Errore Autenticazione: " + e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }

        public static async Task<string> authSP(ILogger log)
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
                //Console.WriteLine("Errore Autenticazione sharepoint online: " + e.Message);
                log.LogError("ERRORE autenticazione a sharepoint online: {0}",e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }
        public static async Task<string> authHubSite(ILogger log)
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
                log.LogError("Errore autenticazione HUb site: {0}",e.Message);
                //Console.WriteLine("Errore Autenticazione sharepoint online: " + e.Message);
                throw new Exception("Errore Autenticazione applicazione");
            }

        }
    }
}
