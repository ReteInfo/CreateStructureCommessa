using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Graph.Auth;

namespace Company.Function
{
  static class Authentication
  {
      private static string _clientId = Environment.GetEnvironmentVariable("ClientIDApp");
      private static string _tenantId = Environment.GetEnvironmentVariable("TenantIDApp");
      private static string _keyApp = Environment.GetEnvironmentVariable("ClientSecretApp");

      //autenticazione
      public static GraphServiceClient auth()
      {
          try
          {
               List<string> scopes = new List<string>
                    {
                        "https://graph.microsoft.com/.default"
                    };
               var app = ConfidentialClientApplicationBuilder
                                  .Create(_clientId)
                                  .WithAuthority(AzureCloudInstance.AzurePublic, _tenantId)
                                  .WithClientSecret(_keyApp)
                                  .Build();
            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(app);
            return new GraphServiceClient(authenticationProvider);

          }
          catch (Exception e)
          {
              throw new Exception(e.Message, e.InnerException);
          }

      }
  }
}
