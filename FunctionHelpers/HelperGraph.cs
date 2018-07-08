using System;
using System.Globalization;
using System.Threading;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace FunctionHelpers
{
    public static class HelperGraph
    {
        public static AuthenticationResult GetAppOnlyToken(string clientIdKey, string appKeyKey, string tenant, string aadInstance, string keyVaultUrl)
        {

            var clientId = HelperSecrets.GetSecretString(clientIdKey, keyVaultUrl).Result;
            var appKey = HelperSecrets.GetSecretString(appKeyKey, keyVaultUrl).Result;
            
            var authority = String.Format(CultureInfo.InvariantCulture, aadInstance ?? throw new InvalidOperationException("aadInstance is not specified"), tenant);

            var authContext = new AuthenticationContext(authority);

            AuthenticationResult result = null;
            int retryCount = 0;
            bool retry = false;
            do
            {
                retry = false;
                try
                {

                    result = authContext.AcquireTokenAsync("https://graph.microsoft.com",
                        new ClientCredential(clientId, appKey)).Result;
                }
                catch (AdalException ex)
                {
                    if (ex.ErrorCode == "temporarily_unavailable")
                    {
                        retry = true;
                        retryCount++;
                        Thread.Sleep(3000);
                    }
                }
            } while ((retry == true) && (retryCount < 3));


            //log.Info(result == null ? "Cancelling attempt..." : "authenticated successfully.. ");

            return result;
        }

    }
}
