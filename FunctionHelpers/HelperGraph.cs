using System;
using System.Globalization;
using System.Threading;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Extensions.Logging;

namespace FunctionHelpers
{
    public static class HelperGraph
    {
        /// <summary>
        /// Get an apponly token to be able to access the graph
        /// </summary>
        /// <param name="clientIdKey"></param>
        /// <param name="appKeyKey"></param>
        /// <param name="tenant"></param>
        /// <param name="aadInstance"></param>
        /// <param name="keyVaultUrl"></param>
        /// <returns></returns>
        public static AuthenticationResult GetAppOnlyToken(string clientIdKey, string appKeyKey, string tenant, string aadInstance, string keyVaultUrl, ILogger logger)
        {
            logger.LogInformation("calling GetAppOnlyToken");

            //Get these from Key Vault
            var clientId = HelperSecrets.GetSecretString(clientIdKey, keyVaultUrl, logger).Result;
            var appKey = HelperSecrets.GetSecretString(appKeyKey, keyVaultUrl, logger).Result;

            var authority = String.Format(CultureInfo.InvariantCulture, aadInstance ?? throw new InvalidOperationException("aadInstance is not specified"), tenant);

            var authContext = new AuthenticationContext(authority);

            AuthenticationResult result = null;
            var retryCount = 0;
            var retry = false;
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


            logger.LogInformation(result == null ? "Cancelling attempt..." : "authenticated successfully.. ");

            return result;
        }

    }
}
