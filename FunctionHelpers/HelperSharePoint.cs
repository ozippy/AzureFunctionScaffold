using System;
using System.Globalization;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;

namespace FunctionHelpers
{
    public static class HelperSharePoint
    {
        /// <summary>
        /// Get a context by retrieving the serialized certificate collection from Azure Key Vault
        /// </summary>
        /// <param name="tenant"></param>
        /// <param name="siteUrl"></param>
        /// <param name="clientIdEnv"></param>
        /// <param name="keyVaultUrl"></param>
        /// <param name="certName"></param>
        /// <returns></returns>
        public static ClientContext GetClientContext(string tenant, string siteUrl, string clientIdEnv,
            string keyVaultUrl, string certName, ILogger logger)
        {
            var clientId = HelperSecrets.GetSecretString(clientIdEnv, keyVaultUrl, logger).Result; ;

            var certificate = HelperSecrets.GetCertificate(keyVaultUrl, certName, logger).Result;
            ClientContext ctx = null;

            ctx = GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, certificate).Result;

            return ctx;
        }

        // <summary>
        // Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        // </summary>
        // <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        // <param name="clientId">The Azure AD Application Client ID</param>
        // <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        // <param name="certificate">Certificate used to authenticate</param>
        // <returns></returns>
        public static async Task<ClientContext> GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId,
            string tenant, X509Certificate2 certificate)
        {
            var clientContext = new ClientContext(siteUrl);
            var authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", "https://login.windows.net", tenant);
            var authContext = new AuthenticationContext(authority);
            var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);
            var host = new Uri(siteUrl);
            var ar = await authContext.AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate);

            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + ar.AccessToken;
            };

            return clientContext;
        }

        /// <summary>
        /// Get an authentication result by retrieving the serialized certificate collection from Azure Key Vault
        /// </summary>
        /// <param name="tenant"></param>
        /// <param name="siteUrl"></param>
        /// <param name="clientIdEnv"></param>
        /// <param name="keyVaultUrl"></param>
        /// <param name="certName"></param>
        /// <returns></returns>
        public static AuthenticationResult GetAuthenticationResult(string tenant, string siteUrl, string clientIdEnv,
            string keyVaultUrl, string certName, ILogger logger)
        {
            var clientId = HelperSecrets.GetSecretString(clientIdEnv, keyVaultUrl, logger).Result; ;

            var certificate = HelperSecrets.GetCertificate(keyVaultUrl, certName, logger).Result;
            AuthenticationResult ar = null;

            ar = GetAzureAdAppOnlyAccessToken(siteUrl, clientId, tenant, certificate, logger).Result;

            return ar;
        }


        // <summary>
        // Returns a SharePoint Authentication Result using Azure Active Directory App Only Authentication. This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        // </summary>
        // <param name="siteUrl">Site for which the ClientContext object will be instantiated</param>
        // <param name="clientId">The Azure AD Application Client ID</param>
        // <param name="tenant">The Azure AD Tenant, e.g. mycompany.onmicrosoft.com</param>
        // <param name="certificate">Certificate used to authenticate</param>
        // <returns></returns>
        public static async Task<AuthenticationResult> GetAzureAdAppOnlyAccessToken(string siteUrl, string clientId,
            string tenant, X509Certificate2 certificate, ILogger logger)
        {
            var clientContext = new ClientContext(siteUrl);
            var authority = string.Format(CultureInfo.InvariantCulture, "{0}/{1}/", "https://login.windows.net", tenant);
            var authContext = new AuthenticationContext(authority);
            var clientAssertionCertificate = new ClientAssertionCertificate(clientId, certificate);
            var host = new Uri(siteUrl);
            var ar = await authContext.AcquireTokenAsync(host.Scheme + "://" + host.Host + "/", clientAssertionCertificate);

            return ar;
        }

        /// <summary>
        /// Returns a SharePoint ClientContext using Azure Active Directory App Only Authentication.This requires that you have a certificated created, and updated the key credentials key in the application manifest in the azure AD accordingly.
        /// </summary>
        /// <param name = "siteUrl" > Site for which the ClientContext object will be instantiated</param>
        /// <param name = "clientId" > The Azure AD Application Client ID</param>
        /// <param name = "tenant" > The Azure AD Tenant, e.g.mycompany.onmicrosoft.com</param>
        /// <param name = "certificatePath" > The path to the certificate (*.pfx) file on the file system</param>
        /// <param name = "certificatePassword" > Password to the certificate</param>
        /// <returns>Client context object</returns>
        public static ClientContext GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId, string tenant, string certificatePath, SecureString certificatePassword, ILogger logger)
        {
            var certfile = System.IO.File.OpenRead(certificatePath);
            var certificateBytes = new byte[certfile.Length];
            certfile.Read(certificateBytes, 0, (int)certfile.Length);
            var cert = new X509Certificate2(
                certificateBytes,
                certificatePassword,
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.PersistKeySet);

            return GetAzureADAppOnlyAuthenticatedContext(siteUrl, clientId, tenant, cert).Result;
        }
    }
}