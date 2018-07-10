using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace FunctionHelpers
{
    public static class HelperSecrets
    {
        /// <summary>
        /// Get a simple string secret value from the specified Azure Key Vault
        /// </summary>
        /// <param name="secretKey"></param>
        /// <param name="keyVaultUrl"></param>
        /// <returns></returns>
        public static async Task<string> GetSecretString(string secretKey, string keyVaultUrl)
        {
 
            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            string message = null;

            try
            {
                var keyVaultClient = new KeyVaultClient(
                    new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));

                var secret = await keyVaultClient.GetSecretAsync(keyVaultUrl + secretKey)
                    .ConfigureAwait(false);

                message = secret.Value;
            }
            catch (Exception ex)
            {
                //log Error
                throw;
            }

            return message;
        }



        /// <summary>
        /// Get a serialised PFX certificate from a secret in the Azure Key Vault
        /// </summary>
        /// <param name="keyVaultUrl"></param>
        /// <param name="certName"></param>
        /// <returns></returns>
        public static async Task<X509Certificate2> GetCertificate(string keyVaultUrl,
            string certName)
        {
            var azureServiceTokenProvider = new AzureServiceTokenProvider();

            X509Certificate2 certificate = null;

            try
            {

                var keyVaultClient = new KeyVaultClient(
                    new KeyVaultClient.AuthenticationCallback(azureServiceTokenProvider.KeyVaultTokenCallback));

                var secretRetrieved = await keyVaultClient.GetSecretAsync(keyVaultUrl + certName);
                var pfxBytes = Convert.FromBase64String(secretRetrieved.Value);
                var x509Certificate2Collection = new X509Certificate2Collection();
                x509Certificate2Collection.Import(pfxBytes, "", X509KeyStorageFlags.Exportable | X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet);
                certificate = x509Certificate2Collection[0];
            }
            catch (Exception ex)
            {
                //log error
                throw;
            }

            return certificate;
        }
    }
}
