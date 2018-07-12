using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using FunctionHelpers;
using Microsoft.Extensions.Logging;

namespace Functions
{
    public static class TestSetWebhook
    {
        /// <summary>
        /// A function that will register a web hook against a defined list in SharePoint Online
        /// </summary>
        /// <param name="req"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        [FunctionName("TestSetWebhook")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, ILogger logger)
        {
            logger.LogInformation("C# HTTP trigger function processed a request for testing the web hook.");


            var siteId = Environment.GetEnvironmentVariable("sharePointSite");
            var listId = Environment.GetEnvironmentVariable("sharePointListId");
            var tenant = Environment.GetEnvironmentVariable("ida:Tenant");
            var webHookEndPoint=Environment.GetEnvironmentVariable("webHookEndPoint");
            
            var keyVaultUrl = Environment.GetEnvironmentVariable("KEYVAULT");
            //Get these from Key Vault
            var clientId = Environment.GetEnvironmentVariable("secretCertClientIdKey");
            var certName = Environment.GetEnvironmentVariable("secretCertName");
            
            //Authenticate with SharePoint using the certificate
            var authenticationResult = HelperSharePoint.GetAuthenticationResult(tenant, siteId, clientId, keyVaultUrl, certName, logger);

            var result = await HelperWebHooks.AddListWebHookAsync(siteId, listId, webHookEndPoint, authenticationResult.AccessToken, 3);

            return result == null
                 ? req.CreateResponse(HttpStatusCode.BadRequest, "Couldn't register the web hook")
                 : req.CreateResponse(HttpStatusCode.OK, result);
        }
    }
}
