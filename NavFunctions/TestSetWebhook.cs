using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using FunctionHelpers;

namespace Functions
{
    public static class TestSetWebhook
    {
        [FunctionName("TestSetWebhook")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");


            var siteId = Environment.GetEnvironmentVariable("NavSite");
            var listId = Environment.GetEnvironmentVariable("NavNodesListId");
            var tenant = Environment.GetEnvironmentVariable("ida:Tenant");
            var clientId = Environment.GetEnvironmentVariable("secretCertClientIdKey");
            var keyVaultUrl = Environment.GetEnvironmentVariable("KEYVAULT");
            var certName = Environment.GetEnvironmentVariable("secretCertName");
            string webHookEndPoint="https://navfunctionappozippy.azurewebsites.net/api/TestWebHookFunction?code=ZJAx3rX1ryXifutlp1dn2KvkF8ajamEjAjNz7YjdvnExoWFO18OtJA==";
            
            var authenticationResult = HelperSharePoint.GetAuthenticationResult(tenant, siteId, clientId, keyVaultUrl, certName);

            var result = await HelperWebHooks.AddListWebHookAsync(siteId, listId, webHookEndPoint, authenticationResult.AccessToken, 3);
            //var result = "test";
            return result == null
                 ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                 : req.CreateResponse(HttpStatusCode.OK, result);
        }
    }
}
