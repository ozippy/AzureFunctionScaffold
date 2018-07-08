using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using FunctionHelpers;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

namespace Functions
{
    public static class TestGraphFunction
    {
        [FunctionName("TestGraphFunction")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            var keyVaultUrl = Environment.GetEnvironmentVariable("KEYVAULT");
            var aadInstance = Environment.GetEnvironmentVariable("ida:AADInstance");
            var tenant = Environment.GetEnvironmentVariable("ida:Tenant");
            //These two get retrieved from the keyvault
            var clientIdKey = Environment.GetEnvironmentVariable("secretClientIdKey");
            var appKeyKey = Environment.GetEnvironmentVariable("secretAppKeyKey");
            
            var token = HelperGraph.GetAppOnlyToken(clientIdKey, appKeyKey, tenant, aadInstance, keyVaultUrl);

            var result = await GraphMethods.MakeHttpsCall(token,log);

        return result == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, result.ToString());
        }
    }
}
