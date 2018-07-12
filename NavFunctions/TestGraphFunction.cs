using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using FunctionHelpers;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Functions
{
    public static class TestGraphFunction
    {
        /// <summary>
        /// A simple function that makes a call to Microsoft Graph
        /// </summary>
        /// <param name="req"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        [FunctionName("TestGraphFunction")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, ILogger logger)
        {
            logger.LogInformation("TestGraphFunction HTTP trigger function processed a request.");

            var keyVaultUrl = Environment.GetEnvironmentVariable("KEYVAULT");
            var aadInstance = Environment.GetEnvironmentVariable("ida:AADInstance");
            var tenant = Environment.GetEnvironmentVariable("ida:Tenant");
            //These two get retrieved from the key vault
            var clientIdKey = Environment.GetEnvironmentVariable("secretClientIdKey");
            var appKeyKey = Environment.GetEnvironmentVariable("secretAppKeyKey");
            
            //Get the access token
            var token = HelperGraph.GetAppOnlyToken(clientIdKey, appKeyKey, tenant, aadInstance, keyVaultUrl);

            //Make a call to the graph to retrieve some data
            var result = await GetDriveContents(token,logger);

        return result == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Couldn't retrieve data from the Microsoft Graph")
                : req.CreateResponse(HttpStatusCode.OK, result.ToString());
        }


        public static async Task<string> GetDriveContents(AuthenticationResult result, ILogger logger)
        {
            HttpClient httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", result.AccessToken);

            HttpResponseMessage response = await httpClient.GetAsync("https://graph.microsoft.com/v1.0/users");

            var graphserviceClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", result.AccessToken);

                        return Task.FromResult(0);
                    }));


            var driveId=Environment.GetEnvironmentVariable("sharePointDriveId");

            IDriveItemChildrenCollectionPage driveItems = await graphserviceClient.Drives[driveId].Root.Children.Request().GetAsync();
            
            if (response.IsSuccessStatusCode)
            {
                string r = await response.Content.ReadAsStringAsync();
                logger.LogInformation(r);
                return r;
            }
            else
            {
                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    //authContext.TokenCache.Clear();
                }

                logger.LogError("Access Denied");
                return "Access Denied";
            }
        }
    }
}
