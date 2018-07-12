using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using FunctionHelpers;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;

namespace Functions
{
    public static class WebHooks
    {
        /// <summary>
        /// Once we have registered the web hook, this can act as the endpoint to call.
        /// </summary>
        /// <param name="req"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        [FunctionName("TestWebHookFunction")]
        public static async Task<object> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
            HttpRequestMessage req,
            ILogger logger)
        {
            logger.LogInformation($"Webhook was triggered!");

            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;

            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                logger.LogInformation($"Validation token {validationToken} received");
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            logger.LogInformation($"SharePoint triggered our webhook");
            var content = await req.Content.ReadAsStringAsync();
            logger.LogInformation($"Received following payload: {content}");

            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
            logger.LogInformation($"Found {notifications.Count} notifications");

            var storageConnection =
                FunctionHelpers.HelperSecrets.GetSecretString(Environment.GetEnvironmentVariable("storageConnection"),Environment.GetEnvironmentVariable("KEYVAULT")).Result;

            if (notifications.Count > 0)
            {
                logger.LogInformation($"Processing notifications...");
                foreach (var notification in notifications)
                {
                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageConnection);
                    // Get queue... create if does not exist.
                    CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    CloudQueue queue = queueClient.GetQueueReference("sharepointlistwebhookeventazuread");
                    queue.CreateIfNotExists();

                    // add message to the queue
                    var message = JsonConvert.SerializeObject(notification);
                    logger.LogInformation($"Before adding a message to the queue. Message content: {message}");
                    queue.AddMessage(new CloudQueueMessage(message));
                    logger.LogInformation($"Message added");
                }
            }

            // if we get here we assume the request was well received
            return req.CreateResponse(HttpStatusCode.OK);
        }
    }
}