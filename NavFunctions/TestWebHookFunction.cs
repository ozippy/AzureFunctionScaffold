using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using FunctionHelpers;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;

namespace NavFunctions
{
    public static class WebHooks
    {
        [FunctionName("TestWebHookFunction")]
        public static async Task<object> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
            HttpRequestMessage req,
            TraceWriter log)
        {
            log.Info($"Webhook was triggered!");

            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;

            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new 
            // web hook is being added
            if (validationToken != null)
            {
                log.Info($"Validation token {validationToken} received");
                var response = req.CreateResponse(HttpStatusCode.OK);
                response.Content = new StringContent(validationToken);
                return response;
            }

            log.Info($"SharePoint triggered our webhook...great :-)");
            var content = await req.Content.ReadAsStringAsync();
            log.Info($"Received following payload: {content}");

            var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
            log.Info($"Found {notifications.Count} notifications");

            var storageConnection =
                FunctionHelpers.HelperSecrets.GetSecretString(Environment.GetEnvironmentVariable("storageConnection"),Environment.GetEnvironmentVariable("KEYVAULT")).Result;

            if (notifications.Count > 0)
            {
                log.Info($"Processing notifications...");
                foreach (var notification in notifications)
                {
                    CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageConnection);
                    // Get queue... create if does not exist.
                    CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                    CloudQueue queue = queueClient.GetQueueReference("sharepointlistwebhookeventazuread");
                    queue.CreateIfNotExists();

                    // add message to the queue
                    string message = JsonConvert.SerializeObject(notification);
                    log.Info($"Before adding a message to the queue. Message content: {message}");
                    queue.AddMessage(new CloudQueueMessage(message));
                    log.Info($"Message added :-)");
                }
            }

            // if we get here we assume the request was well received
            return req.CreateResponse(HttpStatusCode.OK);
        }
    }
}