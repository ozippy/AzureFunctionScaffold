using System;
using System.Net;
using System.Net.Http;
using FunctionHelpers;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;

namespace Functions
{
    public static class TestSharePointFunction
    {
        private static string key = TelemetryConfiguration.Active.InstrumentationKey = 
            System.Environment.GetEnvironmentVariable(
                "APPINSIGHTS_INSTRUMENTATIONKEY", EnvironmentVariableTarget.Process);

        private static TelemetryClient telemetryClient = new TelemetryClient() { InstrumentationKey = key };

        [FunctionName("TestSharePointFunction")]
        public static HttpResponseMessage Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, ExecutionContext context, ILogger logger)
        {
            logger.LogInformation("C# HTTP trigger function processed a request.");

            var siteId = Environment.GetEnvironmentVariable("sharePointSite");
            var tenant = Environment.GetEnvironmentVariable("ida:Tenant");
            var clientId = Environment.GetEnvironmentVariable("secretCertClientIdKey");
            var keyVaultUrl = Environment.GetEnvironmentVariable("KEYVAULT");
            var certName = Environment.GetEnvironmentVariable("secretCertName");

            //Some examples of logging to AppInsights
            logger.LogInformation("101 Azure Function Demo - Logging with ITraceWriter");
            logger.LogTrace("Here is a verbose log message");
            logger.LogWarning("Here is a warning log message");
            logger.LogError("Here is an error log message");
            logger.LogCritical("This is a critical log message => {message}", "We have a big problem");


            // Track an Event
            var evt = new EventTelemetry("Function called");
            UpdateTelemetryContext(evt.Context, context, "CertificateAppOnly");
            telemetryClient.TrackEvent(evt);


            var clientContext = HelperSharePoint.GetClientContext(tenant, siteId, clientId, keyVaultUrl, certName);

            Web ccWeb = clientContext.Web;
  
            clientContext.Load(ccWeb);
            clientContext.ExecuteQuery();

            logger.LogInformation("web title is " + ccWeb.Title);
            logger.LogMetric("TestMetric", 1234); 

            return clientContext == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "We couldn't get the title of the site")
                : req.CreateResponse(HttpStatusCode.OK, "The title of the web is " + clientContext.Web.Title);
        }

        
        // This correllates all telemetry with the current Function invocation
        private static void UpdateTelemetryContext(TelemetryContext context, ExecutionContext functionContext, string userName)
        {
            context.Operation.Id = functionContext.InvocationId.ToString();
            context.Operation.ParentId = functionContext.InvocationId.ToString();
            context.Operation.Name = functionContext.FunctionName;
            context.User.Id = userName;
        }
    }
}
