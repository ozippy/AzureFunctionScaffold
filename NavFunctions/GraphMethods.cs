//using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;


namespace Functions
{
    public static class GraphMethods
    {
        public static async Task<string> MakeHttpsCall(AuthenticationResult result, TraceWriter log)
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

            IDriveItemChildrenCollectionPage driveItems = await graphserviceClient.Drives["b!Rt1723NZBEiSae5CBqlEl4ZRGeBfvo5Oj3jHTKrvSFMjBNezNBi7SoiiSHw9H21J"].Root.Children.Request().GetAsync();

         
            if (response.IsSuccessStatusCode)
            {
                string r = await response.Content.ReadAsStringAsync();
                log.Info(r);
                return r;
            }
            else
            {
                if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                {
                    //authContext.TokenCache.Clear();
                }

                log.Info("Access Denied");
                return "Access Denied";
            }
        }
    }
}
