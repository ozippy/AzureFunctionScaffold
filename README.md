I plan to start using Azure functions, Microsoft Graph, SharePoint Online and web hooks to integrate with my applications.

I have read a lot of blogs, watched a lot of training videos, watched conference sessions and looked at sample code, but nothing seemed to pull together my basic requirements.

These are;
<ol>
	<li>Securely manage all secrets</li>
	<li>Securely connect to SharePoint Online without a user context</li>
	<li>Securely connect to Microsoft Graph without a user context</li>
	<li>Log events and telemetry in a central location</li>
</ol>
You might think this would be simple, but that was not my experience.

I have created a GitHub repo with the scaffold of my project which you are welcome to refer to. <a href="https://github.com/ozippy/AzureFunctionScaffold">https://github.com/ozippy/AzureFunctionScaffold</a>

<img class=" size-full wp-image-204 aligncenter" src="https://ozippy.files.wordpress.com/2018/07/scaffoldsolution.png" alt="ScaffoldSolution" width="281" height="464" />

I'll go through each requirement and reference some of the articles I used.
<h2>Securely managing secrets</h2>
Microsoft recently GA'd the Azure Managed Service Identity.

<a href="https://docs.microsoft.com/en-us/azure/app-service/app-service-managed-service-identity">https://docs.microsoft.com/en-us/azure/app-service/app-service-managed-service-identity</a>

This allows us to connect to Azure resources without needing to manage the authentication to those resources in our code, which is awesome.

It is important that secrets are not stored in our source code and potentially in our source control systems. I don't want developers to have to be concerned with or even have direct access to those secrets.

The scaffold project abstracts all secrets away, so that in the functions, it uses environment variables to reference the key of the secret stored in Azure Key Vault. No keys or other secrets are stored in the source or in the function.
<h2>Securely Connect to SharePoint Online</h2>
In order to be able to have app-only access to SharePoint Online for use with web hooks and bots, I followed this guide for using a certificate to gain secure access.

<a href="https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread">https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread</a>

This approach by itself, did not quite achieve requirement number 1 for me.

I wanted to be able to store all secrets including my serialised PFX certificate in the Azure Key Vault.

This is where things got tricky for me. The dependencies for nuget packages started to clash.

I had planned to use the Authentication Manager in SharePoint PNP Sites Core nuget package. This however seems to have a dependency on ADAL v2.29.

I was receiving exceptions saying that the function "Could not load file or assembly 'Microsoft.IdentityModel.Clients.ActiveDirectory, Version=2.29.0.1078".

It turned out that the Managed Service Identity package Microsoft.Azure.Services.AppAuthentication requires ADAL 3.19.

So to cut a long story short, I adapted one of the methods in SharePoint PNP Sites Core to work with ADAL v3.19.
<pre><code>    public static async Task GetAzureADAppOnlyAuthenticatedContext(string siteUrl, string clientId,
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
    }</code></pre>
This allowed me to authenticate using my deserialised certificate which I retrieve from Key Vault. Strangely I have to store it as a secret rather than a Key Vault certificate. As I understand it, the certificate store in Key Vault won't allow the retrieval of a PFX.

This was another useful reference for serialising and deserialisng my certificate;

<a href="https://stackoverflow.com/questions/33728213/how-to-serialize-and-deserialize-a-pfx-certificate-in-azure-key-vault">https://stackoverflow.com/questions/33728213/how-to-serialize-and-deserialize-a-pfx-certificate-in-azure-key-vault</a>
<h2>Securely connect to Microsoft Graph without a user context</h2>
This one seemed a bit simpler. I could register the app in Azure AD and use the client id and secret to authenticate with Azure AD and access the Graph.

For consistency though, I store the secrets in Key Vault.
<h2>Log events and telemetry in a central location</h2>
Application Insights has direct integration into Azure Functions. I have to admit though that something as simple as finding out how to access the logs through the UI was done.

As it turned out, it was through the Search blade in AppInsights. There you can list and filter the logs as you please.

<img class="alignnone size-full wp-image-203" src="https://ozippy.files.wordpress.com/2018/07/appinsightslogs.png" alt="AppInsightsLogs" width="1095" height="870" />
<h2>Running the code</h2>
I used Visual Studio 2017 to build and test my functions.

You'll need a local.settings.json to configure your environment variables.

This is what I am currently using to test my scaffold;

{
"IsEncrypted": false,
"Values": {
"APPINSIGHTS_INSTRUMENTATIONKEY": "",
"AzureWebJobsStorage": "UseDevelopmentStorage=true",
"AzureWebJobsDashboard": "UseDevelopmentStorage=true",
"ida:AADInstance": "https://login.microsoftonline.com/{0}",
"ida:SharePointInstance": "https://login.windows.net/{0}/",
"ida:Tenant": "", (e.g. <your tenant>.onmicrosoft.com)
"KEYVAULT": "https://<your key vault>.vault.azure.net/secrets/",
"secretClientIdKey": "<your secret key>",
"secretAppKeyKey": "<your secret key>",
"secretCertClientIdKey": "<your secret key>",
"sharepointSite": "https://.sharepoint.com/sites/",
"secretCertName": "<your secret key>",
"webHookEndPoint": <your webhook endpoint>,
"sharePointDriveId": "<the id of the SharePoint drive to query>"
}
}

Things you'll need to do to run the code:
<ul>
	<li>Create an Azure Key Vault</li>
	<li>Create an Azure Function app</li>
	<li>Configure Application Insights for your function app</li>
	<li>Clone the repo</li>
	<li>Add the various secrets to the Key Vault and add the certificate using the powershell scripts</li>
	<li>Setup the Azure AD applications and grant permissions</li>
	<li>Grant permissions to teh SharePoint site collection</li>
	<li>F5</li>
</ul>
 

One thing to watch out for. I spent ages trying to work out why I was getting a "Could not load file or assembly 'Microsoft.IdentityModel.Clients.ActiveDirectory, Version=3.14.2" exception. Then I discovered it worked okay on one of my other PCs. It turned out that I needed to download the Azure CLI and do an 'az login'. After doing that it worked. If I did an 'az logout' it failed again.

I'm hoping this might help others who have similar requirements to me and might reduce the amount of research and experimentation needed.

I anticipate that there are various optimisations that could be made to the code and possibly completely different ways to approach it, so feel free to make suggestions. I do like to keep it simple though.
