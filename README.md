# Microsoft Graph bindings for Azure Functions - Connect() 2017 demo
## About
[Microsoft Graph bindings for Azure Functions](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-microsoft-graph) enables development of serverless solutions that integrate with personal and work/school data in [Microsoft Graph](https://graph.microsoft.com). This demo covers an [auth token input binding](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-microsoft-graph#auth-token-input-binding) that provides access to the Graph token that was included in the request to the Azure function. This function leverages request builders from the [.NET SDK for Microsoft Graph](https://github.com/microsoftgraph/msgraph-sdk-dotnet). Please note that these bindings are in *beta* and not intended for production use.

When triggered by the tenant admin, 'FetchUsers' function does the following:
1) Iterates over users in an organization
2) Checks if any users are missing profile photos 
3) Emails them a reminder to set one.

## Set-up
1) Create a new Function App in Azure Portal 
2) [Configure Function App to support beta functionality](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-microsoft-graph)
3) Upload the function definition into the Azure Portal
4) Navigate to the HTTP trigger URL *as the tenant admin*
5) Verify that the function is running succesfully via the logs
