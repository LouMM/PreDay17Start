// This is a C# Script file, you can copy and paste it directly into the Azure Function Portal Editor

//This syntax is similar to adding a reference in Visual Studio
//This is looking into the /BIN directory of an Azure Function for your assembly.
//More Info: https://aka.ms/AzFuncRefAssem
#r "Newtonsoft.Json"
#r "Microsoft.Graph"
#r "D:\home\site\wwwroot\bin\Microsoft.Graph.Core.dll" //TODO: this is temporary to reference the assembly for the pre-release template.
#r "System.Linq.Expressions"
using System.Net; 
using System.Net.Http; 
using System.Net.Http.Headers; 
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;


// Welcome Ignite Pre Day - this example uses the Microsoft Graph SDK 
public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, string graphToken, TraceWriter log)
{
	//Log some information about the caller.
    log.Info($"Request Successful [Add your own tracing]");    

    var proxy = GetClientInstance(graphToken,log);

    //This is the key line it uses a Client Proxy to give you a Object model and Request builder to work against.
    //In this example, it is getting the logged in users emails, that were received between a specific date and has attachments
    //Notice that this is also using the 'Expand' functionality for relationships, and actually pulling down the attachment data along witht his call.
    //The Select() is also asking for All Attributes, but you can scope it down to the attributes of the mail you want.
    var values = await proxy.Me.MailFolders.Inbox.Messages.Request().Select("").Filter("ReceivedDateTime ge 2017-09-23 and hasAttachments eq true").Expand("Attachments").GetAsync();

    //you could wrao the GetAsync method with a Try/Catch and in the catch, return another HTTP status code, with whatever
    //error you want to send the system that is calling this HTTPTrigger
    HttpResponseMessage s = new HttpResponseMessage(System.Net.HttpStatusCode.OK);
    s.Content = new StringContent(JsonConvert.SerializeObject(values));
    return s;
}


//This method is just a wrapper around the instantiation of the Client Proxy Object
//It passes in the graph token the azure function recieved, in order to create the proxy.
public static GraphServiceClient GetClientInstance(string graphToken, TraceWriter log)
{
	GraphServiceClient _graphClient = null;

		try
		{
			//More details here: https://github.com/microsoftgraph/msgraph-sdk-dotnet/blob/dev/docs/overview.md
			_graphClient = new GraphServiceClient(
				"https://graph.microsoft.com/v1.0",
				new DelegateAuthenticationProvider(
					 async (requestMessage) =>
					{
						
						requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", graphToken);
								

					}));
			return _graphClient;
		}
		catch (Exception ex)
		{
			log.Error(ex.Message);
		}
	

	return _graphClient;
}