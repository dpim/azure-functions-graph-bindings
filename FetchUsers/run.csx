#r "Newtonsoft.Json"
#r "Microsoft.Graph"
#r "D:\home\site\wwwroot\bin\Microsoft.Graph.Core.dll"

using System.Net; 
using System.Net.Http; 
using System.Net.Http.Headers; 
using Newtonsoft.Json;
using Microsoft.Graph;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, string graphToken, TraceWriter log)
{
    //Set up Graph client
    GraphServiceClient graphClient = new GraphServiceClient(
        "https://graph.microsoft.com/v1.0",
        new DelegateAuthenticationProvider(
            async (requestMessage) =>
            {
                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", graphToken);
            }
        )
    );
    
    //Fetch users
    List<Recipient> recipientList = new List<Recipient>();    
    IGraphServiceUsersCollectionPage users = null;
    try 
    {
        users = await graphClient.Users.Request().GetAsync();
    }
    catch (ServiceException e)
    {
        log.Info("Experienced an issue when fetching users: " + e.Error.Message);
    }
    //Iterate through users, see if they have a photo
    foreach (User user in users)
    {
        //filter out non-users in AD
        if (user.GivenName != null)
        {
            //check for photo
            try 
            {
                var photo = await graphClient.Users[user.Id].Photo.Request().GetAsync();
            }  
            catch (ServiceException e)
            {
                
                recipientList.Add(
                    new Recipient
                    {
                        EmailAddress = new EmailAddress 
                        {
                            Address = user.Mail
                        }
                    }
                );
            }
        }
    }

    //Send email to those that don't
    var email = new Message
    {
        Body = new ItemBody
        {
            Content = "",
            ContentType = BodyType.Html,
        },
        Subject = "Hello",
        BccRecipients = recipientList,
    };
    
    try
    {
        await graphClient.Me.SendMail(email, true).Request().PostAsync();
    } 
    catch (ServiceException e)
    {
        log.Info("Experienced an issue when sending mail: " + e.Error.Message);
    }

    //return recipient list contacted about lack of photos
    HttpResponseMessage response = new HttpResponseMessage();
    string serialized = JsonConvert.SerializeObject(recipientList, Formatting.Indented);
    response.Content = new StringContent(serialized, System.Text.Encoding.UTF8, "application/json");
    return response;
}