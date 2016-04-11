
# Project Information
#### This forked Version of ASP.NET MVC sample provides support on subscribing for notifications on specific Mail Folders of logged in user. 


![Preview](readme-images/specific%20mail%20folder%20subscribe.png)

#### Resource value for subscribing to specifc mailFolder

 ```
"me/mailFolders/<MailFolderId>/messages"
 ```
 
#### Sources

* [Microsoft Graph WebHooks Update - January 2016](http://officedevcenter-msprod.azurewebsites.net/Contents/Item/Display/7206)

# Below THIS is the original documentation 




# Microsoft Graph ASP.NET Webhooks

Subscribe for webhooks to get notified when your user's data changes so you don't have to poll for changes.

This ASP.NET MVC sample shows how to start getting notifications from Microsoft Graph. [Microsoft Graph](http://graph.microsoft.io/) provides a unified API endpoint to access data from the Microsoft cloud.

The following are common tasks that a web application performs with Microsoft Graph webhooks.

* Sign-in your users with their work or school account to get an access token.
* Use the access token to create a webhook [subscription](http://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/subscription).
* Send back a validation token to confirm the notification URL.
* Listen for notifications from Microsoft Graph.
* Request for more information in Microsoft Office 365 using data in the notification.
  
![Microsoft Graph Webhook Sample for ASP.NET screenshot](/readme-images/Page1.PNG)

The previous screenshot shows the app's start page. After the app creates a subscription on behalf of the signed-in user, Microsoft Graph sends a notification to the registered endpoint when events happen in the user's data. The app then reacts to the event.

This sample subscribes to the `me/mailFolders('Inbox')/messages` resource for `created` changes. It gets notified when the user receives a mail message, and then updates a page with information about the message. 

## Prerequisites

To use the Microsoft Graph ASP.NET Webhooks sample, you need the following:

* Visual Studio 2015 installed on your development computer. 

* A public HTTPS endpoint to receive and send HTTP requests. You can host this on Microsoft Azure or another service, or you can [use ngrok](#ngrok) or a similar tool while testing.

* The client ID and key from the application that you registered on a Microsoft Azure tenant. Use the [Office 365 app registration tool](http://dev.office.com/app-registration) to quickly register an app with the following parameters:

   |       Parameter | Value                    |
   |----------------:|:-------------------------|
   |        App name | \<any name\>             |
   |        App type | Web App                  |
   |     Sign on URL | https://localhost:44300/ |
   |    Redirect URI | https://localhost:44300/ |
   | App permissions | Mail.Read                |
  
   Copy and store the returned **Client ID** and **Client Secret** values. (Learn more about [setting up your development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment?aspnet).)


<a name="ngrok"></a>
### Set up the ngrok proxy (optional) 
You must expose a public HTTPS endpoint to create a subscription and receive notifications from Microsoft Graph. While testing, you can use ngrok to temporarily allow messages from Microsoft Graph to tunnel to a *localhost* port on your computer. To learn more about using ngrok, see the [ngrok website](https://ngrok.com/).  

1. In Solution Explorer, select the **GraphWebhooks** project.

1. Copy the **URL** port number from the **Properties** window.  If the **Properties** window isn't showing, choose **View > Properties Window**. 

	![The URL port number in the Properties window](readme-images/PortNumber.png)

1. [Download ngrok](https://ngrok.com/download) for Windows.  

1. Unzip the package and run ngrok.exe.

1. Replace the two *\<port-number\>* placeholder values in the following command with the port number you copied, and then run the command in the ngrok console.

   ```
ngrok http <port-number> -host-header=localhost:<port-number>
   ```

	![Example command to run in the ngrok console](readme-images/ngrok1.PNG)

1. Copy the HTTPS URL that's shown in the console. You'll use this to configure your notification URL in the sample.

	![The forwarding HTTPS URL in the ngrok console](readme-images/ngrok2.PNG)

   >Keep the console open while testing. If you close it, the tunnel also closes and you'll need to generate a new URL and update the sample.

See [Hosting without a tunnel](https://github.com/OfficeDev/Microsoft-Graph-Nodejs-Webhooks/wiki/Hosting-the-sample-without-a-tunnel) and [Why do I have to use a tunnel?](https://github.com/OfficeDev/Microsoft-Graph-Nodejs-Webhooks/wiki/Why-do-I-have-to-use-a-tunnel) for more information.


## Configure and run the sample

1. Expose a public HTTPS notification endpoint. It can run on a service such as Microsoft Azure, or you can create a proxy web server by [using ngrok](#ngrok) or a similar tool.

1. Open **GraphWebhooks.sln** in the sample files. 

   >You may be prompted to trust certificates for localhost.

1. In Solution Explorer, open the **Web.config** file in the root directory of the project.  
   a. For the **ClientId** key, replace *ENTER_YOUR_CLIENT_ID* with the client ID of your registered Azure application.  
 
   b. For the **ClientSecret** key, replace *ENTER_YOUR_SECRET* with the key of your registered Azure application.  

   c. For the **NotificationUrl** key, replace *ENTER_YOUR_URL* with the HTTPS URL. Keep the */notification/listen* portion. If you're using ngrok, use the HTTPS URL that you copied. The value will look something like this:

   ```xml
<add key="ida:NotificationUrl" value="https://0f6fd138.ngrok.io/notification/listen" />
   ```

1. Make sure that the ngrok console is still running, then press F5 to build and run the solution in debug mode. 


### Use the app
 
1. Sign in with your Office 365 work or school account. 

1. Choose the **Create subscription** button. The **Subscription** page loads with information about the subscription.

	![App page showing properties of the new subscription](readme-images/Page2.PNG)
	
1. Choose the **Watch for notifications** button.

1. Send an email to your Office 365 account. The **Notification** page displays some message properties. It may take several seconds for the page to update.
   
	![App page showing properties of the new message](readme-images/Page3.PNG)

1. Choose the **Delete subscription and sign out** button. 


## Key components of the sample

**Controllers**  
- [`NotificationController.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Controllers/NotificationController.cs) Receives notifications.  
- [`SubscriptionContoller.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Controllers/SubscriptionController.cs) Creates and receives webhook subscriptions.
 
**Models**  
- [`Message.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Models/Message.cs) Represents an Outlook mail message. 
- [`Notification.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Models/Notification.cs) Represents a change notification. 
- [`Subscription.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Models/Subscription.cs) Represents a webhook subscription. Also defines the **SubscriptionViewModel** that represents the data displayed in the Subscription view. 

**Views**  
- [`Notification/Notification.cshtml`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Views/Notification/Notification.cshtml) Displays information about received messages, and contains the **Delete subscription and sign out** button. 
- [`Subscription/Index.cshtml`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Views/Subscription/Index.cshtml) Landing page that contains the **Create subscription** button. 
- [`Subscription/Subscription.cshtml`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Views/Subscription/Subscription.cshtml) Displays subscription properties, and contains the **Watch for notifications** button. 

**Other**  
- [`Web.config`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Web.config) Contains values used for authentication and authorization. 
- [`Startup.Auth.cs`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/App_Start/Startup.Auth.cs) and [`Controllers/Utils/AuthHelper`](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/blob/master/GraphWebhooks/Controllers/Utils/AuthHelper.cs) Contain code used for authentication and authorization. The sample uses [OpenID Connect](https://msdn.microsoft.com/en-us/library/azure/dn645541.aspx) and [Active Directory Authentication Library .NET (v2)](http://go.microsoft.com/fwlink?LinkId=258232) to authenticate and authorize the user.


## Troubleshooting

| Issue | Resolution |
|:------|:------|
| The app opens to a *Server Error in '/' Application. The resource cannot be found.* browser page. | Make sure that a CSHTML view file isn't the active tab when you run the app from Visual Studio. |
| You're using ngrok and get a *Subscription validation request timed out* response. | Make sure that you used your project's HTTP port for the tunnel (not HTTPS). |


## Questions and comments

We'd love to get your feedback about the Microsoft Graph ASP.NET Webhooks sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/Microsoft-Graph-ASPNET-Webhooks/issues) section of this repository.

Questions about Microsoft Graph in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/MicrosoftGraph). Make sure that your questions or comments are tagged with *[MicrosoftGraph]*.

You can suggest changes for Microsoft Graph on [GitHub](https://github.com/OfficeDev/microsoft-graph-docs).
  

## Additional resources

* [Microsoft Graph Node.js Webhooks sample](https://github.com/OfficeDev/Microsoft-Graph-Nodejs-Webhooks)
* [Subscription resource](http://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/subscription)
* [Microsoft Graph documentation](http://graph.microsoft.io/)
* [Call Microsoft Graph in an ASP.NET MVC app](https://graph.microsoft.io/en-us/docs/platform/aspnetmvc)
* [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment?aspnet)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.

