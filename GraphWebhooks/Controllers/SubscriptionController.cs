/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  See LICENSE in the source repository root for complete license information.
 */

using System;
using System.Web;
using System.Collections.Generic;
using System.Web.Mvc;
using GraphWebhooks.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using GraphWebhooks.Utils;
using GraphWebhooks.Helpers;
namespace GraphWebhooks.Controllers
{
    public class SubscriptionController : Controller
    {
        private List<string> _excludedMailFolderList = new List<string>(new string[] { "Deleted Items", "Drafts", "Junk Email", "Outbox", "Sent Items" });

        [Authorize, HandleAdalException]
        public async Task<ActionResult> Index()
        {
            //TODO: Do We need to always get new token every request?

            // Get an access token and add it to the client. 
            // This sample stores the refreshToken, so get the AuthenticationResult that has the access token and refresh token.
            AuthenticationResult authResult = await AuthHelper.GetAccessTokenAsync();

            string resourceEndpoint = ConfigurationManager.AppSettings["ida:ResourceId"],
                   resourceValue = "/beta/me/mailFolders",
                   resource = resourceEndpoint + resourceValue;

            // Send the 'GET' request.
            HttpResponseMessage responseMailFolder = await GraphHelper.GetMailFolders(authResult.AccessToken, resource);s
            if (responseMailFolder.IsSuccessStatusCode)
            {

                // Parse the JSON response.
                string stringResult = await responseMailFolder.Content.ReadAsStringAsync();
                JObject jsonObject = JObject.Parse(stringResult);

                //get list of Mailfolder
                JArray mailFolderList = JArray.Parse(jsonObject["value"].ToString());

                List<SelectListItem> list = new List<SelectListItem>();

                foreach (var item in mailFolderList)
                {
                    MailFolder folder = JsonConvert.DeserializeObject<MailFolder>(item.ToString());
                    
                    //add selected folders
                    if (!isFolderExcluded(folder.displayName))
                    {
                        list.Add(new SelectListItem
                        {
                            Text = folder.displayName,
                            Value = folder.Id
                        });
                    }
                    
                }

                //set ViewBag from array of MailFolder object
                ViewBag.MailFolders = list;

            }
            else
            {// response status failed for get mailFolder Id
                return RedirectToAction("Index", "Error", new { message = responseMailFolder.StatusCode, debug = await responseMailFolder.Content.ReadAsStringAsync() });
            }

            return View();
        }

        
        [Authorize, HandleAdalException]
        public async Task<ActionResult> CreateSubscription(string mailFolderId)
        {
            string resourceEndpoint = ConfigurationManager.AppSettings["ida:ResourceId"],
                   resourceValue;

            // Get an access token and add it to the client. 
            // This sample stores the refreshToken, so get the AuthenticationResult that has the access token and refresh token.
            AuthenticationResult authResult = await AuthHelper.GetAccessTokenAsync();
           
            if (String.IsNullOrEmpty(mailFolderId))
            {
                //use default resource if no MailFolder was selected
                resourceValue = "me/mailFolders('Inbox')/messages";                
            }
            else
            {
                //use mailFolder Id
                resourceValue = String.Format("me/mailFolders/{0}/messages", mailFolderId);
            }
            
            var subscription = new Subscription
            {
                Resource = resourceValue,
                ChangeType = "created",
                NotificationUrl = ConfigurationManager.AppSettings["ida:NotificationUrl"],
                ClientState = Guid.NewGuid().ToString(),
                ExpirationDateTime = DateTime.UtcNow + new TimeSpan(3, 0, 0, 0)
            };

            HttpResponseMessage subscriptionResponse = await GraphHelper.CreateSubscription(authResult.AccessToken, subscription);

            if (subscriptionResponse.IsSuccessStatusCode)
            {

                // Parse the JSON response.
                string stringResult = await subscriptionResponse.Content.ReadAsStringAsync();
                SubscriptionViewModel viewModel = new SubscriptionViewModel
                {
                    Subscription = JsonConvert.DeserializeObject<Subscription>(stringResult)
                };

                // This app temporarily stores the current subscription ID, refresh token, and client state. 
                // These are required so the NotificationController, which is not authenticated, can retrieve an access token keyed from the subscription ID.
                // Production apps typically use some method of persistent storage.
                HttpRuntime.Cache.Insert("subscriptionId_" + viewModel.Subscription.Id,
                    Tuple.Create(viewModel.Subscription.ClientState, authResult.RefreshToken), null, DateTime.MaxValue, new TimeSpan(24, 0, 0), System.Web.Caching.CacheItemPriority.NotRemovable, null);

                // Save the latest subscription ID, so we can delete it later.
                Session["SubscriptionId"] = viewModel.Subscription.Id;
                return View("Subscription", viewModel);
            }
            else
            {
                return RedirectToAction("Index", "Error", new { message = subscriptionResponse.StatusCode, debug = await subscriptionResponse.Content.ReadAsStringAsync() });
            }

        }

        // Delete the current webhooks subscription and sign out the user.
        [Authorize, HandleAdalException]
        public async Task<ActionResult> DeleteSubscription()
        {
            string subscriptionId = (string)Session["SubscriptionId"];

            if (!string.IsNullOrEmpty(subscriptionId))
            {
                // Get an access token and add it to the client.
                AuthenticationResult authResult = await AuthHelper.GetAccessTokenAsync();
                HttpResponseMessage deleteResponse = await GraphHelper.DeleteSubscription(authResult.AccessToken, subscriptionId);
                if (!deleteResponse.IsSuccessStatusCode)
                {
                    return RedirectToAction("Index", "Error", new { message = deleteResponse.StatusCode, debug = deleteResponse.Content.ReadAsStringAsync() });
                }
            }
            return RedirectToAction("SignOut", "Account");
        }

        bool isFolderExcluded(string folderName)
        {
            return _excludedMailFolderList.Contains(folderName);
        }
    }
}