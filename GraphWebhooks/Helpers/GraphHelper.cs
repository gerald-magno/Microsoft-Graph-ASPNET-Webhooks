using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using GraphWebhooks.Models;
using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace GraphWebhooks.Helpers
{
    public class GraphHelper
    {
        /// <summary>
        /// helper method to get all mailFolders of logged in user 
        /// </summary>
        /// <param name="accessToken">access token value</param>
        /// <param name="resource">resource value</param>
        /// <returns>HttpResponseMessage</returns>
        public static async Task<HttpResponseMessage> GetMailFolders(string accessToken, string resource)
        {
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, resource))
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage response = await client.SendAsync(request);
                    return response;
                }
            }

        }


        /// <summary>
        /// helper method to create subscription 
        /// </summary>
        /// <param name="accessToken">access token value</param>
        /// <param name="subscription">subscription model</param>
        /// <returns>HttpResponseMessage</returns>
        public static async Task<HttpResponseMessage> CreateSubscription(string accessToken, Subscription subscription)
        {
            //TODO: get of subscriptionEndpoint
            string subscriptionsEndpoint = "https://graph.microsoft.com/v1.0/subscriptions/";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, subscriptionsEndpoint))
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    
                    string contentString = JsonConvert.SerializeObject(subscription, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
                    request.Content = new StringContent(contentString, System.Text.Encoding.UTF8, "application/json");

                    // Send the request and parse the response.
                    HttpResponseMessage response = await client.SendAsync(request);
                    return response;
                }

            }

        }

        /// <summary>
        /// helper method to delete subscription 
        /// </summary>
        /// <param name="accessToken">access token value</param>
        /// <param name="subscriptionId">subscription Id</param>
        /// <returns>HttpResponseMessage</returns>
        public static async Task<HttpResponseMessage> DeleteSubscription(string accessToken, string subscriptionId)
        {

            string serviceRootUrl = "https://graph.microsoft.com/v1.0/subscriptions/";
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Delete, serviceRootUrl + subscriptionId))
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                   
                    HttpResponseMessage response = await client.SendAsync(request);
                    return response;
                }
                
            }
            
        }
    }
}
