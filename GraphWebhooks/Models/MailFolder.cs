
using Newtonsoft.Json;
namespace GraphWebhooks.Models
{
    public class MailFolder
    {
        [JsonProperty(PropertyName = "@odata.context")]
        public string OdataContext { get; set; }

        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "displayName")]
        public string displayName { get; set; }

        [JsonProperty(PropertyName = "parentFolderId")]
        public string ParentFolderId { get; set; }

        [JsonProperty(PropertyName = "childFolderCount")]
        public string ChildFolderCount { get; set; }

        [JsonProperty(PropertyName = "unreadItemCount")]
        public string UnreadItemCount { get; set; }

        [JsonProperty(PropertyName = "totalItemCount")]
        public string TotalItemCount { get; set; }

        [JsonProperty(PropertyName = "wellKnownName")]
        public string WellKnownName { get; set; }

    }
}
