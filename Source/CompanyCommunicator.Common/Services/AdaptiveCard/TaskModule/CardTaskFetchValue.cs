

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard.TaskModule
{
    using Newtonsoft.Json;

    public class CardTaskFetchValue<T>
    {
        [JsonProperty("type")]
        public object Type { get; set; } = "task/fetch";

        [JsonProperty("data")]
        public T Data { get; set; }
    }
}
