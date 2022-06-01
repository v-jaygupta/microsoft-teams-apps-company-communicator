

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using Newtonsoft.Json;

    /// <summary>
    /// Tab info action model class.
    /// </summary>
    public class TabInfoAction
    {
        /// <summary>
        /// Gets or sets type of tab.
        /// </summary>
        [JsonProperty("type")]
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets tab info.
        /// </summary>
        [JsonProperty("tabInfo")]
        public TabInfo TabInfo { get; set; }
    }
}
