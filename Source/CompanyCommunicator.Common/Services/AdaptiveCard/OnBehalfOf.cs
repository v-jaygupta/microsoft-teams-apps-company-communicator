

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using Newtonsoft.Json;

    public class OnBehalfOfEntity
    {
        /// <summary>
        /// Gets or sets the ItemId.
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public int ItemId { get; set; }

        /// <summary>
        /// Gets or sets the MentionType.
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string MentionType { get; set; }

        /// <summary>
        /// Gets or sets the MRI for the bot.
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string Mri { get; set; }

        /// <summary>
        /// Gets or sets the Display name of user to be used in header.
        /// </summary>
        [JsonProperty(Required = Required.Always)]
        public string DisplayName { get; set; }
    }
}
