namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard
{
    using Microsoft.Bot.Schema;
    using Newtonsoft.Json;

    /// <summary>
    /// Adaptive card action model class.
    /// </summary>
    public class AdaptiveCardMSTeamsAction
    {
        /// <summary>
        /// Gets or sets Ms Teams card action.
        /// </summary>
        [JsonProperty("msteams")]
        public CardAction MsteamsCardAction { get; set; }
    }
}
