namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    public class ReadReceiptObject
    {
        public string BotSentMessageId; // the id of the message that the bot sent and the bot wants to track the read status on.
        public string ReaderAadId; // AadId of the user whom the bot sent the message to
        public bool IsMessageRead; // track the read status of the botSentMessageId. 
    }
}
