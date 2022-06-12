// <copyright file="SyncCustomUserListActivity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Azure.WebJobs.Extensions.DurableTask;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Recipients;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Teams;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.User;
    using Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func.PreparingToSend.Extensions;

    /// <summary>
    /// Syncs custom user list to SentNotification table.
    /// </summary>
    public class SyncCustomUserListActivity
    {
        private readonly ITeamDataRepository teamDataRepository;
        private readonly ITeamMembersService memberService;
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly IStringLocalizer<Strings> localizer;
        private readonly IUserDataRepository userDataRepository;
        private readonly IUserTypeService userTypeService;

        private readonly IUsersService usersService;
        private readonly IRecipientsService recipientsService;

        /// <summary>
        /// Initializes a new instance of the <see cref="SyncCustomUserListActivity"/> class.
        /// </summary>
        /// <param name="teamDataRepository">Team Data repository.</param>
        /// <param name="memberService">Teams member service.</param>
        /// <param name="notificationDataRepository">Notification data repository.</param>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="localizer">Localization service.</param>
        /// <param name="userDataRepository">User Data repository.</param>
        /// <param name="userTypeService">User Type service.</param>
        public SyncCustomUserListActivity(
            ITeamDataRepository teamDataRepository,
            ITeamMembersService memberService,
            INotificationDataRepository notificationDataRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            IStringLocalizer<Strings> localizer,
            IUserDataRepository userDataRepository,
            IUserTypeService userTypeService,
            IUsersService usersService,
            IRecipientsService recipientsService)
        {
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
            this.memberService = memberService ?? throw new ArgumentNullException(nameof(memberService));
            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.localizer = localizer ?? throw new ArgumentNullException(nameof(localizer));
            this.userDataRepository = userDataRepository ?? throw new ArgumentNullException(nameof(userDataRepository));
            this.userTypeService = userTypeService ?? throw new ArgumentNullException(nameof(userTypeService));

            this.usersService = usersService ?? throw new ArgumentNullException(nameof(usersService));
            this.recipientsService = recipientsService ?? throw new ArgumentNullException(nameof(recipientsService));
        }

        /// <summary>
        /// Syncs Team members to SentNotification table.
        /// </summary>
        /// <param name="input">Input data.</param>
        /// <param name="log">Logging service.</param>
        /// <returns>A <see cref="Task"/> representing the asynchronous operation.</returns>
        [FunctionName(FunctionNames.SyncCustomUserListActivity)]
        public async Task<RecipientsInfo> RunAsync(
            [ActivityTrigger] NotificationDataEntity notification,
            ILogger log)
        {
            if (notification == null)
            {
                throw new ArgumentNullException(nameof(notification));
            }

            if (log == null)
            {
                throw new ArgumentNullException(nameof(log));
            }

            string usersString = await this.notificationDataRepository.GetCSVUsersAsync(notification.UsersFile);
            var userArray = usersString.ToLower().Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries).Distinct();

            if (userArray.IsNullOrEmpty())
            {
                return new RecipientsInfo(notification.Id);
            }

            var users = new List<User>();
            foreach (var user in userArray)
            {
                try
                {
                    var u = await this.usersService.GetUserAsync(user);
                    users.Add(u);
                }
                catch (Exception ex)
                {
                    log.LogError($"Error: Getting user {user} from graph: {ex}");
                }
            }

            // Convert to Recipients
            var recipients = await this.GetRecipientsAsync(notification.Id, users);

            await this.sentNotificationDataRepository.BatchInsertOrMergeAsync(recipients);

            // Store in batches and return batch info.
            return await this.recipientsService.BatchRecipients(recipients);
        }

        /// <summary>
        /// Reads corresponding user entity from User table and creates a recipient for every user.
        /// </summary>
        /// <param name="notificationId">Notification Id.</param>
        /// <param name="users">Users.</param>
        /// <returns>List of recipients.</returns>
        private async Task<IEnumerable<SentNotificationDataEntity>> GetRecipientsAsync(string notificationId, IEnumerable<User> users)
        {
            var recipients = new ConcurrentBag<SentNotificationDataEntity>();

            // Get User Entities.
            var maxParallelism = Math.Min(100, users.Count());
            await Task.WhenAll(users.ForEachAsync(maxParallelism, async user =>
            {
                var userEntity = await this.userDataRepository.GetAsync(UserDataTableNames.UserDataPartition, user.Id);

                // This is to set the type of user(existing only, new ones will be skipped) to identify later if it is member or guest.
                var userType = user.UserPrincipalName.GetUserType();
                if (userEntity == null && userType.Equals(UserType.Guest, StringComparison.OrdinalIgnoreCase))
                {
                    // Skip processing new Guest users.
                    return;
                }

                await this.userTypeService.UpdateUserTypeForExistingUserAsync(userEntity, userType);
                if (userEntity == null)
                {
                    userEntity = new UserDataEntity()
                    {
                        AadId = user.Id,
                        UserType = userType,
                        Name = user.DisplayName,
                    };
                }

                recipients.Add(userEntity.CreateInitialSentNotificationDataEntity(partitionKey: notificationId));
            }));

            return recipients;
        }
    }
}
