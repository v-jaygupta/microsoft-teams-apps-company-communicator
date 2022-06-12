// <copyright file="NotificationRepositoryExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Extensions
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Extensions for the repository of the notification data.
    /// </summary>
    public static class NotificationRepositoryExtensions
    {
        /// <summary>
        /// Create a new draft notification.
        /// </summary>
        /// <param name="notificationRepository">The notification repository.</param>
        /// <param name="notification">Draft Notification model class instance passed in from Web API.</param>
        /// <param name="userName">Name of the user who is running the application.</param>
        /// <returns>The newly created notification's id.</returns>
        public static async Task<string> CreateDraftNotificationAsync(
            this INotificationDataRepository notificationRepository,
            DraftNotification notification,
            string userName)
        {
            var newId = notificationRepository.TableRowKeyGenerator.CreateNewKeyOrderingOldestToMostRecent();

            var notificationEntity = new NotificationDataEntity
            {
                PartitionKey = NotificationDataTableNames.DraftNotificationsPartition,
                RowKey = newId,
                Id = newId,
                Title = notification.Title,
                Summary = notification.Summary,
                Author = notification.Author,
                ButtonTitle = notification.ButtonTitle,
                ButtonLink = notification.ButtonLink,
                CreatedBy = userName,
                CreatedDate = DateTime.UtcNow,
                IsDraft = true,
                Teams = notification.Teams,
                Rosters = notification.Rosters,
                Groups = notification.Groups,
                AllUsers = notification.AllUsers,
                Ack = notification.Ack,
                ScheduledDateTime = notification.ScheduledDateTime,
                InlineTranslation = notification.InlineTranslation,
                NotifyUser = notification.NotifyUser,
                FullWidth = notification.FullWidth,
                OnBehalfOf = notification.OnBehalfOf,
                StageView = notification.StageView,
                PollOptions = notification.PollOptions,
                MessageType = notification.MessageType,
                IsPollQuizMode = notification.IsPollQuizMode,
                PollQuizAnswers = notification.PollQuizAnswers,
                IsPollMultipleChoice = notification.IsPollMultipleChoice,
            };

            if (notification.ImageLink.StartsWith(Constants.ImageBase64Format))
            {
                notificationEntity.ImageLink = await notificationRepository.SaveImageAsync(newId, notification.ImageLink);
                notificationEntity.ImageBase64BlobName = newId;
            }
            else
            {
                notificationEntity.ImageLink = notification.ImageLink;
            }

            if (notification.MessageType == "CustomAC")
            {
                await notificationRepository.SaveCustomAdaptiveCardAsync(newId, notification.Summary);
                notificationEntity.Summary = newId;
            }

            if (notification.UsersList != null && notification.UsersList.Length > 0)
            {
                await notificationRepository.SaveCSVUsersAsync(newId, notification.UsersList);
                notificationEntity.UsersFile = newId;
            }

            await notificationRepository.CreateOrUpdateAsync(notificationEntity);

            return newId;
        }
    }
}
