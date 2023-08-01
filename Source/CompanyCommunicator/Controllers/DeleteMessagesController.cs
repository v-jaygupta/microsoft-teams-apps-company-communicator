// <copyright file="DeleteMessagesController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Net.Http;
    using System.Net.Http.Json;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.CleanUpHistory;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;

    using Microsoft.Teams.Apps.CompanyCommunicator.Models;
    using Newtonsoft.Json;

    /// <summary>
    /// Controller for older deleting messages.
    /// </summary>
    //[Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/deletemessages")]
    public class DeleteMessagesController : Controller
    {
        private readonly ICleanUpHistoryRepository cleanUpHistoryRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly TableRowKeyGenerator tableRowKeyGenerator;
        private readonly IHttpClientFactory clientFactory;
        private readonly IConfiguration configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="DeleteMessagesController"/> class.
        /// </summary>
        /// <param name="cleanUpHistoryRepository">Clean Up History repository instance.</param>
        /// <param name="sentNotificationDataRepository">The SentNotificationData repository.</param>
        /// <param name="tableRowKeyGenerator">Table row key generator service.</param>
        /// <param name="clientFactory">Http client service.</param>
        /// <param name="configuration">Configuration service.</param>
        public DeleteMessagesController(
            ICleanUpHistoryRepository cleanUpHistoryRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            TableRowKeyGenerator tableRowKeyGenerator,
            IHttpClientFactory clientFactory,
            IConfiguration configuration)
        {
            this.cleanUpHistoryRepository = cleanUpHistoryRepository ?? throw new ArgumentNullException(nameof(cleanUpHistoryRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.tableRowKeyGenerator = tableRowKeyGenerator ?? throw new ArgumentNullException(nameof(tableRowKeyGenerator));
            this.clientFactory = clientFactory ?? throw new ArgumentNullException(nameof(clientFactory));
            this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        ///// <summary>
        ///// Initiate the deletion of historical messages.
        ///// </summary>
        ///// <param name="deleteHistoricalMessage">delete historical message request.</param>
        ///// <returns>The result of an action method.</returns>
        //[HttpPost]
        //public async Task<IActionResult> DeleteHistoricalMessagesAsync(
        //    [FromBody] DeleteHistoricalMessages deleteHistoricalMessage)
        //{
        //    if (deleteHistoricalMessage == null)
        //    {
        //        throw new ArgumentNullException(nameof(deleteHistoricalMessage));
        //    }

        //    await Task.WhenAll(
        //       this.sentNotificationDataRepository.EnsureSentNotificationDataTableExistsAsync(),
        //       this.cleanUpHistoryRepository.EnsureCleanUpHistoryTableExistsAsync());

        //    var newId = this.tableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();
        //    deleteHistoricalMessage.RowKeyId = newId;

        //    await this.cleanUpHistoryRepository.CreateOrUpdateAsync(new CleanUpHistoryEntity()
        //    {
        //        PartitionKey = "Delete Messages",
        //        RowKey = newId,
        //        SelectedDateRange = deleteHistoricalMessage.SelectedDateRange,
        //        RecordsDeleted = 0,
        //        DeletedBy = deleteHistoricalMessage.DeletedBy,
        //        Status = CleanUpStatus.InProgress.ToString(),
        //        StartDate = deleteHistoricalMessage.StartDate,
        //        EndDate = deleteHistoricalMessage.EndDate,
        //    });

        //    var httpResponse = await this.httpClient.PutAsJsonAsync(this.configuration["DataFunctionUrl"], deleteHistoricalMessage);

        //    if (!httpResponse.IsSuccessStatusCode)
        //    {
        //        throw new Exception($"Failed to send HTTP PUT request to Azure Function: {httpResponse.StatusCode}");
        //    }

        //    return await Task.FromResult<IActionResult>(this.Ok());
        //}

        /// <summary>
        /// Initiate the deletion of historical messages.
        /// </summary>
        /// <param name="deleteHistoricalMessage">delete historical message request.</param>
        /// <returns>The result of an action method.</returns>
        [HttpPost]
        public async Task DeleteHistoricalMessagesAsync([FromBody] DeleteHistoricalMessages deleteHistoricalMessage)
        {
            if (deleteHistoricalMessage == null)
            {
                throw new ArgumentNullException(nameof(deleteHistoricalMessage));
            }

            await Task.WhenAll(
               this.sentNotificationDataRepository.EnsureSentNotificationDataTableExistsAsync(),
               this.cleanUpHistoryRepository.EnsureCleanUpHistoryTableExistsAsync());

            var newId = this.tableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();
            deleteHistoricalMessage.RowKeyId = newId;
            deleteHistoricalMessage.DeletedBy = this.HttpContext.User?.Identity?.Name ?? "defaultUser";

            await this.cleanUpHistoryRepository.CreateOrUpdateAsync(new CleanUpHistoryEntity()
            {
                PartitionKey = "Delete Messages",
                RowKey = newId,
                SelectedDateRange = deleteHistoricalMessage.SelectedDateRange,
                RecordsDeleted = 0,
                DeletedBy = deleteHistoricalMessage.DeletedBy,
                Status = CleanUpStatus.InProgress.ToString(),
                StartDate = deleteHistoricalMessage.StartDate,
                EndDate = deleteHistoricalMessage.EndDate,
            });

            _ = Task.Run(async () =>
            {
                try
                {
                    string functionUrl = "https://jayschedule-data-function.azurewebsites.net/api/CompanyCommunicatorDataCleanUpFunction";
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, functionUrl);
                    string jsonPayload = JsonConvert.SerializeObject(deleteHistoricalMessage);
                    request.Content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                    this.clientFactory.CreateClient().SendAsync(request);

                    //if (!httpResponse.IsSuccessStatusCode)
                    //{
                    //    throw new Exception($"Failed to send HTTP PUT request to Azure Function: {httpResponse.StatusCode}");
                    //}
                }
                catch (Exception ex)
                {
                    await this.cleanUpHistoryRepository.CreateOrUpdateAsync(new CleanUpHistoryEntity()
                    {
                        PartitionKey = "Delete Messages",
                        RowKey = newId,
                        SelectedDateRange = deleteHistoricalMessage.SelectedDateRange,
                        RecordsDeleted = 0,
                        DeletedBy = deleteHistoricalMessage.DeletedBy,
                        Status = CleanUpStatus.Failed.ToString(),
                        StartDate = deleteHistoricalMessage.StartDate,
                        EndDate = deleteHistoricalMessage.EndDate,
                    });
                    // Log the exception or handle
                    Console.WriteLine($"Exception in background task: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Get the clean up history.
        /// </summary>
        /// <returns>A list of <see cref="CleanUpHistoryEntity"/> instances.</returns>
        [HttpGet]
        public async Task<IEnumerable<CleanUpHistoryEntity>> GetCleanUpHistorySummary()
        {
            var notificationEntities = await this.cleanUpHistoryRepository.GetAllCleanUpHistoryAsync();

            var result = new List<CleanUpHistoryEntity>();
            foreach (var notificationEntity in notificationEntities)
            {
                var summary = new CleanUpHistoryEntity
                {
                    SelectedDateRange = notificationEntity.SelectedDateRange,
                    DeletedBy = notificationEntity.DeletedBy,
                    Status = notificationEntity.Status.AddSpacesToCamelCase(),
                    RecordsDeleted = notificationEntity.RecordsDeleted,
                    Timestamp = notificationEntity.Timestamp,
                    StartDate = notificationEntity.StartDate,
                    EndDate = notificationEntity.EndDate,
                };

                result.Add(summary);
            }

            return result;
        }
    }
}
