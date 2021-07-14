// <copyright file="SentNotificationsController.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.DataQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MessageQueues.PrepareToSendQueue;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Controllers.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the sent notification data.
    /// </summary>
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    [Route("api/sentNotifications")]
    public class SentNotificationsController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;
        private readonly ITeamDataRepository teamDataRepository;
        private readonly IPrepareToSendQueue prepareToSendQueue;
        private readonly IDataQueue dataQueue;
        private readonly double forceCompleteMessageDelayInSeconds;
        private readonly IGroupsService groupsService;
        private readonly IExportDataRepository exportDataRepository;
        private readonly IAppCatalogService appCatalogService;
        private readonly IAppSettingsService appSettingsService;
        private readonly UserAppOptions userAppOptions;
        private readonly ILogger<SentNotificationsController> logger;

        /// <summary>
        /// Initializes a new instance of the <see cref="SentNotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository service that deals with the table storage in azure.</param>
        /// <param name="sentNotificationDataRepository">Sent notification data repository.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        /// <param name="prepareToSendQueue">The service bus queue for preparing to send notifications.</param>
        /// <param name="dataQueue">The service bus queue for the data queue.</param>
        /// <param name="dataQueueMessageOptions">The options for the data queue messages.</param>
        /// <param name="groupsService">The groups service.</param>
        /// <param name="exportDataRepository">The Export data repository instance.</param>
        /// <param name="appCatalogService">App catalog service.</param>
        /// <param name="appSettingsService">App settings service.</param>
        /// <param name="userAppOptions">User app options.</param>
        /// <param name="loggerFactory">The logger factory.</param>
        public SentNotificationsController(
            INotificationDataRepository notificationDataRepository,
            ISentNotificationDataRepository sentNotificationDataRepository,
            ITeamDataRepository teamDataRepository,
            IPrepareToSendQueue prepareToSendQueue,
            IDataQueue dataQueue,
            IOptions<DataQueueMessageOptions> dataQueueMessageOptions,
            IGroupsService groupsService,
            IExportDataRepository exportDataRepository,
            IAppCatalogService appCatalogService,
            IAppSettingsService appSettingsService,
            IOptions<UserAppOptions> userAppOptions,
            ILoggerFactory loggerFactory)
        {
            if (dataQueueMessageOptions is null)
            {
                throw new ArgumentNullException(nameof(dataQueueMessageOptions));
            }

            this.notificationDataRepository = notificationDataRepository ?? throw new ArgumentNullException(nameof(notificationDataRepository));
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
            this.teamDataRepository = teamDataRepository ?? throw new ArgumentNullException(nameof(teamDataRepository));
            this.prepareToSendQueue = prepareToSendQueue ?? throw new ArgumentNullException(nameof(prepareToSendQueue));
            this.dataQueue = dataQueue ?? throw new ArgumentNullException(nameof(dataQueue));
            this.forceCompleteMessageDelayInSeconds = dataQueueMessageOptions.Value.ForceCompleteMessageDelayInSeconds;
            this.groupsService = groupsService ?? throw new ArgumentNullException(nameof(groupsService));
            this.exportDataRepository = exportDataRepository ?? throw new ArgumentNullException(nameof(exportDataRepository));
            this.appCatalogService = appCatalogService ?? throw new ArgumentNullException(nameof(appCatalogService));
            this.appSettingsService = appSettingsService ?? throw new ArgumentNullException(nameof(appSettingsService));
            this.userAppOptions = userAppOptions?.Value ?? throw new ArgumentNullException(nameof(userAppOptions));
            this.logger = loggerFactory?.CreateLogger<SentNotificationsController>() ?? throw new ArgumentNullException(nameof(loggerFactory));
        }

        /// <summary>
        /// Send a notification, which turns a draft to be a sent notification.
        /// </summary>
        /// <param name="draftNotification">An instance of <see cref="DraftNotification"/> class.</param>
        /// <returns>The result of an action method.</returns>
        [HttpPost]
        public async Task<IActionResult> CreateSentNotificationAsync(
            [FromBody] DraftNotification draftNotification)
        {
            if (draftNotification == null)
            {
                throw new ArgumentNullException(nameof(draftNotification));
            }

            var draftNotificationDataEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.DraftNotificationsPartition,
                draftNotification.Id);
            if (draftNotificationDataEntity == null)
            {
                return this.NotFound($"Draft notification, Id: {draftNotification.Id}, could not be found.");
            }

            var newSentNotificationId =
                await this.notificationDataRepository.MoveDraftToSentPartitionAsync(draftNotificationDataEntity);

            // Ensure the data table needed by the Azure Functions to send the notifications exist in Azure storage.
            await this.sentNotificationDataRepository.EnsureSentNotificationDataTableExistsAsync();

            // Update user app id if proactive installation is enabled.
            await this.UpdateUserAppIdAsync();

            var prepareToSendQueueMessageContent = new PrepareToSendQueueMessageContent
            {
                NotificationId = newSentNotificationId,
            };
            await this.prepareToSendQueue.SendAsync(prepareToSendQueueMessageContent);

            // Send a "force complete" message to the data queue with a delay to ensure that
            // the notification will be marked as complete no matter the counts
            var forceCompleteDataQueueMessageContent = new DataQueueMessageContent
            {
                NotificationId = newSentNotificationId,
                ForceMessageComplete = true,
            };
            await this.dataQueue.SendDelayedAsync(
                forceCompleteDataQueueMessageContent,
                this.forceCompleteMessageDelayInSeconds);

            return this.Ok();
        }

        /// <summary>
        /// Get most recently sent notification summaries.
        /// </summary>
        /// <returns>A list of <see cref="SentNotificationSummary"/> instances.</returns>
        [HttpGet]
        public async Task<IEnumerable<SentNotificationSummary>> GetSentNotificationsAsync()
        {
            var notificationEntities = await this.notificationDataRepository.GetMostRecentSentNotificationsAsync();

            var result = new List<SentNotificationSummary>();
            foreach (var notificationEntity in notificationEntities)
            {
                var summary = new SentNotificationSummary
                {
                    Id = notificationEntity.Id,
                    Title = notificationEntity.Title,
                    CreatedDateTime = notificationEntity.CreatedDate,
                    SentDate = notificationEntity.SentDate,
                    Succeeded = notificationEntity.Succeeded,
                    Failed = notificationEntity.Failed,
                    Unknown = this.GetUnknownCount(notificationEntity),
                    TotalMessageCount = notificationEntity.TotalMessageCount,
                    SendingStartedDate = notificationEntity.SendingStartedDate,
                    Status = notificationEntity.GetStatus(),
                };

                result.Add(summary);
            }

            return result;
        }

        /// <summary>
        /// Get a sent notification by Id.
        /// </summary>
        /// <param name="id">Id of the requested sent notification.</param>
        /// <returns>Required sent notification.</returns>
        [HttpGet("{id}")]
        public async Task<IActionResult> GetSentNotificationByIdAsync(string id)
        {
            if (id == null)
            {
                throw new ArgumentNullException(nameof(id));
            }

            var notificationEntity = await this.notificationDataRepository.GetAsync(
                NotificationDataTableNames.SentNotificationsPartition,
                id);
            if (notificationEntity == null)
            {
                return this.NotFound();
            }

            var groupNames = await this.groupsService.
                GetByIdsAsync(notificationEntity.Groups).
                Select(x => x.DisplayName).
                ToListAsync();

            var userId = this.HttpContext.User.FindFirstValue(Common.Constants.ClaimTypeUserId);
            var userNotificationDownload = await this.exportDataRepository.GetAsync(userId, id);

            var result = new SentNotification
            {
                Id = notificationEntity.Id,
                Title = notificationEntity.Title,
                ImageLink = notificationEntity.ImageLink,
                Summary = notificationEntity.Summary,
                Author = notificationEntity.Author,
                ButtonTitle = notificationEntity.ButtonTitle,
                ButtonLink = notificationEntity.ButtonLink,
                Buttons = notificationEntity.Buttons,
                IsScheduled = notificationEntity.IsScheduled,
                IsImportant = notificationEntity.IsImportant,
                CreatedDateTime = notificationEntity.CreatedDate,
                SentDate = notificationEntity.SentDate,
                Succeeded = notificationEntity.Succeeded,
                Failed = notificationEntity.Failed,
                Unknown = this.GetUnknownCount(notificationEntity),
                TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Teams),
                RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(notificationEntity.Rosters),
                GroupNames = groupNames,
                AllUsers = notificationEntity.AllUsers,
                SendingStartedDate = notificationEntity.SendingStartedDate,
                ErrorMessage = notificationEntity.ErrorMessage,
                WarningMessage = notificationEntity.WarningMessage,
                CanDownload = userNotificationDownload == null,
                SendingCompleted = notificationEntity.IsCompleted(),
            };

            return this.Ok(result);
        }

        private int? GetUnknownCount(NotificationDataEntity notificationEntity)
        {
            var unknown = notificationEntity.Unknown;

            // In CC v2, the number of throttled recipients are counted and saved in NotificationDataEntity.Unknown property.
            // However, CC v1 saved the number of throttled recipients in NotificationDataEntity.Throttled property.
            // In order to make it backward compatible, we add the throttled number to the unknown variable.
            var throttled = notificationEntity.Throttled;
            if (throttled > 0)
            {
                unknown += throttled;
            }

            return unknown > 0 ? unknown : (int?)null;
        }

        /// <summary>
        /// Updates user app id if its not already synced.
        /// </summary>
        private async Task UpdateUserAppIdAsync()
        {
            // check if proactive installation is enabled.
            if (!this.userAppOptions.ProactivelyInstallUserApp)
            {
                return;
            }

            // check if we have already synced app id.
            var appId = await this.appSettingsService.GetUserAppIdAsync();
            if (!string.IsNullOrWhiteSpace(appId))
            {
                return;
            }

            try
            {
                // Fetch and store user app id in App Catalog.
                appId = await this.appCatalogService.GetTeamsAppIdAsync(this.userAppOptions.UserAppExternalId);

                // Graph SDK returns empty id if the app is not found.
                if (string.IsNullOrEmpty(appId))
                {
                    this.logger.LogError($"Failed to find an app in AppCatalog with external Id: {this.userAppOptions.UserAppExternalId}");
                    return;
                }

                await this.appSettingsService.SetUserAppIdAsync(appId);
            }
            catch (ServiceException exception)
            {
                // Failed to fetch app id.
                this.logger.LogError(exception, $"Failed to get catalog app id. Error message: {exception.Message}.");
            }
        }
    }
}
