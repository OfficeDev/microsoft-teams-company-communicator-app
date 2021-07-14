// <copyright file="DraftNotificationPreviewService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.DraftNotificationPreview
{
    using System;
    using System.Net;
    using System.Threading;
    using System.Threading.Tasks;
    using AdaptiveCards;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Schema;
    using Microsoft.Extensions.Options;
    using Microsoft.Teams.Apps.CompanyCommunicator.Bot;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AdaptiveCard;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.CommonBot;

    /// <summary>
    /// Draft notification preview service.
    /// </summary>
    public class DraftNotificationPreviewService : IDraftNotificationPreviewService
    {
        private static readonly string MsTeamsChannelId = "msteams";
        private static readonly string ChannelConversationType = "channel";
        private static readonly string ThrottledErrorResponse = "Throttled";

        private readonly string botAppId;
        private readonly AdaptiveCardCreator adaptiveCardCreator;
        private readonly CompanyCommunicatorBotAdapter companyCommunicatorBotAdapter;

        /// <summary>
        /// Initializes a new instance of the <see cref="DraftNotificationPreviewService"/> class.
        /// </summary>
        /// <param name="botOptions">The bot options.</param>
        /// <param name="adaptiveCardCreator">Adaptive card creator service.</param>
        /// <param name="companyCommunicatorBotAdapter">Bot framework http adapter instance.</param>
        public DraftNotificationPreviewService(
            IOptions<BotOptions> botOptions,
            AdaptiveCardCreator adaptiveCardCreator,
            CompanyCommunicatorBotAdapter companyCommunicatorBotAdapter)
        {
            var options = botOptions ?? throw new ArgumentNullException(nameof(botOptions));
            this.botAppId = options.Value.AuthorAppId;
            if (string.IsNullOrEmpty(this.botAppId))
            {
                throw new ApplicationException("AuthorAppId setting is missing in the configuration.");
            }

            this.adaptiveCardCreator = adaptiveCardCreator ?? throw new ArgumentNullException(nameof(adaptiveCardCreator));
            this.companyCommunicatorBotAdapter = companyCommunicatorBotAdapter ?? throw new ArgumentNullException(nameof(companyCommunicatorBotAdapter));
        }

        /// <inheritdoc/>
        public async Task<HttpStatusCode> SendPreview(NotificationDataEntity draftNotificationEntity, TeamDataEntity teamDataEntity, string teamsChannelId)
        {
            if (draftNotificationEntity == null)
            {
                throw new ArgumentException("Null draft notification entity.");
            }

            if (teamDataEntity == null)
            {
                throw new ArgumentException("Null team data entity.");
            }

            if (string.IsNullOrWhiteSpace(teamsChannelId))
            {
                throw new ArgumentException("Null channel id.");
            }

            // Create bot conversation reference.
            var conversationReference = this.PrepareConversationReferenceAsync(teamDataEntity, teamsChannelId);

            // Trigger bot to send the adaptive card.
            try
            {
                await this.companyCommunicatorBotAdapter.ContinueConversationAsync(
                    this.botAppId,
                    conversationReference,
                    async (turnContext, cancellationToken) => await this.SendAdaptiveCardAsync(turnContext, draftNotificationEntity),
                    CancellationToken.None);
                return HttpStatusCode.OK;
            }
            catch (ErrorResponseException e)
            {
                var errorResponse = (ErrorResponse)e.Body;
                if (errorResponse != null
                    && errorResponse.Error.Code.Equals(DraftNotificationPreviewService.ThrottledErrorResponse, StringComparison.OrdinalIgnoreCase))
                {
                    return HttpStatusCode.TooManyRequests;
                }

                throw;
            }
        }

        private ConversationReference PrepareConversationReferenceAsync(TeamDataEntity teamDataEntity, string channelId)
        {
            var channelAccount = new ChannelAccount
            {
                Id = $"28:{this.botAppId}",
            };

            var conversationAccount = new ConversationAccount
            {
                ConversationType = DraftNotificationPreviewService.ChannelConversationType,
                Id = channelId,
                TenantId = teamDataEntity.TenantId,
            };

            var conversationReference = new ConversationReference
            {
                Bot = channelAccount,
                ChannelId = DraftNotificationPreviewService.MsTeamsChannelId,
                Conversation = conversationAccount,
                ServiceUrl = teamDataEntity.ServiceUrl,
            };

            return conversationReference;
        }

        private async Task SendAdaptiveCardAsync(
            ITurnContext turnContext,
            NotificationDataEntity draftNotificationEntity)
        {
            var reply = this.CreateReply(draftNotificationEntity);
            await turnContext.SendActivityAsync(reply);
        }

        private IMessageActivity CreateReply(NotificationDataEntity draftNotificationEntity)
        {
            var adaptiveCard = this.adaptiveCardCreator.CreateAdaptiveCard(
                draftNotificationEntity.Title,
                draftNotificationEntity.ImageLink,
                draftNotificationEntity.Summary,
                draftNotificationEntity.Author,
                draftNotificationEntity.ButtonTitle,
                draftNotificationEntity.ButtonLink,
                draftNotificationEntity.Buttons);

            var attachment = new Attachment
            {
                ContentType = AdaptiveCard.ContentType,
                Content = adaptiveCard,
            };

            var reply = MessageFactory.Attachment(attachment);

            return reply;
        }
    }
}