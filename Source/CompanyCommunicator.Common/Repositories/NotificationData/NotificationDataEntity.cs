// <copyright file="NotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.Cosmos.Table;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Newtonsoft.Json;

    /// <summary>
    /// Notification data entity class.
    /// This entity type holds the data for notifications that are either (depending on partition key):
    ///     drafts
    ///     sent
    /// It holds the data for the content of the notification.
    /// It holds the data for the recipients of the notification.
    /// </summary>
    public class NotificationDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the id of the notification.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the title text of the notification's content.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the image link of the notification's content.
        /// </summary>
        public string ImageLink { get; set; }
        
        /// <summary>
        /// Gets or sets the blob name for the image.
        /// </summary>
        public string ImageBlobName { get; set; }

        /// <summary>
        /// Gets or sets the summary text of the notification's content.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the author text of the notification's content.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the button title of the notification's content.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the button link of the notification's content.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the information for the user that created the notification.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets the DateTime the notification was created.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets the DateTime the notification's sending was completed.
        /// </summary>
        public DateTime? SentDate { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the notification is a draft.
        /// </summary>
        public bool IsDraft { get; set; }

        /// <summary>
        /// Gets or sets the TeamsInString value.
        /// This property helps to save the Teams data in the Azure Table storage.
        /// Table storage doesn't support an array type of the property directly
        /// so this is a comma separated list of the team ids.
        /// </summary>
        public string TeamsInString { get; set; }

        /// <summary>
        /// Gets or sets Teams audience collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> Teams
        {
            get => JsonConvert.DeserializeObject<IEnumerable<string>>(this.TeamsInString.IsNullOrEmpty() ? "[]" : this.TeamsInString);
            set => this.TeamsInString = JsonConvert.SerializeObject(value);
        }

        /// <summary>
        /// Gets or sets the RostersInString value.
        /// This property helps to save the Rosters list in the Azure Table storage.
        /// Table storage doesn't support an array type of the property directly
        /// so this is a comma separated list of the team ids for which the rosters
        /// are the recipients.
        /// </summary>
        public string RostersInString { get; set; }

        /// <summary>
        /// Gets or sets the team ids of the Rosters audience collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> Rosters
        {
            get => JsonConvert.DeserializeObject<IEnumerable<string>>(this.RostersInString.IsNullOrEmpty() ? "[]" : this.RostersInString);
            set => this.RostersInString = JsonConvert.SerializeObject(value);
        }

        /// <summary>
        /// Gets or sets the GroupsInsString value.
        /// This property helps to save the Grousp list in the Azure Table storage.
        /// Table storage doesn't support an array type of the property directly
        /// so this is a comma separated list of the group ids for which the members
        /// are the recipients.
        /// </summary>
        public string GroupsInString { get; set; }

        /// <summary>
        /// Gets or sets the team ids of the Groups audience collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> Groups
        {
            get => JsonConvert.DeserializeObject<IEnumerable<string>>(this.GroupsInString.IsNullOrEmpty() ? "[]" : this.GroupsInString);
            set => this.GroupsInString = JsonConvert.SerializeObject(value);
        }

        /// <summary>
        /// Gets or sets a value indicating whether a notification should be sent to all the
        /// known users - this is equivalent to all of the users stored in the User Data table.
        /// </summary>
        public bool AllUsers { get; set; }

        /// <summary>
        /// Gets or sets the message version number.
        /// </summary>
        public string MessageVersion { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who have received the notification successfully.
        /// </summary>
        public int Succeeded { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who failed to receive the notification because
        /// of a failure response in the API call.
        /// </summary>
        public int Failed { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who did not receive the message because
        /// the API response indicated that the bot was throttled.
        /// [DEPRECATED - because the bot now retries, this should always stay 0].
        /// </summary>
        public int Throttled { get; set; }

        /// <summary>
        /// Gets or sets the number or recipients who have an unknown status - this means a status
        /// that has not changed from the initial initialization status after the notification has
        /// been force completed.
        /// </summary>
        public int Unknown { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the sending process is completed.
        /// [DEPRECATED - kept for backward compatibility].
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// Gets or sets the total number of expected messages to send.
        /// </summary>
        public int TotalMessageCount { get; set; }

        /// <summary>
        /// Gets or sets the DateTime the notification's was queued to be sent.
        /// </summary>
        public DateTime? SendingStartedDate { get; set; }

        /// <summary>
        /// Gets or sets the error message for the notification if there was a failure in
        /// preparing and sending the notification.
        /// Front-end shows the ExceptionMessage value in the "View status" task module.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Gets or sets the warning message for the notification if there was a warning given
        /// when preparing and sending the notification.
        /// Front-end shows the WarningMessage value in the "View status" task module.
        /// </summary>
        public string WarningMessage { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the notification is in the "preparing to send" state.
        /// [DEPRECATED - kept for backward compatibility].
        /// </summary>
        public bool IsPreparingToSend { get; set; }

        /// <summary>
        /// Gets or sets notification status.
        /// </summary>
        public string Status { get; set; }
    }
}
