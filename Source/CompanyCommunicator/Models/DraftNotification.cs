// <copyright file="DraftNotification.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Extensions.Localization;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Resources;

    /// <summary>
    /// Draft notification model class.
    /// </summary>
    public class DraftNotification : BaseNotification
    {
        private static readonly int MaxSelectedTeamNum = 20;

        /// <summary>
        /// Initializes a new instance of the <see cref="DraftNotification"/> class.
        /// </summary>
        public DraftNotification()
        {
            this.Teams = new List<string>();
            this.Rosters = new List<string>();
        }

        /// <summary>
        /// Gets or sets Teams audience id collection.
        /// </summary>
        public IEnumerable<string> Teams { get; set; }

        /// <summary>
        /// Gets or sets Rosters audience id collection.
        /// </summary>
        public IEnumerable<string> Rosters { get; set; }

        /// <summary>
        /// Gets or sets Groups audience id collection.
        /// </summary>
        public IEnumerable<string> Groups { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a notification should be sent to all the users.
        /// </summary>
        public bool AllUsers { get; set; }

        /// <summary>
        /// Gets or sets ScheduledDate value.
        /// </summary>
        public DateTime? ScheduledDate { get; set; }

        /// <summary>
        /// Validates a draft notification.
        /// Teams and Rosters property should not contain more than 20 items.
        /// </summary>
        /// <param name="localizer">The string localizer service.</param>
        /// <param name="errorMessage">It returns the error message found by the method to the callers.</param>
        /// <returns>A flag indicates if a draft notification is valid or not.</returns>
        public bool Validate(IStringLocalizer<Strings> localizer, out string errorMessage)
        {
            var stringBuilder = new StringBuilder();

            var teams = this.Teams.ToList();
            if (teams.Count > DraftNotification.MaxSelectedTeamNum)
            {
                var format = localizer.GetString("NumberOfTeamsExceededLimitWarningFormat");
                stringBuilder.AppendFormat(format, teams.Count, DraftNotification.MaxSelectedTeamNum);
                stringBuilder.AppendLine();
            }

            var rosters = this.Rosters.ToList();
            if (rosters.Count > DraftNotification.MaxSelectedTeamNum)
            {
                var format = localizer.GetString("NumberOfRostersExceededLimitWarningFormat");
                stringBuilder.AppendFormat(format, rosters.Count, DraftNotification.MaxSelectedTeamNum);
                stringBuilder.AppendLine();
            }

            var groups = this.Groups.ToList();
            if (groups.Count > DraftNotification.MaxSelectedTeamNum)
            {
                var format = localizer.GetString("NumberOfGroupsExceededLimitWarningFormat");
                stringBuilder.AppendFormat(format, groups.Count, DraftNotification.MaxSelectedTeamNum);
                stringBuilder.AppendLine();
            }

            errorMessage = stringBuilder.ToString();
            return stringBuilder.Length == 0;
        }
    }
}
