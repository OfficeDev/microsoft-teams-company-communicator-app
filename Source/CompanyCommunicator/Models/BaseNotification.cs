// <copyright file="BaseNotification.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;

    /// <summary>
    /// Base notification model class.
    /// </summary>
    public class BaseNotification
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Buttons value.
        /// </summary>
        public string Buttons { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsScheduled value.
        /// </summary>
        public bool IsScheduled { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether gets or sets the IsImportant value.
        /// </summary>
        public bool IsImportant { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }
    }
}
