﻿// <copyright file="AudienceEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Repositories.Notification
{
    /// <summary>
    /// Audience entity class.
    /// </summary>
    public class AudienceEntity
    {
        /// <summary>
        /// Gets or sets team Id.
        /// </summary>
        public string TeamId { get; set; }

        /// <summary>
        /// Gets or sets delivery state.
        /// </summary>
        public DeliveryStatus DeliveryState { get; set; }

        // other properties
        // Acknowlegement
        // Recation
        // Response
    }
}