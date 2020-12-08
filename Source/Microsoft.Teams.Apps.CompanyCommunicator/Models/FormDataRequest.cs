// <copyright file="FormDataRequest.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using Microsoft.AspNetCore.Http;

    /// <summary>
    /// Form Data Request model class.
    /// </summary>
    public class FormDataRequest
    {
        /// <summary>
        /// Gets or sets the Draft Notification String.
        /// </summary>
        public string DraftMessage { get; set; }

        /// <summary>
        /// Gets or sets the File content value.
        /// </summary>
        public IFormFile File { get; set; } = null;
    }
}