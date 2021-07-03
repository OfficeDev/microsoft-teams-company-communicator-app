// <copyright file="GraphTokenProvider.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Authentication
{
    using System;
    using System.Linq;
    using System.Net.Http;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Graph;
    using Microsoft.Identity.Web;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;

    /// <summary>
    /// Add Access Token to Graph Api.
    /// </summary>
    public class GraphTokenProvider : IAuthenticationProvider
    {
        private readonly ITokenAcquisition tokenAcquisition;

        /// <summary>
        /// Initializes a new instance of the <see cref="GraphTokenProvider"/> class.
        /// </summary>
        /// <param name="tokenAcquisition">MSAL.NET token acquisition service.</param>
        public GraphTokenProvider(ITokenAcquisition tokenAcquisition)
        {
            this.tokenAcquisition = tokenAcquisition ?? throw new ArgumentNullException(nameof(tokenAcquisition));
        }

        /// <summary>
        /// Intercepts HttpRequest and add Bearer token.
        /// </summary>
        /// <param name="request">Represents a HttpRequestMessage.</param>
        /// <returns>asynchronous operation.</returns>
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            var permissionType = this.ExtractPermissionType(request.Headers);
            string accessToken = await this.GetAccessToken(permissionType);
            request.Headers.Remove(Common.Constants.PermissionTypeKey);

            // Append the access token to the request.
            request.Headers.Authorization = new AuthenticationHeaderValue(
                Common.Constants.BearerAuthorizationScheme, accessToken);
        }

        private async Task<string> GetAccessToken(string permissionType)
        {
            string accessToken;
            if (permissionType.Equals(GraphPermissionType.Application.ToString(), StringComparison.CurrentCultureIgnoreCase))
            {
                // we use MSAL.NET to get a token to call the API for application
                accessToken = await this.tokenAcquisition.GetAccessTokenForAppAsync(Common.Constants.ScopeDefault);
            }
            else
            {
                // we use MSAL.NET to get a token to call the API On Behalf Of the current user
                accessToken = await this.tokenAcquisition.GetAccessTokenForUserAsync(new string[] { Common.Constants.ScopeGroupReadAll, Common.Constants.ScopeAppCatalogReadAll });
            }

            return accessToken;
        }

        private string ExtractPermissionType(HttpRequestHeaders headers)
        {
            if (headers != null && headers.Contains(Common.Constants.PermissionTypeKey))
            {
                var permissionType = headers.GetValues(Common.Constants.PermissionTypeKey).FirstOrDefault();
                return permissionType;
            }

            return string.Empty;
        }
    }
}
