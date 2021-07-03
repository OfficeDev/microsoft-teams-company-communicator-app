// <copyright file="AuthenticationServiceCollectionExtensions.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Authentication
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.AspNetCore.Authentication.AzureAD.UI;
    using Microsoft.AspNetCore.Authentication.JwtBearer;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Identity.Web;
    using Microsoft.Identity.Web.TokenCacheProviders.InMemory;
    using Microsoft.IdentityModel.Tokens;

    /// <summary>
    /// Extension class for registering auth services in DI container.
    /// </summary>
    public static class AuthenticationServiceCollectionExtensions
    {
        /// <summary>
        /// Extension method to register the authentication services.
        /// </summary>
        /// <param name="services">IServiceCollection instance.</param>
        /// <param name="configuration">The configuration instance.</param>
        /// <param name="authenticationOptions">The authentication options.</param>
        public static void AddAuthentication(
            this IServiceCollection services,
            IConfiguration configuration,
            AuthenticationOptions authenticationOptions)
        {
            AuthenticationServiceCollectionExtensions.RegisterAuthenticationServices(services, configuration, authenticationOptions);

            AuthenticationServiceCollectionExtensions.RegisterAuthorizationPolicy(services);
        }

        // This method works specifically for single tenant application.
        private static void RegisterAuthenticationServices(
            IServiceCollection services,
            IConfiguration configuration,
            AuthenticationOptions authenticationOptions)
        {
            AuthenticationServiceCollectionExtensions.ValidateAuthenticationOptions(authenticationOptions);

            services.AddMicrosoftIdentityWebApiAuthentication(configuration)
                    .EnableTokenAcquisitionToCallDownstreamApi()
                    .AddInMemoryTokenCaches();
            services.Configure<JwtBearerOptions>(JwtBearerDefaults.AuthenticationScheme, options =>
            {
                var azureADOptions = new AzureADOptions
                {
                    Instance = authenticationOptions.AzureAdInstance,
                    TenantId = authenticationOptions.AzureAdTenantId,
                    ClientId = authenticationOptions.AzureAdClientId,
                };
                options.Authority = $"{azureADOptions.Instance}{azureADOptions.TenantId}/v2.0";
                options.SaveToken = true;
                options.TokenValidationParameters.ValidAudiences = AuthenticationServiceCollectionExtensions.GetValidAudiences(authenticationOptions);
                options.TokenValidationParameters.AudienceValidator = AuthenticationServiceCollectionExtensions.AudienceValidator;
                options.TokenValidationParameters.ValidIssuers = AuthenticationServiceCollectionExtensions.GetValidIssuers(authenticationOptions);
            });
        }

        private static void ValidateAuthenticationOptions(AuthenticationOptions authenticationOptions)
        {
            if (string.IsNullOrWhiteSpace(authenticationOptions?.AzureAdClientId))
            {
                throw new ApplicationException("AzureAd ClientId is missing in the configuration file.");
            }

            if (string.IsNullOrWhiteSpace(authenticationOptions?.AzureAdTenantId))
            {
                throw new ApplicationException("AzureAd TenantId is missing in the configuration file.");
            }

            if (string.IsNullOrWhiteSpace(authenticationOptions?.AzureAdApplicationIdUri))
            {
                throw new ApplicationException("AzureAd ApplicationIdUri is missing in the configuration file.");
            }

            if (string.IsNullOrWhiteSpace(authenticationOptions?.AzureAdValidIssuers))
            {
                throw new ApplicationException("AzureAd ValidIssuers is missing in the configuration file.");
            }
        }

        private static IEnumerable<string> SplitAuthenticationOptionsList(string stringInAuthenticationOptions)
        {
            var settings = stringInAuthenticationOptions
                ?.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                ?.Select(p => p.Trim());
            if (settings == null)
            {
                throw new ApplicationException($"Invalid list of settings in authentication options.");
            }

            return settings;
        }

        private static IEnumerable<string> GetValidAudiences(AuthenticationOptions authenticationOptions)
        {
            var validAudiences = new List<string>
            {
                authenticationOptions.AzureAdClientId,
                authenticationOptions.AzureAdApplicationIdUri.ToLower(),
            };

            return validAudiences;
        }

        private static IEnumerable<string> GetValidIssuers(AuthenticationOptions authenticationOptions)
        {
            var tenantId = authenticationOptions.AzureAdTenantId;

            var validIssuers =
                AuthenticationServiceCollectionExtensions.SplitAuthenticationOptionsList(
                    authenticationOptions.AzureAdValidIssuers);

            validIssuers = validIssuers.Select(validIssuer => validIssuer.Replace("TENANT_ID", tenantId));

            return validIssuers;
        }

        private static void RegisterAuthorizationPolicy(IServiceCollection services)
        {
            services.AddAuthorization(options =>
            {
                var mustContainUpnClaimRequirement = new MustBeValidUpnRequirement();
                options.AddPolicy(
                    PolicyNames.MustBeValidUpnPolicy,
                    policyBuilder => policyBuilder
                    .AddRequirements(mustContainUpnClaimRequirement)
                    .RequireAuthenticatedUser()
                    .Build());
                options.AddPolicy(
                    PolicyNames.MSGraphGroupDataPolicy,
                    policyBuilder => policyBuilder
                    .AddRequirements(new MSGraphScopeRequirement(new string[] { Common.Constants.ScopeGroupReadAll }))
                    .RequireAuthenticatedUser()
                    .Build());
            });

            services.AddScoped<IAuthorizationHandler, MustBeValidUpnHandler>();
            services.AddScoped<IAuthorizationHandler, MSGraphScopeHandler>();
        }

        private static bool AudienceValidator(
            IEnumerable<string> tokenAudiences,
            SecurityToken securityToken,
            TokenValidationParameters validationParameters)
        {
            if (tokenAudiences == null || !tokenAudiences.Any())
            {
                throw new ApplicationException("No audience defined in token!");
            }

            var validAudiences = validationParameters.ValidAudiences;
            if (validAudiences == null || !validAudiences.Any())
            {
                throw new ApplicationException("No valid audiences defined in validationParameters!");
            }

            foreach (var tokenAudience in tokenAudiences)
            {
                if (validAudiences.Any(validAudience => validAudience.Equals(tokenAudience, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
