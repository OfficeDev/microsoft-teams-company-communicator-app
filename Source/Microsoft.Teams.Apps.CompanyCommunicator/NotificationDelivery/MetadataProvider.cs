// <copyright file="MetadataProvider.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.NotificationDelivery
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.ApplicationInsights;
    using Microsoft.ApplicationInsights.DataContracts;
    using Microsoft.Bot.Connector;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Extensions.Configuration;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.UserData;
    using Newtonsoft.Json.Linq;

    /// <summary>
    /// The class provides the user data by using the team data captured in the bot.
    /// </summary>
    public class MetadataProvider
    {
        private readonly IConfiguration configuration;
        private readonly UserDataRepository userDataRepository;
        private readonly TeamDataRepository teamDataRepository;
        private readonly TelemetryClient telemetryClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="MetadataProvider"/> class.
        /// </summary>
        /// <param name="configuration">The configuration.</param>
        /// <param name="userDataRepository">User Data repository service.</param>
        /// <param name="teamDataRepository">Team Data repository service.</param>
        public MetadataProvider(
            IConfiguration configuration,
            UserDataRepository userDataRepository,
            TeamDataRepository teamDataRepository,
            TelemetryClient telemetryClient)
        {
            this.configuration = configuration;
            this.userDataRepository = userDataRepository;
            this.teamDataRepository = teamDataRepository;
            this.telemetryClient = telemetryClient;
        }

        /// <summary>
        /// Get all user data.
        /// </summary>
        /// <returns>User data dictionary.</returns>
        public async Task<Dictionary<string, UserDataEntity>> GetUserDataDictionaryAsync()
        {
            var userDataEntities = await this.userDataRepository.GetAllAsync();
            var result = new Dictionary<string, UserDataEntity>();
            foreach (var userDataEntity in userDataEntities)
            {
                result.Add(userDataEntity.AadId, userDataEntity);
            }

            return result;
        }

        /// <summary>
        /// Get all teams' roster.
        /// </summary>
        /// <returns>All teams' roster.</returns>
        public async Task<IDictionary<string, UserDataEntity>> GetAllTeamsRostersAsync()
        {
            var rosterUserDataEntityDictionary = new Dictionary<string, UserDataEntity>();

            var teams = await this.teamDataRepository.GetAllAsync();
            foreach (var team in teams)
            {
                var roster = await this.GetTeamRosterAsync(team.ServiceUrl, team.TeamId);
                this.AddRosterToUserDataEntityDictionary(roster, rosterUserDataEntityDictionary);
            }

            return rosterUserDataEntityDictionary;
        }

        /// <summary>
        /// Get multiple teams' roster.
        /// </summary>
        /// <param name="teamIds">List of team ids.</param>
        /// <returns>Roster of the multiple teams.</returns>
        public async Task<IDictionary<string, UserDataEntity>> GetTeamsRostersAsync(IEnumerable<string> teamIds)
        {
            var teamDataEntities = await this.teamDataRepository.GetTeamDataEntitiesByIdsAsync(teamIds);

            var rosterUserDataEntityDictionary = new Dictionary<string, UserDataEntity>();

            foreach (var teamDataEntity in teamDataEntities)
            {
                var roster = await this.GetTeamRosterAsync(teamDataEntity.ServiceUrl, teamDataEntity.TeamId);

                this.AddRosterToUserDataEntityDictionary(roster, rosterUserDataEntityDictionary);
            }

            return rosterUserDataEntityDictionary;
        }

        /// <summary>
        /// Merge a roster list to a dictionary of users.
        /// </summary>
        /// <param name="roster">Roster list.</param>
        /// <param name="rosterUserDataEntityDictionary">Dictionary of users.</param>
        public void AddRosterToUserDataEntityDictionary(
            IEnumerable<UserDataEntity> roster,
            IDictionary<string, UserDataEntity> rosterUserDataEntityDictionary)
        {
            foreach (var userDataEntity in roster)
            {
                if (!rosterUserDataEntityDictionary.ContainsKey(userDataEntity.AadId))
                {
                    rosterUserDataEntityDictionary.Add(userDataEntity.AadId, userDataEntity);
                }
            }
        }

        /// <summary>
        /// Get a team's roster.
        /// </summary>
        /// <param name="serviceUrl">The service URL.</param>
        /// <param name="teamId">Team id, e.g. "19:44777361677b439281a0f0cd914cb149@thread.skype".</param>
        /// <returns>Roster of the team with the passed in id.</returns>
        public async Task<IEnumerable<UserDataEntity>> GetTeamRosterAsync(string serviceUrl, string teamId)
        {
            try
            {
                MicrosoftAppCredentials.TrustServiceUrl(serviceUrl);

                var botAppId = this.configuration.GetValue<string>("MicrosoftAppId");
                var botAppPassword = this.configuration.GetValue<string>("MicrosoftAppPassword");

                var connectorClient = new ConnectorClient(
                    new Uri(serviceUrl),
                    botAppId,
                    botAppPassword);

                var members = await connectorClient.Conversations.GetConversationMembersAsync(teamId);

                return members.Select(member =>
                {
                    var userDataEntity = new UserDataEntity
                    {
                        UserId = member.Id,
                        Name = member.Name,
                    };

                    if (member.Properties is JObject jObject)
                    {
                        if (jObject["objectId"] == null)
                        {
                            this.telemetryClient.TrackTrace($"Unexpected Channel Member ID '{member.Id}' getted from '{teamId}' this user probably might left or excluded from tenant.", SeverityLevel.Warning);
                            return null;
                        }

                        userDataEntity.AadId = jObject["objectId"].ToString();
                        userDataEntity.Email = jObject["email"]?.ToString();
                        userDataEntity.Upn = jObject["userPrincipalName"]?.ToString();
                        userDataEntity.TenantId = jObject["tenantId"].ToString();
                        userDataEntity.ConversationId = null;
                        userDataEntity.ServiceUrl = serviceUrl;
                    }

                    return userDataEntity;
                }).Where(member => member != null);
            }
            catch
            {
                throw new ApplicationException("The app is not authorized to access the bot service. Please send a message to the bot, then it will work.");
            }
        }

        /// <summary>
        /// Deduplicate user data.
        /// </summary>
        /// <param name="rosterUserDataEntityDictionary">User Data Dictionary from roster.</param>
        /// <param name="usersUserDataEntityDictionary">User Data Dictionary from user data table.</param>
        public void AddRosterDictionaryToUserDictionary(
            IDictionary<string, UserDataEntity> rosterUserDataEntityDictionary,
            IDictionary<string, UserDataEntity> usersUserDataEntityDictionary)
        {
            foreach (var rosterUserKeyValuePair in rosterUserDataEntityDictionary)
            {
                if (!usersUserDataEntityDictionary.ContainsKey(rosterUserKeyValuePair.Key))
                {
                    usersUserDataEntityDictionary.Add(rosterUserKeyValuePair.Key, rosterUserKeyValuePair.Value);
                }
                else
                {
                    // Want to merge the two user data entities to backfill data that comes from the roster call
                    // (e.g. UPN, email, etc.) to the user data repo while keeping the conversation ID.
                    var conversationId = usersUserDataEntityDictionary[rosterUserKeyValuePair.Key].ConversationId;
                    var rosterUserDataEntity = rosterUserKeyValuePair.Value;
                    rosterUserDataEntity.ConversationId = conversationId;
                    usersUserDataEntityDictionary[rosterUserKeyValuePair.Key] = rosterUserDataEntity;
                }
            }
        }

        /// <summary>
        /// Creates user data entities for a list of team IDs.
        /// </summary>
        /// <param name="teamIds">Team IDs.</param>
        /// <returns>List of user data entities.</returns>
        public async Task<IEnumerable<UserDataEntity>> GetTeamsReceiverEntities(IEnumerable<string> teamIds)
        {
            var teamDataEntities = await this.teamDataRepository.GetTeamDataEntitiesByIdsAsync(teamIds);

            IList<UserDataEntity> teamReceiverEntities = new List<UserDataEntity>();

            foreach (var teamDataEntity in teamDataEntities)
            {
                teamReceiverEntities.Add(
                    new UserDataEntity
                    {
                        AadId = teamDataEntity.TeamId,
                        ConversationId = teamDataEntity.TeamId,
                        ServiceUrl = teamDataEntity.ServiceUrl,
                    });
            }

            return teamReceiverEntities;
        }
    }
}
