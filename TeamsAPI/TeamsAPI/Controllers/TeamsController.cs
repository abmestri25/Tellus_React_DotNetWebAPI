using Azure.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.ExternalConnectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TeamsAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TeamsController : ControllerBase
    {

        private IConfiguration _configuration { get; set; }
        public TeamsController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IEnumerable<TeamsGroup> Get(string code)
        {
            List<TeamsGroup> groups = new List<TeamsGroup>();
            if (!string.IsNullOrEmpty(code))
            {
                var scopes = new[] { "user.read", "user.readbasic.all", "group.read.all", "channel.readbasic.all" };
               
                var tenantId = _configuration.GetValue(typeof(String), "AzureAD:TenantId").ToString();
                var clientId = _configuration.GetValue(typeof(String), "AzureAD:ClientId").ToString();
                var clientSecret = _configuration.GetValue(typeof(String), "AzureAD:ClientSecret").ToString();

                var options = new AuthorizationCodeCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                    RedirectUri = new Uri($"http://localhost:3000")
                };
                var authCodeCredential = new AuthorizationCodeCredential(tenantId, clientId, clientSecret, code, options); //The authorization code flow enables native and web apps to securely obtain tokens in the name of the user.
                var graphClient = new GraphServiceClient(authCodeCredential, scopes);

                /*var me = graphClient.Me.Request().GetAsync().Result;*/

                var allGroups = graphClient.Groups.Request().GetAsync().Result;
                foreach (var group in allGroups)
                {
                    TeamsGroup team = new TeamsGroup();
                    team.Name = group.DisplayName;
                    team.Description = group.Description;
                    team.Visibility = group.Visibility;
                    try
                    {
                        var channels = graphClient.Teams[group.Id].Channels.Request().GetAsync().Result;
                        team.Channels = new List<TeamChannel>();
                        foreach (var channel in channels)
                        {
                            team.Channels.Add(new TeamChannel { Name = channel.DisplayName, MembershipType = channel.MembershipType.ToString() });
                        }
                    }
                    catch (Exception ex) { }
                    groups.Add(team);
                }
            }

            return groups.ToArray();
        }
    }
    public class TeamsGroup
    {
        public string Name { get; set; }
        public string Description { get; set; } //about group
        public string Visibility { get; set; } //private or public
        public List<TeamChannel> Channels { get; set; }
    }
    public class TeamChannel
    {
        public string Name { get; set; }
        public string MembershipType { get; set; } //private or public
    }
}
