using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using TeamGovernance.API.Models;

namespace TeamGovernance.API
{
    public class GraphService
    {

        private const string adminConsentUrlFormat = "https://login.microsoftonline.com/{0}/adminconsent?client_id={1}&redirect_uri={2}";

        private const string tenantIdClaimType = "http://schemas.microsoft.com/identity/claims/tenantid";
        private const string authorityFormat = "https://login.microsoftonline.com/{0}/v2.0";
        private const string msGraphScope = "https://graph.microsoft.com/.default";
        private const string msGraphQuery = "https://graph.microsoft.com/v1.0/users";

        public const string graphBasePath = "https://graph.microsoft.com/v1.0";
        public const string graphBasePathBeta = "https://graph.microsoft.com/beta";

        public async Task<string> GetGraphToken(string tenantId)
        {
            try
            {
                IConfidentialClientApplication daemonClient;
                daemonClient = ConfidentialClientApplicationBuilder.Create("b6be2ec6-eb67-4422-8d57-e673f0cfdfe8")
                    .WithAuthority(string.Format(authorityFormat, tenantId))
                    .WithRedirectUri("https://localhost:44336/")
                    .WithClientSecret("7GUBnA7Gn5gan4/DEFH_G=Thu/3B..VB")
                    .Build();

                AuthenticationResult authResult = await daemonClient.AcquireTokenForClient(new[] { msGraphScope })
                    .ExecuteAsync();

                return authResult.AccessToken;
            }
            catch
            {
                return null;
            }
        }

        public async Task<string> GetTenantId(string domain)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                string tenantID = "";

                var url = "https://login.windows.net/" + domain + "/v2.0/.well-known/openid-configuration";

                var response = await client.GetAsync(url);
                var content = await response.Content.ReadAsStringAsync();
                dynamic json = JsonConvert.DeserializeObject(content);
                tenantID = json.authorization_endpoint;
                tenantID = tenantID.Substring(26, 36);

                return tenantID;
            }
        }

        public async Task<string> CreateGroup(string domain, string siteTitle, string siteAlias, string owner)
        {
            var tenantId = await GetTenantId(domain);
            var token = await GetGraphToken(tenantId);

            using (var client = new HttpClient())
            {
                var request = new HttpRequestMessage(HttpMethod.Post, $"{graphBasePathBeta}/groups");
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                var payload = "{\"description\": \"-\",\"visibility\":\"Private\",\"owners@odata.bind\": [\"https://graph.microsoft.com/beta/users/--param3--\"],\"members@odata.bind\": [\"https://graph.microsoft.com/beta/users/--param3--\"], \"displayName\": \"--param1--\",\"groupTypes\": [\"Unified\"],\"mailEnabled\": true,\"mailNickname\": \"--param2--\", \"resourceBehaviorOptions\": [\"WelcomeEmailDisabled\",\"HideGroupInOutlook\" ] ,\"securityEnabled\": false}";

                payload = payload.Replace("--param1--", siteTitle);
                payload = payload.Replace("--param2--", siteAlias);
                payload = payload.Replace("--param3--", owner);

                HttpContent c = new StringContent(payload, Encoding.UTF8, "application/json");
                request.Content = c;

                var response = await client.SendAsync(request);

                var location = response.Headers.Where(r => r.Key == "Location").FirstOrDefault().Value.FirstOrDefault();

                return location.Substring(location.IndexOf("directoryObjects/")).Replace("/Microsoft.DirectoryServices.Group", "").Replace("directoryObjects/","");
            }

            return string.Empty;
        }

        public async Task<Team> CreateTeamFromGroup(string domain, string groupId)
        {
            var team = new Team();

            var tenantId = await GetTenantId(domain);
            var token = await GetGraphToken(tenantId);

            var client = new HttpClient();

            var request = new HttpRequestMessage(HttpMethod.Put, $"{graphBasePathBeta}/groups/{groupId}/team");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var payload = "{\"memberSettings\": { \"allowCreateUpdateChannels\": true },\"messagingSettings\": {\"allowUserEditMessages\": true,\"allowUserDeleteMessages\": true},\"funSettings\": {\"allowGiphy\": true,\"giphyContentRating\": \"strict\"}}";

            HttpContent c = new StringContent(payload, Encoding.UTF8, "application/json");
            request.Content = c;

            var response = await client.SendAsync(request);

            team.TeamUrl = response.Headers.Where(r => r.Key == "Location").FirstOrDefault().Value.FirstOrDefault();

            team.TeamId = team.TeamUrl.Substring(team.TeamUrl.IndexOf("team('")).Replace("')", "").Replace("team('", "");

            return team;
        }
    }
}
