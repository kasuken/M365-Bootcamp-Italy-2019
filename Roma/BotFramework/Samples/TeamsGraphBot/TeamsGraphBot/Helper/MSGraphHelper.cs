using Microsoft.Bot.Schema;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace TeamsGraphBot.Helper
{
    public class MSGraphHelper
    {
        public static async Task<Site> GetSiteContext(TokenResponse tokenResponse, string aadGroupId)
        {
            var graphClient = new GraphServiceClient(
                                new DelegateAuthenticationProvider(
                                requestMessage =>
                                {
                                    // Append the access token to the request.
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenResponse.Token);

                                    // Get event times in the current time zone.
                                    requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                                    return Task.CompletedTask;
                                })
                            );

            return await graphClient.Groups[aadGroupId].Sites["root"].Request().GetAsync();
        }

        public static MeetingTimeSuggestionsResult GetFreeTime(TokenResponse tokenResponse, string aadGroupId)
        {
            var graphClient = new GraphServiceClient(
                                new DelegateAuthenticationProvider(
                                requestMessage =>
                                {
                                    // Append the access token to the request.
                                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", tokenResponse.Token);

                                    // Get event times in the current time zone.
                                    requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

                                    return Task.CompletedTask;
                                })
                            );
            return graphClient.Me.FindMeetingTimes().Request().PostAsync().Result;
        }
    }
}
