using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using TeamsGraphBot.Helper;

namespace TeamsGraphBot.Dialogs
{
    public class MainDialog : ComponentDialog
    {
        protected readonly ILogger _logger;

        public MainDialog(ILogger<MainDialog> logger, IConfiguration configuration) : base(nameof(MainDialog))
        {
            _logger = logger;

            AddDialog(new OAuthPrompt(
                nameof(OAuthPrompt),
                new OAuthPromptSettings
                {
                    ConnectionName = configuration["ConnectionName"],
                    Text = "Please login",
                    Title = "Login",
                    Timeout = 300000
                }));


            AddDialog(new TextPrompt(nameof(TextPrompt)));
            AddDialog(new ChoicePrompt(nameof(ChoicePrompt)));

            AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
            {
                PromptStepAsync,
                DisplayContextInfoStepAsync,
                CommandStepAsync,
                ProcessStepAsync
            }));

            InitialDialogId = nameof(WaterfallDialog);

        }

        private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }



        private async Task<DialogTurnResult> DisplayContextInfoStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                var tokenResponse = stepContext.Result as TokenResponse;
                if (tokenResponse?.Token != null)
                {
                    var teamsContext = stepContext.Context.TurnState.Get<ITeamsContext>();

                    if (teamsContext != null) // the bot is used inside MS Teams
                    {
                        if (teamsContext.Team != null) // inside team
                        {
                            await stepContext.Context.SendActivityAsync(MessageFactory.Text("We're in MS Teams, inside a Team! :)"), cancellationToken);

                            var team = teamsContext.Team;
                            var teamDetails = await teamsContext.Operations.FetchTeamDetailsWithHttpMessagesAsync(team.Id);
                            var token = tokenResponse.Token;
                            var aadGroupId = teamDetails.Body.AadGroupId;

                            var siteInfo = await MSGraphHelper.GetSiteContext(tokenResponse, aadGroupId);

                            await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Site Id: {siteInfo.Id}, Site Title: {siteInfo.DisplayName}, Site Url: {siteInfo.WebUrl}"), cancellationToken).ConfigureAwait(false);
                            return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Would you like to do? (type 'agenda')") }, cancellationToken);

                        }
                        else // private or group chat
                        {
                            await stepContext.Context.SendActivityAsync(MessageFactory.Text($"We're in MS Teams but not in Team"), cancellationToken).ConfigureAwait(false);
                            return await stepContext.PromptAsync(nameof(TextPrompt), new PromptOptions { Prompt = MessageFactory.Text("Would you like to do? (type 'agenda')") }, cancellationToken);

                        }
                    }
                    else // outside MS Teams
                    {
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text("We're not in MS Teams context"), cancellationToken).ConfigureAwait(false);
                    }
                }
            }

            return await stepContext.EndDialogAsync();
        }
        private async Task<DialogTurnResult> CommandStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            stepContext.Values["command"] = stepContext.Result;

            // Call the prompt again because we need the token. The reasons for this are:
            // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
            // about refreshing it. We can always just call the prompt again to get the token.
            // 2. We never know how long it will take a user to respond. By the time the
            // user responds the token may have expired. The user would then be prompted to login again.
            //
            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed.
            return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
        }

        private async Task<DialogTurnResult> ProcessStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
        {
            if (stepContext.Result != null)
            {
                // We do not need to store the token in the bot. When we need the token we can
                // send another prompt. If the token is valid the user will not need to log back in.
                // The token will be available in the Result property of the task.
                var tokenResponse = stepContext.Result as TokenResponse;

                var teamsContext = stepContext.Context.TurnState.Get<ITeamsContext>();


                // If we have the token use the user is authenticated so we may use it to make API calls.
                if (tokenResponse?.Token != null)
                {
                    var temp = ((string)stepContext.Values["command"] ?? string.Empty).ToLowerInvariant();

                    var parts = teamsContext.GetActivityTextWithoutMentions().Split(' ');
                    //Regex.Replace(HttpUtility.HtmlDecode(temp), "<.*?>", string.Empty).Split(' ');

                    var command = parts[0];

                    if (command == "agenda")
                    {
                        var res = MSGraphHelper.GetFreeTime(tokenResponse, "");

                        IMessageActivity reply = null;

                        if (res.MeetingTimeSuggestions.Any())
                        {
                            var count = res.MeetingTimeSuggestions.Count();
                            if (count > 3)
                            {
                                count = 3;
                            }

                            reply = MessageFactory.Attachment(new List<Microsoft.Bot.Schema.Attachment>());
                            reply.AttachmentLayout = AttachmentLayoutTypes.Carousel;

                            for (var i = 0; i < count; i++)
                            {
                                var slot = res.MeetingTimeSuggestions.ToList()[i];


                                var card = new HeroCard(
                                    $"Proposal number #{i}",
                                    $"{slot.SuggestionReason}",
                                    $"{DateTime.Parse(slot.MeetingTimeSlot.Start.DateTime).ToString("dd/MM/yyyy H:mm")} --> {DateTime.Parse(slot.MeetingTimeSlot.End.DateTime).ToString("dd/MM/yyyy H:mm")}");
                                reply.Attachments.Add(card.ToAttachment());
                            }
                        }
                        else
                        {
                            reply = MessageFactory.Text("Unable to find any recent unread mail.");
                        }

                        await stepContext.Context.SendActivityAsync(reply, cancellationToken);

                    }
                    else 
                    { 
                        await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Your token is: {tokenResponse?.Token}"), cancellationToken);

                    } 
                     
                }
            }
            else
            {
                await stepContext.Context.SendActivityAsync(MessageFactory.Text("We couldn't log you in. Please try again later."), cancellationToken);
            }

            return await stepContext.EndDialogAsync(cancellationToken: cancellationToken);
        }
    }
}
