// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;

namespace TeamsGraphBot.Bots
{
	public class GraphBot<T> : DialogBot<T> where T : Dialog
	{
		public GraphBot(ConversationState conversationState, UserState userState, T dialog, ILogger<DialogBot<T>> logger)
			: base(conversationState, userState, dialog, logger)
		{
		}

		protected override async Task OnMembersAddedAsync(IList<ChannelAccount> membersAdded, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
		{
			foreach (var member in membersAdded)
			{
				if (member.Id != turnContext.Activity.Recipient.Id)
				{
					await turnContext.SendActivityAsync(MessageFactory.Text($"Hello world!"), cancellationToken).ConfigureAwait(false);
				}
			}
		}


	}
}
