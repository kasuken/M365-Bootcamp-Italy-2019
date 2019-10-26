using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Logging;
using System.Threading;
using System.Threading.Tasks;
using TeamsGraphBot.Extensions;

namespace TeamsGraphBot.Bots
{
	public class DialogBot<T> : ActivityHandler where T : Dialog
	{
		protected readonly BotState ConversationState;
		protected readonly Dialog Dialog;
		protected readonly ILogger Logger;
		protected readonly BotState UserState;

		public DialogBot(ConversationState conversationState, UserState userState, T dialog, ILogger<DialogBot<T>> logger)
		{
			ConversationState = conversationState;
			UserState = userState;
			Dialog = dialog;
			Logger = logger;
		}

		public override async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
		{
			if (turnContext?.Activity?.Type == ActivityTypes.Invoke && turnContext.Activity.ChannelId == "msteams")
				await Dialog.Run(turnContext, ConversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
			else
				await base.OnTurnAsync(turnContext, cancellationToken);

			// Save any state changes that might have occured during the turn.
			await ConversationState.SaveChangesAsync(turnContext, false, cancellationToken);
			await UserState.SaveChangesAsync(turnContext, false, cancellationToken);
		}

		protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
		{
			Logger.LogInformation("Running dialog with Message Activity.");

			// Run the Dialog with the new message Activity.
			await Dialog.Run(turnContext, ConversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
		}

		protected override async Task OnTokenResponseEventAsync(ITurnContext<IEventActivity> turnContext, CancellationToken cancellationToken)
		{
			Logger.LogInformation("Running dialog with Token Response Event Activity.");

			// Run the Dialog with the new Token Response Event Activity.
			await Dialog.Run(turnContext, ConversationState.CreateProperty<DialogState>(nameof(DialogState)), cancellationToken);
		}
	}
}
