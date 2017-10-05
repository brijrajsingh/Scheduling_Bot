using BotAuth.AADv2;
using BotAuth.Dialogs;
using BotAuth.Models;
using BotAuth;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using SampleAADv2Bot.Extensions;

namespace SampleAADv2Bot.Dialogs
{
    [Serializable]
    public class RootDialog : IDialog<string>
    {
        private string subject = null;
        private string duration = null;
        private string emails = null;
        private string date = null;

        private int normalizedDuration = 0;
        private string[] normalizedEmails;



        //Scheduling
        AuthResult result = null;


        public async Task StartAsync(IDialogContext context)
        {
            context.Wait(MessageReceivedAsync);
        }

        public virtual async Task MessageReceivedAsync(IDialogContext context, IAwaitable<IMessageActivity> item)
        {
            var message = await item;

            // Initialize AuthenticationOptions and forward to AuthDialog for token
            AuthenticationOptions options = new AuthenticationOptions()
            {
                UseMagicNumber = false,
                Authority = ConfigurationManager.AppSettings["aad:Authority"],
                ClientId = ConfigurationManager.AppSettings["aad:ClientId"],
                ClientSecret = ConfigurationManager.AppSettings["aad:ClientSecret"],
                Scopes = new string[] { "User.Read" },
                RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
            };
            await context.Forward(new AuthDialog(new MSALAuthProvider(), options), ResumeAfterAuth, message, CancellationToken.None);
        }

        public async Task ResumeAfterAuth(IDialogContext authContext, IAwaitable<AuthResult> authResult)
        {
            var result = await authResult;

            // Use token to call into service
            var json = await new HttpClient().GetWithAuthAsync(result.AccessToken, "https://graph.microsoft.com/v1.0/me");
            await authContext.PostAsync($"Hello {json.Value<string>("displayName")}!, I am Schedulo, I will help you schedule Meetings with your colleagues");
            PromptDialog.Text(authContext, this.SubjectMessageReceivedAsync, "Please enter the subject of the meeting.");
        }

        public async Task SubjectMessageReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.subject = await argument;
            await context.PostAsync("I have set the Subject of the meeting as " + subject+" !");
            PromptDialog.Text(context, this.DurationReceivedAsync, "Please enter the duration of the meeting.");
            //PromptDialog.Text(context, this.DateMessageReceivedAsync, "Please enter when you want to have the meeting. e.g. 2017-10-10");
        }

        public async Task DurationReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.duration = await argument;
          
            if (this.duration.IsNaturalNumber())
            {
                await context.PostAsync("The duration of your meeting is set as "+duration+"mins. Now I only need to know the name of the collegues!!");
                normalizedDuration = Int32.Parse(this.duration);
                PromptDialog.Text(context, this.EmailsMessageReceivedAsync, "Please enter emails of the participants separeted by comma.");
            }
            else
            {
                await context.PostAsync("The entered duration is not valid");
                PromptDialog.Text(context, this.DurationReceivedAsync, "Please enter duration of the meeting.");
            }
        }

        public async Task EmailsMessageReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.emails = await argument;
            //remove space
            this.emails = this.emails.Replace(" ", "").Replace("　", "");
            this.emails = this.emails.Replace("&#160;", "").Replace("&#160:^", "");
            this.emails = System.Text.RegularExpressions.Regex.Replace(this.emails, "\\(.+?\\)", "");

            if (this.emails.IsEmailAddressList())
            {
                normalizedEmails = this.emails.Split(',');
                await context.PostAsync("You would like to invite ");
                foreach (var i in normalizedEmails)
                    await context.PostAsync(i);
                PromptDialog.Text(context, this.DateMessageReceivedAsync, "Please enter when you want to have the meeting. e.g. 2017-10-10");
            }
            else
            {
                await context.PostAsync("Please enter only emails.");
                PromptDialog.Text(context, this.EmailsMessageReceivedAsync, "Please enter emails of the participants separeted by comma.");
            }
        }

        public async Task DateMessageReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.date = await argument;
            DateTime dateTime;
            DateTime.TryParse(date, out dateTime);

            if (dateTime != DateTime.MinValue && dateTime != DateTime.MaxValue)
            {
                try
                {
                    await context.PostAsync("Please wait While I search for the available time slots......");
                    //await GetMeetingSuggestions(context, argument);
                }
                catch (Exception e)
                {
                    PromptDialog.Text(context, this.DateMessageReceivedAsync, "Please enter another date.");

                }
            }
            else
            {
                PromptDialog.Text(context, this.DateMessageReceivedAsync, "Please enter when you want to have the meeting. e.g. 2017-10-10");
            }
        }
    }
}