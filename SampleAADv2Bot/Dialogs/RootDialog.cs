using BotAuth.AADv2;
using BotAuth.Dialogs;
using BotAuth.Models;
using BotAuth;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using SampleAADv2Bot.Extensions;
using SampleAADv2Bot.Services;
using Microsoft.Graph;
using System.Text;

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



        private List<DateTime> lstStartDateTme = null;
        // TBD - Replace with dependency injection 
        MeetingService meetingService = new MeetingService(new RoomService());

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
                Scopes = new string[] { "User.Read", "Calendars.ReadWrite", "Calendars.ReadWrite.Shared" },
                RedirectUrl = ConfigurationManager.AppSettings["aad:Callback"]
            };
            await context.Forward(new AuthDialog(new MSALAuthProvider(), options), ResumeAfterAuth, message, CancellationToken.None);
        }

        public async Task ResumeAfterAuth(IDialogContext authContext, IAwaitable<AuthResult> authResult)
        {
            this.result = await authResult;

            // Use token to call into service
            var json = await new HttpClient().GetWithAuthAsync(result.AccessToken, "https://graph.microsoft.com/v1.0/me");
            await authContext.PostAsync($"Hello {json.Value<string>("displayName")}!, I am Schedulo, I will help you schedule Meetings with your colleagues");

            PromptDialog.Text(authContext, this.SubjectMessageReceivedAsync, "Please enter the subject of the meeting.");
        }

        public async Task SubjectMessageReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.subject = await argument;
            await context.PostAsync("I have set the Subject of the meeting as " + subject + " !");
            //await ScheduleMeeitng(context, argument);

            PromptDialog.Text(context, this.DurationReceivedAsync, "Please enter the duration of the meeting.");
            //PromptDialog.Text(context, this.DateMessageReceivedAsync, "Please enter when you want to have the meeting. e.g. 2017-10-10");
        }

        public async Task DurationReceivedAsync(IDialogContext context, IAwaitable<string> argument)
        {
            this.duration = await argument;

            if (this.duration.IsNaturalNumber())
            {
                await context.PostAsync("The duration of your meeting is set as " + duration + "mins. Now I only need to know the name of the collegues!!");
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
                    await GetMeetingSuggestions(context, argument);
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

        private async Task GetMeetingSuggestions(IDialogContext context, IAwaitable<string> argument)
        {
            string startDate = date + "T00:00:00.000Z";
            string endDate = date + "T20: 00:00.00Z";
            List<Attendee> inputAttendee = new List<Attendee>();
            foreach (var i in normalizedEmails)
            {
                inputAttendee.Add(
                     new Attendee()
                     {
                         EmailAddress = new EmailAddress()
                         {
                             Address = i
                         }
                     }
                    );
            }
            Duration inputDuration = new Duration(new TimeSpan(0, normalizedDuration, 0));

            var userFindMeetingTimesRequestBody = new UserFindMeetingTimesRequestBody()
            {
                Attendees = inputAttendee,
                TimeConstraint = new TimeConstraint()
                {
                    Timeslots = new List<TimeSlot>()
                        {
                            new TimeSlot()
                            {
                                Start = new DateTimeTimeZone()
                                {
                                    DateTime = startDate,
                                    TimeZone = "UTC"
                                },
                                End = new DateTimeTimeZone()
                                {
                                    DateTime = endDate,
                                    TimeZone = "UTC"
                                }
                            }
                        }
                },
                MeetingDuration = inputDuration,
                MaxCandidates = 15,
                IsOrganizerOptional = false,
                ReturnSuggestionReasons = true,
                MinimumAttendeePercentage = 100

            };
            var meetingTimeSuggestion = await meetingService.GetMeetingsTimeSuggestions(result.AccessToken, userFindMeetingTimesRequestBody);
            var stringBuilder = new StringBuilder();
            int num = 1;
            lstStartDateTme = new List<DateTime>();
            foreach (var suggestion in meetingTimeSuggestion.MeetingTimeSuggestions)
            {
                DateTime startTime, endTime;
                DateTime.TryParse(suggestion.MeetingTimeSlot.Start.DateTime, out startTime);
                DateTime.TryParse(suggestion.MeetingTimeSlot.End.DateTime, out endTime);
                lstStartDateTme.Add(startTime);
                stringBuilder.AppendLine($"{num} {startTime.ToString()}  - {endTime.ToString()}\n");
                num++;
            }

            await context.PostAsync($"There are the options for meeting");
            await context.PostAsync(stringBuilder.ToString());
            PromptDialog.Text(context, this.ScheduleMeeitng,
                "Please select one slot by giving input such as 1, 2 ,3 .....");


        }

        public async Task ScheduleMeeitng(IDialogContext context, IAwaitable<string> message)
        {

            var optiontime = await (message);
            await context.PostAsync("the time slot you have chosen is " + lstStartDateTme[Convert.ToInt16(optiontime) - 1]);
            var startDate = lstStartDateTme[Convert.ToInt16(optiontime) - 1];
            try
            {
                List<Attendee> inputAttendee = new List<Attendee>();
                foreach (var i in normalizedEmails)
                {
                    inputAttendee.Add(
                         new Attendee()
                         {
                             EmailAddress = new EmailAddress()
                             {
                                 Address = i
                             }
                         }
                        );
                }

                var meeting = new Event()
                {
                    Subject = subject,
                    Body = new ItemBody()
                    {
                        ContentType = BodyType.Html,
                        Content = "Does this schedule work for you?"
                    },
                    Start = new DateTimeTimeZone()
                    {
                        //DateTime = "2017-10-29T08:30:00.000Z",


                        DateTime = startDate.Year.ToString("D4") +"-"+ startDate.Month.ToString("D2") + "-"+
                        startDate.Day.ToString("D2") + "T" + 
                        startDate.Hour.ToString("D2") +":"+ startDate.Minute.ToString("D2") + ":00.000Z",

                        TimeZone = "UTC"
                    },
                    End = new DateTimeTimeZone()
                    {
                        DateTime = startDate.Year.ToString("D4") + "-" + startDate.Month.ToString("D2") + "-" +
                        startDate.Day.ToString("D2") + "T" +


                        startDate.AddMinutes(normalizedDuration).Hour.ToString("D2") + ":" + startDate.AddMinutes(normalizedDuration).Minute.ToString("D2") +":00.000Z",

                        TimeZone = "UTC"
                    },
                    Location = new Location()
                    {
                        DisplayName = "Hoshi"
                    },
                    Attendees = inputAttendee
                };


                var scheduledMeeting = await meetingService.ScheduleMeeting(result.AccessToken, meeting);
                var stringBuilder = new StringBuilder();
                stringBuilder.AppendLine($"Booked the meeting! Subject is {subject}. Participants are");
                foreach (var email in normalizedEmails)
                    stringBuilder.AppendLine(email + " ");
                stringBuilder.AppendLine(". Date is " + startDate.ToShortDateString()+" "+ startDate.ToShortTimeString() + " - " + startDate.AddMinutes(normalizedDuration).ToShortTimeString() + ".");

                await context.PostAsync(stringBuilder.ToString());

            }
            catch (Exception ex)
            {
                var msg = ex.Message;
                throw ex;
            }
        }
    }
}