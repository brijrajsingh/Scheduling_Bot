using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Graph;
using System.Threading.Tasks;
using System.Net.Http;
using BotAuth.Models;
using Newtonsoft.Json;

namespace SampleAADv2Bot.Services
{
    [Serializable]
    public class MeetingService : IMeetingService
    {
        private readonly string FindsMeetingTimeEndpoint = "https://graph.microsoft.com/v1.0/me/findMeetingTimes";
        private readonly string ScheduleMeetingEndpoint = "https://graph.microsoft.com/v1.0/me/events";
        private readonly IRoomService roomService;

        public MeetingService(IRoomService roomService)
        {
            this.roomService = roomService;
        }

        public async Task<MeetingTimeSuggestionsResult> GetMeetingsTimeSuggestions(string accessToken, UserFindMeetingTimesRequestBody userFindMeetingTimesRequestBody)
        {
            try
            {
                //var rooms = roomService.GetRooms();
                //roomService.AddRooms(userFindMeetingTimesRequestBody, rooms);
                var httpResponseMessage = await ApplyOperation(FindsMeetingTimeEndpoint, accessToken, userFindMeetingTimesRequestBody, string.Empty);
                var meetingTimeSuggestionsResult = JsonConvert.DeserializeObject<MeetingTimeSuggestionsResult>(await httpResponseMessage.Content.ReadAsStringAsync());
                return meetingTimeSuggestionsResult;
            }
            catch (Exception ex)
            {
                // TBD - log exception
                var messgae = ex.Message;
                throw ex;
            }
        }


        public async Task<Event> ScheduleMeeting(string accessToken, Event meeting)
        {
            try
            {
                var httpResponseMessage = await ApplyOperation(ScheduleMeetingEndpoint, accessToken, meeting, "UTC");
                var scheduledMeeting = JsonConvert.DeserializeObject<Event>(await httpResponseMessage.Content.ReadAsStringAsync());
                return scheduledMeeting;
            }
            catch (Exception ex)
            {
                // TBD - log exception
                var messgae = ex.Message;
                throw ex;
            }
        }

        private async Task<HttpResponseMessage> ApplyOperation(string endpoint, string accessToken, object payload, string preferTimeZone)
        {
            using (var httpClient = new HttpClient())
            {
                var serializedObject = JsonConvert.SerializeObject(payload);
                var body = new StringContent(serializedObject);
                body.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
                if (!string.IsNullOrEmpty(preferTimeZone))
                {
                    body.Headers.Add("Prefer", preferTimeZone);
                }
                httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
                var httpResponseMessage = await httpClient.PostAsync(endpoint, body);
                httpResponseMessage.EnsureSuccessStatusCode();
                return httpResponseMessage;
            }
        }

    }
}