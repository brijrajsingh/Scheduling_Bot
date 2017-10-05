using System.Threading.Tasks;
using Microsoft.Graph;

namespace SampleAADv2Bot.Services
{
    public interface IMeetingService
    {
        Task<MeetingTimeSuggestionsResult> GetMeetingsTimeSuggestions(string accessToken, UserFindMeetingTimesRequestBody userFindMeetingTimesRequestBody);
        Task<Event> ScheduleMeeting(string accessToken, Event meeting);
    }
}
