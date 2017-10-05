using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace SampleAADv2Bot.Services
{
    public interface IRoomService
    {
        List<Room> GetRooms();
        void AddRooms(UserFindMeetingTimesRequestBody request, List<Room> rooms);

    }
}
