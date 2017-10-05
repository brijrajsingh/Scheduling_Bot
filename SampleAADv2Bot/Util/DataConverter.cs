using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Graph;

namespace SampleAADv2Bot.Util
{
    public class DataConverter
    {
        public static UserFindMeetingTimesRequestBody GetMeetingRequest()
        {
            var request = new UserFindMeetingTimesRequestBody();
            return request;
        }
    }
}