using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SampleAADv2Bot.Extensions
{
    public static class ExpressionCheckingHelper
    {
        public static bool IsNaturalNumber(this string argument)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(argument,
                @"^[1-9][0-9]*$"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsEmailAddress(this string argument)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(argument,
                @"^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool IsEmailAddressList(this string argument)
        {
            string[] separatedEmailAddress;
            //remove space
            string spaceRemovedArgument = argument.Replace(" ", "").Replace("　", "");
            separatedEmailAddress = spaceRemovedArgument.Split(',');
            foreach (var i in separatedEmailAddress)
                Console.WriteLine(i);
            foreach (string s in separatedEmailAddress)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(s,
                                @"^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"))
                    return false;

            }
            return true;
        }

        public static bool IsDatatime(this string argument)
        {
            if (System.Text.RegularExpressions.Regex.IsMatch(argument,
                                    //@"^[0-9]{4}/[0-9]{2}/[0-9]{2}$"))
                                    @"^[0-9]{4}-[0-9]{2}-[0-9]{2}$"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}