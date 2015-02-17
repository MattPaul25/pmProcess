using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PM_News
{
    public static class textUtils
    {
        public static string getTimeNum(int amHour, int pmHour, int timeDiff)
        {
            int dt = DateTime.Now.Hour;
            int retVal = dt + timeDiff >= pmHour ? pmHour : amHour;
            return retVal.ToString();
        }

        public static string getTodayFormatted(int dateAdd, string format)
        {
            //formats date to specified format
            DateTime dt = DateTime.Today.AddDays(dateAdd);
            string myFormattedDate = dt.ToString(format);
            return myFormattedDate;
        }
        public static int CountStrings(string yourString, string yourMarker)
        {
            int myCnt = 0;            
            try
            {                
                string newstring = yourString.Replace(yourMarker, "");
                myCnt = (yourString.Length - newstring.Length) / yourMarker.Length;                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return myCnt;
        }       

    }
}
