using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PM_News
{
    public static class TextUtils
    {
        public static string getParams(string peram, string fileName)
        {
            //returns the paramater value from stored values in text file
            StreamReader sr = new StreamReader(fileName);
            string currentString = sr.ReadLine();
            string returnParam = "";
            if (currentString != "")
            {
                while (currentString != null)
                {
                    int myPoint = Search(currentString, "|", 1) - 1;
                    myPoint = myPoint == -1? 0: myPoint;
                    if (currentString.Substring(0, myPoint) == peram)
                    {
                        returnParam = currentString.Substring(myPoint + 1);
                        break;
                    }
                    currentString = sr.ReadLine();
                }
            }
            sr.Dispose();
            return returnParam;
        }
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
        public static int Search(string yourString, string yourMarker, int yourInst = 1, bool yourCapPref = true)
        {
            //returns the placement of a string in another string
            int num = 1;
            int ginst = 1;
            int mlen = yourMarker.Length;
            int slen = yourString.Length;
            string tString = "";

            try
            {
                if (yourCapPref == false)
                {
                    yourString = yourString.ToLower();
                    yourMarker = yourMarker.ToLower();
                }

                while (num < slen)
                {
                    tString = yourString.Substring(num, mlen);

                    if (tString == yourMarker && ginst == yourInst)
                    {
                        return num + 1;
                    }
                    else if (tString == yourMarker && yourInst != ginst)
                    {
                        ginst += 1;
                        num += 1;
                    }
                    else
                    {
                        num += 1;
                    }
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

    }
}
