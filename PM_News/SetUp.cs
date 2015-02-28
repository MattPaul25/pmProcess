using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace PM_News
{
    class SetUp
    {
        private string myFile;
        private bool fileExists;

        public SetUp()
        {
            myFile = "MyPeramData.txt";
            fileExists = File.Exists(myFile);
            
            if (fileExists)
            {
                Console.WriteLine("I see that the perameter file already exists \n Do you want to continue with setup? Press any key for yes.");
                bool buttonPressed = Task.Factory.StartNew(() => Console.ReadKey()).Wait(TimeSpan.FromSeconds(5.0));
                if (buttonPressed) { runQuestionaire(); }
            }
            else if (!fileExists)
            {
                runQuestionaire();
            }
        }

        private void runQuestionaire()
        {
            Console.WriteLine("Welcome to the set up menu..");
            Console.WriteLine("Each one of your responses is recorded in a log to be used later in execution");
            Console.WriteLine("Press enter to continue!");
            Console.ReadLine();
            Console.WriteLine("fyi, if you want to keep the current value (will display) just hit enter for that to be restored");
            Console.ReadLine();

            string[] myTextValues = new string[Questions.Length];
            if (fileExists)
            {
                //gets current values in text file and dumps in string array
                for (int i = 0; i < Questions.Length; i++)
                {
                    myTextValues[i] = TextUtils.getParams(VariableNames[i], myFile);
                }
            }
            StreamWriter sw = new StreamWriter(myFile);
            for (int i = 0; i < Questions.Length; i++)
            {
                Console.WriteLine(Questions[i]);
                Console.WriteLine("Previous Value: " + myTextValues[i]);                
                string peram = Console.ReadLine();
                if (myTextValues[i] != "" && peram == "")
                {
                    //situation to use value already stored in previous version of text file
                    Console.WriteLine("using previous Values");
                    sw.WriteLine(VariableNames[i] + "|" + myTextValues[i]);
                }
                else
                {
                    //situation to change value in text file
                    sw.WriteLine(VariableNames[i] + "|" + peram);
                }
            }
            sw.Dispose();
        }



       public string[] Questions = 
       {
            "please write the url in which you are downloading the zipfiles - URI prior to the actual zip name..",
            "Please write in your starting directory (uses forward slashes)\n this is the directory you storing excel work and access DB",
            "Please write in the name of the macro file thats in the aforementioned directory - please include .xlsm extension",
            "What is the name of the macro that downloads the file... i.e. DownLoadRoutine",
            "What is the name of the macro that parses the csv",
            "What is the name of the macro that preps the data for the access import -- i.e. PrepareDataForUpload",
            "What is the column you would like to filter out key words called?..",
            "What are the keywords in the description you would like to filter out..",
            "What is the name of the Access Database? Please include the extension.",
            "What is the name of the sub routine in access?"

       };

       public string[] VariableNames = {  "BaseUrl",
                                          "BaseDir",
                                          "ExcelFile", 
                                          "MacrDownloadName",
                                          "MacrCsvParserName",
                                          "MacrImportDataName",
                                          "RemovalStringsColumn",
                                          "RemovalStrings",                                          
                                          "AccessName",
                                          "MacrAccessImport"

                                          };
    }
}
