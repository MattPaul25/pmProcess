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
        public SetUp()
        {
            if(File.Exists("myPeramData.txt"))
            {
                Console.WriteLine("I see that the perameter file already exists \n Do you want to continue with setup? Press any key for yes." );
                bool buttonPressed = Task.Factory.StartNew(() => Console.ReadKey()).Wait(TimeSpan.FromSeconds(5.0));
                if (!buttonPressed)
                {
                    return;
                }
            }
           
           
                Console.WriteLine("Welcome to the set up menu..");
                Console.WriteLine("Each one of your responses is recorded in a log to be used later in execution");
                Console.WriteLine("Press any button to continue!");
                Console.ReadLine();
                StreamWriter sw = new StreamWriter("MyPeramData.txt");
                for (int i = 0; i < Questions.Length; i++)
                {
                    Console.WriteLine(Questions[i]);
                    string peram = Console.ReadLine();
                    sw.WriteLine(VariableNames[i] + "|" + peram);
                }
                sw.Close();
            

        }



       public string[] Questions = 
       {
            "Please write in your starting directory (uses forward slashes)\n this is the directory you storing excel work and access DB",
            "Please write in the name of the macro file thats in the aforementioned directory - please include .xlsm extension",
            "What is the name of the macro that downloads the file... i.e. DownLoadRoutine",
            "What is the name of the macro that parses the csv",
            "What is the name of the macro that preps the data for the access import -- i.e. PrepareDataForUpload",
            "What is the name of the Access Database? Please include the extension.",
            "What is the name of the sub routine in access?"

       };

       public string[] VariableNames = {  "BaseDir",
                                          "ExcelFile", 
                                          "MacrDownloadName",
                                          "MacrCsvParserName",
                                          "MacrImportDataName",
                                          "AccessName",
                                          "MacrAccessImport"
                                          };
    }
}
