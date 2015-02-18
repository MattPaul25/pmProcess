using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace PM_News
{
    class myProgramPerams
    {
        /*class that outputs the perameters for the program to run - we pass an instantiated object of this to other
        classes so they have access to all these locations without passing a million strings */
        public myProgramPerams()
        {
            Console.WriteLine("Getting Params");
            ///constructor assigns values to properties above
            ///todo: will get properties from a text file that is set by setup class
            BaseUrl = "http://ats.factiva.net/fhpnas/ats/Reports/RnC/PM/";
            TimeNum = textUtils.getTimeNum(15, 23, 3);
            ZipDate = textUtils.getTodayFormatted(-1, "yyyyMMdd");
            Name = "rnc_pm_" + ZipDate + "_" + TimeNum;
            Ext = ".zip";
            MyURL = BaseUrl + Name + Ext;
            BaseDir = "V:/Research/Shared/BarcelonaProjects/PM_News/";
            DestZipDir = BaseDir + Name + Ext;
            DestDir = BaseDir + Name;
            WbLoc = BaseDir + "SupportMacro.xlsm";
            FullCsvPath = DestDir + "/mnt/auto/dmpwork/" + ZipDate + "_" + TimeNum + "/" + Name + ".csv";
            FullTxtPath = DestDir + "/mnt/auto/dmpwork/" + ZipDate + "_" + TimeNum + "/" + Name + ".txt";

            RemovalStrings = "rating|buy back|buyback|customers|cars|products|debt|reserves|from hold|brokers|broker|argued|amended|recommends";
            RemovalStringsColumn = "TF Label2 (Extraction Text)";

            MacrDownloadName = "DownLoadRoutine";
            MacrDownloadArg1 = MyURL;
            MacrDownloadArg2 = DestZipDir;

            MacrCsvParserName = "RunCSVParser";
            MacrCsvParserArg1 = FullCsvPath;
            MacrCsvParserArg2 = Name;
            MacrCsvParserArg3 = ".csv";

            ExportLocation = BaseDir + "ExportedData.csv";
            MacrImportDataName = "PrepareDataForUpload";
            AccessLocation = BaseDir + "ThingFinderUpload.accdb";
            MacrAccessImport = "ImportMacro";

            DupsColumn = "DistDoc Accession Number";
            DateColumn = "DistDoc Publication Date";

        }
        public string BaseUrl { get; protected set; }
        public string BaseDir { get; protected set; }
        public string Name { get; protected set; }
        public string ZipDate { get; protected set; }
        public string Ext { get; protected set; }
        public string MyURL { get; protected set; }
        public string DestZipDir { get; protected set; }
        public string DestDir { get; protected set; }
        public string TimeNum { get; protected set; }

        public string DateColumn { get; protected set; }
        public string DupsColumn { get; protected set; }

        public string MacrDownloadName { get; protected set; }
        public string MacrDownloadArg1 { get; protected set; }
        public string MacrDownloadArg2 { get; protected set; }

        public string MacrCsvParserName { get; protected set; }
        public string MacrCsvParserArg1 { get; protected set; }
        public string MacrCsvParserArg2 { get; protected set; }
        public string MacrCsvParserArg3 { get; protected set; }

        public string WbLoc { get; protected set; }
        public string FullCsvPath { get; protected set; }
        public string FullTxtPath { get; protected set; }
        //public string[] macrArgs {get; }

        public string RemovalStrings { get; protected set; }
        public string RemovalStringsColumn { get; protected set; }

        public string ExportLocation { get; protected set; }
        public string MacrImportDataName { get; protected set; }

        public string AccessLocation { get; protected set; }
        public string MacrAccessImport { get; protected set; }
    }
}
