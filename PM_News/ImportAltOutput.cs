using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO.Compression;
using System.IO;


namespace PM_News
{
    class ImportData
    {
        private myProgramPerams perams;
        public DataTable ImportedData { get; set; }

        public ImportData(myProgramPerams p)
        {
            perams = p;
            Console.WriteLine("downloading file via vba urlmon c++ com utility");
            downloadFile();
            Console.WriteLine("pulling downloaded csv into memory");
            GetDataTableFromCsv();

        }

        private void downloadFile()
        {
            if (File.Exists(perams.DestZipDir)) { System.IO.File.Delete(perams.DestZipDir); }
            var xlMac1 = new RunExcelMacro(perams.WbLoc, perams.MacrDownloadName, 2,
                                                     perams.MacrDownloadArg1, perams.MacrDownloadArg2);
            unzipFile();
            var xlMac2 = new RunExcelMacro(perams.WbLoc, perams.MacrCsvParserName, 3, perams.MacrCsvParserArg1,
                                                perams.MacrCsvParserArg2, perams.MacrCsvParserArg3);

        }

        private void unzipFile()
        {
            //something changed
            //if file is alread there will delete
            if (System.IO.Directory.Exists(perams.DestDir)) { System.IO.Directory.Delete(perams.DestDir, true); }
            //extract zipfile and then delete it
            ZipFile.ExtractToDirectory(perams.DestZipDir, perams.DestDir);
            System.IO.File.Delete(perams.DestZipDir);

        }
        private void GetDataTableFromCsv()
        {
            /*grabs data from txt file in unzipped folder - uses perams class to get locations
             * out puts a data table with the data in it */
            StreamReader sr = new StreamReader(perams.FullTxtPath);
            string header = sr.ReadLine();
            var headers = header.Split('|');
            DataTable dataTable = new DataTable();

            foreach (var h in headers)
            {
                dataTable.Columns.Add(h.Trim());
            }

            int myColumnInterval = headers.Count();
            string newLine = sr.ReadLine();
            while (newLine != null)
            {
                var fields = newLine.Split('|'); // csv delimiter              
                string[] adjustedFields = new string[myColumnInterval];
                //want to bring in exactly the column interval amount of columns
                for (int i = 0; i < myColumnInterval; i++)
                {
                    adjustedFields[i] = fields[i].Trim();
                }
                dataTable.Rows.Add(adjustedFields);
                newLine = sr.ReadLine();
            }
            ImportedData = dataTable;
        }
      
    }
    class AlterData
    {
             //manipulates imported data
        public DataTable DataImport { get; set; }

        public AlterData(DataTable dt, string criteriaColumn, string WordsToRemove,
            string datesColumn, string dupsColumn)
        {
            Console.WriteLine("altering data");
            DataImport = dt;
            string dateToday = textUtils.getTodayFormatted(0, "M/d/yyyy");
            string dateYesterday = textUtils.getTodayFormatted(-1, "M/d/yyyy");

            /* because i want to keep both today and yesterday i use concatenate the dates which will check if
             one item exist within item*/
            string dateBoth = dateYesterday + dateToday;
            Console.WriteLine("removing items with key words we want to remove: " + WordsToRemove);
            removeItemsFromDataTable(criteriaColumn, WordsToRemove, false);
            Console.WriteLine("removing dates that are neither today nor yesterday...");
            removeItemsFromDataTable(datesColumn, dateBoth, true, true);
            Console.WriteLine("deleting duplicates");
            deleteDuplicatesFromDataTable(dupsColumn);
            Console.WriteLine("sorting...");
            SortDataTable("TF Label1 (Rule Name)");
            SortDataTable("Priority");
            Console.WriteLine("adding sharepoint columns");
            addColumns();
        }
        private void deleteDuplicatesFromDataTable(string duplicatesColumn) 
        {
            //removes duplicates from a sorted list
            SortDataTable(duplicatesColumn);
            for (int i = 0; i + 1 < DataImport.Rows.Count; i++)
             {
                 DataRow r = DataImport.Rows[i];                 
                 string currentString = r[duplicatesColumn].ToString().ToLower();
                 DataRow r2 = DataImport.Rows[i + 1];
                 string nextString = r2[duplicatesColumn].ToString().ToLower();
                 if (currentString == nextString)
                 {
                     DataImport.Rows[i].Delete();
                     i--; //move index back one after deletion
                 }
             }
        }
        private void SortDataTable(string sortingColumn)
        {
            DataView dvDataImport = DataImport.DefaultView;
            dvDataImport.Sort = sortingColumn;
            DataImport = dvDataImport.ToTable();
        }
        private void removeItemsFromDataTable(string columnName, string wordCriteria, 
                                               bool isKeepable = true, bool isReversableLookUp = false)
        {
            /*method removes item from data table that match or dont match a criteria
            * based onthe keep variable being true or false - true dictates that 
            * isReversalableLookup reverses check by seeing if item in datatable is in wordCriteria*/
            string[] criteriaStrings = wordCriteria.Split('|');
            for (int i = 0; i < DataImport.Rows.Count; i++)
			{
                foreach (string criteriaString in criteriaStrings)
                {
                    //loops through removal string array to see if cell (current string) contains any items
                    DataRow r = DataImport.Rows[i];
                    string currentString = r[columnName].ToString();
                    bool isStringThere;
                    if (isReversableLookUp)
                    {
                        isStringThere = textUtils.CountStrings(criteriaString, currentString) > 0;
                    }
                    else
                    {
                        isStringThere = textUtils.CountStrings(currentString, criteriaString) > 0;
                    }
                    bool isDeletable = ((isKeepable & !isStringThere) == true) || ((!isKeepable & isStringThere) == true);
                    if (isDeletable) 
                    {
                        DataImport.Rows[i].Delete();  //if item is deletable then delete the entire row
                        i = i == 0? 0: i - 1; //subtract 1 from i unless its 0
                    }
                }
            }            
        }

        private void addColumns()
        {
            //adds columns to match sharepoint
            DataImport.Columns.Add("ID", typeof(int)).SetOrdinal(0);
            DataImport.Columns.Add("Entity", typeof(string)).SetOrdinal(1);
            DataImport.Columns.Add("Researcher", typeof(string)).SetOrdinal(2);
            DataImport.Columns.Add("Story Status", typeof(string)).SetOrdinal(3);
            DataImport.Columns.Add("Transaction Type", typeof(string)).SetOrdinal(20);
            DataImport.Columns.Add("Content Type", typeof(string)).SetOrdinal(21);
            DataImport.Columns.Add("Attachments", typeof(string)).SetOrdinal(22);
        }
    }
    class OutputData
    {
        private DataTable DataExport { get; set; }
        private myProgramPerams perams { get; set; }

        public OutputData(DataTable dt, myProgramPerams p)
        {
            perams = p;
            DataExport = dt;            
            Console.WriteLine("creating the csv file");
            createCsvFile();
            Console.WriteLine("uploading the file into sharepoint via access");
            uploadFileToSharePoint();
        }

        private void createCsvFile()
        {
            //pushes data table to csv file
            if (File.Exists(perams.ExportLocation)) { File.Delete(perams.ExportLocation); }
            var csv = new StringBuilder();
            string headerString = "";
            for (int i = 0; i < DataExport.Columns.Count; i++)
            {
                headerString = headerString + DataExport.Columns[i] + "|";
            }
            headerString = headerString.Substring(0, headerString.Length - 1) + "\n"; //removes last pipe
            csv.Append(headerString);

            foreach (DataRow row in DataExport.Rows)
            {
                for (int j = 0; j < DataExport.Columns.Count; j++)
                {
                    csv.Append(row[j].ToString());
                    csv.Append(j == DataExport.Columns.Count - 1 ? "\n" : "|");
                }
            }
            File.AppendAllText(perams.ExportLocation, csv.ToString());
        }

        private void uploadFileToSharePoint()
        {
            RunExcelMacro xlMac3 = new RunExcelMacro(perams.WbLoc, perams.MacrImportDataName);
            RunAccessMacro acMac1 = new RunAccessMacro(perams.AccessLocation, perams.MacrAccessImport);

        }

    }        
      
}
