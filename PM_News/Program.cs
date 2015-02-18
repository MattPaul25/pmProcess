using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;


namespace PM_News
{
    class Program
    {
        static void Main(string[] args)
        {   
           Console.WriteLine("Hello, Process Starting");
           var ProgramPerams = new myProgramPerams(); //todo: get file perams from folder to pass object to import
           Console.WriteLine("Downloading and Importing Data");
           var import = new ImportData(ProgramPerams);
           var alterData = new AlterData(import.ImportedData, ProgramPerams.RemovalStringsColumn,
                                          ProgramPerams.RemovalStringsColumn, ProgramPerams.DateColumn,
                                          ProgramPerams.DupsColumn);
           Console.WriteLine("creating new csv file for SharePoint List");
           var outputData = new OutputData(alterData.DataImport, ProgramPerams);
           Directory.Delete(ProgramPerams.DestDir, true);
        }
    }       
      
}
