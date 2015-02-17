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
           var myProgramPerams = new myProgramPerams(); //todo: get file perams from folder to pass object to import
           Console.WriteLine("Downloading and Importing Data");
           var import = new ImportData(myProgramPerams);
           var alterData = new AlterData(import.ImportedData, myProgramPerams.RemovalStringsColumn,
                                          myProgramPerams.RemovalStringsColumn, myProgramPerams.DateColumn, 
                                          myProgramPerams.DupsColumn);
           Console.WriteLine("creating new csv file for SharePoint List");
           var outputData = new OutputData(alterData.DataImport, myProgramPerams.BaseDir + "ExportedData.csv");
         
        }
    }       
   
    class RunAccessMacro
    {
        private string loc { get; set; }
        public RunAccessMacro(string Loc)
        {
            loc = Loc;
        }
    }
    

   
}
