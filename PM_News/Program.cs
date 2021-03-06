﻿using System;
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
           Console.WriteLine("Hello...");
           var setUp = new SetUp();
           Console.WriteLine("process starting");
           var ProgramPerams = new myProgramPerams();
           Console.WriteLine("Downloading and Importing Data");
           var import = new ImportData(ProgramPerams);
           var alterData = new AlterData(import.ImportedData, ProgramPerams.RemovalStringsColumn,
                                          ProgramPerams.RemovalStringsColumn, ProgramPerams.DateColumn,
                                          ProgramPerams.DupsColumn);
           Console.WriteLine("creating new csv file for SharePoint List");
           var outputData = new OutputData(alterData.DataImport, ProgramPerams);
           try
           {
               Directory.Delete(ProgramPerams.DestDir, true);
           }
           catch (Exception x)
           {
               Console.WriteLine(x.Message);
           }
        }
    }       
      
}
