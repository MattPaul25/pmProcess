using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;
using System.Net;
using System.IO;

namespace PM_News
{
    class RunAccessMacro
    {
        private string macroName { get; set; }
        private string fileName { get; set; }
        private int attemptNum;

        public RunAccessMacro(string FileName, string MacroName)
        {
            fileName = FileName;
            macroName = MacroName;
            attemptNum = 0;
            RunMacro();
        }
        private void RunMacro()
        {
            Console.WriteLine("uploading...");
            Access.Application oAccess = new Access.Application();
            oAccess.Visible = false;
            oAccess.OpenCurrentDatabase(fileName, false);
           
            try
            {
                oAccess.Run("ImportMacro");
                oAccess.Quit();
            }
            catch (Exception x)
            {
                oAccess.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;
                if (attemptNum < 3)
                {
                    attemptNum++;
                    Console.WriteLine("attempting time " + attemptNum);
                    RunMacro();
                }
                Console.WriteLine("something went wrong with the access macro: \n" + x.Message);
            }
            finally
            {                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oAccess);
                oAccess = null;
            }
        } 
    }

    class RunExcelMacro
    {
        private string macroName { get; set; }
        private string fileName { get; set; }
        private int argCount { get; set; }
        private string arg1 { get; set; }
        private string arg2 { get; set; }
        private string arg3 { get; set; }

        public RunExcelMacro(string FileName, string MacroName, int ArgCount = 0, 
                                string Arg1 = "", string Arg2 = "", string Arg3 = "")
        {                                                       
            fileName = FileName;                                
            macroName = MacroName;
            arg1 = Arg1;
            arg2 = Arg2;
            arg3 = Arg3;
            argCount =  ArgCount > 3? 3: ArgCount;
            RunMacro();
        }

        private void RunMacro()
        {           
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWB;
            xlApp.Visible = false;
            xlWB = xlApp.Workbooks.Open(fileName);
            try
            {
                switch (argCount)
                {
                    case 3:
                        xlApp.Run(macroName, arg1, arg2, arg3);
                        break;
                    case 2:
                        xlApp.Run(macroName, arg1, arg2);
                        break;
                    case 1:
                        xlApp.Run(macroName, arg1);
                        break;
                    default:
                        xlApp.Run(macroName);
                        break;
                }               
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                xlWB.Close(false);
                xlApp.Quit();
                releaseObject(xlApp);
                releaseObject(xlWB);
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }


    }
}
