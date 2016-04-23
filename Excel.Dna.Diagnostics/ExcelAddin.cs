using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace Excel.Dna.Diagnostics
{
    [ComVisible(true)]
    public class ExcelAddin : ExcelRibbon, IExcelAddIn
    {
        private IRibbonUI ribbon = null;

        public ExcelAddin() {
            Console.WriteLine("Hello");
        
        }

        public void OnLoad(IRibbonUI ribbon)
        {
            this.ribbon = ribbon;
        }


        public void RunTest(IRibbonControl control)
        {
            Console.WriteLine("Hello");
        }

        public void AutoClose()
        {
             
        }

        public void AutoOpen()
        {
            Console.WriteLine("Hello");
             
        }

        /*
         * "C:\Users\Chris W\Documents\GitHub\ExcelDnaTest\Excel.Dna.Diagnostics\bin\Debug\Excel.Dna.Diagnostics-AddIn.xll"
         * 
         * C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE
         * 
         xcopy "$(SolutionDir)\packages\ExcelDna.AddIn.0.33.9\tools\ExcelDna.xll" "$(TargetDir)Excel.Dna.Diagnostics-AddIn.xll*" /C /Y
        xcopy "$(TargetDir)Excel.Dna.Diagnostics-AddIn.dna*" "$(TargetDir)Excel.Dna.Diagnostics-AddIn64.dna*" /C /Y
        xcopy "$(SolutionDir)\packages\ExcelDna.AddIn.0.33.9\tools\ExcelDna64.xll" "$(TargetDir)Excel.Dna.Diagnostics-AddIn64.xll*" /C /Y
        "$(SolutionDir)\packages\ExcelDna.AddIn.0.33.9\tools\ExcelDnaPack.exe" "$(TargetDir)Excel.Dna.Diagnostics-AddIn.dna" /Y
        "$(SolutionDir)\packages\ExcelDna.AddIn.0.33.9\tools\ExcelDnaPack.exe" "$(TargetDir)Excel.Dna.Diagnostics-AddIn64.dna" /Y
         
         * 
         * https://groups.google.com/forum/#!topic/exceldna/IhqXaK-avWg
         * 
         * xcopy "$(SolutionDir)packages\ExcelDna.AddIn.0.33.9\tools\ExcelDna.xll" "$(TargetDir)Excel.Dna.Diagnostics-AddIn.xll*" /C /Y
            "$(SolutionDir)packages\ExcelDna.AddIn.0.33.9\tools\ExcelDnaPack.exe" "$(TargetDir)Excel.Dna.Diagnostics-AddIn.dna" /Y

         * 
         * https://msdn.microsoft.com/en-us/library/aa730920%28v=office.12%29.aspx
         */
    }
}
