using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using AddinX.Ribbon.ExcelDna;
using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using Microsoft.Office.Interop.Excel;

namespace Excel.Dna.Diagnostics
{
    [ComVisible(true)]
    public class ExcelAddin : ExcelRibbon, IExcelAddIn
    {
        private IRibbonUI ribbon = null;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public ExcelAddin() {
            //var config = Path.Combine(GetConfigurationPath(), "Excel.Dna.Diagnostics-AddIn.xll.config");
            var config = string.Format("{0}{1}",GetConfigurationPath(), ".config");
            log4net.Config.XmlConfigurator.Configure(new FileInfo(config));
            log.Error("This is Ctor");

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
            // The Excel Application object
            AddinContext.ExcelApp = new Application(null, ExcelDnaUtil.Application);
            log.Error("This is Auto Open");
            Console.WriteLine("Hello");
        }

        private string GetConfigurationPath()
        {
            
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            return Uri.UnescapeDataString(uri.Path);
            //string path = Uri.UnescapeDataString(uri.Path);
            //return Path.GetDirectoryName(path);
        
        }

        /// <summary>
        /// Test Function
        /// Go to excel type in =CoolFunction("Name")
        /// </summary>
        [ExcelFunction(Description = "Cool Name Function")]
        public static string CoolFunction(string name)
        {
            return string.Format("Hello {0} You are Cool", name);
        }
    
    }
}



/*
 * https://groups.google.com/forum/#!searchin/exceldna/Get$20Ribbon$20to$20show/exceldna/FzkZz9giA4M/mTNkqLSPP1MJ
 * 
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
 * 
 * 
 * Custom Tabs?
 * https://xldennis.wordpress.com/2009/03/11/sharing-custom-tabs-in-the-ribbon-ui/
 */