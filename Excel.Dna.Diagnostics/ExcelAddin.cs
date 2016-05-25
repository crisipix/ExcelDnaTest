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

            //var config = string.Format("{0}{1}",GetConfigurationPath(), ".config");
            //log4net.Config.XmlConfigurator.Configure(new FileInfo(config));
            log4net.Config.XmlConfigurator.Configure();
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
            MakeGraph();
            FillColors();
            FillRange();
        }

        /*
         http://stackoverflow.com/questions/11223641/how-do-i-create-a-new-worksheet-and-populate-it-with-rows-of-data-using-excel-dn
         */
        private void MakeGraph()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            //var app = new Application();
            app.Visible = true;
            var workbook = app.Workbooks.Add(1);

            Sheets excelSheets = workbook.Worksheets;
            string currentSheet = "Sheet1";
            Worksheet worksheet1 = (Worksheet)excelSheets.get_Item(currentSheet);


            worksheet1.Cells[1, 1] = "";
            worksheet1.Cells[1, 2] = "Year 1";
            worksheet1.Cells[1, 3] = "Year 2";
            worksheet1.Cells[1, 4] = "Year 3";
            worksheet1.Cells[1, 5] = "Year 4";
            worksheet1.Cells[1, 6] = "Year 5";

            worksheet1.Cells[2, 1] = "Company A";
            worksheet1.Cells[2, 2] = "10";
            worksheet1.Cells[2, 3] = "50";
            worksheet1.Cells[2, 4] = "70";
            worksheet1.Cells[2, 5] = "70";
            worksheet1.Cells[2, 6] = "70";

            worksheet1.Cells[3, 1] = "Company B";
            worksheet1.Cells[3, 2] = "30";
            worksheet1.Cells[3, 3] = "70";
            worksheet1.Cells[3, 4] = "80";
            worksheet1.Cells[3, 5] = "80";
            worksheet1.Cells[3, 6] = "80";

            worksheet1.Cells[4, 1] = "Company C";
            worksheet1.Cells[4, 2] = "55";
            worksheet1.Cells[4, 3] = "65";
            worksheet1.Cells[4, 4] = "75";
            worksheet1.Cells[4, 5] = "75";
            worksheet1.Cells[4, 6] = "75";

            ChartObjects xlCharts = (ChartObjects)worksheet1.ChartObjects(Type.Missing);
            ChartObject myChart = (ChartObject)xlCharts.Add(60, 10, 300, 300);
            Range chartRange = worksheet1.get_Range("A1", "F4");

            Chart chartPage = myChart.Chart;
            chartPage.SetSourceData(chartRange, System.Reflection.Missing.Value);
            chartPage.ChartType = XlChartType.xlLine;


        }

        private void FillColors()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            //var app = new Application();
            app.Visible = true;
            var workbook = app.ActiveWorkbook;

            Sheets excelSheets = workbook.Worksheets;
            string currentSheet = "Sheet1";
            Worksheet worksheet1 = (Worksheet)excelSheets.get_Item(currentSheet);
            worksheet1.Cells[6, 1] = "April 1st";
            worksheet1.Cells[6, 2] = "April 2nd";
            worksheet1.Cells[6, 3] = "April 3rd";
            worksheet1.Cells[6, 4] = "April 4th";

            // fill in the starting and ending range programmatically this is just an example. 
            string startRange = "A6";
            string endRange = "A6";
            Range currentRange = worksheet1.get_Range(startRange, endRange);

            var text = currentRange.Text.ToString();
            int length = text.Length;
            int index = 0;
            if (text.Contains("st"))
            {
                index = text.IndexOf("st");
            }
            //The other checks for "nd", "rd", "th" obviously check to see a # precedes these.

            if (index > 0)
            {
                currentRange.get_Characters(index + 1, 2).Font.Superscript = true;
            }           
        }

        //http://stackoverflow.com/questions/2692979/how-to-speed-up-dumping-a-datatable-into-an-excel-worksheet
        private void FillRange()
        {
            Application app = (Application)ExcelDnaUtil.Application;
            //var app = new Application();
            app.Visible = true;
            var workbook = app.ActiveWorkbook;

            Sheets excelSheets = workbook.Worksheets;
            string currentSheet = "Sheet1";
            Worksheet worksheet1 = (Worksheet)excelSheets.get_Item(currentSheet);
            Range range = worksheet1.get_Range("A10", Missing.Value);
            range = range.get_Resize(5, 5);

            
                //Create an array.
                double[,] saRet = new double[5, 5];
                
                //Fill the array.
                for (long iRow = 0; iRow < 5; iRow++)
                {
                    for (long iCol = 0; iCol < 5; iCol++)
                    {
                        //Put a counter in the cell.
                        saRet[iRow, iCol] = iRow * iCol;
                    }
                }

                //Set the range value to the array.
                range.set_Value(Missing.Value, saRet);
            
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