using AddinX.Ribbon.Contract;
using AddinX.Ribbon.Contract.Command;
using AddinX.Ribbon.ExcelDna;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Dna.Diagnostics
{
    //http://www.addinx.org/addinx/example_wcf.html
    [ComVisible(true)]
    public class ExcelAddinFluent : RibbonFluent, IExcelAddIn
    {
        
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public ExcelAddinFluent() {
            var config = string.Format("{0}{1}",GetConfigurationPath(), ".config");
            log4net.Config.XmlConfigurator.Configure(new FileInfo(config));
            log.Error("This is Ctor");

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

         public override void OnOpening()
        {
            AddinContext.ExcelApp.SheetActivate += (e) => Ribbon.Invalidate();
            AddinContext.ExcelApp.SheetChange += (a, e) => Ribbon.Invalidate();
        }


        protected override void CreateFluentRibbon(IRibbonBuilder build)
        {
            build.CustomUi.Ribbon.Tabs(c =>
            {
                c.AddTab("Sample").SetId("SampleTab")
                    .Groups(g =>
                    {
                        g.AddGroup("Reporting").SetId("ReportingGroup")
                            .Items(d =>
                            {
                                d.AddButton("Button 1")
                                    .SetId("button1")
                                    .LargeSize()
                                    .ImageMso("Repeat");

                                d.AddBox().SetId("ReportingBox")
                                    .HorizontalDisplay()
                                    .AddItems(i =>
                                    {
                                        i.AddButton("Button 2").SetId("button2")
                                            .NormalSize().NoImage().ShowLabel()
                                            .Screentip("Button 2")
                                            .Supertip("Displays a message box");

                                        i.AddButton("Button 3")
                                           .SetId("button3")
                                           .NormalSize()
                                           .ImageMso("Bold");
                                    });
                            });
                    });
            });
        }

        public override void OnClosing()
        {
            
        }

        protected override void CreateRibbonCommand(IRibbonCommands cmds)
        {
            cmds.AddButtonCommand("button1")
            .Action(() => ExcelAddin.CoolFunction("Add one more sheet"));
            // Reporting Group
            //cmds.AddButtonCommand("button3")
            //    .IsEnabled(() => AddinContext.ExcelApp.Worksheets.Count > 2);

            
            //cmds.AddBoxCommand("ReportingBox")
            //    .IsVisible(() => AddinContext.ExcelApp.Worksheets.Count > 1);
        }
    }

    public static class AddinContext
    {
        public static Application ExcelApp { get; set; }
    }
}
