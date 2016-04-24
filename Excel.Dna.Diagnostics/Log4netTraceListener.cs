using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Dna.Diagnostics
{
   /// <summary>
    /// http://geertvanhorrik.com/2010/06/18/log4net-tracelistener/
    /// https://github.com/Excel-DNA/ExcelDna/wiki/Diagnostic-Logging
   /// </summary>
    public class Log4netTraceListener : System.Diagnostics.TraceListener
    {
        //private static readonly log4net.ILog _log = log4net.LogManager.GetLogger("System.Diagnostics.Redirection");

        //public Log4netTraceListener()
        //{
        //    var config = string.Format("{0}{1}", GetConfigurationPath(), ".config");
        //    log4net.Config.XmlConfigurator.Configure(new FileInfo(config));
        //    _log.Error("This is Ctor");
        //}

        ////public Log4netTraceListener(log4net.ILog log)
        ////{
        ////    _log = log;
        ////}

        //public override void Write(string message)
        //{
        //    if (_log != null)
        //    {
        //        _log.Debug(message);
        //    }
        //}

        //public override void WriteLine(string message)
        //{
        //    if (_log != null)
        //    {
        //        _log.Debug(message);
        //    }
        //}

        //private string GetConfigurationPath()
        //{

        //    string codeBase = Assembly.GetExecutingAssembly().CodeBase;
        //    UriBuilder uri = new UriBuilder(codeBase);
        //    return Uri.UnescapeDataString(uri.Path);
        //    //string path = Uri.UnescapeDataString(uri.Path);
        //    //return Path.GetDirectoryName(path);

        //}

        private readonly log4net.ILog _log;

        /// <summary>
        /// The trace listener will log system diagnostics messages. 
        /// This instance is picked up in the xll.config file and attached as a listener of the ExcelDna.Integration Source
        /// <source name="ExcelDna.Integration" switchValue="Verbose"> since this log4net is associated with
        /// the Excel.Dna.Diagnostics class it will also listen for anything that happens there as well. 
        /// We will be able to intercept logging that is called inside the Excel.Integration source as well as the main Diagnostics class. 
        /// 
        /// 
        /// </summary>
        public Log4netTraceListener()
        {
            
            _log = log4net.LogManager.GetLogger("System.Diagnostics.Redirection");
        }

        public Log4netTraceListener(log4net.ILog log)
        {
            _log = log;
        }

        public override void Write(string message)
        {
            if (_log != null)
            {
                _log.Warn(message);
            }
        }

        public override void WriteLine(string message)
        {
            if (_log != null)
            {
                _log.Warn(message);
            }
        }
    }
}


