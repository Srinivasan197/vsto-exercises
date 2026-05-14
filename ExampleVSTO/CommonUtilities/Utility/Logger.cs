using System;
using System.Diagnostics;
namespace ExampleVSTO.CommonUtilities.Utility
{
    public class Logger
    {
        private readonly string _className;

        public Logger(Type type)
        {
            _className = type.Name;
        }
        public void Info(string message)
        {
            WriteLog("INFO", message);
        }
        public void Warn(string message)
        {
            WriteLog("WARN", message);
        }
        public void Error(string message, Exception ex = null)
        {
            if (ex != null)
            {
                message += Environment.NewLine +
                           ex.Message +
                           Environment.NewLine +
                           ex.StackTrace;
            }

            WriteLog("ERROR", message);
        }
        public void Debug(string message)
        {
            WriteLog("DEBUG", message);
        }
        private void WriteLog(string level,
                              string message)
        {
            string logMessage =
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} " +
                $"[{level}] " +
                $"[{_className}] " +
                $"{message}";

            // Output Window
            System.Diagnostics.Debug.WriteLine(logMessage);

            // Console Output
            Console.WriteLine(logMessage);
        }
    }
}
