using System;
using System.Diagnostics;
using System.IO;

namespace ExampleVSTO.CommonUtilities.Utility
{
    public class Logger
    {
        private readonly string _className;
        private readonly string _logFolderPath;
        private readonly string _logFilePath;
        public Logger(Type type)
        {
            _className = type.Name;

            _logFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

            // Create folder if not exists
            if (!Directory.Exists(_logFolderPath))
            {
                Directory.CreateDirectory(_logFolderPath);
            }
            // Log file name based on current date
            _logFilePath = Path.Combine(_logFolderPath, $"Log_{DateTime.Now:yyyyMMdd}.txt");
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
        private void WriteLog(string level, string message)
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

            // Write log to file
            try
            {
                File.AppendAllText(
                    _logFilePath,
                    logMessage + Environment.NewLine
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to write log file: " + ex.Message);
            }
        }
    }
}