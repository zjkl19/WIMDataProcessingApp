// Utils/Logger.cs
using System;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace WIMDataProcessingApp.Utils
{
    public static class Logger
    {
        private static readonly object _initLock = new object();


        private static volatile bool _initialized;
        private static string _logDir;
        private static string _logFile;
        private static StreamWriter _writer;

        // UI 订阅这个事件即可实时接收日志
        public static event Action<string> OnMessage;

        public static string LogDirectory => _logDir;
        public static string LogFilePath => _logFile;

        public static void Initialize(string baseDir = null)
        {
            if (_initialized) return;
            lock (_initLock)
            {
                if (_initialized) return;

                var root = string.IsNullOrWhiteSpace(baseDir)
                    ? AppDomain.CurrentDomain.BaseDirectory
                    : baseDir;

                _logDir = Path.Combine(root, "Logs");
                Directory.CreateDirectory(_logDir);

                var date = DateTime.Now.ToString("yyyyMMdd");
                _logFile = Path.Combine(_logDir, $"hs_proc_{date}.log");

                // 附加写入（每日同一文件），UTF-8 无 BOM
                _writer = new StreamWriter(new FileStream(_logFile, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    AutoFlush = true,
                    NewLine = Environment.NewLine
                };

                _initialized = true;
                Info($"Logger initialized. LogFile = {_logFile}");
            }
        }

        public static void Info(string msg) => Write("INFO", msg);
        public static void Warn(string msg) => Write("WARN", msg);
        public static void Error(string msg) => Write("ERROR", msg);

        public static void Error(Exception ex, string hint = null)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(hint)) sb.AppendLine(hint);
            sb.AppendLine(ex.GetType().FullName);
            sb.AppendLine(ex.Message);
            sb.AppendLine(ex.StackTrace);
            if (ex.InnerException != null)
            {
                sb.AppendLine("Inner:");
                sb.AppendLine(ex.InnerException.GetType().FullName);
                sb.AppendLine(ex.InnerException.Message);
                sb.AppendLine(ex.InnerException.StackTrace);
            }
            Write("ERROR", sb.ToString());
        }

        private static void Write(string level, string msg)
        {
            if (!_initialized) Initialize(); // 兜底
            var line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {level} {msg}";
            lock (_initLock)
            {
                _writer.WriteLine(line);
            }
            try
            {
                OnMessage?.Invoke(line); // 推到 UI
                Debug.WriteLine(line);   // 输出到调试窗口
            }
            catch { /* 避免 UI 事件异常影响日志 */ }
        }
    }
}
