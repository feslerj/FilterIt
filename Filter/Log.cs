using System;

namespace Filter
{
    public class LogState
    {
        public string CurrentLogMessage { private set; get; }
        public string CurrentErrorMessage { private set; get; }
        public string CurrentErrorStackTrace { private set; get; } 

        public LogState()
        {
            Clear();
        }

        public void Log(string msg)
        {
            Clear();
            CurrentLogMessage = msg;
        }

        public void Error(string msg, Exception ex)
        {
            Clear();
            CurrentLogMessage = msg;
            CurrentErrorMessage = ex.Message;
            CurrentErrorStackTrace = ex.StackTrace;
        }

        public void Error(string msg)
        {
            Clear();
            CurrentLogMessage = CurrentErrorMessage = msg;
        }

        public void Clear()
        {
            CurrentLogMessage = CurrentErrorMessage = CurrentErrorStackTrace = String.Empty;
        }

    }
}