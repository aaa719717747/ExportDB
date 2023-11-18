namespace ExportDB.Core;

public enum LogLevel
{
    INFO,
    ERROR,
}

public static class LogHelper
{
    public static LogWindow logWindow;
    public static List<string> logEntries = new List<string>();

    public static void Init()
    {
        logWindow = new LogWindow();
        Clear();
    }

    public static void Log(string log,LogLevel logLevel= LogLevel.INFO)
    {
        logEntries.Add($"[{logLevel.ToString()}]{log}");
    }

    public static void Clear()
    {
        logEntries.Clear();
    }

    public static void ShowLogWindow()
    {
        // 在这里添加传递日志的逻辑
        logWindow.SetLogEntries(logEntries); // 假设 logEntries 是你存储日志的集合
        logWindow.ShowDialog();
    }
}