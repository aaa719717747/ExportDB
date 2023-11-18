using System.Collections.Generic;
using System.Windows;
using ExportDB.Core;

namespace ExportDB
{
    public partial class LogWindow : Window
    {
        public LogWindow()
        {
            InitializeComponent();
        }

        private void OnCopyLogClick(object sender, RoutedEventArgs e)
        {
            // 将日志列表中的内容复制到剪切板
            Clipboard.SetText(GetLogContent());
            MessageBox.Show("已复制到剪切板");
        }

        // 获取日志内容的方法
        private string GetLogContent()
        {
            // 这里返回日志列表中的内容
            List<string> logEntries = new List<string>();
            foreach (var item in LogListBox.Items)
            {
                logEntries.Add(item.ToString());
            }

            return string.Join("\n", logEntries);
        }

        // 设置日志列表
        public void SetLogEntries(List<string> logEntries)
        {
            LogListBox.ItemsSource = logEntries;
        }
    }
}