using System.Windows;

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

        // 可以根据实际情况获取日志内容的方法
        private string GetLogContent()
        {
            // 这里返回日志列表中的内容
            return "这是一个示例日志。";
        }
    }
}