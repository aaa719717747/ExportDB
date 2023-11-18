using System.Collections.ObjectModel;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ExportDB.Core;
using ExportDB.Tools;
using Microsoft.Win32;

namespace ExportDB
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<FileItem> Files { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            Files = new ObservableCollection<FileItem>();
            FileList.ItemsSource = Files;
        }
        private void OnClikExport(object sender, RoutedEventArgs e)
        {
            
        }
        private void OnSelectDBExportPathClick(object sender, RoutedEventArgs e)
        {
            string selectedFolderPath = FolderBrowser.ShowDialog();

            if (!string.IsNullOrEmpty(selectedFolderPath))
            {
                // 更新 TextBlock 的内容
                DBExportPath.Text = selectedFolderPath;
            }
        }
        private void OnRemoveButtonClick(object sender, RoutedEventArgs e)
        {
            Button removeButton = (Button)sender;
            FileItem fileToRemove = (FileItem)removeButton.DataContext;
            Files.Remove(fileToRemove);
        }
        private void OnShowLogClick(object sender, RoutedEventArgs e)
        {
            LogWindow logWindow = new LogWindow();
            logWindow.ShowDialog();
        }

        private void OnSelectExcelFileClick(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                foreach (string selectedFilePath in openFileDialog.FileNames)
                {
                    // 检查文件是否已经存在于列表中
                    if (Files.Any(file => file.FilePath == selectedFilePath))
                    {
                        // 文件已存在，可以显示一条消息或者采取其他适当的操作
                        continue; // 跳过继续处理下一个文件
                    }

                    try
                    {
                        // 获取文件图标
                        Icon icon = System.Drawing.Icon.ExtractAssociatedIcon(selectedFilePath);

                        // 转换为WPF的ImageSource
                        ImageSource imageSource = Imaging.CreateBitmapSourceFromHIcon(
                            icon.Handle,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions()
                        );

                        // 添加到列表
                        Files.Add(new FileItem
                        {
                            FilePath = selectedFilePath,
                            Icon = imageSource
                        });
                    }
                    catch (Exception ex)
                    {
                        // 处理异常，你可以输出异常信息到日志或者显示一个错误消息框
                        Console.WriteLine($"Error extracting icon for {selectedFilePath}: {ex.Message}");
                    }
                }

            }
        }


    }
}