using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ExportDB.Core;
using ExportDB.Tools;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.Util.Collections;
using NPOI.XSSF.UserModel;
using LogLevel = ExportDB.Core.LogLevel;

namespace ExportDB
{
    public partial class MainWindow : Window
    {
        private SQLiteDBHelper dbHelper;
        public ObservableCollection<FileItem> Files { get; set; }

        public MainWindow()
        {
            LogHelper.Init();
            InitializeComponent();
            Files = new ObservableCollection<FileItem>();
            FileList.ItemsSource = Files;
            DBExportPath.Text = LoadPersonFromConfig();
        }

        // 保存数据到配置文件
        private void SavePersonToConfig()
        {
            var configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);  
            var settings = configFile.AppSettings.Settings;  
            settings["Path"].Value = DBExportPath.Text;  
            configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(configFile.AppSettings.SectionInformation.Name);
        }

        // 从配置文件中加载数据
        private string LoadPersonFromConfig()
        {
            
            try  
            {  
                var appSettings = ConfigurationManager.AppSettings;  
                return appSettings["Path"] ?? "Not Found";
            }
            catch (Exception e)
            {
                // 处理异常，比如配置文件不存在或数据格式错误
                LogHelper.Log($"加载数据时发生错误: {e.Message}", LogLevel.ERROR);
                throw;
            }
        }
        

        private void OnClikExport(object sender, RoutedEventArgs e)
        {
            LogHelper.Clear();
            if (Files.Count <= 0)
            {
                MessageBox.Show("导出Excel列表为空!");
                return;
            }

            try
            {
                // 获取DB导出路径
                string exportPath = DBExportPath.Text;

                // 检查导出路径是否存在，不存在则创建
                if (!Directory.Exists(exportPath))
                {
                    Directory.CreateDirectory(exportPath);
                }

                foreach (var fileItem in Files)
                {
                    string dbName = $"{fileItem.ExcelName}.db";
                    string dbFilePath = Path.Combine(exportPath, dbName);
                    //如果目标路径中已经存在指定名称的DB文件，那么应该删除它
                    if (File.Exists(dbFilePath))
                    {
                        LogHelper.Log($"删除{dbName}...");
                        File.Delete(dbFilePath);
                    }
                }


                // 遍历Excel文件列表，逐个转换为SQLite的DB文件
                foreach (var fileItem in Files)
                {
                    // 初始化SQLite数据库
                    string dbName = $"{fileItem.ExcelName}.db";
                    string dbFilePath = Path.Combine(exportPath, dbName);
                    dbHelper = new SQLiteDBHelper(dbFilePath);

                    string excelFilePath = fileItem.FilePath;
                    string tableName = Path.GetFileNameWithoutExtension(excelFilePath);

                    // 添加日志记录
                    LogHelper.Log($"开始转换{tableName}...");

                    // 执行Excel到SQLite的转换
                    SQLDB.Change2DB(dbHelper, excelFilePath);
                }

                // 导出成功的日志记录
                LogHelper.Log("数据库导出成功！");
            }
            catch (Exception ex)
            {
                // 导出失败的日志记录
                LogHelper.Log($"导出数据库时发生错误: {ex.Message}", LogLevel.ERROR);
                LogHelper.Log($"导出数据库时发生错误: {ex.StackTrace}", LogLevel.ERROR);
                LogHelper.Log("导出失败!!!");
            }
            finally
            {
                // 不论导出是否成功，弹出日志界面
                LogHelper.ShowLogWindow();
            }
        }


        private void OnSelectDBExportPathClick(object sender, RoutedEventArgs e)
        {
            string selectedFolderPath = FolderBrowser.ShowDialog();

            if (!string.IsNullOrEmpty(selectedFolderPath))
            {
                // 更新 TextBlock 的内容
                DBExportPath.Text = selectedFolderPath;
                SavePersonToConfig();
            }
        }

        private void OnRemoveButtonClick(object sender, RoutedEventArgs e)
        {
            Button removeButton = (Button)sender;
            FileItem fileToRemove = (FileItem)removeButton.DataContext;
            Files.Remove(fileToRemove);
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
                        LogHelper.Log($"Error extracting icon for {selectedFilePath}: {ex.Message}", LogLevel.ERROR);
                        LogHelper.Log("导出失败!!!");
                    }
                }
            }
        }
    }
}