using System.Collections.ObjectModel;
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

                // 初始化SQLite数据库
                string dbName = "GameDB.db3";
                string dbFilePath = Path.Combine(exportPath, dbName);
                dbHelper = new SQLiteDBHelper(dbFilePath);

                // 遍历Excel文件列表，逐个转换为SQLite的DB文件
                foreach (var fileItem in Files)
                {
                    string excelFilePath = fileItem.FilePath;
                    string tableName = Path.GetFileNameWithoutExtension(excelFilePath);

                    // 添加日志记录
                    LogHelper.Log($"开始转换{tableName}...");

                    // 执行Excel到SQLite的转换
                    dbHelper.CreateTable(tableName, GetColumnNames(excelFilePath), GetColumnTypes(excelFilePath));
                    ConvertExcelToSQLite(excelFilePath, tableName);
                }

                // 导出成功的日志记录
                LogHelper.Log("数据库导出成功！");
            }
            catch (Exception ex)
            {
                // 导出失败的日志记录
                LogHelper.Log($"导出数据库时发生错误: {ex.Message}",LogLevel.ERROR);
            }
            finally
            {
                // 不论导出是否成功，弹出日志界面
                LogHelper.ShowLogWindow();
            }
        }


        private void ConvertExcelToSQLite(string excelFilePath, string tableName)
        {
            IWorkbook workbook;
            using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            int rowCount = sheet.LastRowNum;

            // 获取表头的列名
            IRow headerRow = sheet.GetRow(2);
            List<string> columnNames = new List<string>();
            foreach (ICell cell in headerRow.Cells)
            {
                columnNames.Add(cell.ToString());
            }

            for (int rowIndex = 4; rowIndex <= rowCount; rowIndex++) // 注意修改起始行为第四行
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    List<string> values = new List<string>();
                    for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                    {
                        ICell cell = row.GetCell(columnIndex);
                        string columnName = columnNames[columnIndex];

                        // 这里添加根据列名将 Excel 单元格的值映射到相应的列的逻辑
                        // 例如：values.Add(MapExcelCellValueToDBType(cell, columnName));

                        values.Add(cell.ToString()); // 简单示例，直接使用 ToString
                    }

                    string insertSql = dbHelper.GetInsertValueStr(tableName, values.ToArray());
                    dbHelper.ExecuteNonQuery(insertSql, null);
                }
            }
        }


        private string[] GetColumnNames(string excelFilePath)
        {
            IWorkbook workbook;
            using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(2); // 获取第三行作为表头
            List<string> columnNames = new List<string>();
            foreach (ICell cell in headerRow.Cells)
            {
                columnNames.Add(cell.ToString());
            }

            return columnNames.ToArray();
        }


        private string[] GetColumnTypes(string excelFilePath)
        {
            IWorkbook workbook;
            using (FileStream file = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }

            ISheet sheet = workbook.GetSheetAt(0);
            IRow dataTypeRow = sheet.GetRow(1);

            List<string> columnTypes = new List<string>();
            foreach (ICell cell in dataTypeRow.Cells)
            {
                // 根据你的需要将 Excel 中的数据类型映射为 SQLite 数据库类型
                string dataType = MapExcelDataTypeToSQLite(cell.ToString());
                columnTypes.Add(dataType);
            }

            return columnTypes.ToArray();
        }

        private string MapExcelDataTypeToSQLite(string excelDataType)
        {
            // 根据需要添加其他映射
            switch (excelDataType.ToLower())
            {
                case "int":
                    return "INT";
                case "string":
                    return "TEXT";
                case "decimal":
                    return "DECIMAL";
                case "boolean":
                    return "BOOL";
                case "single":
                    return "REAL";
                default:
                    return "TEXT";
            }
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
                        LogHelper.Log($"Error extracting icon for {selectedFilePath}: {ex.Message}",LogLevel.ERROR);
                    }
                }
                
            }
        }
    }
}