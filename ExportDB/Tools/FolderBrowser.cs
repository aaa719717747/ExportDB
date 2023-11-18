using Ookii.Dialogs.Wpf;

namespace ExportDB.Tools
{
    public static class FolderBrowser
    {
        public static string ShowDialog()
        {
            VistaFolderBrowserDialog folderDialog = new VistaFolderBrowserDialog();

            bool? result = folderDialog.ShowDialog();

            if (result == true)
            {
                return folderDialog.SelectedPath;
            }

            return null;
        }
    }
}