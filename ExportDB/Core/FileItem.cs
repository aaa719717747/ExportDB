using System.Windows.Media;

namespace ExportDB.Core;

public class FileItem
{
    public string FilePath { get; set; }
    public ImageSource Icon { get; set; }
    
    public string DisplayName
    {
        get { return System.IO.Path.GetFileName(FilePath); }
    }
}