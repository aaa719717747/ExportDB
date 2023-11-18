using System.IO;
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
    public string ExcelName
    {
        get
        {
            string pathName = System.IO.Path.GetFileName(FilePath);
            return Path.GetFileNameWithoutExtension(pathName);;
        }
    }
}