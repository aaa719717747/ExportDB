# ExportDB
这是一个简易的将Excel转换为DB的WPF桌面程序
this is a excel to SQLite  DB  file tool

![UPM_Image](https://github.com/aaa719717747/ExportDB/blob/master/Images/show.png?raw=true)

### 文件目录：
- Setup  包含工具exe可执行程序包
- Tool   包含SQLiteDB查看工具
- ExportDBFolder   包含导出的示例DB文件
- StandardSampleExcel  包含示例的标准的Excel格式源文件
- ExportDB   此工具的WPF程序源码工程

Excel2DB支持的数据类型说明：
| Excel数据类型 | SQLite数据类型 | C#数据类型 |
|----------|:---------|:--------:|
| Int   | INT  | System.Int32   |
| String   | TEXT  | System.String  |
| Single   | REAL  | float   |
| Boolean   | BOOL  | System.Boolean  |
| Vector3   | TEXT  | System.String   |

### 1.0.2 预发布功能
- 1.增加一键导出C#代码API功能
