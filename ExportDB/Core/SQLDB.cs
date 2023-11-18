using System.Data.Common;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExportDB.Core;

public class SQLDB
{
    ///2转DB
    public static void Change2DB(SQLiteDBHelper sqliteDBHelper, string ExcelPath)
    {
        IWorkbook workbook = null;
        FileStream fileStream = new FileStream(ExcelPath, FileMode.Open, FileAccess.Read);
        string tableName = Path.GetFileNameWithoutExtension(ExcelPath);
        if (ExcelPath.IndexOf(".xlsx") > 0) // 2007版本  
        {
            workbook = new XSSFWorkbook(fileStream); //xlsx数据读入workbook  
        }
        else if (ExcelPath.IndexOf(".xls") > 0) // 2003版本  
        {
            workbook = new HSSFWorkbook(fileStream); //xls数据读入workbook  
        }
        else
        {
            return;
        }

        ISheet sheet = workbook.GetSheetAt(0);
        if (sheet != null)
        {
            //创建标投
            IRow rowItem1 = sheet.GetRow(2);
            IRow rowItem2 = sheet.GetRow(3);

            List<string> ColumnsName = new List<string>();
            List<string> ColumnsType = new List<string>();


            List<string> CSType = new List<string>();
            int cellCount = rowItem1.LastCellNum;
            try
            {
                for (int k = 0; k < cellCount; k++)
                {
                    ICell cell = rowItem1.GetCell(k);
                    ICell cel2 = rowItem2.GetCell(k);
                    if (cell != null)
                    {
                        string cellValue = cell.StringCellValue;
                        ColumnsName.Add(cel2.StringCellValue);
                        ColumnsType.Add(CS2DbType(cell.StringCellValue));
                        CSType.Add(GetCSharpType(cell.StringCellValue));
                    }
                }
            }
            catch (Exception e)
            {
                LogHelper.Log($"导出数据库时发生错误: {e.Message}", LogLevel.ERROR);
                LogHelper.Log($"导出数据库时发生错误: {e.StackTrace}", LogLevel.ERROR);
                throw;
            }


            sqliteDBHelper.CreateTable(tableName, ColumnsName.ToArray(), ColumnsType.ToArray());

            using (DbTransaction transaction = sqliteDBHelper.connection.BeginTransaction())
            {
                using (DbCommand cmd = sqliteDBHelper.connection.CreateCommand())
                {
                    int startRow = 4;
                    int rowCount = sheet.LastRowNum;
                    for (int i = startRow; i <= rowCount; i++)
                    {
                        IRow rowItem = sheet.GetRow(i);
                        List<string> ContentStr = new List<string>();
                        bool caInsert = true;
                        for (int k = 0; k < rowItem.LastCellNum; k++)
                        {
                            ICell cellItem = rowItem.GetCell(k);
                            if (cellItem != null && !string.IsNullOrEmpty(cellItem.ToString()))
                            {
                                ContentStr.Add(cellItem.ToString());
                            }
                            else
                            {
                                caInsert = false;
                                break;
                            }
                        }

                        if (caInsert)
                        {
                            //解决SQL注入漏洞  cmd是SqlCommand对象
                            cmd.CommandText =
                                sqliteDBHelper.GetInsertValueStr(tableName,
                                    ColumnsName
                                        .ToArray()); // @"select count(*) from UserInfo where UserName=@UserName and UserPwd=@UserPwd";

                            for (int k = 0; k < ColumnsName.Count; k++)
                            {
                                DbParameter para = cmd.CreateParameter();
                                para.ParameterName = "@" + ColumnsName[k] + ""; //SQL参数化
                                para.Value = ContentStr[k]; //SQL参数化
                                cmd.Parameters.Add(para);
                            }

                            cmd.ExecuteScalar();
                        }
                    }
                }

                transaction.Commit();
            }
        }
    }

    /// <summary>
    /// 转化为为C#对应的类型
    /// </summary>
    /// <param name="_type"></param>
    /// <returns></returns>
    private static string GetCSharpType(string _type)
    {
        string DbType = "";
        if (_type == "Int")
        {
            DbType = "System.Int32";
        }
        else if (_type == "String")
        {
            DbType = "System.String";
        }
        else if (_type == "Vector3")
        {
            DbType = "Vector3";
        }
        else if (_type == "Decimal")
        {
            DbType = "System.Decimal";
        }
        else if (_type == "Boolean")
        {
            DbType = "System.Boolean";
        }
        else if (_type == "Single")
        {
            DbType = "System.Single";
        }
        else
        {
            LogHelper.Log($"没有找到合适的类型：{_type}");
        }

        return DbType;
    }

    /// <summary>
    /// 把类型转化为数据库类型
    /// </summary>
    /// <param name="_type"></param>
    /// <returns></returns>
    private static string CS2DbType(string _type)
    {
        string DbType = "";
        if (_type == "Int")
        {
            DbType = "INT";
        }
        else if (_type == "String" || _type == "Vector3")
        {
            DbType = "Text";
        }
        else if (_type == "Decimal")
        {
            DbType = "DECIMAL";
        }
        else if (_type == "Boolean")
        {
            DbType = "BOOL";
        }
        else if (_type == "Single")
        {
            DbType = "REAL";
        }
        else
        {
            LogHelper.Log($"没有找到合适的类型：{_type}");
        }

        return DbType;
    }
}