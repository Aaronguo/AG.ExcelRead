using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelRead
{
    public class ExcelRead
    {
        /// <summary>
        /// 读取Excel文件到DataSet中
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public List<ProResult> ToDataTable(string filePath, string name)
        {
            List<ProResult> ProResultlist = new List<ProResult>();

            var fileName = name;
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(fileName);


            if (string.IsNullOrEmpty(fileType)) return null;
            if (fileType == ".xls" || fileType == ".xlt")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";
            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            System.Data.DataTable dtSheetName = null;
            DataSet ds = new DataSet();
            try
            {
                // 初始化连接，并打开
                conn = new OleDbConnection(connStr);
                conn.Open();
                // 获取数据源的表定义元数据 
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                // 初始化适配器
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];
                    if ((SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$")) || SheetName.Contains("评估$") || SheetName.Contains("健康指导$"))
                    {
                        continue;
                    }
                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);


                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, SheetName);
                    if (SheetName == "'1$'")
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            AddProResult(ProResultlist, j, 0, 1, 3, dsItem, SheetName, SheetName);
                        }
                    }
                    var newSheetName = (Regex.Replace(SheetName, @"((')|(\d)|(\d)|(#)|(\$))+", ""));

                    for (int j = 0; j < dsItem.Tables[SheetName].Rows.Count; j++)
                    {
                        if (j < 9) continue;
                        if (string.IsNullOrEmpty(newSheetName)) continue;
                        else if (newSheetName == "心理情绪")
                        {
                            AddProResult(ProResultlist, j, 0, 1, 4, dsItem, SheetName, newSheetName);
                            AddProResult(ProResultlist, j, 4, 5, 8, dsItem, SheetName, newSheetName);
                        }
                        else
                        {
                            AddProResult(ProResultlist, j, 0, 1, dsItem.Tables[SheetName].Columns.Count, dsItem, SheetName, newSheetName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log(ex.Message);
            }
            finally
            {
                // 关闭连接
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return ProResultlist;
        }

        /// <summary>
        /// 项目结果
        /// </summary>
        public class ProResult
        {
            /// <summary>
            /// 检查名称
            /// </summary>
            public string examine { get; set; }

            /// <summary>
            /// 检查结果
            /// </summary>
            public string result { get; set; }
            /// <summary>
            /// 项目名称
            /// </summary>
            public string proName { get; set; }
        }

        /// <summary>
        /// 提取DataSet中的检测项目信息
        /// </summary>
        /// <param name="proResultlist">检测结果集合</param>
        /// <param name="x">当前循环的行数</param>
        /// <param name="y">当前循环的列数</param>
        /// <param name="z">起始循环的列数</param>
        /// <param name="w">结束循环的列数</param>
        /// <param name="dsItem">DataSet数据</param>
        /// <param name="SheetName">Sheet名称</param>
        /// <param name="newSheetName">新的Sheet名称</param>
        /// <returns></returns>
        public static List<ProResult> AddProResult(List<ProResult> proResultlist, int x, int y, int z, int w, DataSet dsItem, string SheetName, string newSheetName)
        {
            var examine = dsItem.Tables[SheetName].Rows[x][y].ToString();
            var result = "";
            for (int k = z; k < w; k++)
            {
                if (!string.IsNullOrEmpty(dsItem.Tables[SheetName].Rows[x][k].ToString()))
                {
                    result += dsItem.Tables[SheetName].Rows[x][k].ToString() + ";";
                }
            }
            if (!string.IsNullOrEmpty(examine) && !string.IsNullOrEmpty(result) && !examine.Contains("文字") && !examine.Contains("评估"))
            {
                ProResult proResult = new ProResult()
                {
                    examine = examine,
                    result = result,
                    proName = newSheetName
                };
                proResultlist.Add(proResult);
            }
            return proResultlist;
        }

        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="path"></param>
        public void log(string msg, string path = "C://Medical")
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                if (!dir.Exists)
                {
                    dir.Create();
                }
                string filename = path + "/" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                System.IO.FileInfo file = new System.IO.FileInfo(filename);
                StreamWriter writer = null;
                if (!file.Exists)
                {
                    writer = file.CreateText();
                }
                else
                {
                    writer = file.AppendText();
                }
                writer.Write(msg + "\r\n");
                writer.Flush();
                writer.Close();
            }
            catch { }
        }
    }
}
