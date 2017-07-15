using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 批量导入帮助类
    /// </summary>
    public class BulkHelper
    {
        /// <summary>
        /// 批量导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="connectionStr">链接字符串</param>
        /// <param name="tableName">表名</param>
        /// <param name="list">集合</param>
        /// <param name="batchSize">批量大小</param>
        public static void BulkInsert<T>(string connectionStr, string tableName, ICollection<T> list, int batchSize)
        {
            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connectionStr, SqlBulkCopyOptions.UseInternalTransaction))
            {
                try
                {
                    bulkCopy.BatchSize = batchSize;
                    int count = list.Count;
                    bulkCopy.NotifyAfter = count % 10 == 0 ? count / 10 : count % 10;
                    bulkCopy.DestinationTableName = tableName;
                    bulkCopy.SqlRowsCopied += new SqlRowsCopiedEventHandler(OnSqlRowsCopied);

                    var props = TypeDescriptor.GetProperties(typeof(T))
                        .Cast<PropertyDescriptor>()
                        .Where(propertyInfo => propertyInfo.PropertyType.Namespace.Equals("System")).ToArray();

                    var table = new DataTable();
                    foreach (var propertyInfo in props)
                    {
                        bulkCopy.ColumnMappings.Add(propertyInfo.Name, propertyInfo.Name);
                        table.Columns.Add(propertyInfo.Name, Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType);
                    }

                    var values = new object[props.Length];
                    foreach (var item in list)
                    {
                        for (var i = 0; i < values.Length; i++)
                        {
                            values[i] = props[i].GetValue(item);
                        }

                        table.Rows.Add(values);
                    }
                    bulkCopy.WriteToServer(table);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        private static void OnSqlRowsCopied(object sender, SqlRowsCopiedEventArgs e)
        {
            //处理完成后的回调事件
            Console.WriteLine(e.RowsCopied);
        }
    }
}