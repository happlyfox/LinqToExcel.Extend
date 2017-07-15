using LinqToExcel.Attributes;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 行验证
    /// </summary>
    public class RowValidate
    {
        public static string GetCellStation(int rowIndex, int columnIndex)
        {
            int i = columnIndex % 26;
            string cellRef = Convert.ToChar(65 + i).ToString() + (rowIndex + 1);
            return cellRef;
        }

        public static List<ErrCell> Validate<T>(int rowIndex, List<string> colNames, List<int> colIndexs, List<string> rowCellValues)
        {
            List<ErrCell> errCells = new List<ErrCell>();
            T singleT = Activator.CreateInstance<T>();

            foreach (PropertyInfo pi in singleT.GetType().GetProperties())
            {
                var propertyAttribute = (Attribute.GetCustomAttribute(pi, typeof(ExcelColumnAttribute)));
                if (propertyAttribute == null)
                {
                    continue;
                }
                var proName = ((ExcelColumnAttribute)propertyAttribute).ColumnName;
                for (int colIndex = 0; colIndex < colNames.Count; colIndex++)
                {
                    try
                    {
                        if (proName.Equals(colNames[colIndex], StringComparison.OrdinalIgnoreCase))
                        {
                            string fieldName = pi.PropertyType.GetUnderlyingType().Name;
                            string cellValue = rowCellValues[colIndex];

                            if (!String.IsNullOrWhiteSpace(cellValue))
                            {
                                //如果是日期类型,特殊判断
                                if (fieldName.Equals("DateTime"))
                                {
                                    string data = "";
                                    try
                                    {
                                        data = cellValue.ToDateTimeValue();
                                    }
                                    catch (Exception)
                                    {
                                        data = DateTime.Parse(cellValue).ToString();
                                    }
                                }
                                cellValue.CastTo(pi.PropertyType);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errCells.Add(new ErrCell()
                        {
                            RowIndex = rowIndex,
                            ColumnIndex = colIndexs[colIndex],
                            Name = GetCellStation(rowIndex, colIndexs[colIndex]),
                            ErrMsg = ex.Message
                        });
                    }
                }
            }
            return errCells;
        }
    }
}