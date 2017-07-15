using LinqToExcel.Attributes;
using LinqToExcel.Validate.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 工作表验证
    /// </summary>
    public class WorkSheetValidate
    {
        #region 工作表参数

        public string FilePath { set; get; }
        public string SheetName { set; get; }
        public int SheetIndex { set; get; }

        #endregion 工作表参数

        #region 构造函数

        public WorkSheetValidate(string filePath, string sheetName, int sheetIndex)
        {
            FilePath = filePath;
            SheetName = sheetName;
            SheetIndex = sheetIndex;

            TootarIndex = 0;
        }

        #endregion 构造函数

        #region 标题行索引

        public int TootarIndex { get; set; }

        #endregion 标题行索引

        #region 实体属性对应[列标题][标题对应表格中的实际索引]

        private Dictionary<string, PropertyColIndex> _propertyCollection;

        public Dictionary<string, PropertyColIndex> PropertyCollection
        {
            get
            {
                return _propertyCollection;
            }
        }

        #endregion 实体属性对应[列标题][标题对应表格中的实际索引]

        #region 属性映射

        private Dictionary<string, string> _propertyMappings = new Dictionary<string, string>();

        public Dictionary<string, string> PropertyMappings
        {
            get
            {
                return _propertyMappings;
            }
            set
            {
                _propertyMappings = value;
            }
        }

        public void AddMapping(string propertyName, string excelColName)
        {
            if (PropertyMappings.ContainsKey(propertyName))
            {
                PropertyMappings[propertyName] = excelColName;
            }
            else
            {
                PropertyMappings.Add(propertyName, excelColName);
            }
        }

        public void AddMapping<T>(Expression<Func<T, object>> propertyExpression, string excelColName)
        {
            PropertyExpressionParser<T> property = new PropertyExpressionParser<T>(propertyExpression);
            string propertyName = property.Name;

            if (PropertyMappings.ContainsKey(propertyName))
            {
                PropertyMappings[propertyName] = excelColName;
            }
            else
            {
                PropertyMappings.Add(propertyName, excelColName);
            }
        }

        #endregion 属性映射

        public Verification StartValidate<T>(List<CellMatching<T>> rowValidates = null)
        {
            List<ErrCell> errCells = this.ValidateParameter<T>(TootarIndex);
            if (!errCells.Any())
            {
                TootarIndex += 1;
                errCells.AddRange(this.ValidateMatching<T>(rowValidates, TootarIndex));
            }

            Verification validate = new Verification();

            if (errCells.Any())
            {
                validate = new Verification()
                {
                    IfPass = false,
                    ErrCells = errCells
                };
            }
            else
            {
                validate = new Verification()
                {
                    IfPass = true,
                    ErrCells = errCells
                };
            }

            return validate;
        }

        public List<T> GetRows<T>(List<Transformation<T>> trans = null)
        {
            var excel = new ExcelQueryFactory(FilePath);
            foreach (var mapping in PropertyMappings)
            {
                excel.AddMapping(mapping.Key, mapping.Value);
            }

            if (trans != null)
            {
                foreach (var transformation in trans)
                {
                    excel.AddTransformation<T>(transformation.propertyExpression, transformation.transformation);
                }
            }

            List<T> rows = (from c in excel.Worksheet<T>(SheetIndex)
                            select c).ToList();
            return rows;
        }

        private List<ErrCell> ValidateParameter<T>(int startRowIndex)
        {
            var excel = new ExcelQueryFactory(FilePath);
            var rows = (from c in excel.WorksheetNoHeader(SheetIndex)
                        select c).ToList();

            var headerColNames = rows[startRowIndex].ToList(); //标题行

            //实体属性标签列表
            var propertys = typeof(T).GetProperties().Select(u => new
            {
                ProPertyName = u.Name,
                AttributeName = Attribute.IsDefined(u, typeof(ExcelColumnAttribute)) ? ((ExcelColumnAttribute)(Attribute.GetCustomAttribute(u, typeof(ExcelColumnAttribute)))).ColumnName : ""
            }).ToList();

            Dictionary<string, string> propertyAttributes = new Dictionary<string, string>();

            propertys.ForEach(property =>
            {
                if (PropertyMappings.ContainsKey(property.ProPertyName))
                {
                    propertyAttributes.Add(property.ProPertyName, PropertyMappings[property.ProPertyName]);
                }
                else
                {
                    if (!string.IsNullOrEmpty(property.AttributeName))
                    {
                        propertyAttributes.Add(property.ProPertyName, property.AttributeName);
                    }
                }
            });

            _propertyCollection = new Dictionary<string, PropertyColIndex>();
            for (int excelHeaderColIndex = 0; excelHeaderColIndex < headerColNames.Count; excelHeaderColIndex++)
            {
                string colName = headerColNames[excelHeaderColIndex];

                foreach (var property in propertyAttributes)
                {
                    if (property.Value == colName)
                    {
                        _propertyCollection.Add(property.Key, new PropertyColIndex
                        {
                            ColName = colName,
                            ColIndex = excelHeaderColIndex
                        });
                        break;
                    }
                }
            }

            startRowIndex++;
            return GetErrCellByParameter<T>(rows, startRowIndex);
        }

        private List<ErrCell> ValidateMatching<T>(List<CellMatching<T>> rowValidates, int startRowIndex)
        {
            if (rowValidates == null)
                return new List<ErrCell>();

            foreach (var rowValidate in rowValidates)
            {
                PropertyExpressionParser<T> property = new PropertyExpressionParser<T>(rowValidate.paramater);
                if (!_propertyCollection.Any(u => u.Key == property.Name))
                {
                    throw new Exception(string.Format("列必须拥有映射关系{0}", property.Name));
                }
            }

            var excel = new ExcelQueryFactory(FilePath);
            foreach (var mapping in PropertyMappings)
            {
                excel.AddMapping(mapping.Key, mapping.Value);
            }
            var rows = (from c in excel.Worksheet<T>(SheetIndex)
                        select c).ToList();

            List<ErrCell> errCells = new List<ErrCell>();
            foreach (var row in rows)
            {
                errCells.AddRange(GetErrCellByMatching(row, startRowIndex++, rowValidates));
            }
            return errCells;
        }

        private List<ErrCell> GetErrCellByMatching<T>(T box, int rowIndex, List<CellMatching<T>> rowValidates)
        {
            List<T> row = new List<T>();
            row.Add(box);

            List<ErrCell> errCells = new List<ErrCell>();
            for (int i = 0; i < rowValidates.Count; i++)
            {
                var rowValidate = rowValidates[i];
                if (!row.Any(rowValidate.matchCondition))
                {
                    string errMsg = rowValidate.errMsg;

                    PropertyExpressionParser<T> property = new PropertyExpressionParser<T>(rowValidate.paramater);
                    var colDescription = (Attribute.GetCustomAttribute(property.GetPropertyInfo(rowValidate.paramater), typeof(ExcelColumnAttribute))) as ExcelColumnAttribute;
                    int ColIndex = _propertyCollection[property.Name].ColIndex;

                    if (!errCells.Any(u => u.RowIndex == rowIndex && u.ColumnIndex == ColIndex))
                    {
                        errCells.Add(new ErrCell()
                        {
                            RowIndex = rowIndex,
                            ColumnIndex = ColIndex,
                            Name = RowValidate.GetCellStation(rowIndex, ColIndex),
                            ErrMsg = errMsg
                        });
                    }
                    else
                    {
                        errCells.Find(u => u.RowIndex == rowIndex && u.ColumnIndex == ColIndex).ErrMsg += ";" + errMsg;
                    }
                }
            }
            return errCells;
        }

        private List<ErrCell> GetErrCellByParameter<T>(List<RowNoHeader> rows, int startRowIndex)
        {
            List<string> colNames = _propertyCollection.Values.Select(u => u.ColName).ToList();
            List<int> colIndexs = _propertyCollection.Values.Select(u => u.ColIndex).ToList();

            List<ErrCell> errCells = new List<ErrCell>();
            for (int rowIndex = startRowIndex; rowIndex < rows.Count; rowIndex++)
            {
                List<string> rowValues = rows[rowIndex].Where((u, index) => colIndexs.Any(p => p == index)).Select(u => u.ToString()).ToList();
                errCells.AddRange(RowValidate.Validate<T>(rowIndex, colNames, colIndexs, rowValues));
            }
            return errCells;
        }
    }
}