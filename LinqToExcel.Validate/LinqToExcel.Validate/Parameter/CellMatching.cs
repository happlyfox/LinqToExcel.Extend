using System;
using System.Linq.Expressions;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 单元格匹配条件
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class CellMatching<T>
    {
        /// <summary>
        /// 验证属性
        /// </summary>
        public Expression<Func<T, object>> paramater { get; set; }

        /// <summary>
        /// 匹配的条件
        /// </summary>
        public Func<T, bool> matchCondition { get; set; }

        /// <summary>
        /// 定义出错信息
        /// </summary>
        public string errMsg { get; set; }
    }
}