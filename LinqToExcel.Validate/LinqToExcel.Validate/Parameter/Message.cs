using System.Collections.Generic;

namespace LinqToExcel.Validate
{
    /// <summary>
    /// 验证返回实体
    /// </summary>
    public class Verification
    {
        /// <summary>
        /// 是否通过验证
        /// </summary>
        public bool IfPass { get; set; }

        /// <summary>
        /// 错误单元格集合
        /// </summary>
        public List<ErrCell> ErrCells { get; set; }
    }
}