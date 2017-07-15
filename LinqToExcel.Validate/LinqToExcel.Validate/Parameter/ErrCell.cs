namespace LinqToExcel.Validate
{
    /// <summary>
    /// 错误单元
    /// </summary>
    public class ErrCell
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 名称 such as A1
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrMsg { get; set; }

        public override string ToString()
        {
            return string.Format("{0}\t{1}\t{2}\t{3}", RowIndex, ColumnIndex, Name, ErrMsg);
        }
    }
}