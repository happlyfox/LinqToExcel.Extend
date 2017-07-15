using LinqToExcel.Attributes;
using System;

namespace LinqToExcel.Validate.Test.Model
{
    public class User
    {
        [ExcelColumn("Id")]
        public string ID { get; set; }

        [ExcelColumn("名称")]
        public string Name { get; set; }

        [ExcelColumn("年龄")]
        public int Age { get; set; }

        [ExcelColumn("出生日期")]
        public DateTime BirthDay { get; set; }

        public override string ToString()
        {
            return string.Format("{0}\t{1}\t{2}", ID, Name, Age);
        }
    }
}