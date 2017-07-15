using LinqToExcel.Validate.Test.Model;
using System;
using System.Collections.Generic;

namespace LinqToExcel.Validate.Test
{
    public class ModelValidate
    {
        /// <summary>
        /// 参数有效性验证
        /// </summary>
        public static void ValidateParameter()
        {
            WorkBookValidate workbook = new WorkBookValidate("Default.xlsx");
            var errlist = workbook[0].StartValidate<User>();
            Validate(errlist);
        }

        /// <summary>
        /// 基础验证 通过实体特性
        /// </summary>
        public static void BasicValidate()
        {
            WorkBookValidate workbook = new WorkBookValidate("Default.xlsx");
            var errlist = workbook[0].StartValidate<User>();
            Validate(errlist);
        }

        /// <summary>
        /// 基础验证 自定义对应关系
        /// </summary>
        public static void BasicValidate2()
        {
            WorkBookValidate workbook = new WorkBookValidate("Other.xlsx");

            workbook[0].AddMapping("ID", "主键");
            var errlist = workbook[0].StartValidate<User>();
            Validate(errlist);
        }

        /// <summary>
        /// 带条件验证
        /// </summary>
        public static void ValidateWithCondition()
        {
            WorkBookValidate workbook = new WorkBookValidate("Default.xlsx");

            List<CellMatching<User>> rowValidate = new List<CellMatching<User>>();
            rowValidate.Add(new CellMatching<User>()
            {
                paramater = u => u.Age,
                matchCondition = u => u.Age > 10,
                errMsg = "请选择年龄大于10的人员"
            });

            var errlist = workbook[0].StartValidate<User>(rowValidate);
            Validate(errlist);
        }

        private static void Validate(Verification verifity)
        {
            Console.WriteLine(string.Format("是否通过{0}", verifity.IfPass));
            //出错列
            foreach (var item in verifity.ErrCells)
            {
                Console.WriteLine(item);
            }
        }
    }
}