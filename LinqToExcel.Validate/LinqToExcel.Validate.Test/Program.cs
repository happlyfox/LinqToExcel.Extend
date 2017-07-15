using System;

namespace LinqToExcel.Validate.Test
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //LinqToExcel在取数据方面给予了及大的帮助，但是在测试中我发现其本身具有的问题
            //没有实现验证功能。当Excel列和实体关系字段对应类型不同时报错

            //本代码在LinqToExcel的基础上扩展了验证功能，实例化验证对象时，传入对应的excel路径，自动完成验证。
            //自动验证为参数有效性验证，用户可自定义逻辑有效性验证，作为参数实例化时传入
            // 实例：定义验证对象
            // WorkBookValidate workbook = new WorkBookValidate("Default.xlsx");
            //自定义工作簿验证
            //workbook[0].StartValidate<User>();
            //验证结束返回Verification对象，对象包含 俩个属性，一个为是否验证成功，一个为验证出错的集合信息

            ModelValidate.BasicValidate();
            Console.Read();
        }
    }
}