using System;
using System.Linq.Expressions;

namespace LinqToExcel.Validate
{
    public class Transformation<T>
    {
        public Expression<Func<T, object>> propertyExpression;
        public Func<string, object> transformation;
    }
}