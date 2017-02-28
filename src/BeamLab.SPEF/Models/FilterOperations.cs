using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace BeamLab.SPEF.Models
{
    public interface ISPEFQueryNode<T>
    {

    }

    public class SPEFQueryNode<T> : ISPEFQueryNode<T>
    {
        public static SPEFQueryNode<T> Where<TP2>(Expression<Func<T, TP2>> expression, Op op, object value)
        {
            return new SPEFExpression<T, TP2>(op)
            {
                Expression = expression,
                Value = value
            };
        }

        public SPEFQueryNode<T> And(SPEFQueryNode<T> operation2)
        {
            return new SPEFOperation<T>(Operators.And)
            {
                Operation1 = this,
                Operation2 = operation2
            };
        }

        public SPEFQueryNode<T> And<TP2>(Expression<Func<T, TP2>> expression2, Op op, object value)
        {
            return new SPEFOperation<T>(Operators.And)
            {
                Operation1 = this,
                Operation2 = new SPEFExpression<T, TP2>(op)
                {
                    Expression = expression2,
                    Value = value
                }
            };
        }

        public SPEFQueryNode<T> Or(SPEFQueryNode<T> operation2)
        {
            return new SPEFOperation<T>(Operators.Or)
            {
                Operation1 = this,
                Operation2 = operation2
            };
        }

        public SPEFQueryNode<T> Or<TP2>(Expression<Func<T, TP2>> expression2, Op op, object value)
        {
            return new SPEFOperation<T>(Operators.Or)
            {
                Operation1 = this,
                Operation2 = new SPEFExpression<T, TP2>(op)
                {
                    Expression = expression2,
                    Value = value
                }
            };
        }

        public SPEFSortNode<T, TP2> OrderBy<TP2>(Expression<Func<T, TP2>> expression, bool ascending = true)
        {
            return new SPEFSortNode<T, TP2>()
            {
                Query = this,
                Expression = expression,
                Ascending = ascending
            };
        }
    }

    //public class SPEFQueryNode<T> : ISPEFQueryNode<T>
    //{
    //    public static SPEFQueryNode<T> Where<TP2>(Expression<Func<T, TP2>> expression, Op op, TP2 value)
    //    {
    //        return new SPEFExpression<T, TP2>(op)
    //        {
    //            Expression = expression,
    //            Value = value
    //        };
    //    }

    //    public SPEFQueryNode<T> And(SPEFQueryNode<T> operation2)
    //    {
    //        return new SPEFOperation<T>(Operators.And)
    //        {
    //            Operation1 = this,
    //            Operation2 = operation2
    //        };
    //    }

    //    public SPEFQueryNode<T> And<TP2>(Expression<Func<T, TP2>> expression2, Op op, TP2 value)
    //    {
    //        return new SPEFOperation<T>(Operators.And)
    //        {
    //            Operation1 = this,
    //            Operation2 = new SPEFExpression<T, TP2>(op)
    //            {
    //                Expression = expression2,
    //                Value = value
    //            }
    //        };
    //    }

    //    public SPEFQueryNode<T> Or(SPEFQueryNode<T> operation2)
    //    {
    //        return new SPEFOperation<T>(Operators.Or)
    //        {
    //            Operation1 = this,
    //            Operation2 = operation2
    //        };
    //    }

    //    public SPEFQueryNode<T> Or<TP2>(Expression<Func<T, TP2>> expression2, Op op, TP2 value)
    //    {
    //        return new SPEFOperation<T>(Operators.Or)
    //        {
    //            Operation1 = this,
    //            Operation2 = new SPEFExpression<T, TP2>(op)
    //            {
    //                Expression = expression2,
    //                Value = value
    //            }
    //        };
    //    }

    //    public SPEFSortNode<T, TP2> OrderBy<TP2>(Expression<Func<T, TP2>> expression, bool ascending = true)
    //    {
    //        return new SPEFSortNode<T, TP2>()
    //        {
    //            Query = this,
    //            Expression = expression,
    //            Ascending = ascending
    //        };
    //    }
    //}

    public enum Operators
    {
        And,
        Or
    }
    public class SPEFOperation<T> : SPEFQueryNode<T>
    {
        public SPEFQueryNode<T> Operation1 { get; set; }
        public Operators Operator { get; private set; }
        public SPEFQueryNode<T> Operation2 { get; set; }

        public SPEFOperation(Operators op)
        {
            Operator = op;
        }
    }

    public enum Op
    {
        Eq,
        Neq,
        Leq,
        Lt,
        Geq,
        Gt,
        IsNull,
        IsNotNull,
        BeginsWith,
        Contains,
    }
    public class SPEFExpression<T, TP> : SPEFQueryNode<T>
    {
        public SPEFExpression(Op op)
        {
            this.Op = op;
        }
        public Expression<Func<T, TP>> Expression { get; set; }
        public Op Op { get; private set; }
        public object Value { get; set; }
    }
    
    public class SPEFSortNode<T, TP> : ISPEFQueryNode<T>
    {
        public SPEFQueryNode<T> Query { get; set; }
        public  bool Ascending { get; set; }
        public Expression<Func<T, TP>> Expression { get; set; }
    }

    public class SPEFCount<TP>
    {
        public TP Value { get; set; }
        public int Count { get; set; }
    }

}
