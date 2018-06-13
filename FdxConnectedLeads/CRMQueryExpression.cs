using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxConnectedLeads
{
    class CRMQueryExpression
    {
        string filterAttribute;
        ConditionOperator filterOperator;
        object filterValue;

        public CRMQueryExpression(string _attr, ConditionOperator _opr, object _val)
        {
            this.filterAttribute = _attr;
            this.filterOperator = _opr;
            this.filterValue = _val;
        }
        public static QueryExpression getQueryExpression(string _entityName, ColumnSet _columnSet, CRMQueryExpression[] _exp, LogicalOperator _filterOperator = LogicalOperator.And)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = _columnSet;
            query.Distinct = false;
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = _filterOperator;
            for (int i = 0; i < _exp.Length; i++)
            {
                query.Criteria.AddCondition(_exp[i].filterAttribute, _exp[i].filterOperator, _exp[i].filterValue);
            }

            return query;
        }
    }
}
