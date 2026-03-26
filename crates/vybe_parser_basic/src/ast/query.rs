use crate::ast::expr::Expression;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct QueryExpression {
    pub from_clause: FromClause,
    pub body: QueryBody,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct FromClause {
    pub ranges: Vec<RangeVariable>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct RangeVariable {
    pub name: String,
    pub collection: Expression,
    pub type_name: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct QueryBody {
    pub clauses: Vec<QueryClause>,
    pub select_or_group: SelectOrGroupClause,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum QueryClause {
    Where(Expression),
    OrderBy(Vec<Ordering>),
    Let { name: String, value: Expression },
    // Join not yet implemented
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Ordering {
    pub expression: Expression,
    pub direction: OrderDirection,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum OrderDirection {
    Ascending,
    Descending,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum SelectOrGroupClause {
    Select(Vec<Expression>), // Usually one, but can be multiple for anonymous types? In VB it's usually Select explicit_expr
    Group(Box<GroupClause>),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GroupClause {
    pub item: Expression,
    pub key: Expression,
}
