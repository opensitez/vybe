use super::Identifier;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Expression {
    // Literals
    IntegerLiteral(i32),
    DoubleLiteral(f64),
    StringLiteral(String),
    BooleanLiteral(bool),
    DateLiteral(String),
    ArrayLiteral(Vec<Expression>),
    Nothing,

    // Variables and access
    Variable(Identifier),
    MemberAccess(Box<Expression>, Identifier),
    ArrayAccess(Identifier, Vec<Expression>),

    // Binary operations
    Add(Box<Expression>, Box<Expression>),
    Subtract(Box<Expression>, Box<Expression>),
    Multiply(Box<Expression>, Box<Expression>),
    Divide(Box<Expression>, Box<Expression>),
    IntegerDivide(Box<Expression>, Box<Expression>),
    Modulo(Box<Expression>, Box<Expression>),
    Exponent(Box<Expression>, Box<Expression>),
    Concatenate(Box<Expression>, Box<Expression>),

    // Comparison
    Equal(Box<Expression>, Box<Expression>),
    NotEqual(Box<Expression>, Box<Expression>),
    LessThan(Box<Expression>, Box<Expression>),
    LessThanOrEqual(Box<Expression>, Box<Expression>),
    GreaterThan(Box<Expression>, Box<Expression>),
    GreaterThanOrEqual(Box<Expression>, Box<Expression>),

    // Logical
    And(Box<Expression>, Box<Expression>),
    AndAlso(Box<Expression>, Box<Expression>),
    Or(Box<Expression>, Box<Expression>),
    OrElse(Box<Expression>, Box<Expression>),
    Xor(Box<Expression>, Box<Expression>),
    Not(Box<Expression>),

    // Reference equality
    Is(Box<Expression>, Box<Expression>),
    IsNot(Box<Expression>, Box<Expression>),
    
    // Pattern matching
    Like(Box<Expression>, Box<Expression>),

    // Type checking
    TypeOf {
        expr: Box<Expression>,
        type_name: String,
    },
    
    // Bitwise Shift
    BitShiftLeft(Box<Expression>, Box<Expression>),
    BitShiftRight(Box<Expression>, Box<Expression>),

    // Unary
    Negate(Box<Expression>),

    // Function/method calls
    Call(Identifier, Vec<Expression>),
    MethodCall(Box<Expression>, Identifier, Vec<Expression>),
    
    // Instantiation
    New(Identifier, Vec<Expression>),

    // Collection initializer: New List(Of T) From { expr, expr, ... }
    NewFromInitializer(Identifier, Vec<Expression>, Vec<Expression>),

    // Object initializer: New Type() With { .Prop = expr, ... }
    NewWithInitializer(Identifier, Vec<Expression>, Vec<(String, Expression)>),

    // Lambda
    Lambda {
        params: Vec<super::decl::Parameter>,
        body: Box<LambdaBody>,
    },
    
    // Async
    Await(Box<Expression>),

    // Self-reference
    Me,

    // Base class reference (MyBase.Member)
    MyBase,

    // With block implicit target (for .Property syntax)
    WithTarget,

    // Inline If expression: If(cond, true, false) or If(value, default)
    IfExpression(Box<Expression>, Box<Expression>, Option<Box<Expression>>),

    // AddressOf (delegate reference - stored as string for now)
    AddressOf(String),

    // Type cast: CType(expr, Type), DirectCast(expr, Type), TryCast(expr, Type)
    Cast {
        kind: CastKind,
        expr: Box<Expression>,
        target_type: String,
    },

    // LINQ Query
    Query(Box<super::query::QueryExpression>),

    // XML Literals
    XmlLiteral(Box<super::xml::XmlNode>),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CastKind {
    CType,
    DirectCast,
    TryCast,
}

impl Expression {
    pub fn binary(op: BinaryOp, left: Expression, right: Expression) -> Self {
        match op {
            BinaryOp::Add => Expression::Add(Box::new(left), Box::new(right)),
            BinaryOp::Subtract => Expression::Subtract(Box::new(left), Box::new(right)),
            BinaryOp::Multiply => Expression::Multiply(Box::new(left), Box::new(right)),
            BinaryOp::Divide => Expression::Divide(Box::new(left), Box::new(right)),
            BinaryOp::IntegerDivide => Expression::IntegerDivide(Box::new(left), Box::new(right)),
            BinaryOp::Modulo => Expression::Modulo(Box::new(left), Box::new(right)),
            BinaryOp::Exponent => Expression::Exponent(Box::new(left), Box::new(right)),
            BinaryOp::Concatenate => Expression::Concatenate(Box::new(left), Box::new(right)),
            BinaryOp::Equal => Expression::Equal(Box::new(left), Box::new(right)),
            BinaryOp::NotEqual => Expression::NotEqual(Box::new(left), Box::new(right)),
            BinaryOp::LessThan => Expression::LessThan(Box::new(left), Box::new(right)),
            BinaryOp::LessThanOrEqual => Expression::LessThanOrEqual(Box::new(left), Box::new(right)),
            BinaryOp::GreaterThan => Expression::GreaterThan(Box::new(left), Box::new(right)),
            BinaryOp::GreaterThanOrEqual => Expression::GreaterThanOrEqual(Box::new(left), Box::new(right)),
            BinaryOp::And => Expression::And(Box::new(left), Box::new(right)),
            BinaryOp::AndAlso => Expression::AndAlso(Box::new(left), Box::new(right)),
            BinaryOp::Or => Expression::Or(Box::new(left), Box::new(right)),
            BinaryOp::OrElse => Expression::OrElse(Box::new(left), Box::new(right)),
            BinaryOp::Xor => Expression::Xor(Box::new(left), Box::new(right)),
            BinaryOp::BitShiftLeft => Expression::BitShiftLeft(Box::new(left), Box::new(right)),
            BinaryOp::BitShiftRight => Expression::BitShiftRight(Box::new(left), Box::new(right)),
            BinaryOp::Is => Expression::Is(Box::new(left), Box::new(right)),
            BinaryOp::IsNot => Expression::IsNot(Box::new(left), Box::new(right)),
            BinaryOp::Like => Expression::Like(Box::new(left), Box::new(right)),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinaryOp {
    Add,
    Subtract,
    Multiply,
    Divide,
    IntegerDivide,
    Modulo,
    Exponent,
    Concatenate,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
    And,
    AndAlso,
    Or,
    OrElse,
    Xor,
    BitShiftLeft,
    BitShiftRight,
    Is,
    IsNot,
    Like,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum LambdaBody {
    Expression(Box<Expression>),
    Statement(Box<super::stmt::Statement>),
    Block(Vec<super::stmt::Statement>),
}
