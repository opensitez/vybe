/// C# AST — covers the core language needed for WinForms apps.

#[derive(Debug, Clone)]
pub struct CompilationUnit {
    pub usings: Vec<String>,
    pub namespace: Option<String>,
    pub members: Vec<TypeDecl>,
    /// Top-level statements (C# 9+)
    pub top_level_statements: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum TypeDecl {
    Class(ClassDecl),
    Struct(StructDecl),
    Enum(EnumDecl),
    Interface(InterfaceDecl),
}

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub name: String,
    pub is_partial: bool,
    pub is_static: bool,
    pub is_abstract: bool,
    pub is_sealed: bool,
    pub access: Access,
    pub base_type: Option<String>,
    pub interfaces: Vec<String>,
    pub members: Vec<MemberDecl>,
}

#[derive(Debug, Clone)]
pub struct StructDecl {
    pub name: String,
    pub access: Access,
    pub members: Vec<MemberDecl>,
}

#[derive(Debug, Clone)]
pub struct EnumDecl {
    pub name: String,
    pub access: Access,
    pub members: Vec<(String, Option<Expression>)>,
}

#[derive(Debug, Clone)]
pub struct InterfaceDecl {
    pub name: String,
    pub access: Access,
    pub members: Vec<MemberDecl>,
}

#[derive(Debug, Clone)]
pub enum MemberDecl {
    Field {
        name: String,
        type_name: Option<String>,
        initializer: Option<Expression>,
        is_static: bool,
        is_readonly: bool,
        is_const: bool,
        access: Access,
    },
    Method(MethodDecl),
    Constructor(ConstructorDecl),
    Property(PropertyDecl),
    Event {
        name: String,
        type_name: String,
        access: Access,
    },
}

#[derive(Debug, Clone)]
pub struct MethodDecl {
    pub name: String,
    pub return_type: Option<String>,
    pub params: Vec<Parameter>,
    pub body: Vec<Statement>,
    pub is_static: bool,
    pub is_override: bool,
    pub is_virtual: bool,
    pub is_abstract: bool,
    pub is_async: bool,
    pub access: Access,
}

#[derive(Debug, Clone)]
pub struct ConstructorDecl {
    pub params: Vec<Parameter>,
    pub body: Vec<Statement>,
    pub base_args: Option<Vec<Expression>>,
    pub access: Access,
}

#[derive(Debug, Clone)]
pub struct PropertyDecl {
    pub name: String,
    pub type_name: Option<String>,
    pub getter: Option<Vec<Statement>>,
    pub setter: Option<(String, Vec<Statement>)>, // (value_param, body)
    pub is_auto: bool,
    pub access: Access,
}

#[derive(Debug, Clone)]
pub struct Parameter {
    pub name: String,
    pub type_name: Option<String>,
    pub default: Option<Expression>,
    pub is_params: bool,
    pub is_ref: bool,
    pub is_out: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Access {
    Public,
    Private,
    Protected,
    Internal,
}

// ============================================================
// Statements
// ============================================================

#[derive(Debug, Clone)]
pub enum Statement {
    // Declarations
    LocalDecl {
        name: String,
        type_name: Option<String>,
        initializer: Option<Expression>,
        is_var: bool,
    },
    // Assignments
    Assignment {
        target: Expression,
        value: Expression,
    },
    CompoundAssignment {
        target: Expression,
        op: CompoundOp,
        value: Expression,
    },
    // Control flow
    If {
        condition: Expression,
        then_body: Vec<Statement>,
        else_if: Vec<(Expression, Vec<Statement>)>,
        else_body: Option<Vec<Statement>>,
    },
    For {
        init: Option<Box<Statement>>,
        condition: Option<Expression>,
        update: Option<Box<Statement>>,
        body: Vec<Statement>,
    },
    ForEach {
        var_name: String,
        iterable: Expression,
        body: Vec<Statement>,
    },
    While {
        condition: Expression,
        body: Vec<Statement>,
    },
    DoWhile {
        body: Vec<Statement>,
        condition: Expression,
    },
    Switch {
        expr: Expression,
        cases: Vec<SwitchCase>,
    },
    // Jump
    Return(Option<Expression>),
    Break,
    Continue,
    Throw(Expression),
    // Exception
    TryCatchFinally {
        try_body: Vec<Statement>,
        catches: Vec<CatchClause>,
        finally_body: Option<Vec<Statement>>,
    },
    // Expressions as statements
    Expression(Expression),
    // Using statement
    Using {
        var_name: String,
        initializer: Expression,
        body: Vec<Statement>,
    },
    // Lock
    Lock {
        lock_object: Expression,
        body: Vec<Statement>,
    },
    // Block
    Block(Vec<Statement>),
    // Empty
    Empty,
}

#[derive(Debug, Clone)]
pub struct SwitchCase {
    pub labels: Vec<Option<Expression>>, // None = default
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub type_name: Option<String>,
    pub var_name: Option<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CompoundOp {
    AddAssign, SubAssign, MulAssign, DivAssign, ModAssign,
    AndAssign, OrAssign, XorAssign, ShlAssign, ShrAssign,
}

// ============================================================
// Expressions
// ============================================================

#[derive(Debug, Clone)]
pub enum Expression {
    // Literals
    IntLiteral(i64),
    DoubleLiteral(f64),
    StringLiteral(String),
    CharLiteral(char),
    BoolLiteral(bool),
    NullLiteral,
    InterpolatedString(Vec<StringPart>),

    // Names
    Identifier(String),
    This,
    Base,

    // Binary
    Binary(BinaryOp, Box<Expression>, Box<Expression>),

    // Unary
    Unary(UnaryOp, Box<Expression>),
    PostIncrement(Box<Expression>),
    PostDecrement(Box<Expression>),
    PreIncrement(Box<Expression>),
    PreDecrement(Box<Expression>),

    // Member access
    MemberAccess(Box<Expression>, String),
    NullConditionalAccess(Box<Expression>, String), // ?.

    // Indexer
    Index(Box<Expression>, Box<Expression>),

    // Invocation
    Call(Box<Expression>, Vec<Expression>),

    // Object creation
    New(String, Vec<Expression>),
    NewArray(String, Box<Expression>),
    ArrayInit(Vec<Expression>),
    ObjectInit(Box<Expression>, Vec<(String, Expression)>),

    // Cast / type check
    Cast(String, Box<Expression>),
    Is(Box<Expression>, String),
    As(Box<Expression>, String),
    TypeOf(String),

    // Ternary
    Conditional(Box<Expression>, Box<Expression>, Box<Expression>),

    // Null coalescing
    NullCoalescing(Box<Expression>, Box<Expression>),

    // Lambda
    Lambda(Vec<String>, Box<Expression>),
    LambdaBlock(Vec<String>, Vec<Statement>),

    // Await
    Await(Box<Expression>),

    // Nameof
    NameOf(String),

    // Default
    Default(Option<String>),
}

#[derive(Debug, Clone)]
pub enum StringPart {
    Text(String),
    Expr(Expression),
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinaryOp {
    Add, Sub, Mul, Div, Mod,
    Eq, Neq, Lt, Gt, Le, Ge,
    And, Or, // logical
    BitAnd, BitOr, BitXor, Shl, Shr,
    NullCoalescing,
    Range, // ..
}

#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOp {
    Neg, Not, BitNot,
}
