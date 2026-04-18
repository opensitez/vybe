// Python 3.x AST definitions

#[derive(Debug, Clone)]
pub struct Module {
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Statement {
    Expression(Expression),

    Assign {
        targets: Vec<Expression>,
        value: Expression,
    },
    AugAssign {
        target: Expression,
        op: AugOp,
        value: Expression,
    },
    AnnAssign {
        target: Expression,
        annotation: Expression,
        value: Option<Expression>,
    },

    If {
        test: Expression,
        body: Vec<Statement>,
        elif_clauses: Vec<(Expression, Vec<Statement>)>,
        else_body: Option<Vec<Statement>>,
    },
    While {
        test: Expression,
        body: Vec<Statement>,
        else_body: Option<Vec<Statement>>,
    },
    For {
        target: Expression,
        iter: Expression,
        body: Vec<Statement>,
        else_body: Option<Vec<Statement>>,
        is_async: bool,
    },
    Break,
    Continue,
    Pass,
    Return(Option<Expression>),

    FunctionDef {
        name: String,
        params: Parameters,
        body: Vec<Statement>,
        decorators: Vec<Expression>,
        returns: Option<Expression>,
        is_async: bool,
    },
    ClassDef {
        name: String,
        bases: Vec<Expression>,
        keywords: Vec<Keyword>,
        body: Vec<Statement>,
        decorators: Vec<Expression>,
    },

    Try {
        body: Vec<Statement>,
        handlers: Vec<ExceptHandler>,
        else_body: Option<Vec<Statement>>,
        finally_body: Option<Vec<Statement>>,
    },
    Raise {
        exc: Option<Expression>,
        cause: Option<Expression>,
    },

    Import {
        names: Vec<Alias>,
    },
    ImportFrom {
        module: Option<String>,
        names: Vec<Alias>,
        level: usize,
    },

    With {
        items: Vec<WithItem>,
        body: Vec<Statement>,
        is_async: bool,
    },

    Global(Vec<String>),
    Nonlocal(Vec<String>),
    Delete(Vec<Expression>),
    Assert {
        test: Expression,
        msg: Option<Expression>,
    },

    Match {
        subject: Expression,
        cases: Vec<MatchCase>,
    },
}

#[derive(Debug, Clone)]
pub enum Expression {
    // Literals
    Int(i64),
    Float(f64),
    Str(String),
    FString { parts: Vec<FStringPart> },
    Bool(bool),
    None,
    Ellipsis,

    // Collections
    List(Vec<Expression>),
    Tuple(Vec<Expression>),
    Dict {
        keys: Vec<Option<Expression>>,
        values: Vec<Expression>,
    },
    Set(Vec<Expression>),

    // Comprehensions
    ListComp {
        element: Box<Expression>,
        generators: Vec<Comprehension>,
    },
    SetComp {
        element: Box<Expression>,
        generators: Vec<Comprehension>,
    },
    DictComp {
        key: Box<Expression>,
        value: Box<Expression>,
        generators: Vec<Comprehension>,
    },
    GeneratorExp {
        element: Box<Expression>,
        generators: Vec<Comprehension>,
    },

    // Identifier
    Name(String),

    // Operations
    BinOp {
        op: BinOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    UnaryOp {
        op: UnaryOp,
        operand: Box<Expression>,
    },
    BoolOp {
        op: BoolOp,
        values: Vec<Expression>,
    },
    Compare {
        left: Box<Expression>,
        ops: Vec<CmpOp>,
        comparators: Vec<Expression>,
    },

    // Access
    Attribute {
        value: Box<Expression>,
        attr: String,
    },
    Subscript {
        value: Box<Expression>,
        slice: Box<Expression>,
    },
    Slice {
        lower: Option<Box<Expression>>,
        upper: Option<Box<Expression>>,
        step: Option<Box<Expression>>,
    },
    Starred(Box<Expression>),

    // Calls
    Call {
        func: Box<Expression>,
        args: Vec<Expression>,
        keywords: Vec<Keyword>,
    },

    // Conditional
    IfExp {
        test: Box<Expression>,
        body: Box<Expression>,
        orelse: Box<Expression>,
    },

    // Lambda
    Lambda {
        params: Parameters,
        body: Box<Expression>,
    },

    // Walrus
    NamedExpr {
        target: Box<Expression>,
        value: Box<Expression>,
    },

    // Async
    Await(Box<Expression>),
    Yield(Option<Box<Expression>>),
    YieldFrom(Box<Expression>),
}

// --- Supporting types ---

#[derive(Debug, Clone)]
pub struct Parameters {
    pub args: Vec<Param>,
    pub vararg: Option<Param>,
    pub kwonly_args: Vec<Param>,
    pub kwarg: Option<Param>,
    pub defaults: Vec<Expression>,
    pub kw_defaults: Vec<Option<Expression>>,
}

impl Default for Parameters {
    fn default() -> Self {
        Self {
            args: Vec::new(),
            vararg: None,
            kwonly_args: Vec::new(),
            kwarg: None,
            defaults: Vec::new(),
            kw_defaults: Vec::new(),
        }
    }
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub annotation: Option<Box<Expression>>,
}

#[derive(Debug, Clone)]
pub struct Keyword {
    pub name: Option<String>,
    pub value: Expression,
}

#[derive(Debug, Clone)]
pub struct Comprehension {
    pub target: Expression,
    pub iter: Expression,
    pub ifs: Vec<Expression>,
    pub is_async: bool,
}

#[derive(Debug, Clone)]
pub struct ExceptHandler {
    pub exc_type: Option<Expression>,
    pub name: Option<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct WithItem {
    pub context_expr: Expression,
    pub optional_vars: Option<Expression>,
}

#[derive(Debug, Clone)]
pub struct Alias {
    pub name: String,
    pub asname: Option<String>,
}

#[derive(Debug, Clone)]
pub struct MatchCase {
    pub pattern: Pattern,
    pub guard: Option<Expression>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Pattern {
    Value(Expression),
    Singleton(Expression),
    Sequence(Vec<Pattern>),
    Mapping(Vec<(Expression, Pattern)>),
    Class {
        cls: Expression,
        patterns: Vec<Pattern>,
        kw_patterns: Vec<(String, Pattern)>,
    },
    Star(Option<String>),
    As {
        pattern: Option<Box<Pattern>>,
        name: Option<String>,
    },
    Or(Vec<Pattern>),
    Wildcard,
}

#[derive(Debug, Clone)]
pub enum FStringPart {
    Literal(String),
    Expr(Expression),
    /// Expression with format spec, e.g. f"{x:.2f}" → FormattedExpr(x, ".2f")
    FormattedExpr(Expression, String),
}

// Operator enums

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    Add,
    Sub,
    Mul,
    Div,
    FloorDiv,
    Mod,
    Pow,
    LShift,
    RShift,
    BitOr,
    BitXor,
    BitAnd,
    MatMul,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp {
    Invert,
    Not,
    UAdd,
    USub,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BoolOp {
    And,
    Or,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CmpOp {
    Eq,
    NotEq,
    Lt,
    LtE,
    Gt,
    GtE,
    Is,
    IsNot,
    In,
    NotIn,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum AugOp {
    Add,
    Sub,
    Mul,
    Div,
    FloorDiv,
    Mod,
    Pow,
    LShift,
    RShift,
    BitOr,
    BitXor,
    BitAnd,
    MatMul,
}
