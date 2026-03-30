/// Complete Dart AST for the vybe bytecode compiler.

#[derive(Debug, Clone)]
pub struct Program {
    pub body: Vec<TopLevel>,
}

/// Top-level declarations in a Dart file.
#[derive(Debug, Clone)]
pub enum TopLevel {
    Import(ImportDecl),
    Function(FunctionDecl),
    Class(ClassDecl),
    Variable(VarDecl),
    Statement(Statement), // top-level expressions / bare statements
}

#[derive(Debug, Clone)]
pub struct ImportDecl {
    pub uri: String,
    pub prefix: Option<String>,   // import 'x' as prefix
    pub show: Vec<String>,         // show a, b
    pub hide: Vec<String>,         // hide a, b
}

// ── Statements ────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Statement {
    Block(Vec<Statement>),
    VarDecl(VarDecl),
    FunctionDecl(FunctionDecl),
    Expression(Expression),
    If {
        condition: Expression,
        then_branch: Box<Statement>,
        else_branch: Option<Box<Statement>>,
    },
    While {
        condition: Expression,
        body: Box<Statement>,
    },
    DoWhile {
        body: Box<Statement>,
        condition: Expression,
    },
    For(ForStatement),
    ForIn {
        is_final: bool,
        var_type: Option<String>,
        var_name: String,
        iterable: Expression,
        body: Box<Statement>,
    },
    Switch {
        expr: Expression,
        cases: Vec<SwitchCase>,
    },
    Return(Option<Expression>),
    Break(Option<String>),
    Continue(Option<String>),
    Throw(Expression),
    Try {
        body: Vec<Statement>,
        catches: Vec<CatchClause>,
        finally: Option<Vec<Statement>>,
    },
    Assert(Expression, Option<Expression>),
    Empty,
}

#[derive(Debug, Clone)]
pub struct VarDecl {
    pub is_final: bool,
    pub is_const: bool,
    pub is_late: bool,
    pub type_name: Option<String>,
    pub name: String,
    pub initializer: Option<Expression>,
}

#[derive(Debug, Clone)]
pub struct ForStatement {
    pub init: Option<ForInit>,
    pub condition: Option<Expression>,
    pub update: Vec<Expression>,
    pub body: Box<Statement>,
}

#[derive(Debug, Clone)]
pub enum ForInit {
    VarDecl(VarDecl),
    Expression(Expression),
}

#[derive(Debug, Clone)]
pub struct SwitchCase {
    pub label: Option<Expression>, // None = default
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub on_type: Option<String>,   // on ExceptionType
    pub var_name: Option<String>,  // catch (e)
    pub stack_name: Option<String>, // catch (e, s)
    pub body: Vec<Statement>,
}

// ── Functions ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct FunctionDecl {
    pub name: String,
    pub type_params: Vec<String>,
    pub params: Params,
    pub return_type: Option<TypeAnnotation>,
    pub body: FunctionBody,
    pub is_async: bool,
    pub is_generator: bool,
}

#[derive(Debug, Clone)]
pub struct Params {
    pub positional: Vec<Param>,       // required positional
    pub optional_pos: Vec<Param>,     // [optional positional]
    pub named: Vec<Param>,            // {named params}
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub type_ann: Option<TypeAnnotation>,
    pub default_value: Option<Expression>,
    pub is_required: bool,   // required keyword
    pub is_this: bool,       // this.field shorthand
}

#[derive(Debug, Clone)]
pub enum FunctionBody {
    Block(Vec<Statement>),
    Expression(Expression),  // => expr
    Empty,                   // abstract / native
}

// ── Classes ───────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub name: String,
    pub type_params: Vec<String>,
    pub extends: Option<String>,
    pub implements: Vec<String>,
    pub mixins: Vec<String>,
    pub is_abstract: bool,
    pub members: Vec<ClassMember>,
}

#[derive(Debug, Clone)]
pub enum ClassMember {
    Field {
        is_static: bool,
        is_final: bool,
        is_late: bool,
        type_ann: Option<TypeAnnotation>,
        name: String,
        initializer: Option<Expression>,
    },
    Method {
        is_static: bool,
        is_abstract: bool,
        is_override: bool,
        kind: MethodKind,
        decl: FunctionDecl,
    },
    Constructor {
        name: Option<String>,   // None = default, Some = named ctor
        params: Params,
        initializers: Vec<CtorInitializer>,
        body: Option<Vec<Statement>>,
        is_const: bool,
        is_factory: bool,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub enum MethodKind { Method, Getter, Setter }

#[derive(Debug, Clone)]
pub enum CtorInitializer {
    SuperCall(Vec<Expression>),
    FieldInit(String, Expression),
    AssertInit(Expression),
}

// ── Expressions ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Expression {
    // Literals
    Int(i64),
    Double(f64),
    Bool(bool),
    Null,
    String(StringExpr),
    List { type_arg: Option<Box<TypeAnnotation>>, elements: Vec<Expression> },
    Map { type_args: Option<(Box<TypeAnnotation>, Box<TypeAnnotation>)>, entries: Vec<(Expression, Expression)> },
    Set { type_arg: Option<Box<TypeAnnotation>>, elements: Vec<Expression> },

    // Identifiers
    Identifier(String),
    This,
    Super,

    // Operations
    Binary { op: BinOp, left: Box<Expression>, right: Box<Expression> },
    Unary { op: UnaryOp, expr: Box<Expression> },
    PostfixUnary { op: PostfixOp, expr: Box<Expression> },
    Assign { op: AssignOp, left: Box<Expression>, right: Box<Expression> },
    Ternary { cond: Box<Expression>, then: Box<Expression>, else_: Box<Expression> },
    NullCoalesce { left: Box<Expression>, right: Box<Expression> },

    // Access
    Member { object: Box<Expression>, member: String, null_safe: bool },
    Index { object: Box<Expression>, index: Box<Expression> },
    Cascade { object: Box<Expression>, ops: Vec<CascadeOp> },

    // Calls
    Call { callee: Box<Expression>, type_args: Vec<TypeAnnotation>, args: Vec<Argument>, null_safe: bool },
    New { class: String, constructor: Option<String>, type_args: Vec<TypeAnnotation>, args: Vec<Argument> },
    Const { class: String, constructor: Option<String>, args: Vec<Argument> },

    // Functions
    Lambda { params: Params, body: Box<FunctionBody>, is_async: bool },

    // Type operations
    Is { expr: Box<Expression>, type_ann: TypeAnnotation, negated: bool },
    As { expr: Box<Expression>, type_ann: TypeAnnotation },

    // Async
    Await(Box<Expression>),

    // Spread
    Spread(Box<Expression>),

    // If-null assignment shorthand
    IfNull { left: Box<Expression>, right: Box<Expression> },
}

#[derive(Debug, Clone)]
pub enum StringExpr {
    Simple(String),
    Interpolated(Vec<StringPart>),
}

#[derive(Debug, Clone)]
pub enum StringPart {
    Literal(String),
    Expr(Expression),
}

#[derive(Debug, Clone)]
pub struct Argument {
    pub label: Option<String>,
    pub value: Expression,
}

#[derive(Debug, Clone)]
pub enum CascadeOp {
    Method(String, Vec<Argument>),
    Field(String),
    Index(Expression),
    Assign(String, Expression),
}

// ── Types ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub struct TypeAnnotation {
    pub name: String,
    pub args: Vec<TypeAnnotation>,
    pub nullable: bool,
}

// ── Operators ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    Add, Sub, Mul, Div, IntDiv, Mod,
    Eq, NotEq, Lt, Gt, Le, Ge,
    And, Or,
    BitAnd, BitOr, BitXor, Shl, Shr, UShr,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp { Neg, Not, BitNot, PreInc, PreDec }

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum PostfixOp { PostInc, PostDec }

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum AssignOp {
    Assign, AddAssign, SubAssign, MulAssign, DivAssign,
    ModAssign, IntDivAssign, AndAssign, OrAssign,
    BitAndAssign, BitOrAssign, BitXorAssign,
    ShlAssign, ShrAssign, NullAssign,
}
