/// A complete Ruby program.
#[derive(Debug, Clone)]
pub struct Program {
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Statement {
    Expression(Expression),
    Block(Vec<Statement>),
    /// puts expr, expr, ...
    Puts(Vec<Expression>),
    /// print expr
    Print(Vec<Expression>),
    /// p expr
    P(Vec<Expression>),
    /// var = value (local assignment)
    Assignment {
        target: Expression,
        op: AssignOp,
        value: Expression,
    },
    MethodDef(MethodDecl),
    ClassDef(ClassDecl),
    ModuleDef(ModuleDecl),
    If {
        test: Expression,
        body: Vec<Statement>,
        elsifs: Vec<ElsIf>,
        else_body: Option<Vec<Statement>>,
    },
    Unless {
        test: Expression,
        body: Vec<Statement>,
        else_body: Option<Vec<Statement>>,
    },
    While {
        test: Expression,
        body: Vec<Statement>,
    },
    Until {
        test: Expression,
        body: Vec<Statement>,
    },
    For {
        var: String,
        iterable: Expression,
        body: Vec<Statement>,
    },
    Case {
        subject: Option<Expression>,
        whens: Vec<WhenClause>,
        else_body: Option<Vec<Statement>>,
    },
    Return(Option<Expression>),
    Break(Option<Expression>),
    Next(Option<Expression>),
    Raise(Option<Expression>),
    Begin {
        body: Vec<Statement>,
        rescues: Vec<RescueClause>,
        else_body: Option<Vec<Statement>>,
        ensure: Option<Vec<Statement>>,
    },
    Require(String),
    Empty,
}

#[derive(Debug, Clone)]
pub struct ElsIf {
    pub test: Expression,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct WhenClause {
    pub conditions: Vec<Expression>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct RescueClause {
    pub types: Vec<String>,
    pub var: Option<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct MethodDecl {
    pub name: String,
    pub params: Vec<Param>,
    pub body: Vec<Statement>,
    pub is_self: bool,  // def self.method_name
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub default: Option<Expression>,
    pub splat: bool,      // *args
    pub double_splat: bool, // **kwargs
    pub block: bool,      // &block
}

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub name: String,
    pub parent: Option<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct ModuleDecl {
    pub name: String,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Expression {
    Number(f64),
    Str(String),
    Symbol(String),
    Bool(bool),
    Nil,
    SelfExpr,
    /// Local variable or method call without parens
    Identifier(String),
    /// @name
    InstanceVar(String),
    /// @@name
    ClassVar(String),
    /// $name
    GlobalVar(String),
    /// CONSTANT
    ConstantRef(String),
    /// [1, 2, 3]
    Array(Vec<Expression>),
    /// {key => val, ...} or {key: val, ...}
    Hash(Vec<(Expression, Expression)>),
    /// a..b or a...b
    Range {
        start: Box<Expression>,
        end: Box<Expression>,
        exclusive: bool,
    },
    /// obj.method(args)
    MethodCall {
        receiver: Option<Box<Expression>>,
        method: String,
        args: Vec<Expression>,
        block: Option<Box<BlockArg>>,
    },
    /// obj[idx]
    IndexAccess {
        object: Box<Expression>,
        index: Box<Expression>,
    },
    /// Binary op
    Binary {
        op: BinaryOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    /// Unary op
    Unary {
        op: UnaryOp,
        expr: Box<Expression>,
    },
    /// condition ? then : else
    Ternary {
        test: Box<Expression>,
        consequent: Box<Expression>,
        alternate: Box<Expression>,
    },
    /// { |x| body } or do |x| body end
    Block {
        params: Vec<String>,
        body: Vec<Statement>,
    },
    /// -> (x) { body }  lambda
    Lambda {
        params: Vec<Param>,
        body: Vec<Statement>,
    },
    /// Proc.new { |x| body }
    ProcNew {
        params: Vec<String>,
        body: Vec<Statement>,
    },
    /// yield(args)
    Yield(Vec<Expression>),
    /// block_given?
    BlockGiven,
    /// super(args)
    Super(Vec<Expression>),
    /// Class::Member
    ScopeResolution {
        left: Box<Expression>,
        name: String,
    },
    /// Splat: *expr
    Splat(Box<Expression>),
    /// String interpolation: "hello #{expr}"
    Interpolated(Vec<InterpolPart>),
    /// attr_reader :name, :age  etc
    AttrDecl {
        kind: AttrKind,
        names: Vec<String>,
    },
    /// include ModuleName
    Include(String),
    /// extend ModuleName
    Extend(String),
}

#[derive(Debug, Clone)]
pub enum InterpolPart {
    Lit(String),
    Expr(Expression),
}

#[derive(Debug, Clone)]
pub struct BlockArg {
    pub params: Vec<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum AttrKind {
    Reader,
    Writer,
    Accessor,
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinaryOp {
    Add, Sub, Mul, Div, Mod, Pow,
    Eq, Ne, Lt, Gt, Le, Ge,
    Spaceship,
    And, Or,           // &&, ||
    BitAnd, BitOr, BitXor,
    Shl, Shr,
    RangeIncl, RangeExcl,  // .., ...
}

#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOp {
    Neg, Pos, Not, BitNot,
}

#[derive(Debug, Clone, PartialEq)]
pub enum AssignOp {
    Assign,
    AddAssign, SubAssign, MulAssign, DivAssign, ModAssign,
    AndAssign, OrAssign,
    BitAndAssign, BitOrAssign, BitXorAssign,
    ShlAssign, ShrAssign,
}
