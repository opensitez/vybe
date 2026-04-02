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
    /// a, b = 1, 2
    MultiAssign {
        targets: Vec<Expression>,
        splat_index: Option<usize>,
        values: Vec<Expression>,
    },
    /// alias new_name old_name
    Alias {
        new_name: String,
        old_name: String,
    },
    /// private / protected / public
    AccessModifier(AccessLevel),
    /// retry (in rescue)
    Retry,
    /// loop { body }
    Loop(Vec<Statement>),
    /// catch(:tag) { body }
    CatchThrow {
        tag: Expression,
        body: Vec<Statement>,
    },
    /// redo
    Redo,
    /// at_exit { body }
    AtExit(Vec<Statement>),
    Empty,
}

#[derive(Debug, Clone, PartialEq)]
pub enum AccessLevel {
    Public,
    Private,
    Protected,
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
    pub keyword: bool,    // name: (keyword arg)
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
    /// /pattern/ regex literal
    Regex(String),
    /// defined?(expr)
    Defined(Box<Expression>),
    /// &:method_name (symbol-to-proc)
    SymbolProc(String),
    /// proc { |x| body }
    ProcLiteral {
        params: Vec<String>,
        body: Vec<Statement>,
    },
    /// Struct.new(:field1, :field2)
    StructNew {
        name: Option<String>,
        fields: Vec<String>,
    },
    /// throw :tag, value
    Throw {
        tag: Box<Expression>,
        value: Option<Box<Expression>>,
    },
    /// obj.freeze
    Freeze(Box<Expression>),
    /// obj.frozen?
    FrozenCheck(Box<Expression>),
    /// obj.respond_to?(:method)
    RespondTo {
        object: Box<Expression>,
        method: String,
    },
    /// obj.send(:method, args)
    Send {
        object: Box<Expression>,
        method: Box<Expression>,
        args: Vec<Expression>,
    },
    /// Chained assignment: a = b = c = 1
    ChainedAssign {
        targets: Vec<Expression>,
        value: Box<Expression>,
    },
    /// case expr; in pattern => body; end (Ruby 3)
    PatternMatch {
        subject: Box<Expression>,
        arms: Vec<PatternArm>,
        else_body: Option<Vec<Statement>>,
    },
    /// if/unless/case/begin as expression (returns value of last expr in branch)
    IfExpr {
        test: Box<Expression>,
        body: Vec<Statement>,
        elsifs: Vec<ElsIf>,
        else_body: Option<Vec<Statement>>,
    },
    UnlessExpr {
        test: Box<Expression>,
        body: Vec<Statement>,
        else_body: Option<Vec<Statement>>,
    },
    BeginExpr {
        body: Vec<Statement>,
        rescues: Vec<RescueClause>,
        else_body: Option<Vec<Statement>>,
        ensure: Option<Vec<Statement>>,
    },
    /// expr rescue default
    InlineRescue {
        expr: Box<Expression>,
        rescue_val: Box<Expression>,
    },
    /// `command` backtick shell
    Backtick(String),
    /// obj&.method (safe navigation)
    SafeNav {
        receiver: Box<Expression>,
        method: String,
        args: Vec<Expression>,
        block: Option<Box<BlockArg>>,
    },
    /// __FILE__, __LINE__, __dir__, __method__
    MagicConstant(MagicConst),
    /// obj.instance_variable_get/set
    IvarGet {
        object: Box<Expression>,
        name: String,
    },
    IvarSet {
        object: Box<Expression>,
        name: String,
        value: Box<Expression>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub enum MagicConst {
    File,
    Line,
    Dir,
    Method,
}

#[derive(Debug, Clone)]
pub struct PatternArm {
    pub pattern: Expression,
    pub guard: Option<Expression>,
    pub body: Vec<Statement>,
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
