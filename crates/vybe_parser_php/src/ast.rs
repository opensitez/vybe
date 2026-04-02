/// A complete PHP program.
#[derive(Debug, Clone)]
pub struct Program {
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum Statement {
    Expression(Expression),
    Block(Vec<Statement>),
    /// echo expr, expr, ...;
    Echo(Vec<Expression>),
    /// $var = ...  or  function/const declarations
    VariableDeclaration {
        name: String,
        value: Option<Expression>,
    },
    /// const NAME = expr;
    ConstDeclaration {
        name: String,
        value: Expression,
    },
    /// global $var, $var2;
    Global(Vec<String>),
    FunctionDeclaration(FunctionDecl),
    ClassDeclaration(ClassDecl),
    If {
        test: Expression,
        consequent: Box<Statement>,
        alternates: Vec<ElseIf>,
        alternate: Option<Box<Statement>>,
    },
    While {
        test: Expression,
        body: Box<Statement>,
    },
    DoWhile {
        body: Box<Statement>,
        test: Expression,
    },
    For {
        init: Vec<Expression>,
        test: Option<Expression>,
        update: Vec<Expression>,
        body: Box<Statement>,
    },
    ForEach {
        array: Expression,
        key: Option<String>,
        value: String,
        body: Box<Statement>,
    },
    Switch {
        discriminant: Expression,
        cases: Vec<SwitchCase>,
    },
    Return(Option<Expression>),
    Break(Option<Expression>),
    Continue(Option<Expression>),
    Throw(Expression),
    Try {
        block: Vec<Statement>,
        catches: Vec<CatchClause>,
        finalizer: Option<Vec<Statement>>,
    },
    Empty,
}

#[derive(Debug, Clone)]
pub struct ElseIf {
    pub test: Expression,
    pub body: Box<Statement>,
}

#[derive(Debug, Clone)]
pub struct SwitchCase {
    /// None = default:
    pub test: Option<Expression>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct CatchClause {
    /// Caught type hints (may be multiple: `catch (Foo|Bar $e)`)
    pub types: Vec<String>,
    pub var: Option<String>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct FunctionDecl {
    pub name: String,
    pub params: Vec<Param>,
    pub body: Vec<Statement>,
    pub is_static: bool,
    pub visibility: Visibility,
    pub return_by_ref: bool,
}

#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub default: Option<Expression>,
    pub by_ref: bool,
    pub variadic: bool,
    pub type_hint: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Visibility {
    Public,
    Private,
    Protected,
    None,
}

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub name: String,
    pub parent: Option<String>,
    pub interfaces: Vec<String>,
    pub traits: Vec<String>,
    pub members: Vec<ClassMember>,
}

#[derive(Debug, Clone)]
pub enum ClassMember {
    Method(FunctionDecl),
    Property {
        name: String,
        visibility: Visibility,
        is_static: bool,
        default: Option<Expression>,
    },
    Constant {
        name: String,
        value: Expression,
    },
}

#[derive(Debug, Clone)]
pub enum Expression {
    // Literals
    Number(f64),
    Str(String),
    Bool(bool),
    Null,
    /// $name
    Variable(String),
    /// Bare identifier used as a constant or function name
    Identifier(String),
    /// array(...) or [...]
    Array(Vec<ArrayElement>),
    /// function($x) use ($y) { }  or  fn($x) => expr
    Closure {
        params: Vec<Param>,
        uses: Vec<String>,
        body: Box<ClosureBody>,
        is_arrow: bool,
    },
    /// $obj->prop  or  $obj?->prop
    Property {
        object: Box<Expression>,
        name: Box<Expression>,
        nullsafe: bool,
    },
    /// $obj->method(...)  or  $obj?->method(...)
    MethodCall {
        object: Box<Expression>,
        method: Box<Expression>,
        args: Vec<Argument>,
        nullsafe: bool,
    },
    /// ClassName::method(...)  or  ClassName::CONST  or  $obj::method()
    StaticAccess {
        class: Box<Expression>,
        member: Box<Expression>,
    },
    StaticCall {
        class: Box<Expression>,
        method: Box<Expression>,
        args: Vec<Argument>,
    },
    /// Regular function call: foo(...)
    Call {
        callee: Box<Expression>,
        args: Vec<Argument>,
    },
    /// new ClassName(...)
    New {
        class: Box<Expression>,
        args: Vec<Argument>,
    },
    Binary {
        op: BinaryOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Unary {
        op: UnaryOp,
        expr: Box<Expression>,
    },
    PreUpdate {
        op: UpdateOp,
        expr: Box<Expression>,
    },
    PostUpdate {
        op: UpdateOp,
        expr: Box<Expression>,
    },
    Assign {
        op: AssignOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Ternary {
        test: Box<Expression>,
        consequent: Option<Box<Expression>>,
        alternate: Box<Expression>,
    },
    NullCoalesce {
        left: Box<Expression>,
        right: Box<Expression>,
    },
    /// $arr[$idx]  or  $arr["key"]
    ArrayAccess {
        array: Box<Expression>,
        index: Box<Expression>,
    },
    /// list($a, $b) = expr  or  [$a, $b] = expr  (LHS only)
    List(Vec<Option<Expression>>),
    /// match($x) { val => expr, ... }
    Match {
        subject: Box<Expression>,
        arms: Vec<MatchArm>,
    },
    /// isset($var), empty($var), unset($var[, ...]) — treated as special calls
    Isset(Vec<Expression>),
    Empty(Box<Expression>),
    Unset(Vec<Expression>),
    Cast { cast: CastKind, expr: Box<Expression> },
    /// $this
    This,
    /// static / self / parent keyword used as a class name
    ClassKeyword(String),
    /// Spread: ...$args
    Spread(Box<Expression>),
}

#[derive(Debug, Clone)]
pub struct ArrayElement {
    pub key: Option<Expression>,
    pub value: Expression,
    pub by_ref: bool,
    pub spread: bool,
}

#[derive(Debug, Clone)]
pub struct Argument {
    pub value: Expression,
    pub by_ref: bool,
    pub spread: bool,
    /// PHP 8 named arg: name: value
    pub name: Option<String>,
}

#[derive(Debug, Clone)]
pub enum ClosureBody {
    Block(Vec<Statement>),
    Expr(Expression),
}

#[derive(Debug, Clone)]
pub struct MatchArm {
    /// None = default arm
    pub conditions: Option<Vec<Expression>>,
    pub body: Expression,
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinaryOp {
    Add, Sub, Mul, Div, Mod, Pow,
    Concat,         // .
    Eq, SEq,        // ==, ===
    Ne, SNe,        // !=, !==
    Lt, Gt, Le, Ge,
    Spaceship,      // <=>
    And, Or,        // &&, ||
    BitAnd, BitOr, BitXor,
    Shl, Shr,
    InstanceOf,
}

#[derive(Debug, Clone, PartialEq)]
pub enum UnaryOp {
    Neg, Pos, Not, BitNot,
}

#[derive(Debug, Clone, PartialEq)]
pub enum UpdateOp { Inc, Dec }

#[derive(Debug, Clone, PartialEq)]
pub enum AssignOp {
    Assign,
    AddAssign, SubAssign, MulAssign, DivAssign, ModAssign, PowAssign,
    ConcatAssign,
    AndAssign, OrAssign, XorAssign,
    ShlAssign, ShrAssign,
    NullCoalesceAssign,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CastKind {
    Int, Float, String, Bool, Array, Object,
}
