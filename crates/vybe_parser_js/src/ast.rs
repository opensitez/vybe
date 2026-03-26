/// A complete JavaScript program.
#[derive(Debug, Clone)]
pub struct Program {
    pub body: Vec<Statement>,
}

/// Statement nodes.
#[derive(Debug, Clone)]
pub enum Statement {
    /// Expression followed by semicolon
    Expression(Expression),
    /// { ... }
    Block(Vec<Statement>),
    /// var/let/const declarations
    VariableDeclaration {
        kind: VarKind,
        declarations: Vec<VarDeclarator>,
    },
    /// function name(...) { ... }
    FunctionDeclaration(FunctionDecl),
    /// class Name { ... }
    ClassDeclaration(ClassDecl),
    /// if (test) consequent else alternate
    If {
        test: Expression,
        consequent: Box<Statement>,
        alternate: Option<Box<Statement>>,
    },
    /// for (init; test; update) body
    For {
        init: Option<ForInit>,
        test: Option<Expression>,
        update: Option<Expression>,
        body: Box<Statement>,
    },
    /// for (left in right) body
    ForIn {
        left: ForInTarget,
        right: Expression,
        body: Box<Statement>,
    },
    /// for (left of right) body
    ForOf {
        left: ForInTarget,
        right: Expression,
        body: Box<Statement>,
    },
    /// while (test) body
    While {
        test: Expression,
        body: Box<Statement>,
    },
    /// do body while (test)
    DoWhile {
        body: Box<Statement>,
        test: Expression,
    },
    /// switch (discriminant) { cases }
    Switch {
        discriminant: Expression,
        cases: Vec<SwitchCase>,
    },
    /// return [expr]
    Return(Option<Expression>),
    /// break [label]
    Break(Option<String>),
    /// continue [label]
    Continue(Option<String>),
    /// throw expr
    Throw(Expression),
    /// try { } catch { } finally { }
    Try {
        block: Vec<Statement>,
        handler: Option<CatchClause>,
        finalizer: Option<Vec<Statement>>,
    },
    /// label: statement
    Labeled {
        label: String,
        body: Box<Statement>,
    },
    /// Empty statement (;)
    Empty,

    // -- Modules --

    /// import { a, b } from "module"
    /// import * as name from "module"
    /// import name from "module"
    Import {
        specifiers: Vec<ImportSpecifier>,
        source: String,
    },
    /// export function foo() {}
    /// export let x = 5
    /// export { a, b }
    /// export default expr
    Export {
        declaration: Option<Box<Statement>>,
        specifiers: Vec<ExportSpecifier>,
        default: Option<Box<Expression>>,
    },
}

/// import { foo as bar } from "mod"
#[derive(Debug, Clone)]
pub enum ImportSpecifier {
    /// import { name } or import { name as alias }
    Named { name: String, alias: Option<String> },
    /// import * as name
    Namespace(String),
    /// import defaultName from "mod"
    Default(String),
}

/// export { foo as bar }
#[derive(Debug, Clone)]
pub struct ExportSpecifier {
    pub name: String,
    pub alias: Option<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VarKind {
    Var,
    Let,
    Const,
}

/// A binding pattern — simple name or destructuring.
#[derive(Debug, Clone)]
pub enum BindingPattern {
    Identifier(String),
    Object(Vec<ObjectPatternProp>),
    Array(Vec<ArrayPatternElem>),
}

#[derive(Debug, Clone)]
pub struct ObjectPatternProp {
    pub key: String,
    pub value: Option<BindingPattern>,  // None = shorthand { x }
    pub default: Option<Expression>,
}

#[derive(Debug, Clone)]
pub enum ArrayPatternElem {
    Pattern(BindingPattern, Option<Expression>),  // element with optional default
    Rest(String),                                  // ...rest
    Hole,                                          // empty slot in [,x]
}

#[derive(Debug, Clone)]
pub struct VarDeclarator {
    pub pattern: BindingPattern,
    pub init: Option<Expression>,
}

impl VarDeclarator {
    /// Convenience for simple `let x = ...` declarations.
    pub fn simple(name: String, init: Option<Expression>) -> Self {
        VarDeclarator { pattern: BindingPattern::Identifier(name), init }
    }
}

#[derive(Debug, Clone)]
pub enum ForInit {
    VarDecl(VarKind, Vec<VarDeclarator>),
    Expression(Expression),
}

#[derive(Debug, Clone)]
pub enum ForInTarget {
    VarDecl(VarKind, String),
    Identifier(String),
}

#[derive(Debug, Clone)]
pub struct SwitchCase {
    pub test: Option<Expression>,
    pub consequent: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct CatchClause {
    pub param: Option<String>,
    pub body: Vec<Statement>,
}

/// Expression nodes.
#[derive(Debug, Clone)]
pub enum Expression {
    // Literals
    Number(f64),
    String(String),
    Boolean(bool),
    Null,
    Undefined,
    Array(Vec<Expression>),
    Object(Vec<PropertyDef>),
    TemplateLiteral {
        quasis: Vec<String>,
        expressions: Vec<Expression>,
    },

    // Identifiers
    Identifier(String),
    This,
    Super,

    // Operations
    Binary {
        op: BinaryOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Logical {
        op: LogicalOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    Unary {
        op: UnaryOp,
        argument: Box<Expression>,
    },
    Update {
        op: UpdateOp,
        prefix: bool,
        argument: Box<Expression>,
    },
    Assignment {
        op: AssignOp,
        left: Box<Expression>,
        right: Box<Expression>,
    },
    /// cond ? then : else
    Conditional {
        test: Box<Expression>,
        consequent: Box<Expression>,
        alternate: Box<Expression>,
    },

    // Access
    Member {
        object: Box<Expression>,
        property: String,
        optional: bool,  // true for ?.
    },
    ComputedMember {
        object: Box<Expression>,
        property: Box<Expression>,
    },
    /// function call
    Call {
        callee: Box<Expression>,
        arguments: Vec<Expression>,
        optional: bool,  // true for ?.()
    },
    /// new Foo(...)
    New {
        callee: Box<Expression>,
        arguments: Vec<Expression>,
    },

    // Functions as expressions
    Function(FunctionDecl),
    ArrowFunction {
        params: Vec<Param>,
        body: ArrowBody,
        is_async: bool,
    },

    // Operators
    Typeof(Box<Expression>),
    Void(Box<Expression>),
    Delete(Box<Expression>),

    // Spread
    Spread(Box<Expression>),

    // Async
    Await(Box<Expression>),

    // Comma
    Sequence(Vec<Expression>),
}

#[derive(Debug, Clone)]
pub enum ArrowBody {
    Expression(Box<Expression>),
    Block(Vec<Statement>),
}

// -- Shared structures --

/// A function parameter — name with optional default value and rest marker.
#[derive(Debug, Clone)]
pub struct Param {
    pub name: String,
    pub default: Option<Expression>,
    pub rest: bool,
}

impl Param {
    pub fn simple(name: impl Into<String>) -> Self {
        Param { name: name.into(), default: None, rest: false }
    }
}

#[derive(Debug, Clone)]
pub struct FunctionDecl {
    pub name: Option<String>,
    pub params: Vec<Param>,
    pub body: Vec<Statement>,
    pub is_async: bool,
}

#[derive(Debug, Clone)]
pub struct ClassDecl {
    pub name: Option<String>,
    pub super_class: Option<Box<Expression>>,
    pub body: Vec<ClassMember>,
}

#[derive(Debug, Clone)]
pub enum ClassMember {
    Method {
        key: String,
        value: FunctionDecl,
        kind: MethodKind,
        is_static: bool,
    },
    Property {
        key: String,
        value: Option<Expression>,
        is_static: bool,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MethodKind {
    Method,
    Get,
    Set,
    Constructor,
}

#[derive(Debug, Clone)]
pub enum PropertyDef {
    KeyValue {
        key: String,
        value: Expression,
    },
    Computed {
        key: Expression,
        value: Expression,
    },
    Shorthand(String),
    Spread(Expression),
    Method {
        key: String,
        value: FunctionDecl,
    },
}

// -- Operator enums --

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOp {
    Add, Sub, Mul, Div, Mod, Exp,
    BitAnd, BitOr, BitXor, Shl, Shr, UShr,
    Eq, Neq, SEq, SNeq,
    Lt, Gt, Le, Ge,
    In, InstanceOf,
    NullishCoalescing,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum LogicalOp {
    And, Or,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnaryOp {
    Neg, Pos, Not, BitNot,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UpdateOp {
    Increment, Decrement,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AssignOp {
    Assign,
    AddAssign, SubAssign, MulAssign, DivAssign, ModAssign,
    BitAndAssign, BitOrAssign, BitXorAssign,
    ShlAssign, ShrAssign, UShrAssign,
    ExpAssign,
    AndAssign, OrAssign, NullishAssign,
}
