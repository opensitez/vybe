/// Pascal AST for the vybe bytecode compiler.

#[derive(Debug, Clone)]
pub struct Program {
    pub name: String,
    pub uses: Vec<String>,
    pub decls: Vec<Decl>,
    pub body: Vec<Statement>,
}

// ── Declarations ──────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Decl {
    Var(Vec<VarDecl>),
    Const(Vec<ConstDecl>),
    Type(Vec<TypeDecl>),
    Procedure(ProcDecl),
    Function(FuncDecl),
    /// Method implementation: `constructor TFoo.Create(...); begin ... end;`
    Method(MethodImpl),
}

#[derive(Debug, Clone)]
pub struct VarDecl {
    pub names: Vec<String>,
    pub type_name: TypeRef,
    pub init: Option<Expression>,
}

#[derive(Debug, Clone)]
pub struct ConstDecl {
    pub name: String,
    pub type_name: Option<TypeRef>,
    pub value: Expression,
}

#[derive(Debug, Clone)]
pub struct TypeDecl {
    pub name: String,
    pub def: TypeDef,
}

#[derive(Debug, Clone)]
pub enum TypeDef {
    Alias(TypeRef),
    Record(RecordDef),
    Enum(Vec<EnumValue>),
    Array { index: Option<(Expression, Expression)>, element: Box<TypeRef> },
    Pointer(Box<TypeRef>),
    /// Object Pascal class declaration: `TFoo = class(TParent) ... end;`
    Class(ClassDef),
    /// Interface declaration: `IFoo = interface ... end;`
    InterfaceDef(InterfaceDecl),
}

/// Enum value: `Red` or `Red = 0`
#[derive(Debug, Clone)]
pub struct EnumValue {
    pub name: String,
    pub value: Option<Expression>,
}

/// Record definition — can have fields and methods (advanced records)
#[derive(Debug, Clone)]
pub struct RecordDef {
    pub fields: Vec<VarDecl>,
    pub methods: Vec<MethodSig>,
}

/// Object Pascal class definition.
#[derive(Debug, Clone)]
pub struct ClassDef {
    pub parent: Option<String>,
    pub interfaces: Vec<String>,
    pub members: Vec<ClassMember>,
}

/// A member inside a class declaration.
#[derive(Debug, Clone)]
pub enum ClassMember {
    Field(VarDecl),
    MethodDecl(MethodSig),
    PropertyDecl(PropertyDef),
    ClassVar(VarDecl),
}

/// Method signature declared inside the class body (implementation is separate).
#[derive(Debug, Clone)]
pub struct MethodSig {
    pub kind: MethodKind,
    pub name: String,
    pub params: Vec<Param>,
    pub return_type: Option<TypeRef>,
    pub directives: Vec<MethodDirective>,
    pub is_static: bool,
    pub is_operator: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum MethodKind { Constructor, Destructor, Procedure, Function }

#[derive(Debug, Clone, PartialEq)]
pub enum MethodDirective { Virtual, Override, Abstract, Reintroduce }

/// Property definition: `property Name: Type read Getter write Setter;`
#[derive(Debug, Clone)]
pub struct PropertyDef {
    pub name: String,
    pub type_name: TypeRef,
    pub reader: Option<String>,
    pub writer: Option<String>,
    pub default: bool,
    pub index_type: Option<TypeRef>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Visibility { Private, Protected, Public, Published }

/// Interface declaration
#[derive(Debug, Clone)]
pub struct InterfaceDecl {
    pub parent: Option<String>,
    pub methods: Vec<MethodSig>,
    pub properties: Vec<PropertyDef>,
}

/// Method implementation: `constructor TFoo.Create(...); begin ... end;`
#[derive(Debug, Clone)]
pub struct MethodImpl {
    pub kind: MethodKind,
    pub class_name: String,
    pub method_name: String,
    pub params: Vec<Param>,
    pub return_type: Option<TypeRef>,
    pub decls: Vec<Decl>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct TypeRef {
    pub name: String,
    pub generic: Option<Box<TypeRef>>,
}

impl TypeRef {
    pub fn simple(name: &str) -> Self { TypeRef { name: name.to_string(), generic: None } }
}

#[derive(Debug, Clone)]
pub struct ProcDecl {
    pub name: String,
    pub params: Vec<Param>,
    pub decls: Vec<Decl>,
    pub body: Vec<Statement>,
    pub is_forward: bool,
}

#[derive(Debug, Clone)]
pub struct FuncDecl {
    pub name: String,
    pub params: Vec<Param>,
    pub return_type: TypeRef,
    pub decls: Vec<Decl>,
    pub body: Vec<Statement>,
    pub is_forward: bool,
}

#[derive(Debug, Clone)]
pub struct Param {
    pub names: Vec<String>,
    pub type_name: TypeRef,
    pub pass_by: PassBy,
    pub default: Option<Expression>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PassBy { Value, Var, Const, Out }

// ── Statements ────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Statement {
    Assign { target: Expression, value: Expression },
    /// Compound assignment: target += value, target -= value, etc.
    CompoundAssign { target: Expression, op: CompoundOp, value: Expression },
    Call { name: Expression, args: Vec<Expression> },
    Block(Vec<Statement>),
    If { cond: Expression, then: Box<Statement>, else_: Option<Box<Statement>> },
    For { var: String, from: Expression, to: Expression, downto: bool, body: Box<Statement> },
    /// for item in collection do body
    ForIn { var: String, collection: Expression, body: Box<Statement> },
    While { cond: Expression, body: Box<Statement> },
    Repeat { body: Vec<Statement>, until: Expression },
    Case { expr: Expression, arms: Vec<CaseArm>, else_: Option<Vec<Statement>> },
    With { vars: Vec<Expression>, body: Box<Statement> },
    Try { body: Vec<Statement>, handler: TryHandler },
    Raise(Option<Expression>),
    Exit(Option<Expression>),
    Break,
    Continue,
    Empty,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CompoundOp { Add, Sub, Mul, Div }

#[derive(Debug, Clone)]
pub struct CaseArm {
    pub values: Vec<CaseValue>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub enum CaseValue {
    Single(Expression),
    Range(Expression, Expression),
}

#[derive(Debug, Clone)]
pub enum TryHandler {
    Except(Vec<ExceptClause>, Option<Vec<Statement>>),
    Finally(Vec<Statement>),
}

#[derive(Debug, Clone)]
pub struct ExceptClause {
    pub on_type: Option<String>,
    pub var_name: Option<String>,
    pub body: Vec<Statement>,
}

// ── Expressions ───────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
pub enum Expression {
    Int(i64),
    Real(f64),
    Bool(bool),
    Nil,
    Str(String),
    Char(char),
    Identifier(String),
    Binary { op: BinOp, left: Box<Expression>, right: Box<Expression> },
    Unary { op: UnaryOp, expr: Box<Expression> },
    Call { callee: Box<Expression>, args: Vec<Expression> },
    Index { array: Box<Expression>, index: Box<Expression> },
    Field { record: Box<Expression>, field: String },
    Deref(Box<Expression>),
    AddrOf(Box<Expression>),
    Cast { type_name: String, expr: Box<Expression> },
    SetLiteral(Vec<Expression>),
    ArrayLiteral(Vec<Expression>),
    /// `inherited Create(args)` or bare `inherited`
    Inherited { method: Option<String>, args: Vec<Expression> },
    /// Anonymous procedure/function: `procedure(x: Integer) begin ... end`
    Lambda { params: Vec<Param>, return_type: Option<TypeRef>, body: Vec<Statement> },
    /// `obj is TClassName`
    IsCheck { expr: Box<Expression>, type_name: String },
    /// `obj as TClassName`
    AsCast { expr: Box<Expression>, type_name: String },
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum BinOp {
    Add, Sub, Mul, Div, IDiv, Mod,
    Eq, NotEq, Lt, Gt, Le, Ge,
    And, Or, Xor, Shl, Shr,
    In,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum UnaryOp { Neg, Not, Deref, AddrOf }
