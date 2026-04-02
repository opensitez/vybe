/// A complete COBOL program.
#[derive(Debug, Clone)]
pub struct Program {
    pub program_id: String,
    pub author: Option<String>,
    pub data_items: Vec<DataItem>,
    pub paragraphs: Vec<Paragraph>,
    pub main_body: Vec<Statement>,
}

/// A data item declaration (WORKING-STORAGE, LOCAL-STORAGE, LINKAGE)
#[derive(Debug, Clone)]
pub struct DataItem {
    pub level: u8,
    pub name: String,
    pub pic: Option<String>,
    pub value: Option<Literal>,
    pub occurs: Option<u32>,
    pub redefines: Option<String>,
    pub usage: Option<String>,
    pub children: Vec<DataItem>,
    pub conditions: Vec<Condition88>,
}

/// An 88-level condition
#[derive(Debug, Clone)]
pub struct Condition88 {
    pub name: String,
    pub values: Vec<Literal>,
}

/// Literal value
#[derive(Debug, Clone)]
pub enum Literal {
    Num(f64),
    Str(String),
    Spaces,
    Zeros,
    LowValues,
    HighValues,
    True,
    False,
}

/// A paragraph (section of code)
#[derive(Debug, Clone)]
pub struct Paragraph {
    pub name: String,
    pub body: Vec<Statement>,
}

/// COBOL statements
#[derive(Debug, Clone)]
pub enum Statement {
    /// DISPLAY expr [expr ...]
    Display(Vec<Expr>),
    /// ACCEPT var
    Accept(String),
    /// MOVE src TO dst [dst ...]
    Move { src: Expr, dsts: Vec<String> },
    /// MOVE CORRESPONDING src TO dst
    MoveCorresponding { src: String, dst: String },
    /// ADD a [a ...] TO b [GIVING c]
    Add { srcs: Vec<Expr>, to: String, giving: Option<String> },
    /// SUBTRACT a FROM b [GIVING c]
    Subtract { src: Expr, from: String, giving: Option<String> },
    /// MULTIPLY a BY b [GIVING c]
    Multiply { src: Expr, by: String, giving: Option<String> },
    /// DIVIDE a BY b GIVING c [REMAINDER r]
    Divide { src: Expr, by: Expr, giving: String, remainder: Option<String> },
    /// COMPUTE dst = expr
    Compute { dst: String, expr: Expr },
    /// IF cond THEN body [ELSE body] END-IF
    If { test: Expr, body: Vec<Statement>, else_body: Option<Vec<Statement>> },
    /// EVALUATE subject WHEN ... END-EVALUATE
    Evaluate { subject: Expr, whens: Vec<WhenClause>, other: Option<Vec<Statement>> },
    /// PERFORM n TIMES body END-PERFORM
    PerformTimes { count: Expr, body: Vec<Statement> },
    /// PERFORM UNTIL cond body END-PERFORM
    PerformUntil { test: Expr, body: Vec<Statement> },
    /// PERFORM VARYING var FROM start BY step UNTIL cond body END-PERFORM
    PerformVarying { var: String, from: Expr, by: Expr, until: Expr, body: Vec<Statement> },
    /// PERFORM paragraph-name
    PerformParagraph(String),
    /// STRING ... INTO dst END-STRING
    StringConcat { sources: Vec<StringSource>, into: String },
    /// UNSTRING src DELIMITED BY delim INTO dst1 dst2 ... END-UNSTRING
    Unstring { src: String, delimiters: Vec<String>, into: Vec<String> },
    /// INSPECT var TALLYING counter FOR ALL/LEADING/FIRST char
    InspectTallying { var: String, counter: String, mode: InspectMode, target: String },
    /// INSPECT var REPLACING ALL/LEADING/FIRST old BY new
    InspectReplacing { var: String, mode: InspectMode, old: String, new: String },
    /// CALL "name" USING args
    Call { name: String, args: Vec<String> },
    /// INITIALIZE var
    Initialize(String),
    /// SET condition TO TRUE/FALSE
    Set { target: String, value: bool },
    /// STOP RUN
    StopRun,
    /// GOBACK
    Goback,
    /// CONTINUE
    Continue,
    /// GO TO paragraph
    GoTo(String),
    /// RAISE EXCEPTION msg
    Raise(Expr),
    /// JSON GENERATE dst FROM src
    JsonGenerate { dst: String, src: String },
    /// JSON PARSE src INTO dst
    JsonParse { src: String, dst: String },
    /// OPEN INPUT/OUTPUT/EXTEND file
    Open { mode: FileMode, file: String },
    /// CLOSE file
    Close(String),
    /// READ file INTO var
    ReadFile { file: String, into: Option<String> },
    /// WRITE record FROM var
    WriteFile { record: String, from: Option<String> },
    /// SORT file ON ASCENDING/DESCENDING KEY field
    Sort { file: String, ascending: bool, key: String },
    /// PERFORM para1 THRU para2
    PerformThru { from: String, thru: String },
    /// SEARCH table AT END stmts WHEN cond stmts END-SEARCH
    SearchTable { table: String, at_end: Vec<Statement>, when_cond: Expr, when_body: Vec<Statement> },
    /// ACCEPT var FROM DATE/TIME/DAY
    AcceptFrom { var: String, source: AcceptSource },
    /// REWRITE record FROM var
    Rewrite { record: String, from: Option<String> },
    /// DELETE file
    DeleteFile(String),
    /// START file KEY = var
    StartFile { file: String, key: Option<String> },
    /// EXIT PERFORM / EXIT PARAGRAPH
    ExitPerform,
    ExitParagraph,
    /// MERGE file ON ASCENDING/DESCENDING KEY field
    Merge { file: String, ascending: bool, key: String },
    /// COPY copybook
    Copy(String),
    /// INSPECT var CONVERTING from-chars TO to-chars
    InspectConverting { var: String, from: String, to: String },
    /// EVALUATE subject ALSO subject2 ...
    EvaluateAlso { subjects: Vec<Expr>, whens: Vec<WhenAlsoClause>, other: Option<Vec<Statement>> },
    /// INVOKE object method USING args [RETURNING result]
    Invoke { object: String, method: String, args: Vec<String>, returning: Option<String> },
    /// TYPEDEF type-name AS ...
    TypeDef { name: String, pic: Option<String> },
    /// VALIDATE var
    ValidateStmt(String),
    /// FREE var
    FreeStmt(String),
    /// ALLOCATE var
    AllocateStmt(String),
}

#[derive(Debug, Clone)]
pub struct WhenAlsoClause {
    pub values: Vec<Vec<Expr>>,  // one value per subject
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FileMode {
    Input,
    Output,
    IoMode,
    Extend,
}

#[derive(Debug, Clone, PartialEq)]
pub enum AcceptSource {
    Console,
    Date,
    Time,
    Day,
    DayOfWeek,
}

#[derive(Debug, Clone)]
pub struct WhenClause {
    pub values: Vec<Expr>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone)]
pub struct StringSource {
    pub value: Expr,
    pub delimited_by: DelimitedBy,
}

#[derive(Debug, Clone)]
pub enum DelimitedBy {
    Size,
    Value(String),
}

#[derive(Debug, Clone, PartialEq)]
pub enum InspectMode {
    All,
    Leading,
    First,
}

/// COBOL expressions (for COMPUTE, conditions, FUNCTION calls)
#[derive(Debug, Clone)]
pub enum Expr {
    Lit(Literal),
    Ident(String),
    /// Subscripted: WS-ITEM(1)
    Subscript(String, Box<Expr>),
    /// Qualified: X OF Y
    Qualified(String, String),
    /// Binary operation
    BinOp { op: BinOp, left: Box<Expr>, right: Box<Expr> },
    /// Unary NOT
    Not(Box<Expr>),
    /// Comparison: a > b, a = b, etc.
    Compare { op: CmpOp, left: Box<Expr>, right: Box<Expr> },
    /// Logical AND/OR
    Logic { op: LogicOp, left: Box<Expr>, right: Box<Expr> },
    /// FUNCTION call
    FunctionCall { name: String, args: Vec<Expr> },
    /// TRUE / FALSE
    Bool(bool),
    /// Reference modification: name(start:length)
    RefMod { name: String, start: Box<Expr>, length: Option<Box<Expr>> },
    /// Class condition: var IS NUMERIC / ALPHABETIC
    ClassTest { var: Box<Expr>, class: ClassCondition },
    /// Sign condition: var IS POSITIVE / NEGATIVE / ZERO
    SignTest { var: Box<Expr>, sign: SignCondition },
}

#[derive(Debug, Clone, PartialEq)]
pub enum ClassCondition {
    Numeric,
    Alphabetic,
    AlphabeticLower,
    AlphabeticUpper,
}

#[derive(Debug, Clone, PartialEq)]
pub enum SignCondition {
    Positive,
    Negative,
    Zero,
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinOp { Add, Sub, Mul, Div, Pow }

#[derive(Debug, Clone, PartialEq)]
pub enum CmpOp { Eq, Ne, Lt, Gt, Le, Ge }

#[derive(Debug, Clone, PartialEq)]
pub enum LogicOp { And, Or }
