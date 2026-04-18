/// A complete COBOL program.
#[derive(Debug, Clone)]
pub struct Program {
    pub program_id: String,
    pub author: Option<String>,
    pub data_items: Vec<DataItem>,
    pub file_descriptions: Vec<FileDescription>,
    pub paragraphs: Vec<Paragraph>,
    pub main_body: Vec<Statement>,
    pub classes: Vec<ClassDef>,
    pub interfaces: Vec<InterfaceDef>,
    pub special_names: Vec<SpecialName>,
}

/// FD/SD entry in FILE SECTION
#[derive(Debug, Clone)]
pub struct FileDescription {
    pub name: String,
    pub record_size: Option<u32>,
    pub records: Vec<DataItem>,
    pub is_sort: bool,  // SD vs FD
}

/// SPECIAL-NAMES paragraph entries
#[derive(Debug, Clone)]
pub enum SpecialName {
    DecimalPointIsComma,
    CurrencySign(String),
    ClassName { name: String, values: String },
}

/// COBOL 2023 Class definition (CLASS-ID ... END CLASS)
#[derive(Debug, Clone)]
pub struct ClassDef {
    pub name: String,
    pub inherits: Option<String>,
    pub implements: Vec<String>,
    pub data_items: Vec<DataItem>,
    pub factory_methods: Vec<MethodDef>,
    pub instance_methods: Vec<MethodDef>,
}

/// COBOL 2023 Method definition (METHOD-ID ... END METHOD)
#[derive(Debug, Clone)]
pub struct MethodDef {
    pub name: String,
    pub params: Vec<DataItem>,
    pub returning: Option<DataItem>,
    pub body: Vec<Statement>,
    pub is_property_get: bool,
    pub is_property_set: bool,
}

/// COBOL 2023 Interface definition (INTERFACE-ID ... END INTERFACE)
#[derive(Debug, Clone)]
pub struct InterfaceDef {
    pub name: String,
    pub inherits: Vec<String>,
    pub methods: Vec<MethodSignature>,
}

/// Method signature for interfaces
#[derive(Debug, Clone)]
pub struct MethodSignature {
    pub name: String,
    pub params: Vec<DataItem>,
    pub returning: Option<DataItem>,
}

/// A data item declaration (WORKING-STORAGE, LOCAL-STORAGE, LINKAGE)
#[derive(Debug, Clone)]
pub struct DataItem {
    pub level: u8,
    pub name: String,
    pub pic: Option<String>,
    pub value: Option<Literal>,
    pub occurs: Option<u32>,
    pub occurs_depending: Option<String>,  // OCCURS DEPENDING ON var
    pub redefines: Option<String>,
    pub usage: Option<String>,
    pub is_national: bool,  // USAGE NATIONAL (UTF-8)
    pub is_global: bool,    // GLOBAL clause
    pub is_external: bool,  // EXTERNAL clause
    pub children: Vec<DataItem>,
    pub conditions: Vec<Condition88>,
}

/// An 88-level condition
#[derive(Debug, Clone)]
pub struct Condition88 {
    pub name: String,
    pub values: Vec<Literal>,
    pub thru: Option<(Literal, Literal)>,  // VALUE low THRU high
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
    PerformUntil { test: Expr, body: Vec<Statement>, test_after: bool },
    /// PERFORM VARYING var FROM start BY step UNTIL cond body END-PERFORM
    PerformVarying { var: String, from: Expr, by: Expr, until: Expr, body: Vec<Statement> },
    /// PERFORM paragraph-name
    PerformParagraph(String),
    /// STRING ... INTO dst [WITH POINTER ptr] END-STRING
    StringConcat { sources: Vec<StringSource>, into: String, pointer: Option<String> },
    /// UNSTRING src DELIMITED BY delim INTO dst1 [COUNT c1] [DELIMITER d1] ... [WITH POINTER p]
    Unstring { src: String, delimiters: Vec<String>, into: Vec<UnstringTarget>, pointer: Option<String> },
    /// INSPECT var TALLYING counter FOR ALL/LEADING/FIRST char
    InspectTallying { var: String, counter: String, mode: InspectMode, target: String },
    /// INSPECT var REPLACING ALL/LEADING/FIRST old BY new
    InspectReplacing { var: String, mode: InspectMode, old: String, new: String },
    /// CALL "name" USING args
    Call { name: String, args: Vec<String> },
    /// INITIALIZE var
    Initialize(String),
    /// SET condition TO TRUE/FALSE (properly sets parent to 88's value)
    Set { target: String, value: bool },
    /// ADD CORRESPONDING src TO dst
    AddCorresponding { src: String, dst: String },
    /// SUBTRACT CORRESPONDING src FROM dst
    SubtractCorresponding { src: String, dst: String },
    /// COPY copybook REPLACING ==old== BY ==new== ...
    CopyReplacing { copybook: String, replacements: Vec<(String, String)> },
    /// ACCEPT var FROM COMMAND-LINE
    AcceptCommandLine(String),
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
    /// CALL "program" ASYNC [USING args] → spawn thread
    CallAsync { name: String, args: Vec<String>, handle: Option<String> },
    /// WAIT FOR handle → join thread
    Wait(String),
    /// RUN UNIT "program" → run as separate thread
    RunUnit { name: String, args: Vec<String> },
    /// LOCK monitor-name → acquire lock
    LockMonitor(String),
    /// UNLOCK monitor-name → release lock
    UnlockMonitor(String),
    /// PERFORM paragraph ASYNC → run paragraph as fiber
    PerformAsync(String),
    /// YIELD → suspend current fiber
    YieldStmt,
    /// SUSPEND → pause execution
    SuspendStmt,

    // ── Embedded SQL ───────────────────────────────────────
    /// EXEC SQL CONNECT dsn END-EXEC
    SqlConnect { dsn: String, handle_var: Option<String> },
    /// EXEC SQL SELECT ... INTO :var1, :var2 FROM ... END-EXEC
    SqlSelect { sql: String, into_vars: Vec<String>, from_vars: Vec<String> },
    /// EXEC SQL INSERT/UPDATE/DELETE ... END-EXEC
    SqlExecute { sql: String, host_vars: Vec<String> },
    /// EXEC SQL COMMIT END-EXEC
    SqlCommit,
    /// EXEC SQL ROLLBACK END-EXEC
    SqlRollback,
    /// EXEC SQL DECLARE cursor CURSOR FOR SELECT ... END-EXEC
    SqlDeclareCursor { cursor_name: String, sql: String, host_vars: Vec<String> },
    /// EXEC SQL OPEN cursor END-EXEC
    SqlOpenCursor(String),
    /// EXEC SQL FETCH cursor INTO :var1, :var2 END-EXEC
    SqlFetch { cursor_name: String, into_vars: Vec<String> },
    /// EXEC SQL CLOSE cursor END-EXEC
    SqlCloseCursor(String),

    // ── CICS / DLI ─────────────────────────────────────────
    /// EXEC CICS command END-EXEC
    CicsCommand { command: String, params: Vec<(String, String)> },
    /// EXEC DLI command END-EXEC
    DliCommand { command: String, params: Vec<(String, String)> },

    // ── Arithmetic clauses ─────────────────────────────────
    /// ADD ... ROUNDED
    AddRounded { srcs: Vec<Expr>, to: String, giving: Option<String> },
    /// COMPUTE ... ON SIZE ERROR ... END-COMPUTE
    ComputeWithError {
        dst: String,
        expr: Expr,
        rounded: bool,
        on_error: Vec<Statement>,
        not_on_error: Vec<Statement>,
    },
    /// READ ... AT END ... NOT AT END ... END-READ
    ReadFileAtEnd {
        file: String,
        into: Option<String>,
        at_end: Vec<Statement>,
        not_at_end: Vec<Statement>,
    },

    // ── Nested programs ────────────────────────────────────
    /// Nested program definition
    NestedProgram(Box<Program>),

    // ── Editing PIC display ────────────────────────────────
    /// DISPLAY with PIC editing (format number with PIC mask)
    DisplayFormatted { var: String, pic: String },
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
    CommandLine,
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

/// Target in UNSTRING with optional COUNT and DELIMITER
#[derive(Debug, Clone)]
pub struct UnstringTarget {
    pub name: String,
    pub count: Option<String>,
    pub delimiter: Option<String>,
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
