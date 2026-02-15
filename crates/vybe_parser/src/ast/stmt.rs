use super::{Expression, Identifier, VBType};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Statement {
    // Local declaration
    Dim(super::VariableDecl),

    // Constant declaration
    Const(super::ConstDecl),

    // Variable assignment
    Assignment {
        target: Identifier,
        value: Expression,
    },

    // Object assignment (Set keyword)
    SetAssignment {
        target: Identifier,
        value: Expression,
    },

    // Member assignment (e.g., obj.prop = value)
    MemberAssignment {
        object: Expression,
        member: Identifier,
        value: Expression,
    },

    // Array assignment (e.g., arr(5) = value)
    ArrayAssignment {
        array: Identifier,
        indices: Vec<Expression>,
        value: Expression,
    },

    // Array redimensioning
    ReDim {
        preserve: bool,
        array: Identifier,
        bounds: Vec<Expression>,
    },

    // Compound assignment (+=, -=, etc.)
    CompoundAssignment {
        target: Identifier,
        members: Vec<Identifier>,
        indices: Vec<Expression>,
        operator: CompoundOp,
        value: Expression,
    },

    // RaiseEvent
    RaiseEvent {
        event_name: Identifier,
        arguments: Vec<Expression>,
    },

    // Control flow
    If {
        condition: Expression,
        then_branch: Vec<Statement>,
        elseif_branches: Vec<(Expression, Vec<Statement>)>,
        else_branch: Option<Vec<Statement>>,
    },

    For {
        variable: Identifier,
        start: Expression,
        end: Expression,
        step: Option<Expression>,
        body: Vec<Statement>,
    },

    While {
        condition: Expression,
        body: Vec<Statement>,
    },

    DoLoop {
        pre_condition: Option<(LoopConditionType, Expression)>,
        body: Vec<Statement>,
        post_condition: Option<(LoopConditionType, Expression)>,
    },

    Select {
        test_expr: Expression,
        cases: Vec<CaseBlock>,
        else_block: Option<Vec<Statement>>,
    },

    // For Each
    ForEach {
        variable: Identifier,
        collection: Expression,
        body: Vec<Statement>,
    },

    // With block
    With {
        object: Expression,
        body: Vec<Statement>,
    },

    // Using block (resource disposal)
    Using {
        variable: Identifier,
        resource: Expression,
        body: Vec<Statement>,
    },

    // Exit statements
    ExitSub,
    ExitFunction,
    ExitFor,
    ExitDo,
    ExitWhile,
    ExitSelect,
    ExitTry,
    ExitProperty,

    // Return
    Return(Option<Expression>),

    // Procedure call
    Call {
        name: Identifier,
        arguments: Vec<Expression>,
    },

    // Expression statement (for calls with side effects)
    ExpressionStatement(Expression),

    // Exception handling
    Try {
        body: Vec<Statement>,
        catches: Vec<CatchBlock>,
        finally: Option<Vec<Statement>>,
    },

    // Throw exception
    Throw(Option<Expression>),

    // AddHandler obj.Event, AddressOf handler
    AddHandler {
        event_target: String,  // e.g. "Button1.Click"
        handler: String,       // e.g. "HandleClick" or "Me.HandleClick"
    },

    // RemoveHandler obj.Event, AddressOf handler
    RemoveHandler {
        event_target: String,
        handler: String,
    },

    // Continue
    Continue(ContinueType),

    /// `Static x As Integer = 0` — local but persists across calls
    StaticVar {
        name: Identifier,
        var_type: Option<VBType>,
        initializer: Option<Expression>,
    },

    // GoTo / Labels
    GoTo(String),
    Label(String),

    // VB6/VB.NET error handling
    OnErrorResumeNext,
    OnErrorGoTo(String),    // Label name, or "0" to disable
    Resume(ResumeTarget),

    // VB6 File I/O
    Open {
        file_path: Expression,
        mode: FileOpenMode,
        file_number: Expression,
    },
    CloseFile {
        file_number: Option<Expression>,
    },
    PrintFile {
        file_number: Expression,
        items: Vec<Expression>,
        newline: bool,
    },
    WriteFile {
        file_number: Expression,
        items: Vec<Expression>,
    },
    InputFile {
        file_number: Expression,
        variables: Vec<Identifier>,
    },
    LineInput {
        file_number: Expression,
        variable: Identifier,
    },

    // SyncLock
    SyncLock {
        lock_object: Expression,
        body: Vec<Statement>,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum LoopConditionType {
    While,
    Until,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CaseBlock {
    pub conditions: Vec<CaseCondition>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum CaseCondition {
    Value(Expression),
    Range { from: Expression, to: Expression },
    Comparison { op: CompOp, expr: Expression },
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CatchBlock {
    pub variable: Option<(Identifier, Option<super::VBType>)>,
    pub when_clause: Option<Expression>,
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum ContinueType {
    Do,
    For,
    While,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum CompOp {
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum FileOpenMode {
    Input,
    Output,
    Append,
    Binary,
    Random,
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum CompoundOp {
    AddAssign,       // +=
    SubtractAssign,  // -=
    MultiplyAssign,  // *=
    DivideAssign,    // /=
    IntDivideAssign, // \=
    ConcatAssign,    // &=
    ExponentAssign,  // ^=
    ShiftLeftAssign, // <<=
    ShiftRightAssign,// >>=
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ResumeTarget {
    /// Resume (retry the statement that caused the error)
    Implicit,
    /// Resume Next (continue from the next statement)
    Next,
    /// Resume <label> (jump to the label)
    Label(String),
}