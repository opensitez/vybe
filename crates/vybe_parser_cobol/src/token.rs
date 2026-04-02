#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,
    pub line: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    // Literals
    Number(f64),
    Str(String),
    Ident(String),       // data-name, paragraph-name

    // Division keywords
    Identification, Data, Procedure, Environment,
    Division, Section,

    // Identification division
    ProgramId, Author, DateWritten,

    // Data division
    WorkingStorage, LocalStorage, FileSection, Linkage,
    Pic, Picture, Value, Occurs, Times, Redefines,
    Indexed, Usage, Binary, Comp, Comp3, Display, Pointer,
    Spaces, Zeros, Zeroes, LowValues, HighValues,

    // Level number (01, 05, 10, 15, 66, 77, 88)
    Level(u8),

    // Procedure division statements
    Move, To, Corresponding, Corr,
    Add, Subtract, From, Giving, Multiply, By, Divide, Remainder,
    Compute,
    DisplayKw, Accept,
    If, Else, EndIf, Then,
    Evaluate, When, Other, EndEvaluate, True, False,
    Perform, EndPerform, Until, Varying, Thru, Through, With, Test, Before, After,
    String_, Delimited, Size, Into, EndString,
    Unstring, EndUnstring,
    Inspect, Tallying, Replacing, All, Leading, First, For, Characters,
    Search, EndSearch, At, EndKey,
    Call, Using, EndCall, Returning,
    GoTo, Go,
    StopRun, Goback,
    Initialize,
    Set, Up, Down,
    Continue,
    Raise, Exception,
    Json, Generate, Parse,
    Xml,
    Not,
    Open, Close, Read, Write, EndRead, EndWrite,
    Input, Output, Extend, IoMode,
    Sort, EndSort, On, Ascending, Descending, Key,
    Release, Return_, EndReturn,
    Select, Assign, FileStatus, Organization, Sequential, Relative, Line,
    Filler, Blank, When_, Zero, Justified, Right, Left,
    Also, Thru88,  // for 88-level VALUE ranges
    Reference, Content,  // CALL BY REFERENCE/CONTENT
    DependingOn,  // OCCURS DEPENDING ON
    Colon,  // : for reference modification
    // Additional verbs
    Rewrite, EndRewrite, Delete, EndDelete, Start, EndStart,
    Exit, Paragraph, Merge,
    Copy, Converting,
    Numeric, Alphabetic, AlphabeticLower, AlphabeticUpper,
    Positive, Negative,
    PointerKw, Count,
    // COBOL 2023
    ClassId, MethodId, Invoke, EndClass, EndMethod,
    Typedef, Validate, EndValidate,
    Free, Allocate,
    Boolean, FloatLong, FloatShort, National,
    Property, Get, Set2, EndInvoke,
    Resume,
    Inherits, Implements, InterfaceId, EndInterface,
    Factory, Object_, EndFactory, EndObject,
    New, Self_,
    Override,

    // Intrinsic functions
    Function,
    Length, UpperCase, LowerCase, Trim, Reverse,
    CurrentDate, Max, Min, Mod, Rem,
    Numval, NumvalC, Ord, Char,
    Substitute, Sqrt, Sum, Integer,
    Abs, Log, Log10, Exp, Sin, Cos, Tan,
    Asin, Acos, Atan, Ceiling, Floor, Sign, Power,
    Random, Mean, Median, Variance,
    DateOfInteger, IntegerOfDate, DayOfInteger,
    Annuity, PresentValue,
    Concatenate, FormattedDate, FormattedTime,
    TestNumval, WhenCompiled,

    // Operators
    Plus,        // +
    Minus,       // -
    Star,        // *
    Slash,       // /
    StarStar,    // **
    Eq,          // =
    Gt,          // >
    Lt,          // <
    GtEq,        // >=
    LtEq,        // <=
    NotEq,       // NOT =

    // Logical
    And, Or,

    // Delimiters
    LParen,
    RParen,
    Period,      // .  (statement terminator)
    Comma,

    // Special
    Of, In,      // qualified names: X OF Y

    Eof,
}
