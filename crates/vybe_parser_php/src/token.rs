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
    // $name
    Variable(String),
    // bare word (function name, class name, etc.)
    Identifier(String),

    // Keywords
    If,
    ElseIf,
    Else,
    While,
    For,
    ForEach,
    As,
    Do,
    Switch,
    Case,
    Default,
    Break,
    Continue,
    Return,
    Function,
    Class,
    Extends,
    Implements,
    Interface,
    Trait,
    New,
    Echo,
    Print,
    Null,
    True,
    False,
    InstanceOf,
    Throw,
    Try,
    Catch,
    Finally,
    Static,
    Public,
    Private,
    Protected,
    Abstract,
    Final,
    Const,
    Match,
    Fn,
    Use,
    Namespace,
    Yield,
    List,
    Global,
    Enum,
    Readonly,
    Clone,

    // Arithmetic / concat
    Plus,        // +
    Minus,       // -
    Star,        // *
    Slash,       // /
    Percent,     // %
    StarStar,    // **
    Dot,         // .  (concat)

    // Assignment
    Eq,          // =
    PlusEq,      // +=
    MinusEq,     // -=
    StarEq,      // *=
    SlashEq,     // /=
    PercentEq,   // %=
    DotEq,       // .=
    AmpAmpEq,    // &&=
    PipePipeEq,  // ||=
    QuestionQuestionEq, // ??=

    // Inc/Dec
    PlusPlus,    // ++
    MinusMinus,  // --

    // Comparison
    EqEq,        // ==
    EqEqEq,      // ===
    BangEq,      // !=
    BangEqEq,    // !==
    Lt,          // <
    Gt,          // >
    LtEq,        // <=
    GtEq,        // >=
    Spaceship,   // <=>

    // Logical
    AmpAmp,      // &&
    PipePipe,    // ||
    Bang,        // !

    // Bitwise
    Amp,         // &
    Pipe,        // |
    Caret,       // ^
    Tilde,       // ~
    LtLt,        // <<
    GtGt,        // >>

    // Member access
    Arrow,       // ->
    NullsafeArrow, // ?->
    ColonColon,  // ::

    // Misc
    Question,    // ?
    QuestionQuestion, // ??
    Colon,       // :
    FatArrow,    // =>
    Ellipsis,    // ...
    At,          // @ (suppress errors — we just ignore it)
    Backslash,   // \

    // Delimiters
    LParen,
    RParen,
    LBrace,
    RBrace,
    LBracket,
    RBracket,
    Semicolon,
    Comma,

    Eof,
}
