#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Span {
    pub start: u32,
    pub end: u32,
    pub line: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    // Literals
    Number(f64),
    String(String),
    TemplateLiteral(String),       // Full template string (no interpolation)
    TemplateHead(String),          // `text${
    TemplateMiddle(String),        // }text${
    TemplateTail(String),          // }text`
    RegExp(String, String),        // pattern, flags

    // Identifier
    Identifier(String),

    // Keywords
    Var,
    Let,
    Const,
    Function,
    Return,
    If,
    Else,
    For,
    While,
    Do,
    Break,
    Continue,
    Switch,
    Case,
    Default,
    New,
    Delete,
    Typeof,
    Void,
    Instanceof,
    In,
    Of,
    This,
    Super,
    Class,
    Extends,
    Static,
    Get,
    Set,
    Throw,
    Try,
    Catch,
    Finally,
    True,
    False,
    Null,
    Undefined,
    Import,
    Export,
    From,
    As,
    Async,
    Await,
    Yield,
    Debugger,

    // Punctuation
    LParen,     // (
    RParen,     // )
    LBrace,     // {
    RBrace,     // }
    LBracket,   // [
    RBracket,   // ]
    Semicolon,  // ;
    Comma,      // ,
    Dot,        // .
    DotDotDot,  // ...
    Colon,      // :
    Question,   // ?
    QuestionDot, // ?.
    QuestionQuestion, // ??
    Arrow,      // =>

    // Operators
    Plus,       // +
    Minus,      // -
    Star,       // *
    Slash,      // /
    Percent,    // %
    StarStar,   // **
    Amp,        // &
    Pipe,       // |
    Caret,      // ^
    Tilde,      // ~
    LtLt,       // <<
    GtGt,       // >>
    GtGtGt,     // >>>
    Bang,       // !
    Lt,         // <
    Gt,         // >
    LtEq,      // <=
    GtEq,      // >=
    EqEq,       // ==
    BangEq,     // !=
    EqEqEq,     // ===
    BangEqEq,   // !==
    AmpAmp,     // &&
    PipePipe,   // ||
    Eq,         // =
    PlusEq,     // +=
    MinusEq,    // -=
    StarEq,     // *=
    SlashEq,    // /=
    PercentEq,  // %=
    AmpEq,      // &=
    PipeEq,     // |=
    CaretEq,    // ^=
    LtLtEq,     // <<=
    GtGtEq,     // >>=
    GtGtGtEq,   // >>>=
    StarStarEq,  // **=
    AmpAmpEq,   // &&=
    PipePipeEq,  // ||=
    QuestionQuestionEq, // ??=
    PlusPlus,    // ++
    MinusMinus,  // --

    // Special
    Eof,
}

impl TokenKind {
    pub fn is_assignment_op(&self) -> bool {
        matches!(self,
            TokenKind::Eq | TokenKind::PlusEq | TokenKind::MinusEq |
            TokenKind::StarEq | TokenKind::SlashEq | TokenKind::PercentEq |
            TokenKind::AmpEq | TokenKind::PipeEq | TokenKind::CaretEq |
            TokenKind::LtLtEq | TokenKind::GtGtEq | TokenKind::GtGtGtEq |
            TokenKind::StarStarEq | TokenKind::AmpAmpEq | TokenKind::PipePipeEq |
            TokenKind::QuestionQuestionEq
        )
    }
}
