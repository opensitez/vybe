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
    Symbol(String),       // :name
    Identifier(String),
    InstanceVar(String),  // @name
    ClassVar(String),     // @@name
    GlobalVar(String),    // $name
    Constant(String),     // Name (uppercase start)

    // Keywords
    If,
    Unless,
    Elsif,
    Else,
    Then,
    End,
    While,
    Until,
    For,
    In,
    Do,
    Case,
    When,
    Break,
    Next,
    Return,
    Def,
    Class,
    Module,
    Self_,
    Super,
    Yield,
    Block_given,
    Begin,
    Rescue,
    Ensure,
    Raise,
    Retry,
    Nil,
    True,
    False,
    And,        // and (low precedence)
    Or,         // or (low precedence)
    Not,        // not (low precedence)
    Require,
    Include,
    Extend,
    Attr_reader,
    Attr_writer,
    Attr_accessor,
    Lambda,
    Proc,
    Puts,
    Print,
    P,
    Private,
    Protected,
    Public,
    Alias,
    Defined,
    Loop,
    Catch,
    Throw,
    Freeze,
    Frozen,

    Pp,
    Format,
    Sprintf,
    Redo,
    AtExit,

    // Regex / special
    Regex(String),
    EqTilde,     // =~
    Backtick(String), // `command`
    AmpDot,      // &. (safe navigation)

    // Arithmetic
    Plus,        // +
    Minus,       // -
    Star,        // *
    Slash,       // /
    Percent,     // %
    StarStar,    // **

    // Assignment
    Eq,          // =
    PlusEq,      // +=
    MinusEq,     // -=
    StarEq,      // *=
    SlashEq,     // /=
    PercentEq,   // %=
    AmpAmpEq,    // &&=
    PipePipeEq,  // ||=
    PipeEq,      // |=
    AmpEq,       // &=
    CaretEq,     // ^=
    LtLtEq,      // <<=
    GtGtEq,      // >>=

    // Comparison
    EqEq,        // ==
    BangEq,      // !=
    Lt,          // <
    Gt,          // >
    LtEq,        // <=
    GtEq,        // >=
    Spaceship,   // <=>
    EqEqEq,      // ===

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

    // Member / range
    Dot,         // .
    DotDot,      // ..
    DotDotDot,   // ...
    ColonColon,  // ::
    FatArrow,    // =>
    Arrow,       // ->
    Question,    // ?
    Colon,       // :

    // Delimiters
    LParen,
    RParen,
    LBrace,
    RBrace,
    LBracket,
    RBracket,
    Semicolon,
    Comma,
    Newline,
    Pipe2,       // | (for block params)

    Eof,
}
