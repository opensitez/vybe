#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,
    pub line: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    // Literals
    IntLit(i64),
    DoubleLit(f64),
    StringLit(String),
    CharLit(char),
    InterpolatedStart(String), // $"text{
    InterpolatedMid(String),   // }text{
    InterpolatedEnd(String),   // }text"
    Identifier(String),

    // Keywords
    Using, Namespace, Class, Struct, Interface, Enum,
    Public, Private, Protected, Internal,
    Static, Abstract, Virtual, Override, Sealed, Partial, Readonly, Const,
    New, This, Base, Null, True, False, Void,
    If, Else, For, ForEach, In, While, Do, Switch, Case, Default,
    Return, Break, Continue, Throw, Try, Catch, Finally,
    Var, Int, String_, Double, Float, Bool, Char, Long, Byte, Object,
    Is, As, TypeOf, NameOf, Async, Await, Params, Ref, Out, Event, Lock,

    // Operators
    Plus, Minus, Star, Slash, Percent,
    Eq, Neq, Lt, Gt, Le, Ge,
    Assign, PlusAssign, MinusAssign, StarAssign, SlashAssign, PercentAssign,
    And, Or, Not, BitAnd, BitOr, BitXor, Tilde, Shl, Shr,
    AndAssign, OrAssign, XorAssign, ShlAssign, ShrAssign,
    Increment, Decrement,
    Arrow, // =>
    Dot, DotDot, QuestionDot, // . .. ?.
    Question, QuestionQuestion, // ? ??
    Colon,
    Semicolon, Comma,

    // Delimiters
    LParen, RParen, LBrace, RBrace, LBracket, RBracket,

    // Special
    Eof,
}
