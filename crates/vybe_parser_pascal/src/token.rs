#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Keywords
    Program, Unit, Uses, Interface, Implementation,
    Begin, End, Var, Const, Type,
    Procedure, Function, Forward,
    If, Then, Else,
    For, To, Downto, Do,
    While, Repeat, Until,
    Case, Of, Otherwise,
    Record, Array, Set, File, Object, Class,
    Inherited, Override, Virtual, Abstract,
    With, In, Is, As,
    Try, Except, Finally, Raise, On,
    And, Or, Not, Xor, Div, Mod, Shl, Shr,
    Nil, True, False,
    String, Integer, Real, Boolean, Char, Byte, Word,
    LongInt, ShortInt, Cardinal, Int64, Single, Double, Extended,
    Pointer, Void,
    Exit, Break, Continue, Halt,
    New, Dispose, Inherited_,
    Result,
    Constructor, Destructor,
    Public, Private, Protected, Published,

    // Literals
    Identifier(String),
    IntLiteral(i64),
    RealLiteral(f64),
    StringLiteral(String),
    CharLiteral(char),

    // Operators
    Plus, Minus, Star, Slash,
    Eq,          // =
    NotEq,       // <>
    Lt, Gt, Le, Ge,
    Assign,      // :=
    At,          // @
    Caret,       // ^
    DotDot,      // ..

    // Punctuation
    LParen, RParen,
    LBracket, RBracket,
    Dot, Comma, Semicolon, Colon,

    EOF,
}
