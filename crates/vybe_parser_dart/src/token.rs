#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    // Keywords
    Abstract, As, Assert, Async, Await,
    Break, Case, Catch, Class, Const, Continue,
    Covariant, Default, Deferred, Do, Dynamic,
    Else, Enum, Export, Extends, Extension,
    External, Factory, False, Final, Finally,
    For, Function, Get, Hide, If, Implements,
    Import, In, Interface, Is, Late, Library,
    Mixin, Native, New, Null, On, Operator,
    Part, Required, Rethrow, Return, Set, Show,
    Static, Super, Switch, Sync, This, Throw,
    True, Try, Typedef, Var, Void, While, With, Yield,
    Override,

    // Literals
    Identifier(String),
    StringLiteral(String),
    IntLiteral(i64),
    DoubleLiteral(f64),

    // String interpolation
    StringStart(String),   // "text${
    StringMiddle(String),  // }text${
    StringEnd(String),     // }text"

    // Operators
    Plus, Minus, Star, Slash, Percent, TildeSlash,
    PlusPlus, MinusMinus,
    Eq, PlusEq, MinusEq, StarEq, SlashEq, PercentEq, TildeSlashEq,
    AmpEq, BarEq, CaretEq, LessLessEq, GreaterGreaterEq, GreaterGreaterGreaterEq,
    QuestionQuestionEq,
    EqEq, BangEq,
    Greater, Less, GreaterEq, LessEq,
    Amp, Bar, Caret, Tilde, LessLess, GreaterGreater, GreaterGreaterGreater,
    AmpAmp, BarBar, Bang,
    Question, QuestionQuestion, QuestionDot,
    Arrow,   // =>
    At,      // @

    // Punctuation
    Dot, DotDot, DotDotQuestion, DotDotDot,
    Comma, Semicolon, Colon,
    LParen, RParen,
    LBracket, RBracket,
    LBrace, RBrace,

    EOF,
    Error(String),
}

impl Token {
    pub fn is_assign_op(&self) -> bool {
        matches!(self,
            Token::Eq | Token::PlusEq | Token::MinusEq | Token::StarEq |
            Token::SlashEq | Token::PercentEq | Token::TildeSlashEq |
            Token::AmpEq | Token::BarEq | Token::CaretEq |
            Token::LessLessEq | Token::GreaterGreaterEq | Token::GreaterGreaterGreaterEq |
            Token::QuestionQuestionEq
        )
    }

    pub fn to_ident_str(&self) -> Option<String> {
        match self {
            Token::Identifier(s) => Some(s.clone()),
            Token::Abstract => Some("abstract".into()),
            Token::As => Some("as".into()),
            Token::Async => Some("async".into()),
            Token::Await => Some("await".into()),
            Token::Covariant => Some("covariant".into()),
            Token::Deferred => Some("deferred".into()),
            Token::Dynamic => Some("dynamic".into()),
            Token::Export => Some("export".into()),
            Token::Extension => Some("extension".into()),
            Token::External => Some("external".into()),
            Token::Factory => Some("factory".into()),
            Token::Function => Some("Function".into()),
            Token::Get => Some("get".into()),
            Token::Hide => Some("hide".into()),
            Token::Implements => Some("implements".into()),
            Token::Import => Some("import".into()),
            Token::Interface => Some("interface".into()),
            Token::Late => Some("late".into()),
            Token::Library => Some("library".into()),
            Token::Mixin => Some("mixin".into()),
            Token::Native => Some("native".into()),
            Token::On => Some("on".into()),
            Token::Operator => Some("operator".into()),
            Token::Part => Some("part".into()),
            Token::Required => Some("required".into()),
            Token::Set => Some("set".into()),
            Token::Show => Some("show".into()),
            Token::Static => Some("static".into()),
            Token::Sync => Some("sync".into()),
            Token::Typedef => Some("typedef".into()),
            Token::Yield => Some("yield".into()),
            Token::Override => Some("override".into()),
            _ => None,
        }
    }
}
