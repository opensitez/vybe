use crate::token::{Token, TokenKind};

pub fn tokenize(source: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let chars: Vec<char> = source.chars().collect();
    let mut pos = 0;
    let mut line = 1u32;

    while pos < chars.len() {
        let c = chars[pos];

        // Whitespace
        if c == '\n' { line += 1; pos += 1; continue; }
        if c.is_whitespace() { pos += 1; continue; }

        // Line comments
        if c == '/' && pos + 1 < chars.len() && chars[pos + 1] == '/' {
            while pos < chars.len() && chars[pos] != '\n' { pos += 1; }
            continue;
        }
        // Block comments
        if c == '/' && pos + 1 < chars.len() && chars[pos + 1] == '*' {
            pos += 2;
            while pos + 1 < chars.len() && !(chars[pos] == '*' && chars[pos + 1] == '/') {
                if chars[pos] == '\n' { line += 1; }
                pos += 1;
            }
            pos += 2;
            continue;
        }

        // Interpolated string $"..."
        if c == '$' && pos + 1 < chars.len() && chars[pos + 1] == '"' {
            pos += 2;
            let mut parts: Vec<Token> = Vec::new();
            let mut text = String::new();
            let mut first = true;
            while pos < chars.len() && chars[pos] != '"' {
                if chars[pos] == '{' && pos + 1 < chars.len() && chars[pos + 1] != '{' {
                    if first {
                        tokens.push(Token { kind: TokenKind::InterpolatedStart(text.clone()), line });
                        first = false;
                    } else {
                        tokens.push(Token { kind: TokenKind::InterpolatedMid(text.clone()), line });
                    }
                    text.clear();
                    pos += 1;
                    // Tokenize expression inside {}
                    let mut depth = 1;
                    let inner_start = pos;
                    while pos < chars.len() && depth > 0 {
                        if chars[pos] == '{' { depth += 1; }
                        if chars[pos] == '}' { depth -= 1; }
                        if depth > 0 { pos += 1; }
                    }
                    let inner: String = chars[inner_start..pos].iter().collect();
                    // Tokenize inner expression and add to tokens
                    let inner_tokens = tokenize(&inner);
                    for t in inner_tokens {
                        if t.kind != TokenKind::Eof {
                            tokens.push(Token { kind: t.kind, line });
                        }
                    }
                    pos += 1; // skip }
                } else if chars[pos] == '{' && pos + 1 < chars.len() && chars[pos + 1] == '{' {
                    text.push('{');
                    pos += 2;
                } else {
                    if chars[pos] == '\n' { line += 1; }
                    text.push(chars[pos]);
                    pos += 1;
                }
            }
            if first {
                // No interpolation — just a regular string
                tokens.push(Token { kind: TokenKind::StringLit(text), line });
            } else {
                tokens.push(Token { kind: TokenKind::InterpolatedEnd(text), line });
            }
            if pos < chars.len() { pos += 1; } // skip "
            continue;
        }

        // String literal
        if c == '"' {
            pos += 1;
            let mut s = String::new();
            while pos < chars.len() && chars[pos] != '"' {
                if chars[pos] == '\\' && pos + 1 < chars.len() {
                    pos += 1;
                    match chars[pos] {
                        'n' => s.push('\n'),
                        't' => s.push('\t'),
                        'r' => s.push('\r'),
                        '\\' => s.push('\\'),
                        '"' => s.push('"'),
                        '0' => s.push('\0'),
                        _ => { s.push('\\'); s.push(chars[pos]); }
                    }
                } else {
                    s.push(chars[pos]);
                }
                pos += 1;
            }
            if pos < chars.len() { pos += 1; }
            tokens.push(Token { kind: TokenKind::StringLit(s), line });
            continue;
        }

        // Verbatim string @"..."
        if c == '@' && pos + 1 < chars.len() && chars[pos + 1] == '"' {
            pos += 2;
            let mut s = String::new();
            while pos < chars.len() {
                if chars[pos] == '"' {
                    if pos + 1 < chars.len() && chars[pos + 1] == '"' {
                        s.push('"');
                        pos += 2;
                    } else {
                        pos += 1;
                        break;
                    }
                } else {
                    if chars[pos] == '\n' { line += 1; }
                    s.push(chars[pos]);
                    pos += 1;
                }
            }
            tokens.push(Token { kind: TokenKind::StringLit(s), line });
            continue;
        }

        // Char literal
        if c == '\'' {
            pos += 1;
            let ch = if pos < chars.len() && chars[pos] == '\\' {
                pos += 1;
                match chars.get(pos) {
                    Some('n') => '\n', Some('t') => '\t', Some('r') => '\r',
                    Some('\\') => '\\', Some('\'') => '\'', Some('0') => '\0',
                    _ => '?',
                }
            } else if pos < chars.len() {
                chars[pos]
            } else { '?' };
            pos += 1;
            if pos < chars.len() && chars[pos] == '\'' { pos += 1; }
            tokens.push(Token { kind: TokenKind::CharLit(ch), line });
            continue;
        }

        // Numbers
        if c.is_ascii_digit() || (c == '.' && pos + 1 < chars.len() && chars[pos + 1].is_ascii_digit()) {
            let start = pos;
            let mut is_float = false;
            while pos < chars.len() && (chars[pos].is_ascii_digit() || chars[pos] == '.' || chars[pos] == '_') {
                // Stop at '..' (range operator)
                if chars[pos] == '.' && pos + 1 < chars.len() && chars[pos + 1] == '.' {
                    break;
                }
                if chars[pos] == '.' { is_float = true; }
                pos += 1;
            }
            // Skip suffix (f, d, m, L, etc.)
            if pos < chars.len() && matches!(chars[pos], 'f' | 'F' | 'd' | 'D' | 'm' | 'M' | 'L' | 'l' | 'u' | 'U') {
                if chars[pos] == 'f' || chars[pos] == 'F' { is_float = true; }
                pos += 1;
            }
            let num_str: String = chars[start..pos].iter().filter(|c| **c != '_').collect();
            if is_float {
                let v = num_str.parse::<f64>().unwrap_or(0.0);
                tokens.push(Token { kind: TokenKind::DoubleLit(v), line });
            } else {
                let v = num_str.parse::<i64>().unwrap_or(0);
                tokens.push(Token { kind: TokenKind::IntLit(v), line });
            }
            continue;
        }

        // Identifiers and keywords
        if c.is_alphabetic() || c == '_' {
            let start = pos;
            while pos < chars.len() && (chars[pos].is_alphanumeric() || chars[pos] == '_') { pos += 1; }
            let word: String = chars[start..pos].iter().collect();
            let kind = match word.as_str() {
                "using" => TokenKind::Using,
                "namespace" => TokenKind::Namespace,
                "class" => TokenKind::Class,
                "struct" => TokenKind::Struct,
                "interface" => TokenKind::Interface,
                "enum" => TokenKind::Enum,
                "public" => TokenKind::Public,
                "private" => TokenKind::Private,
                "protected" => TokenKind::Protected,
                "internal" => TokenKind::Internal,
                "static" => TokenKind::Static,
                "abstract" => TokenKind::Abstract,
                "virtual" => TokenKind::Virtual,
                "override" => TokenKind::Override,
                "sealed" => TokenKind::Sealed,
                "partial" => TokenKind::Partial,
                "readonly" => TokenKind::Readonly,
                "const" => TokenKind::Const,
                "new" => TokenKind::New,
                "this" => TokenKind::This,
                "base" => TokenKind::Base,
                "null" => TokenKind::Null,
                "true" => TokenKind::True,
                "false" => TokenKind::False,
                "void" => TokenKind::Void,
                "if" => TokenKind::If,
                "else" => TokenKind::Else,
                "for" => TokenKind::For,
                "foreach" => TokenKind::ForEach,
                "in" => TokenKind::In,
                "while" => TokenKind::While,
                "do" => TokenKind::Do,
                "switch" => TokenKind::Switch,
                "case" => TokenKind::Case,
                "default" => TokenKind::Default,
                "return" => TokenKind::Return,
                "break" => TokenKind::Break,
                "continue" => TokenKind::Continue,
                "throw" => TokenKind::Throw,
                "try" => TokenKind::Try,
                "catch" => TokenKind::Catch,
                "finally" => TokenKind::Finally,
                "var" => TokenKind::Var,
                "int" => TokenKind::Int,
                "string" => TokenKind::String_,
                "double" => TokenKind::Double,
                "float" => TokenKind::Float,
                "bool" => TokenKind::Bool,
                "char" => TokenKind::Char,
                "long" => TokenKind::Long,
                "byte" => TokenKind::Byte,
                "object" => TokenKind::Object,
                "is" => TokenKind::Is,
                "as" => TokenKind::As,
                "typeof" => TokenKind::TypeOf,
                "nameof" => TokenKind::NameOf,
                "async" => TokenKind::Async,
                "await" => TokenKind::Await,
                "params" => TokenKind::Params,
                "ref" => TokenKind::Ref,
                "out" => TokenKind::Out,
                "event" => TokenKind::Event,
                _ => TokenKind::Identifier(word),
            };
            tokens.push(Token { kind, line });
            continue;
        }

        // Operators and punctuation
        let kind = match c {
            '+' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '+' => { pos += 1; TokenKind::Increment }
                        '=' => { pos += 1; TokenKind::PlusAssign }
                        _ => TokenKind::Plus,
                    }
                } else { TokenKind::Plus }
            }
            '-' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '-' => { pos += 1; TokenKind::Decrement }
                        '=' => { pos += 1; TokenKind::MinusAssign }
                        _ => TokenKind::Minus,
                    }
                } else { TokenKind::Minus }
            }
            '*' => if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::StarAssign } else { TokenKind::Star },
            '/' => if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::SlashAssign } else { TokenKind::Slash },
            '%' => if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::PercentAssign } else { TokenKind::Percent },
            '=' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '=' => { pos += 1; TokenKind::Eq }
                        '>' => { pos += 1; TokenKind::Arrow }
                        _ => TokenKind::Assign,
                    }
                } else { TokenKind::Assign }
            }
            '!' => if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::Neq } else { TokenKind::Not },
            '<' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '=' => { pos += 1; TokenKind::Le }
                        '<' => { pos += 1; if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::ShlAssign } else { TokenKind::Shl } }
                        _ => TokenKind::Lt,
                    }
                } else { TokenKind::Lt }
            }
            '>' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '=' => { pos += 1; TokenKind::Ge }
                        '>' => { pos += 1; if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::ShrAssign } else { TokenKind::Shr } }
                        _ => TokenKind::Gt,
                    }
                } else { TokenKind::Gt }
            }
            '&' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '&' => { pos += 1; TokenKind::And }
                        '=' => { pos += 1; TokenKind::AndAssign }
                        _ => TokenKind::BitAnd,
                    }
                } else { TokenKind::BitAnd }
            }
            '|' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '|' => { pos += 1; TokenKind::Or }
                        '=' => { pos += 1; TokenKind::OrAssign }
                        _ => TokenKind::BitOr,
                    }
                } else { TokenKind::BitOr }
            }
            '^' => if pos + 1 < chars.len() && chars[pos + 1] == '=' { pos += 1; TokenKind::XorAssign } else { TokenKind::BitXor },
            '~' => TokenKind::Tilde,
            '?' => {
                if pos + 1 < chars.len() {
                    match chars[pos + 1] {
                        '?' => { pos += 1; TokenKind::QuestionQuestion }
                        '.' => { pos += 1; TokenKind::QuestionDot }
                        _ => TokenKind::Question,
                    }
                } else { TokenKind::Question }
            }
            '.' => {
                if pos + 1 < chars.len() && chars[pos + 1] == '.' {
                    pos += 1;
                    TokenKind::DotDot
                } else {
                    TokenKind::Dot
                }
            }
            ',' => TokenKind::Comma,
            ':' => TokenKind::Colon,
            ';' => TokenKind::Semicolon,
            '(' => TokenKind::LParen,
            ')' => TokenKind::RParen,
            '{' => TokenKind::LBrace,
            '}' => TokenKind::RBrace,
            '[' => TokenKind::LBracket,
            ']' => TokenKind::RBracket,
            _ => { pos += 1; continue; }
        };
        tokens.push(Token { kind, line });
        pos += 1;
    }

    tokens.push(Token { kind: TokenKind::Eof, line });
    tokens
}
