use super::token::Token;

pub struct Lexer {
    chars: Vec<char>,
    pos: usize,
    line: u32,
}

impl Lexer {
    pub fn new(input: &str) -> Self {
        Self { chars: input.chars().collect(), pos: 0, line: 1 }
    }

    pub fn line(&self) -> u32 { self.line }

    fn peek(&self) -> Option<char> { self.chars.get(self.pos).copied() }
    fn peek2(&self) -> Option<char> { self.chars.get(self.pos + 1).copied() }

    fn advance(&mut self) -> Option<char> {
        let ch = self.chars.get(self.pos).copied();
        if let Some(c) = ch {
            self.pos += 1;
            if c == '\n' { self.line += 1; }
        }
        ch
    }

    fn eat(&mut self, ch: char) -> bool {
        if self.peek() == Some(ch) { self.advance(); true } else { false }
    }

    fn skip_whitespace_and_comments(&mut self) {
        loop {
            // whitespace
            while matches!(self.peek(), Some(c) if c.is_whitespace()) { self.advance(); }
            // line comment
            if self.peek() == Some('/') && self.peek2() == Some('/') {
                while !matches!(self.peek(), None | Some('\n')) { self.advance(); }
                continue;
            }
            // block comment
            if self.peek() == Some('/') && self.peek2() == Some('*') {
                self.advance(); self.advance();
                loop {
                    match self.advance() {
                        None => break,
                        Some('*') if self.peek() == Some('/') => { self.advance(); break; }
                        _ => {}
                    }
                }
                continue;
            }
            break;
        }
    }

    pub fn next_token(&mut self) -> Token {
        self.skip_whitespace_and_comments();
        let ch = match self.advance() {
            Some(c) => c,
            None => return Token::EOF,
        };

        match ch {
            '(' => Token::LParen,
            ')' => Token::RParen,
            '{' => Token::LBrace,
            '}' => Token::RBrace,
            '[' => Token::LBracket,
            ']' => Token::RBracket,
            ',' => Token::Comma,
            ';' => Token::Semicolon,
            ':' => Token::Colon,
            '@' => Token::At,
            '~' => {
                if self.eat('/') {
                    if self.eat('=') { Token::TildeSlashEq } else { Token::TildeSlash }
                } else { Token::Tilde }
            }
            '.' => {
                if self.peek() == Some('.') {
                    self.advance();
                    if self.eat('.') { Token::DotDotDot }
                    else { Token::DotDot }
                } else { Token::Dot }
            }
            '+' => {
                if self.eat('+') { Token::PlusPlus }
                else if self.eat('=') { Token::PlusEq }
                else { Token::Plus }
            }
            '-' => {
                if self.eat('-') { Token::MinusMinus }
                else if self.eat('=') { Token::MinusEq }
                else { Token::Minus }
            }
            '*' => { if self.eat('=') { Token::StarEq } else { Token::Star } }
            '/' => { if self.eat('=') { Token::SlashEq } else { Token::Slash } }
            '%' => { if self.eat('=') { Token::PercentEq } else { Token::Percent } }
            '=' => {
                if self.eat('=') { Token::EqEq }
                else if self.eat('>') { Token::Arrow }
                else { Token::Eq }
            }
            '!' => { if self.eat('=') { Token::BangEq } else { Token::Bang } }
            '<' => {
                if self.eat('<') {
                    if self.eat('=') { Token::LessLessEq } else { Token::LessLess }
                } else if self.eat('=') { Token::LessEq }
                else { Token::Less }
            }
            '>' => {
                if self.peek() == Some('>') {
                    self.advance();
                    if self.peek() == Some('>') {
                        self.advance();
                        if self.eat('=') { Token::GreaterGreaterGreaterEq } else { Token::GreaterGreaterGreater }
                    } else if self.eat('=') { Token::GreaterGreaterEq }
                    else { Token::GreaterGreater }
                } else if self.eat('=') { Token::GreaterEq }
                else { Token::Greater }
            }
            '&' => {
                if self.eat('&') { Token::AmpAmp }
                else if self.eat('=') { Token::AmpEq }
                else { Token::Amp }
            }
            '|' => {
                if self.eat('|') { Token::BarBar }
                else if self.eat('=') { Token::BarEq }
                else { Token::Bar }
            }
            '^' => { if self.eat('=') { Token::CaretEq } else { Token::Caret } }
            '?' => {
                if self.peek() == Some('?') {
                    self.advance();
                    if self.eat('=') { Token::QuestionQuestionEq } else { Token::QuestionQuestion }
                } else if self.eat('.') { 
                    if self.eat('.') { Token::QuestionDotDot }
                    else { Token::QuestionDot }
                } else { Token::Question }
            }
            '\'' | '"' => self.read_string(ch, false),
            'r' if matches!(self.peek(), Some('\'') | Some('"')) => {
                let q = self.advance().unwrap();
                self.read_raw_string(q)
            }
            '0'..='9' => self.read_number(ch),
            'a'..='z' | 'A'..='Z' | '_' | '$' => self.read_identifier(ch),
            _ => Token::Error(format!("Unexpected char: {}", ch)),
        }
    }

    fn read_string(&mut self, quote: char, raw: bool) -> Token {
        // Check for triple-quoted
        let triple = self.peek() == Some(quote) && self.peek2() == Some(quote);
        if triple { self.advance(); self.advance(); }

        let mut s = String::new();
        let mut encoded = String::new();
        let mut has_interp = false;

        loop {
            match self.advance() {
                None => return Token::Error("Unterminated string".into()),
                Some(c) if c == quote => {
                    if triple {
                        if self.peek() == Some(quote) && self.peek2() == Some(quote) {
                            self.advance(); self.advance();
                            break;
                        }
                        s.push(c);
                    } else {
                        break;
                    }
                }
                Some('$') if !raw => {
                    has_interp = true;
                    if !s.is_empty() {
                        encoded.push_str("L");
                        encoded.push_str(&s.replace('\x01', "\x01\x01"));
                        encoded.push('\x01');
                        s.clear();
                    }
                    if self.peek() == Some('{') {
                        self.advance();
                        // We can't parse the expression here in the lexer — emit a marker
                        // The parser handles interpolation by re-lexing
                        // For now, collect until matching }
                        let mut depth = 1;
                        let mut expr_src = String::new();
                        while depth > 0 {
                            match self.advance() {
                                None => break,
                                Some('{') => { depth += 1; expr_src.push('{'); }
                                Some('}') => { depth -= 1; if depth > 0 { expr_src.push('}'); } }
                                Some(c) => expr_src.push(c),
                            }
                        }
                        encoded.push_str("E");
                        encoded.push_str(&expr_src);
                        encoded.push('\x01');
                    } else {
                        // $identifier
                        let mut id = String::new();
                        while matches!(self.peek(), Some(c) if c.is_alphanumeric() || c == '_') {
                            id.push(self.advance().unwrap());
                        }
                        encoded.push_str("E");
                        encoded.push_str(&id);
                        encoded.push('\x01');
                    }
                }
                Some('\\') if !raw => {
                    match self.advance() {
                        Some('n') => s.push('\n'),
                        Some('r') => s.push('\r'),
                        Some('t') => s.push('\t'),
                        Some('\\') => s.push('\\'),
                        Some('\'') => s.push('\''),
                        Some('"') => s.push('"'),
                        Some('$') => s.push('$'),
                        Some('u') => {
                            // \uXXXX or \u{XXXX}
                            let hex = if self.peek() == Some('{') {
                                self.advance();
                                let mut h = String::new();
                                while self.peek() != Some('}') { h.push(self.advance().unwrap_or('0')); }
                                self.advance();
                                h
                            } else {
                                (0..4).map(|_| self.advance().unwrap_or('0')).collect()
                            };
                            if let Ok(n) = u32::from_str_radix(&hex, 16) {
                                if let Some(c) = char::from_u32(n) { s.push(c); }
                            }
                        }
                        Some(c) => s.push(c),
                        None => {}
                    }
                }
                Some(c) => s.push(c),
            }
        }

        if has_interp {
            if !s.is_empty() {
                encoded.push_str("L");
                encoded.push_str(&s.replace('\x01', "\x01\x01"));
            } else if encoded.ends_with('\x01') {
                encoded.pop(); // remove trailing separator
            }
            Token::StringLiteral(format!("\x00INTERP\x00{}", encoded))
        } else {
            Token::StringLiteral(s)
        }
    }

    fn read_raw_string(&mut self, quote: char) -> Token {
        self.read_string(quote, true)
    }

    fn read_number(&mut self, first: char) -> Token {
        let mut s = String::from(first);
        // Hex
        if first == '0' && matches!(self.peek(), Some('x') | Some('X')) {
            s.push(self.advance().unwrap());
            while matches!(self.peek(), Some(c) if c.is_ascii_hexdigit()) {
                s.push(self.advance().unwrap());
            }
            return Token::IntLiteral(i64::from_str_radix(&s[2..], 16).unwrap_or(0));
        }
        let mut is_double = false;
        while let Some(c) = self.peek() {
            if c.is_ascii_digit() { s.push(self.advance().unwrap()); }
            else if c == '.' && !is_double && matches!(self.peek2(), Some(d) if d.is_ascii_digit()) {
                is_double = true; s.push(self.advance().unwrap());
            } else if (c == 'e' || c == 'E') && !is_double {
                is_double = true; s.push(self.advance().unwrap());
                if matches!(self.peek(), Some('+') | Some('-')) { s.push(self.advance().unwrap()); }
            } else { break; }
        }
        if is_double {
            Token::DoubleLiteral(s.parse().unwrap_or(0.0))
        } else {
            Token::IntLiteral(s.parse().unwrap_or(0))
        }
    }

    fn read_identifier(&mut self, first: char) -> Token {
        let mut s = String::from(first);
        while matches!(self.peek(), Some(c) if c.is_alphanumeric() || c == '_' || c == '$') {
            s.push(self.advance().unwrap());
        }
        keyword_or_ident(s)
    }
}

fn keyword_or_ident(s: String) -> Token {
    match s.as_str() {
        "abstract" => Token::Abstract,
        "as" => Token::As,
        "assert" => Token::Assert,
        "async" => Token::Async,
        "await" => Token::Await,
        "break" => Token::Break,
        "case" => Token::Case,
        "catch" => Token::Catch,
        "class" => Token::Class,
        "const" => Token::Const,
        "continue" => Token::Continue,
        "covariant" => Token::Covariant,
        "default" => Token::Default,
        "deferred" => Token::Deferred,
        "do" => Token::Do,
        "dynamic" => Token::Dynamic,
        "else" => Token::Else,
        "enum" => Token::Enum,
        "export" => Token::Export,
        "extends" => Token::Extends,
        "extension" => Token::Extension,
        "external" => Token::External,
        "factory" => Token::Factory,
        "false" => Token::False,
        "final" => Token::Final,
        "finally" => Token::Finally,
        "for" => Token::For,
        "get" => Token::Get,
        "hide" => Token::Hide,
        "if" => Token::If,
        "implements" => Token::Implements,
        "import" => Token::Import,
        "in" => Token::In,
        "interface" => Token::Interface,
        "is" => Token::Is,
        "late" => Token::Late,
        "library" => Token::Library,
        "mixin" => Token::Mixin,
        "native" => Token::Native,
        "new" => Token::New,
        "null" => Token::Null,
        "on" => Token::On,
        "operator" => Token::Operator,
        "override" => Token::Override,
        "part" => Token::Part,
        "required" => Token::Required,
        "rethrow" => Token::Rethrow,
        "return" => Token::Return,
        "set" => Token::Set,
        "show" => Token::Show,
        "static" => Token::Static,
        "super" => Token::Super,
        "switch" => Token::Switch,
        "sync" => Token::Sync,
        "this" => Token::This,
        "throw" => Token::Throw,
        "true" => Token::True,
        "try" => Token::Try,
        "typedef" => Token::Typedef,
        "var" => Token::Var,
        "void" => Token::Void,
        "while" => Token::While,
        "with" => Token::With,
        "yield" => Token::Yield,
        _ => Token::Identifier(s),
    }
}
