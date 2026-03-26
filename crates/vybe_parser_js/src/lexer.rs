use crate::token::{Span, Token, TokenKind};

pub struct Lexer {
    source: Vec<char>,
    pos: usize,
    line: u32,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        Lexer {
            source: source.chars().collect(),
            pos: 0,
            line: 1,
        }
    }

    pub fn tokenize(&mut self) -> Result<Vec<Token>, String> {
        let mut tokens = Vec::new();
        loop {
            let tok = self.next_token()?;
            let is_eof = tok.kind == TokenKind::Eof;
            tokens.push(tok);
            if is_eof {
                break;
            }
        }
        Ok(tokens)
    }

    fn next_token(&mut self) -> Result<Token, String> {
        self.skip_whitespace_and_comments();

        if self.pos >= self.source.len() {
            return Ok(self.make_token(TokenKind::Eof, self.pos));
        }

        let start = self.pos;
        let ch = self.current();

        // Numbers
        if ch.is_ascii_digit() || (ch == '.' && self.peek_is_digit()) {
            return self.lex_number(start);
        }

        // Strings
        if ch == '"' || ch == '\'' {
            return self.lex_string(start, ch);
        }

        // Template literals
        if ch == '`' {
            return self.lex_template(start);
        }

        // Identifiers and keywords
        if is_ident_start(ch) {
            return self.lex_identifier(start);
        }

        // Punctuation and operators
        self.lex_punctuation(start)
    }

    fn lex_number(&mut self, start: usize) -> Result<Token, String> {
        // Hex
        if self.current() == '0' && self.pos + 1 < self.source.len()
            && (self.source[self.pos + 1] == 'x' || self.source[self.pos + 1] == 'X')
        {
            self.pos += 2;
            while self.pos < self.source.len() && self.source[self.pos].is_ascii_hexdigit() {
                self.pos += 1;
            }
            let s: String = self.source[start..self.pos].iter().collect();
            let n = u64::from_str_radix(&s[2..], 16).map_err(|e| format!("Invalid hex: {}", e))?;
            return Ok(self.make_token(TokenKind::Number(n as f64), start));
        }

        // Binary
        if self.current() == '0' && self.pos + 1 < self.source.len()
            && (self.source[self.pos + 1] == 'b' || self.source[self.pos + 1] == 'B')
        {
            self.pos += 2;
            while self.pos < self.source.len() && (self.source[self.pos] == '0' || self.source[self.pos] == '1') {
                self.pos += 1;
            }
            let s: String = self.source[start + 2..self.pos].iter().collect();
            let n = u64::from_str_radix(&s, 2).map_err(|e| format!("Invalid binary: {}", e))?;
            return Ok(self.make_token(TokenKind::Number(n as f64), start));
        }

        // Octal (0o)
        if self.current() == '0' && self.pos + 1 < self.source.len()
            && (self.source[self.pos + 1] == 'o' || self.source[self.pos + 1] == 'O')
        {
            self.pos += 2;
            while self.pos < self.source.len() && self.source[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
            let s: String = self.source[start + 2..self.pos].iter().collect();
            let n = u64::from_str_radix(&s, 8).map_err(|e| format!("Invalid octal: {}", e))?;
            return Ok(self.make_token(TokenKind::Number(n as f64), start));
        }

        // Decimal
        while self.pos < self.source.len() && self.source[self.pos].is_ascii_digit() {
            self.pos += 1;
        }
        // Decimal point
        if self.pos < self.source.len() && self.source[self.pos] == '.'
            && self.pos + 1 < self.source.len() && self.source[self.pos + 1].is_ascii_digit()
        {
            self.pos += 1;
            while self.pos < self.source.len() && self.source[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
        }
        // Exponent
        if self.pos < self.source.len() && (self.source[self.pos] == 'e' || self.source[self.pos] == 'E') {
            self.pos += 1;
            if self.pos < self.source.len() && (self.source[self.pos] == '+' || self.source[self.pos] == '-') {
                self.pos += 1;
            }
            while self.pos < self.source.len() && self.source[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
        }

        let s: String = self.source[start..self.pos].iter().collect();
        let n: f64 = s.parse().map_err(|e| format!("Invalid number '{}': {}", s, e))?;
        Ok(self.make_token(TokenKind::Number(n), start))
    }

    fn lex_string(&mut self, start: usize, quote: char) -> Result<Token, String> {
        self.pos += 1; // skip opening quote
        let mut value = String::new();
        while self.pos < self.source.len() && self.source[self.pos] != quote {
            if self.source[self.pos] == '\\' {
                self.pos += 1;
                if self.pos >= self.source.len() {
                    return Err(format!("Unterminated string at line {}", self.line));
                }
                match self.source[self.pos] {
                    'n' => value.push('\n'),
                    'r' => value.push('\r'),
                    't' => value.push('\t'),
                    '\\' => value.push('\\'),
                    '\'' => value.push('\''),
                    '"' => value.push('"'),
                    '`' => value.push('`'),
                    '0' => value.push('\0'),
                    'u' => {
                        self.pos += 1;
                        value.push(self.lex_unicode_escape()?);
                        continue;
                    }
                    c => { value.push('\\'); value.push(c); }
                }
            } else {
                if self.source[self.pos] == '\n' {
                    self.line += 1;
                }
                value.push(self.source[self.pos]);
            }
            self.pos += 1;
        }
        if self.pos >= self.source.len() {
            return Err(format!("Unterminated string at line {}", self.line));
        }
        self.pos += 1; // skip closing quote
        Ok(self.make_token(TokenKind::String(value), start))
    }

    fn lex_template(&mut self, start: usize) -> Result<Token, String> {
        self.pos += 1; // skip `
        let mut value = String::new();
        while self.pos < self.source.len() {
            if self.source[self.pos] == '`' {
                self.pos += 1;
                return Ok(self.make_token(TokenKind::TemplateLiteral(value), start));
            }
            if self.source[self.pos] == '$' && self.pos + 1 < self.source.len() && self.source[self.pos + 1] == '{' {
                self.pos += 2; // skip ${
                return Ok(self.make_token(TokenKind::TemplateHead(value), start));
            }
            if self.source[self.pos] == '\\' {
                self.pos += 1;
                if self.pos < self.source.len() {
                    match self.source[self.pos] {
                        'n' => value.push('\n'),
                        'r' => value.push('\r'),
                        't' => value.push('\t'),
                        '\\' => value.push('\\'),
                        '`' => value.push('`'),
                        '$' => value.push('$'),
                        c => { value.push('\\'); value.push(c); }
                    }
                }
            } else {
                if self.source[self.pos] == '\n' {
                    self.line += 1;
                }
                value.push(self.source[self.pos]);
            }
            self.pos += 1;
        }
        Err(format!("Unterminated template literal at line {}", self.line))
    }

    /// Continue lexing template after an interpolation expression (after `}`).
    pub fn lex_template_continuation(&mut self) -> Result<Token, String> {
        let start = self.pos;
        let mut value = String::new();
        while self.pos < self.source.len() {
            if self.source[self.pos] == '`' {
                self.pos += 1;
                return Ok(self.make_token(TokenKind::TemplateTail(value), start));
            }
            if self.source[self.pos] == '$' && self.pos + 1 < self.source.len() && self.source[self.pos + 1] == '{' {
                self.pos += 2;
                return Ok(self.make_token(TokenKind::TemplateMiddle(value), start));
            }
            if self.source[self.pos] == '\\' {
                self.pos += 1;
                if self.pos < self.source.len() {
                    match self.source[self.pos] {
                        'n' => value.push('\n'),
                        't' => value.push('\t'),
                        '\\' => value.push('\\'),
                        '`' => value.push('`'),
                        '$' => value.push('$'),
                        c => { value.push('\\'); value.push(c); }
                    }
                }
            } else {
                if self.source[self.pos] == '\n' { self.line += 1; }
                value.push(self.source[self.pos]);
            }
            self.pos += 1;
        }
        Err(format!("Unterminated template literal at line {}", self.line))
    }

    fn lex_unicode_escape(&mut self) -> Result<char, String> {
        if self.pos < self.source.len() && self.source[self.pos] == '{' {
            self.pos += 1;
            let start = self.pos;
            while self.pos < self.source.len() && self.source[self.pos] != '}' {
                self.pos += 1;
            }
            let hex: String = self.source[start..self.pos].iter().collect();
            self.pos += 1; // skip }
            let code = u32::from_str_radix(&hex, 16).map_err(|_| "Invalid unicode escape")?;
            char::from_u32(code).ok_or_else(|| "Invalid unicode codepoint".to_string())
        } else {
            let hex: String = self.source[self.pos..self.pos + 4].iter().collect();
            self.pos += 4;
            let code = u32::from_str_radix(&hex, 16).map_err(|_| "Invalid unicode escape")?;
            char::from_u32(code).ok_or_else(|| "Invalid unicode codepoint".to_string())
        }
    }

    fn lex_identifier(&mut self, start: usize) -> Result<Token, String> {
        while self.pos < self.source.len() && is_ident_part(self.source[self.pos]) {
            self.pos += 1;
        }
        let name: String = self.source[start..self.pos].iter().collect();
        let kind = match name.as_str() {
            "var" => TokenKind::Var,
            "let" => TokenKind::Let,
            "const" => TokenKind::Const,
            "function" => TokenKind::Function,
            "return" => TokenKind::Return,
            "if" => TokenKind::If,
            "else" => TokenKind::Else,
            "for" => TokenKind::For,
            "while" => TokenKind::While,
            "do" => TokenKind::Do,
            "break" => TokenKind::Break,
            "continue" => TokenKind::Continue,
            "switch" => TokenKind::Switch,
            "case" => TokenKind::Case,
            "default" => TokenKind::Default,
            "new" => TokenKind::New,
            "delete" => TokenKind::Delete,
            "typeof" => TokenKind::Typeof,
            "void" => TokenKind::Void,
            "instanceof" => TokenKind::Instanceof,
            "in" => TokenKind::In,
            "of" => TokenKind::Of,
            "this" => TokenKind::This,
            "super" => TokenKind::Super,
            "class" => TokenKind::Class,
            "extends" => TokenKind::Extends,
            "static" => TokenKind::Static,
            "get" => TokenKind::Get,
            "set" => TokenKind::Set,
            "throw" => TokenKind::Throw,
            "try" => TokenKind::Try,
            "catch" => TokenKind::Catch,
            "finally" => TokenKind::Finally,
            "true" => TokenKind::True,
            "false" => TokenKind::False,
            "null" => TokenKind::Null,
            "undefined" => TokenKind::Undefined,
            "import" => TokenKind::Import,
            "export" => TokenKind::Export,
            "from" => TokenKind::From,
            "as" => TokenKind::As,
            "async" => TokenKind::Async,
            "await" => TokenKind::Await,
            "yield" => TokenKind::Yield,
            "debugger" => TokenKind::Debugger,
            _ => TokenKind::Identifier(name),
        };
        Ok(self.make_token(kind, start))
    }

    fn lex_punctuation(&mut self, start: usize) -> Result<Token, String> {
        let ch = self.source[self.pos];
        self.pos += 1;

        let kind = match ch {
            '(' => TokenKind::LParen,
            ')' => TokenKind::RParen,
            '{' => TokenKind::LBrace,
            '}' => TokenKind::RBrace,
            '[' => TokenKind::LBracket,
            ']' => TokenKind::RBracket,
            ';' => TokenKind::Semicolon,
            ',' => TokenKind::Comma,
            '~' => TokenKind::Tilde,
            ':' => TokenKind::Colon,
            '.' => {
                if self.check('.') && self.pos < self.source.len() && self.source[self.pos] == '.' {
                    self.pos += 1; // skip second .
                    TokenKind::DotDotDot
                } else {
                    TokenKind::Dot
                }
            }
            '?' => {
                if self.check('.') {
                    TokenKind::QuestionDot
                } else if self.check('?') {
                    if self.check('=') {
                        TokenKind::QuestionQuestionEq
                    } else {
                        TokenKind::QuestionQuestion
                    }
                } else {
                    TokenKind::Question
                }
            }
            '+' => {
                if self.check('+') { TokenKind::PlusPlus }
                else if self.check('=') { TokenKind::PlusEq }
                else { TokenKind::Plus }
            }
            '-' => {
                if self.check('-') { TokenKind::MinusMinus }
                else if self.check('=') { TokenKind::MinusEq }
                else { TokenKind::Minus }
            }
            '*' => {
                if self.check('*') {
                    if self.check('=') { TokenKind::StarStarEq }
                    else { TokenKind::StarStar }
                } else if self.check('=') { TokenKind::StarEq }
                else { TokenKind::Star }
            }
            '/' => {
                if self.check('=') { TokenKind::SlashEq }
                else { TokenKind::Slash }
            }
            '%' => {
                if self.check('=') { TokenKind::PercentEq }
                else { TokenKind::Percent }
            }
            '&' => {
                if self.check('&') {
                    if self.check('=') { TokenKind::AmpAmpEq }
                    else { TokenKind::AmpAmp }
                } else if self.check('=') { TokenKind::AmpEq }
                else { TokenKind::Amp }
            }
            '|' => {
                if self.check('|') {
                    if self.check('=') { TokenKind::PipePipeEq }
                    else { TokenKind::PipePipe }
                } else if self.check('=') { TokenKind::PipeEq }
                else { TokenKind::Pipe }
            }
            '^' => {
                if self.check('=') { TokenKind::CaretEq }
                else { TokenKind::Caret }
            }
            '!' => {
                if self.check('=') {
                    if self.check('=') { TokenKind::BangEqEq }
                    else { TokenKind::BangEq }
                } else { TokenKind::Bang }
            }
            '=' => {
                if self.check('=') {
                    if self.check('=') { TokenKind::EqEqEq }
                    else { TokenKind::EqEq }
                } else if self.check('>') { TokenKind::Arrow }
                else { TokenKind::Eq }
            }
            '<' => {
                if self.check('<') {
                    if self.check('=') { TokenKind::LtLtEq }
                    else { TokenKind::LtLt }
                } else if self.check('=') { TokenKind::LtEq }
                else { TokenKind::Lt }
            }
            '>' => {
                if self.check('>') {
                    if self.check('>') {
                        if self.check('=') { TokenKind::GtGtGtEq }
                        else { TokenKind::GtGtGt }
                    } else if self.check('=') { TokenKind::GtGtEq }
                    else { TokenKind::GtGt }
                } else if self.check('=') { TokenKind::GtEq }
                else { TokenKind::Gt }
            }
            _ => return Err(format!("Unexpected character '{}' at line {}", ch, self.line)),
        };

        Ok(self.make_token(kind, start))
    }

    // -- Helpers --

    fn current(&self) -> char {
        self.source[self.pos]
    }

    fn check(&mut self, expected: char) -> bool {
        if self.pos < self.source.len() && self.source[self.pos] == expected {
            self.pos += 1;
            true
        } else {
            false
        }
    }

    fn peek_is_digit(&self) -> bool {
        self.pos + 1 < self.source.len() && self.source[self.pos + 1].is_ascii_digit()
    }

    fn skip_whitespace_and_comments(&mut self) {
        while self.pos < self.source.len() {
            let ch = self.source[self.pos];
            match ch {
                ' ' | '\t' | '\r' => self.pos += 1,
                '\n' => {
                    self.pos += 1;
                    self.line += 1;
                }
                '/' if self.pos + 1 < self.source.len() => {
                    if self.source[self.pos + 1] == '/' {
                        // Line comment
                        self.pos += 2;
                        while self.pos < self.source.len() && self.source[self.pos] != '\n' {
                            self.pos += 1;
                        }
                    } else if self.source[self.pos + 1] == '*' {
                        // Block comment
                        self.pos += 2;
                        while self.pos + 1 < self.source.len() {
                            if self.source[self.pos] == '\n' {
                                self.line += 1;
                            }
                            if self.source[self.pos] == '*' && self.source[self.pos + 1] == '/' {
                                self.pos += 2;
                                break;
                            }
                            self.pos += 1;
                        }
                    } else {
                        break;
                    }
                }
                _ => break,
            }
        }
    }

    fn make_token(&self, kind: TokenKind, start: usize) -> Token {
        Token {
            kind,
            span: Span {
                start: start as u32,
                end: self.pos as u32,
                line: self.line,
            },
        }
    }
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '$'
}

fn is_ident_part(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '$'
}
