use crate::token::{Token, TokenKind};

pub struct Lexer {
    src: Vec<char>,
    pos: usize,
    line: u32,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        let mut src: Vec<char> = source.chars().collect();
        if src.first() == Some(&'\u{FEFF}') {
            src.remove(0);
        }
        Lexer { src, pos: 0, line: 1 }
    }

    pub fn tokenize(&mut self) -> Result<Vec<Token>, String> {
        // Skip shebang line
        if self.peek_str("#!") {
            while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                self.pos += 1;
            }
        }
        let mut tokens = Vec::new();
        let mut last_was_newline = false;
        loop {
            self.skip_spaces_and_comments();
            if self.pos >= self.src.len() {
                tokens.push(Token { kind: TokenKind::Eof, line: self.line });
                break;
            }
            let ch = self.src[self.pos];

            // Newlines are significant in Ruby (statement terminators)
            if ch == '\n' || ch == ';' {
                if ch == '\n' { self.line += 1; }
                self.pos += 1;
                if !last_was_newline {
                    // Don't emit newline after operators or opening delimiters
                    if let Some(last) = tokens.last() {
                        if !Self::suppresses_newline(&last.kind) {
                            tokens.push(Token { kind: TokenKind::Newline, line: self.line });
                            last_was_newline = true;
                        }
                    }
                }
                continue;
            }
            last_was_newline = false;

            // Backslash line continuation
            if ch == '\\' && self.pos + 1 < self.src.len() && self.src[self.pos + 1] == '\n' {
                self.pos += 2;
                self.line += 1;
                continue;
            }

            let tok = self.next_token()?;
            tokens.push(tok);
        }
        Ok(tokens)
    }

    fn suppresses_newline(kind: &TokenKind) -> bool {
        matches!(kind,
            TokenKind::Plus | TokenKind::Minus | TokenKind::Star | TokenKind::Slash |
            TokenKind::Percent | TokenKind::StarStar |
            TokenKind::Eq | TokenKind::PlusEq | TokenKind::MinusEq |
            TokenKind::StarEq | TokenKind::SlashEq | TokenKind::PercentEq |
            TokenKind::AmpAmpEq | TokenKind::PipePipeEq |
            TokenKind::EqEq | TokenKind::BangEq | TokenKind::Lt | TokenKind::Gt |
            TokenKind::LtEq | TokenKind::GtEq | TokenKind::Spaceship |
            TokenKind::AmpAmp | TokenKind::PipePipe |
            TokenKind::Amp | TokenKind::Pipe | TokenKind::Caret |
            TokenKind::LtLt | TokenKind::GtGt |
            TokenKind::Dot | TokenKind::ColonColon | TokenKind::FatArrow | TokenKind::Arrow |
            TokenKind::Comma | TokenKind::LParen | TokenKind::LBrace | TokenKind::LBracket |
            TokenKind::Colon | TokenKind::Do | TokenKind::Then |
            TokenKind::Newline | TokenKind::And | TokenKind::Or | TokenKind::Not
        )
    }

    fn skip_spaces_and_comments(&mut self) {
        while self.pos < self.src.len() {
            let ch = self.src[self.pos];
            if ch == ' ' || ch == '\t' || ch == '\r' {
                self.pos += 1;
            } else if ch == '#' {
                // Single-line comment
                while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    fn peek_str(&self, s: &str) -> bool {
        let chars: Vec<char> = s.chars().collect();
        if self.pos + chars.len() > self.src.len() { return false; }
        self.src[self.pos..self.pos + chars.len()] == chars[..]
    }

    fn advance(&mut self) -> char {
        let ch = self.src[self.pos];
        self.pos += 1;
        ch
    }

    fn peek(&self) -> Option<char> {
        if self.pos < self.src.len() { Some(self.src[self.pos]) } else { None }
    }

    fn peek2(&self) -> Option<char> {
        if self.pos + 1 < self.src.len() { Some(self.src[self.pos + 1]) } else { None }
    }

    fn next_token(&mut self) -> Result<Token, String> {
        let line = self.line;
        let ch = self.advance();

        match ch {
            // String literals
            '"' => self.read_double_quoted_string(line),
            '\'' => self.read_single_quoted_string(line),

            // Numbers
            '0'..='9' => self.read_number(ch, line),

            // Symbols
            ':' => {
                if self.pos < self.src.len() {
                    let next = self.src[self.pos];
                    if next.is_alphabetic() || next == '_' {
                        let name = self.read_identifier_str();
                        return Ok(Token { kind: TokenKind::Symbol(name), line });
                    }
                    if next == ':' {
                        self.pos += 1;
                        return Ok(Token { kind: TokenKind::ColonColon, line });
                    }
                }
                Ok(Token { kind: TokenKind::Colon, line })
            }

            // Instance variables @name, class variables @@name
            '@' => {
                if self.peek() == Some('@') {
                    self.pos += 1;
                    let name = self.read_identifier_str();
                    Ok(Token { kind: TokenKind::ClassVar(name), line })
                } else {
                    let name = self.read_identifier_str();
                    Ok(Token { kind: TokenKind::InstanceVar(name), line })
                }
            }

            // Global variables $name
            '$' => {
                let name = self.read_identifier_str();
                Ok(Token { kind: TokenKind::GlobalVar(name), line })
            }

            // Operators
            '+' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::PlusEq, line }) }
                else { Ok(Token { kind: TokenKind::Plus, line }) }
            }
            '-' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::MinusEq, line }) }
                else if self.peek() == Some('>') { self.pos += 1; Ok(Token { kind: TokenKind::Arrow, line }) }
                else { Ok(Token { kind: TokenKind::Minus, line }) }
            }
            '*' => {
                if self.peek() == Some('*') {
                    self.pos += 1;
                    Ok(Token { kind: TokenKind::StarStar, line })
                } else if self.peek() == Some('=') {
                    self.pos += 1;
                    Ok(Token { kind: TokenKind::StarEq, line })
                } else {
                    Ok(Token { kind: TokenKind::Star, line })
                }
            }
            '/' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::SlashEq, line }) }
                else { Ok(Token { kind: TokenKind::Slash, line }) }
            }
            '%' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::PercentEq, line }) }
                else { Ok(Token { kind: TokenKind::Percent, line }) }
            }
            '=' => {
                if self.peek() == Some('=') {
                    self.pos += 1;
                    if self.peek() == Some('=') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::EqEqEq, line })
                    } else {
                        Ok(Token { kind: TokenKind::EqEq, line })
                    }
                } else if self.peek() == Some('>') {
                    self.pos += 1;
                    Ok(Token { kind: TokenKind::FatArrow, line })
                } else if self.peek_str("=begin") {
                    // Multi-line comment =begin ... =end
                    while self.pos < self.src.len() {
                        if self.src[self.pos] == '\n' {
                            self.line += 1;
                            self.pos += 1;
                            if self.peek_str("=end") {
                                self.pos += 4;
                                break;
                            }
                        } else {
                            self.pos += 1;
                        }
                    }
                    // Return next token after comment
                    self.skip_spaces_and_comments();
                    if self.pos >= self.src.len() {
                        Ok(Token { kind: TokenKind::Eof, line: self.line })
                    } else {
                        self.next_token()
                    }
                } else {
                    Ok(Token { kind: TokenKind::Eq, line })
                }
            }
            '!' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::BangEq, line }) }
                else { Ok(Token { kind: TokenKind::Bang, line }) }
            }
            '<' => {
                if self.peek() == Some('=') {
                    self.pos += 1;
                    if self.peek() == Some('>') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::Spaceship, line })
                    } else {
                        Ok(Token { kind: TokenKind::LtEq, line })
                    }
                } else if self.peek() == Some('<') {
                    self.pos += 1;
                    if self.peek() == Some('=') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::LtLtEq, line })
                    } else {
                        Ok(Token { kind: TokenKind::LtLt, line })
                    }
                } else {
                    Ok(Token { kind: TokenKind::Lt, line })
                }
            }
            '>' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::GtEq, line }) }
                else if self.peek() == Some('>') {
                    self.pos += 1;
                    if self.peek() == Some('=') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::GtGtEq, line })
                    } else {
                        Ok(Token { kind: TokenKind::GtGt, line })
                    }
                }
                else { Ok(Token { kind: TokenKind::Gt, line }) }
            }
            '&' => {
                if self.peek() == Some('&') {
                    self.pos += 1;
                    if self.peek() == Some('=') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::AmpAmpEq, line })
                    } else {
                        Ok(Token { kind: TokenKind::AmpAmp, line })
                    }
                } else if self.peek() == Some('=') {
                    self.pos += 1;
                    Ok(Token { kind: TokenKind::AmpEq, line })
                } else {
                    Ok(Token { kind: TokenKind::Amp, line })
                }
            }
            '|' => {
                if self.peek() == Some('|') {
                    self.pos += 1;
                    if self.peek() == Some('=') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::PipePipeEq, line })
                    } else {
                        Ok(Token { kind: TokenKind::PipePipe, line })
                    }
                } else if self.peek() == Some('=') {
                    self.pos += 1;
                    Ok(Token { kind: TokenKind::PipeEq, line })
                } else {
                    Ok(Token { kind: TokenKind::Pipe, line })
                }
            }
            '^' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::CaretEq, line }) }
                else { Ok(Token { kind: TokenKind::Caret, line }) }
            }
            '~' => Ok(Token { kind: TokenKind::Tilde, line }),
            '?' => Ok(Token { kind: TokenKind::Question, line }),

            '.' => {
                if self.peek() == Some('.') {
                    self.pos += 1;
                    if self.peek() == Some('.') {
                        self.pos += 1;
                        Ok(Token { kind: TokenKind::DotDotDot, line })
                    } else {
                        Ok(Token { kind: TokenKind::DotDot, line })
                    }
                } else {
                    Ok(Token { kind: TokenKind::Dot, line })
                }
            }

            // Delimiters
            '(' => Ok(Token { kind: TokenKind::LParen, line }),
            ')' => Ok(Token { kind: TokenKind::RParen, line }),
            '{' => Ok(Token { kind: TokenKind::LBrace, line }),
            '}' => Ok(Token { kind: TokenKind::RBrace, line }),
            '[' => Ok(Token { kind: TokenKind::LBracket, line }),
            ']' => Ok(Token { kind: TokenKind::RBracket, line }),
            ',' => Ok(Token { kind: TokenKind::Comma, line }),

            // Identifiers / keywords
            'a'..='z' | '_' => {
                let mut name = String::new();
                name.push(ch);
                while self.pos < self.src.len() {
                    let c = self.src[self.pos];
                    if c.is_alphanumeric() || c == '_' || c == '?' || c == '!' {
                        name.push(c);
                        self.pos += 1;
                    } else {
                        break;
                    }
                }
                Ok(Token { kind: Self::keyword_or_ident(name), line })
            }

            // Constants (uppercase start)
            'A'..='Z' => {
                let mut name = String::new();
                name.push(ch);
                while self.pos < self.src.len() {
                    let c = self.src[self.pos];
                    if c.is_alphanumeric() || c == '_' {
                        name.push(c);
                        self.pos += 1;
                    } else {
                        break;
                    }
                }
                Ok(Token { kind: TokenKind::Constant(name), line })
            }

            _ => Err(format!("Unexpected character '{}' at line {}", ch, line)),
        }
    }

    fn read_identifier_str(&mut self) -> String {
        let mut name = String::new();
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c.is_alphanumeric() || c == '_' || c == '?' || c == '!' {
                name.push(c);
                self.pos += 1;
            } else {
                break;
            }
        }
        name
    }

    fn read_number(&mut self, first: char, line: u32) -> Result<Token, String> {
        let mut s = String::new();
        s.push(first);
        let mut has_dot = false;
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c.is_ascii_digit() || c == '_' {
                if c != '_' { s.push(c); }
                self.pos += 1;
            } else if c == '.' && !has_dot {
                // Check that next char is a digit (not .. range)
                if self.pos + 1 < self.src.len() && self.src[self.pos + 1].is_ascii_digit() {
                    has_dot = true;
                    s.push('.');
                    self.pos += 1;
                } else {
                    break;
                }
            } else {
                break;
            }
        }
        let val: f64 = s.parse().map_err(|_| format!("Invalid number at line {}", line))?;
        Ok(Token { kind: TokenKind::Number(val), line })
    }

    fn read_double_quoted_string(&mut self, line: u32) -> Result<Token, String> {
        let mut s = String::new();
        let mut has_interpolation = false;
        let mut parts: Vec<(bool, String)> = Vec::new(); // (is_expr, content)

        while self.pos < self.src.len() && self.src[self.pos] != '"' {
            if self.src[self.pos] == '\\' && self.pos + 1 < self.src.len() {
                self.pos += 1;
                let esc = self.advance();
                match esc {
                    'n' => s.push('\n'),
                    't' => s.push('\t'),
                    'r' => s.push('\r'),
                    '\\' => s.push('\\'),
                    '"' => s.push('"'),
                    '#' => s.push('#'),
                    _ => { s.push('\\'); s.push(esc); }
                }
            } else if self.src[self.pos] == '#' && self.pos + 1 < self.src.len() && self.src[self.pos + 1] == '{' {
                // String interpolation #{...}
                has_interpolation = true;
                if !s.is_empty() {
                    parts.push((false, std::mem::take(&mut s)));
                }
                self.pos += 2; // skip #{
                let mut depth = 1;
                let mut expr = String::new();
                while self.pos < self.src.len() && depth > 0 {
                    if self.src[self.pos] == '{' { depth += 1; }
                    if self.src[self.pos] == '}' { depth -= 1; }
                    if depth > 0 {
                        expr.push(self.src[self.pos]);
                    }
                    if self.src[self.pos] == '\n' { self.line += 1; }
                    self.pos += 1;
                }
                parts.push((true, expr));
            } else {
                if self.src[self.pos] == '\n' { self.line += 1; }
                s.push(self.src[self.pos]);
                self.pos += 1;
            }
        }
        if self.pos < self.src.len() { self.pos += 1; } // closing "

        if has_interpolation {
            if !s.is_empty() {
                parts.push((false, s));
            }
            // Encode interpolated parts as special string: \x01lit\x02expr\x01lit...
            let mut encoded = String::from("\x01");
            for (is_expr, content) in &parts {
                if *is_expr {
                    encoded.push('\x02');
                    encoded.push_str(content);
                    encoded.push('\x01');
                } else {
                    encoded.push_str(content);
                }
            }
            Ok(Token { kind: TokenKind::Str(encoded), line })
        } else {
            Ok(Token { kind: TokenKind::Str(s), line })
        }
    }

    fn read_single_quoted_string(&mut self, line: u32) -> Result<Token, String> {
        let mut s = String::new();
        while self.pos < self.src.len() && self.src[self.pos] != '\'' {
            if self.src[self.pos] == '\\' && self.pos + 1 < self.src.len() {
                let next = self.src[self.pos + 1];
                if next == '\'' || next == '\\' {
                    self.pos += 1;
                    s.push(self.src[self.pos]);
                    self.pos += 1;
                } else {
                    s.push(self.src[self.pos]);
                    self.pos += 1;
                }
            } else {
                if self.src[self.pos] == '\n' { self.line += 1; }
                s.push(self.src[self.pos]);
                self.pos += 1;
            }
        }
        if self.pos < self.src.len() { self.pos += 1; }
        Ok(Token { kind: TokenKind::Str(s), line })
    }

    fn keyword_or_ident(name: String) -> TokenKind {
        match name.as_str() {
            "if" => TokenKind::If,
            "unless" => TokenKind::Unless,
            "elsif" => TokenKind::Elsif,
            "else" => TokenKind::Else,
            "then" => TokenKind::Then,
            "end" => TokenKind::End,
            "while" => TokenKind::While,
            "until" => TokenKind::Until,
            "for" => TokenKind::For,
            "in" => TokenKind::In,
            "do" => TokenKind::Do,
            "case" => TokenKind::Case,
            "when" => TokenKind::When,
            "break" => TokenKind::Break,
            "next" => TokenKind::Next,
            "return" => TokenKind::Return,
            "def" => TokenKind::Def,
            "class" => TokenKind::Class,
            "module" => TokenKind::Module,
            "self" => TokenKind::Self_,
            "super" => TokenKind::Super,
            "yield" => TokenKind::Yield,
            "block_given?" => TokenKind::Block_given,
            "begin" => TokenKind::Begin,
            "rescue" => TokenKind::Rescue,
            "ensure" => TokenKind::Ensure,
            "raise" => TokenKind::Raise,
            "retry" => TokenKind::Retry,
            "nil" => TokenKind::Nil,
            "true" => TokenKind::True,
            "false" => TokenKind::False,
            "and" => TokenKind::And,
            "or" => TokenKind::Or,
            "not" => TokenKind::Not,
            "require" => TokenKind::Require,
            "include" => TokenKind::Include,
            "extend" => TokenKind::Extend,
            "attr_reader" => TokenKind::Attr_reader,
            "attr_writer" => TokenKind::Attr_writer,
            "attr_accessor" => TokenKind::Attr_accessor,
            "lambda" => TokenKind::Lambda,
            "puts" => TokenKind::Puts,
            "print" => TokenKind::Print,
            "p" => TokenKind::P,
            _ => TokenKind::Identifier(name),
        }
    }
}
