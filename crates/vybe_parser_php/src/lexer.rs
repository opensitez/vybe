use crate::token::{Token, TokenKind};

pub struct Lexer {
    src: Vec<char>,
    pos: usize,
    line: u32,
    pending_tokens: Vec<Token>,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        let mut src: Vec<char> = source.chars().collect();
        // Remove UTF-8 BOM if present
        if src.first() == Some(&'\u{FEFF}') {
            src.remove(0);
        }
        Lexer { src, pos: 0, line: 1, pending_tokens: Vec::new() }
    }

    pub fn tokenize(&mut self) -> Result<Vec<Token>, String> {
        // Skip <?php opening tag if present
        self.skip_open_tag();
        let mut tokens = Vec::new();
        loop {
            self.skip_whitespace_and_comments();
            if self.pos >= self.src.len() {
                tokens.push(Token { kind: TokenKind::Eof, line: self.line });
                break;
            }
            // Skip close tag
            if self.peek_str("?>") {
                self.pos += 2;
                continue;
            }
            // Drain any pending tokens from interpolated strings
            if !self.pending_tokens.is_empty() {
                let pending = std::mem::take(&mut self.pending_tokens);
                tokens.extend(pending);
                continue;
            }
            let tok = self.next_token()?;
            tokens.push(tok);
        }
        Ok(tokens)
    }

    fn skip_open_tag(&mut self) {
        // skip optional whitespace before <?php
        while self.pos < self.src.len() && self.src[self.pos].is_whitespace() {
            if self.src[self.pos] == '\n' { self.line += 1; }
            self.pos += 1;
        }
        if self.peek_str("<?php") {
            self.pos += 5;
            // must be followed by whitespace or EOF
        } else if self.peek_str("<?") {
            self.pos += 2;
        }
    }

    fn skip_whitespace_and_comments(&mut self) {
        while self.pos < self.src.len() {
            let ch = self.src[self.pos];
            if ch == '\n' {
                self.line += 1;
                self.pos += 1;
            } else if ch.is_whitespace() {
                self.pos += 1;
            } else if self.peek_str("//") || self.peek_str("#") {
                // single-line comment
                while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else if self.peek_str("/*") {
                self.pos += 2;
                while self.pos + 1 < self.src.len() {
                    if self.src[self.pos] == '\n' { self.line += 1; }
                    if self.src[self.pos] == '*' && self.src[self.pos + 1] == '/' {
                        self.pos += 2;
                        break;
                    }
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    fn peek_str(&self, s: &str) -> bool {
        let bytes: Vec<char> = s.chars().collect();
        if self.pos + bytes.len() > self.src.len() { return false; }
        self.src[self.pos..self.pos + bytes.len()] == bytes[..]
    }

    fn current(&self) -> char {
        self.src[self.pos]
    }

    fn advance(&mut self) -> char {
        let ch = self.src[self.pos];
        if ch == '\n' { self.line += 1; }
        self.pos += 1;
        ch
    }

    fn next_token(&mut self) -> Result<Token, String> {
        let line = self.line;
        let ch = self.current();

        // Variable: $name
        if ch == '$' {
            self.pos += 1;
            let name = self.read_identifier();
            return Ok(Token { kind: TokenKind::Variable(name), line });
        }

        // String literals — double-quoted with interpolation
        if ch == '"' {
            self.pos += 1;
            let toks = self.read_double_quoted_interpolated()?;
            if toks.len() == 1 {
                return Ok(toks.into_iter().next().unwrap());
            }
            // Multiple tokens (interpolated) — return first, queue rest
            let mut iter = toks.into_iter();
            let first = iter.next().unwrap();
            self.pending_tokens = iter.collect();
            return Ok(first);
        }
        if ch == '\'' {
            self.pos += 1;
            let s = self.read_single_quoted_string()?;
            return Ok(Token { kind: TokenKind::Str(s), line });
        }

        // Heredoc/Nowdoc: <<<LABEL or <<<'LABEL'
        if ch == '<' && self.pos + 2 < self.src.len() && self.src[self.pos + 1] == '<' && self.src[self.pos + 2] == '<' {
            self.pos += 3;
            // Skip optional whitespace
            while self.pos < self.src.len() && self.src[self.pos] == ' ' { self.pos += 1; }
            let is_nowdoc = self.pos < self.src.len() && self.src[self.pos] == '\'';
            if is_nowdoc { self.pos += 1; }
            // Read label
            let mut label = String::new();
            while self.pos < self.src.len() && (self.src[self.pos].is_alphanumeric() || self.src[self.pos] == '_') {
                label.push(self.src[self.pos]); self.pos += 1;
            }
            if is_nowdoc && self.pos < self.src.len() && self.src[self.pos] == '\'' { self.pos += 1; }
            // Skip to next line
            while self.pos < self.src.len() && self.src[self.pos] != '\n' { self.pos += 1; }
            if self.pos < self.src.len() { self.pos += 1; self.line += 1; }
            // Read body until line starting with label
            let mut body = String::new();
            loop {
                if self.pos >= self.src.len() { break; }
                // Check if current line starts with the label
                let remaining: String = self.src[self.pos..].iter().collect();
                if remaining.starts_with(&label) {
                    let after = self.pos + label.len();
                    if after >= self.src.len() || self.src[after] == ';' || self.src[after] == '\n' || self.src[after] == '\r' {
                        self.pos = after;
                        // Do NOT consume ; — the parser needs it for statement termination
                        break;
                    }
                }
                let c = self.src[self.pos]; self.pos += 1;
                if c == '\n' { self.line += 1; }
                body.push(c);
            }
            // Remove trailing newline
            if body.ends_with('\n') { body.pop(); }
            if body.ends_with('\r') { body.pop(); }
            if is_nowdoc {
                return Ok(Token { kind: TokenKind::Str(body), line });
            } else {
                // Heredoc — treat like double-quoted (with interpolation)
                // For now, return as plain string (interpolation TODO)
                return Ok(Token { kind: TokenKind::Str(body), line });
            }
        }

        // Numbers
        if ch.is_ascii_digit() || (ch == '.' && self.pos + 1 < self.src.len() && self.src[self.pos + 1].is_ascii_digit()) {
            return self.read_number(line);
        }

        // Identifiers / keywords
        if ch.is_alphabetic() || ch == '_' {
            let id = self.read_identifier();
            let kind = Self::keyword_or_ident(id);
            return Ok(Token { kind, line });
        }

        // Backslash (namespace separator)
        if ch == '\\' {
            self.pos += 1;
            return Ok(Token { kind: TokenKind::Backslash, line });
        }

        // @ (error suppressor — ignore semantically)
        if ch == '@' {
            self.pos += 1;
            return Ok(Token { kind: TokenKind::At, line });
        }

        // Operators and punctuation
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
            '^' => TokenKind::Caret,
            '+' => {
                if self.pos < self.src.len() && self.src[self.pos] == '+' { self.pos += 1; TokenKind::PlusPlus }
                else if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::PlusEq }
                else { TokenKind::Plus }
            }
            '-' => {
                if self.pos < self.src.len() && self.src[self.pos] == '-' { self.pos += 1; TokenKind::MinusMinus }
                else if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::MinusEq }
                else if self.pos < self.src.len() && self.src[self.pos] == '>' { self.pos += 1; TokenKind::Arrow }
                else { TokenKind::Minus }
            }
            '*' => {
                if self.pos < self.src.len() && self.src[self.pos] == '*' { self.pos += 1; TokenKind::StarStar }
                else if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::StarEq }
                else { TokenKind::Star }
            }
            '/' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::SlashEq }
                else { TokenKind::Slash }
            }
            '%' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::PercentEq }
                else { TokenKind::Percent }
            }
            '.' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::DotEq }
                else if self.pos + 1 < self.src.len() && self.src[self.pos] == '.' && self.src[self.pos + 1] == '.' {
                    self.pos += 2; TokenKind::Ellipsis
                }
                else { TokenKind::Dot }
            }
            '=' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::EqEqEq }
                    else { TokenKind::EqEq }
                } else if self.pos < self.src.len() && self.src[self.pos] == '>' {
                    self.pos += 1; TokenKind::FatArrow
                } else {
                    TokenKind::Eq
                }
            }
            '!' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::BangEqEq }
                    else { TokenKind::BangEq }
                } else {
                    TokenKind::Bang
                }
            }
            '<' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '>' { self.pos += 1; TokenKind::Spaceship }
                    else { TokenKind::LtEq }
                } else if self.pos < self.src.len() && self.src[self.pos] == '<' {
                    self.pos += 1; TokenKind::LtLt
                } else {
                    TokenKind::Lt
                }
            }
            '>' => {
                if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::GtEq }
                else if self.pos < self.src.len() && self.src[self.pos] == '>' { self.pos += 1; TokenKind::GtGt }
                else { TokenKind::Gt }
            }
            '&' => {
                if self.pos < self.src.len() && self.src[self.pos] == '&' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::AmpAmpEq }
                    else { TokenKind::AmpAmp }
                } else {
                    TokenKind::Amp
                }
            }
            '|' => {
                if self.pos < self.src.len() && self.src[self.pos] == '|' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::PipePipeEq }
                    else { TokenKind::PipePipe }
                } else {
                    TokenKind::Pipe
                }
            }
            '?' => {
                if self.pos < self.src.len() && self.src[self.pos] == '?' {
                    self.pos += 1;
                    if self.pos < self.src.len() && self.src[self.pos] == '=' { self.pos += 1; TokenKind::QuestionQuestionEq }
                    else { TokenKind::QuestionQuestion }
                } else if self.pos < self.src.len() && self.src[self.pos] == '-'
                    && self.pos + 1 < self.src.len() && self.src[self.pos + 1] == '>' {
                    self.pos += 2; TokenKind::NullsafeArrow
                } else {
                    TokenKind::Question
                }
            }
            ':' => {
                if self.pos < self.src.len() && self.src[self.pos] == ':' { self.pos += 1; TokenKind::ColonColon }
                else { TokenKind::Colon }
            }
            _ => return Err(format!("Unexpected character {:?} at line {}", ch, self.line)),
        };

        Ok(Token { kind, line })
    }

    fn read_identifier(&mut self) -> String {
        let mut s = String::new();
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c.is_alphanumeric() || c == '_' {
                s.push(c);
                self.pos += 1;
            } else {
                break;
            }
        }
        s
    }

    fn read_number(&mut self, line: u32) -> Result<Token, String> {
        let start = self.pos;
        // Hex
        if self.pos + 1 < self.src.len() && self.src[self.pos] == '0'
            && (self.src[self.pos + 1] == 'x' || self.src[self.pos + 1] == 'X') {
            self.pos += 2;
            while self.pos < self.src.len() && self.src[self.pos].is_ascii_hexdigit() {
                self.pos += 1;
            }
            let hex: String = self.src[start + 2..self.pos].iter().collect();
            let n = u64::from_str_radix(&hex, 16).map_err(|e| e.to_string())?;
            return Ok(Token { kind: TokenKind::Number(n as f64), line });
        }
        // Octal 0b
        if self.pos + 1 < self.src.len() && self.src[self.pos] == '0'
            && (self.src[self.pos + 1] == 'b' || self.src[self.pos + 1] == 'B') {
            self.pos += 2;
            while self.pos < self.src.len() && (self.src[self.pos] == '0' || self.src[self.pos] == '1') {
                self.pos += 1;
            }
            let bin: String = self.src[start + 2..self.pos].iter().collect();
            let n = u64::from_str_radix(&bin, 2).map_err(|e| e.to_string())?;
            return Ok(Token { kind: TokenKind::Number(n as f64), line });
        }
        // Decimal / float
        while self.pos < self.src.len() && self.src[self.pos].is_ascii_digit() {
            self.pos += 1;
        }
        if self.pos < self.src.len() && self.src[self.pos] == '.' {
            self.pos += 1;
            while self.pos < self.src.len() && self.src[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
        }
        if self.pos < self.src.len() && (self.src[self.pos] == 'e' || self.src[self.pos] == 'E') {
            self.pos += 1;
            if self.pos < self.src.len() && (self.src[self.pos] == '+' || self.src[self.pos] == '-') {
                self.pos += 1;
            }
            while self.pos < self.src.len() && self.src[self.pos].is_ascii_digit() {
                self.pos += 1;
            }
        }
        // Strip _ separators (PHP 7.4+)
        let raw: String = self.src[start..self.pos].iter().filter(|&&c| c != '_').collect();
        let n: f64 = raw.parse().map_err(|e: std::num::ParseFloatError| e.to_string())?;
        Ok(Token { kind: TokenKind::Number(n), line })
    }

    /// Lex a double-quoted string with interpolation support.
    /// "Hello $name!" → [Str("Hello "), Dot, Variable("name"), Dot, Str("!")]
    /// "plain string"  → [Str("plain string")]
    fn read_double_quoted_interpolated(&mut self) -> Result<Vec<Token>, String> {
        let line = self.line;
        let mut parts: Vec<Token> = Vec::new();
        let mut buf = String::new();

        while self.pos < self.src.len() {
            let ch = self.src[self.pos];
            if ch == '"' { self.pos += 1; break; }

            if ch == '\\' && self.pos + 1 < self.src.len() {
                self.pos += 1;
                let esc = self.src[self.pos]; self.pos += 1;
                match esc {
                    'n' => buf.push('\n'),
                    't' => buf.push('\t'),
                    'r' => buf.push('\r'),
                    '\\' => buf.push('\\'),
                    '"' => buf.push('"'),
                    '$' => buf.push('$'),
                    '0'..='9' => {
                        let mut oct = esc.to_string();
                        for _ in 0..2 {
                            if self.pos < self.src.len() && self.src[self.pos].is_ascii_digit() {
                                oct.push(self.src[self.pos]); self.pos += 1;
                            }
                        }
                        let n = u8::from_str_radix(&oct, 8).unwrap_or(0);
                        buf.push(n as char);
                    }
                    _ => { buf.push('\\'); buf.push(esc); }
                }
                continue;
            }

            // Variable interpolation: $name or ${expr} or {$expr}
            if ch == '$' && self.pos + 1 < self.src.len() && (self.src[self.pos + 1].is_alphabetic() || self.src[self.pos + 1] == '_') {
                // Flush text buffer
                if !buf.is_empty() {
                    if !parts.is_empty() { parts.push(Token { kind: TokenKind::Dot, line }); }
                    parts.push(Token { kind: TokenKind::Str(std::mem::take(&mut buf)), line });
                    parts.push(Token { kind: TokenKind::Dot, line });
                }
                self.pos += 1; // skip $
                // Read variable name
                let mut name = String::new();
                while self.pos < self.src.len() && (self.src[self.pos].is_alphanumeric() || self.src[self.pos] == '_') {
                    name.push(self.src[self.pos]); self.pos += 1;
                }
                // Handle $var->prop or $var[idx]
                if !parts.is_empty() || !buf.is_empty() {
                    // already pushed Dot above
                } else {
                    // first part — no leading Dot needed
                }
                parts.push(Token { kind: TokenKind::Variable(name), line });
                continue;
            }

            // {$expr} — curly brace interpolation
            if ch == '{' && self.pos + 1 < self.src.len() && self.src[self.pos + 1] == '$' {
                if !buf.is_empty() {
                    if !parts.is_empty() { parts.push(Token { kind: TokenKind::Dot, line }); }
                    parts.push(Token { kind: TokenKind::Str(std::mem::take(&mut buf)), line });
                    parts.push(Token { kind: TokenKind::Dot, line });
                }
                self.pos += 1; // skip {
                // Lex tokens until matching }
                // Simple version: just read $varname
                self.pos += 1; // skip $
                let mut name = String::new();
                while self.pos < self.src.len() && self.src[self.pos] != '}' && (self.src[self.pos].is_alphanumeric() || self.src[self.pos] == '_') {
                    name.push(self.src[self.pos]); self.pos += 1;
                }
                if self.pos < self.src.len() && self.src[self.pos] == '}' { self.pos += 1; }
                parts.push(Token { kind: TokenKind::Variable(name), line });
                continue;
            }

            if ch == '\n' { self.line += 1; }
            buf.push(ch);
            self.pos += 1;
        }

        // Flush remaining buffer
        if !buf.is_empty() {
            if !parts.is_empty() { parts.push(Token { kind: TokenKind::Dot, line }); }
            parts.push(Token { kind: TokenKind::Str(buf), line });
        }

        // If no interpolation, return single Str token
        if parts.is_empty() {
            return Ok(vec![Token { kind: TokenKind::Str(String::new()), line }]);
        }

        Ok(parts)
    }

    fn read_single_quoted_string(&mut self) -> Result<String, String> {
        let mut s = String::new();
        while self.pos < self.src.len() {
            let ch = self.advance();
            if ch == '\'' { break; }
            if ch == '\\' && self.pos < self.src.len() {
                let esc = self.advance();
                match esc {
                    '\\' => s.push('\\'),
                    '\'' => s.push('\''),
                    _ => { s.push('\\'); s.push(esc); }
                }
            } else {
                s.push(ch);
            }
        }
        Ok(s)
    }

    fn keyword_or_ident(s: String) -> TokenKind {
        match s.to_lowercase().as_str() {
            "if" => TokenKind::If,
            "elseif" => TokenKind::ElseIf,
            "else" => TokenKind::Else,
            "while" => TokenKind::While,
            "for" => TokenKind::For,
            "foreach" => TokenKind::ForEach,
            "as" => TokenKind::As,
            "do" => TokenKind::Do,
            "switch" => TokenKind::Switch,
            "case" => TokenKind::Case,
            "default" => TokenKind::Default,
            "break" => TokenKind::Break,
            "continue" => TokenKind::Continue,
            "return" => TokenKind::Return,
            "function" => TokenKind::Function,
            "class" => TokenKind::Class,
            "extends" => TokenKind::Extends,
            "implements" => TokenKind::Implements,
            "interface" => TokenKind::Interface,
            "trait" => TokenKind::Trait,
            "new" => TokenKind::New,
            "echo" => TokenKind::Echo,
            "print" => TokenKind::Print,
            "null" => TokenKind::Null,
            "true" => TokenKind::True,
            "false" => TokenKind::False,
            "instanceof" => TokenKind::InstanceOf,
            "throw" => TokenKind::Throw,
            "try" => TokenKind::Try,
            "catch" => TokenKind::Catch,
            "finally" => TokenKind::Finally,
            "static" => TokenKind::Static,
            "public" => TokenKind::Public,
            "private" => TokenKind::Private,
            "protected" => TokenKind::Protected,
            "abstract" => TokenKind::Abstract,
            "final" => TokenKind::Final,
            "const" => TokenKind::Const,
            "match" => TokenKind::Match,
            "fn" => TokenKind::Fn,
            "use" => TokenKind::Use,
            "namespace" => TokenKind::Namespace,
            "yield" => TokenKind::Yield,
            "list" => TokenKind::List,
            "global" => TokenKind::Global,
            "enum" => TokenKind::Enum,
            "readonly" => TokenKind::Readonly,
            "and" => TokenKind::AmpAmp,
            "or" => TokenKind::PipePipe,
            "not" => TokenKind::Bang,
            _ => TokenKind::Identifier(s),
        }
    }
}
