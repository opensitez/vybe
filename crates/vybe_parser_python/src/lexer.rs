use crate::token::{Token, TokenKind};
use std::collections::VecDeque;

pub struct Lexer {
    chars: Vec<char>,
    pos: usize,
    line: u32,
    indent_stack: Vec<usize>,
    pending: VecDeque<Token>,
    at_line_start: bool,
    bracket_depth: i32,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        Self {
            chars: source.chars().collect(),
            pos: 0,
            line: 1,
            indent_stack: vec![0],
            pending: VecDeque::new(),
            at_line_start: true,
            bracket_depth: 0,
        }
    }

    pub fn tokenize(&mut self) -> Result<Vec<Token>, String> {
        let mut tokens = Vec::new();
        loop {
            let tok = self.next_token()?;
            let is_eof = tok.kind == TokenKind::Eof;
            tokens.push(tok);
            if is_eof { break; }
        }
        Ok(tokens)
    }

    fn next_token(&mut self) -> Result<Token, String> {
        // Drain pending INDENT/DEDENT/NEWLINE tokens first
        if let Some(tok) = self.pending.pop_front() {
            return Ok(tok);
        }

        // At line start: handle indentation
        if self.at_line_start {
            self.at_line_start = false;
            self.handle_indentation()?;
            if let Some(tok) = self.pending.pop_front() {
                return Ok(tok);
            }
        }

        self.skip_spaces_and_comments();

        if self.is_at_end() {
            // EOF: emit remaining DEDENTs
            while self.indent_stack.len() > 1 {
                self.indent_stack.pop();
                self.pending.push_back(Token { kind: TokenKind::Dedent, line: self.line });
            }
            if let Some(tok) = self.pending.pop_front() {
                return Ok(tok);
            }
            return Ok(Token { kind: TokenKind::Eof, line: self.line });
        }

        let ch = self.peek_char();

        // Newline
        if ch == '\n' {
            return self.lex_newline();
        }
        if ch == '\r' {
            self.advance_char();
            if self.peek_char() == '\n' { self.advance_char(); }
            return self.emit_newline();
        }

        // Explicit line continuation
        if ch == '\\' && self.peek_ahead(1) == '\n' {
            self.advance_char(); // skip backslash
            self.advance_char(); // skip newline
            self.line += 1;
            return self.next_token();
        }
        if ch == '\\' && self.peek_ahead(1) == '\r' {
            self.advance_char();
            self.advance_char();
            if self.peek_char() == '\n' { self.advance_char(); }
            self.line += 1;
            return self.next_token();
        }

        // Numbers
        if ch.is_ascii_digit() || (ch == '.' && self.peek_ahead(1).is_ascii_digit()) {
            return self.lex_number();
        }

        // Strings
        if ch == '\'' || ch == '"' {
            return self.lex_string(false, false);
        }

        // String prefixes: f, r, b, rb, br, u
        if (ch == 'f' || ch == 'F' || ch == 'r' || ch == 'R' || ch == 'b' || ch == 'B' || ch == 'u' || ch == 'U')
            && self.is_string_prefix()
        {
            return self.lex_prefixed_string();
        }

        // Identifiers and keywords
        if ch.is_alphabetic() || ch == '_' {
            return self.lex_identifier();
        }

        // Operators and delimiters
        self.lex_operator()
    }

    // --- Indentation ---

    fn handle_indentation(&mut self) -> Result<(), String> {
        // Count leading whitespace
        let mut indent = 0usize;
        while self.pos < self.chars.len() {
            match self.chars[self.pos] {
                ' ' => { indent += 1; self.pos += 1; }
                '\t' => { indent = (indent / 8 + 1) * 8; self.pos += 1; }
                _ => break,
            }
        }

        // Blank line or comment-only line: skip entirely
        if self.pos >= self.chars.len() || self.chars[self.pos] == '\n' || self.chars[self.pos] == '\r' || self.chars[self.pos] == '#' {
            return Ok(());
        }

        let current = *self.indent_stack.last().unwrap();
        if indent > current {
            self.indent_stack.push(indent);
            self.pending.push_back(Token { kind: TokenKind::Indent, line: self.line });
        } else if indent < current {
            while *self.indent_stack.last().unwrap() > indent {
                self.indent_stack.pop();
                self.pending.push_back(Token { kind: TokenKind::Dedent, line: self.line });
            }
            if *self.indent_stack.last().unwrap() != indent {
                return Err(format!("line {}: inconsistent indentation", self.line));
            }
        }
        Ok(())
    }

    // --- Newline ---

    fn lex_newline(&mut self) -> Result<Token, String> {
        self.advance_char(); // consume \n
        self.line += 1;
        self.emit_newline()
    }

    fn emit_newline(&mut self) -> Result<Token, String> {
        if self.bracket_depth > 0 {
            // Implicit line continuation: skip newline, skip indentation
            self.skip_blank_lines();
            return self.next_token();
        }
        self.at_line_start = true;
        Ok(Token { kind: TokenKind::Newline, line: self.line - 1 })
    }

    fn skip_blank_lines(&mut self) {
        while self.pos < self.chars.len() {
            let ch = self.chars[self.pos];
            if ch == ' ' || ch == '\t' {
                self.pos += 1;
            } else if ch == '\n' {
                self.pos += 1;
                self.line += 1;
            } else if ch == '\r' {
                self.pos += 1;
                if self.pos < self.chars.len() && self.chars[self.pos] == '\n' {
                    self.pos += 1;
                }
                self.line += 1;
            } else if ch == '#' {
                // Skip comment
                while self.pos < self.chars.len() && self.chars[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    // --- Spaces and comments ---

    fn skip_spaces_and_comments(&mut self) {
        while self.pos < self.chars.len() {
            let ch = self.chars[self.pos];
            if ch == ' ' || ch == '\t' {
                self.pos += 1;
            } else if ch == '#' {
                while self.pos < self.chars.len() && self.chars[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    // --- Numbers ---

    fn lex_number(&mut self) -> Result<Token, String> {
        let line = self.line;
        let start = self.pos;

        // Hex, octal, binary
        if self.peek_char() == '0' && self.pos + 1 < self.chars.len() {
            let next = self.chars[self.pos + 1];
            if next == 'x' || next == 'X' {
                self.pos += 2;
                while self.pos < self.chars.len() && (self.chars[self.pos].is_ascii_hexdigit() || self.chars[self.pos] == '_') {
                    self.pos += 1;
                }
                let s: String = self.chars[start..self.pos].iter().filter(|c| **c != '_').collect();
                let val = i64::from_str_radix(&s[2..], 16).map_err(|e| format!("line {}: bad hex literal: {}", line, e))?;
                return Ok(Token { kind: TokenKind::Int(val), line });
            }
            if next == 'o' || next == 'O' {
                self.pos += 2;
                while self.pos < self.chars.len() && ((self.chars[self.pos] >= '0' && self.chars[self.pos] <= '7') || self.chars[self.pos] == '_') {
                    self.pos += 1;
                }
                let s: String = self.chars[start..self.pos].iter().filter(|c| **c != '_').collect();
                let val = i64::from_str_radix(&s[2..], 8).map_err(|e| format!("line {}: bad octal literal: {}", line, e))?;
                return Ok(Token { kind: TokenKind::Int(val), line });
            }
            if next == 'b' || next == 'B' {
                self.pos += 2;
                while self.pos < self.chars.len() && (self.chars[self.pos] == '0' || self.chars[self.pos] == '1' || self.chars[self.pos] == '_') {
                    self.pos += 1;
                }
                let s: String = self.chars[start..self.pos].iter().filter(|c| **c != '_').collect();
                let val = i64::from_str_radix(&s[2..], 2).map_err(|e| format!("line {}: bad binary literal: {}", line, e))?;
                return Ok(Token { kind: TokenKind::Int(val), line });
            }
        }

        // Decimal integer or float
        while self.pos < self.chars.len() && (self.chars[self.pos].is_ascii_digit() || self.chars[self.pos] == '_') {
            self.pos += 1;
        }
        let mut is_float = false;
        if self.pos < self.chars.len() && self.chars[self.pos] == '.' {
            // Check it's not just a dot operator (e.g., 1.method)
            if self.pos + 1 < self.chars.len() && self.chars[self.pos + 1].is_ascii_digit() {
                is_float = true;
                self.pos += 1; // skip .
                while self.pos < self.chars.len() && (self.chars[self.pos].is_ascii_digit() || self.chars[self.pos] == '_') {
                    self.pos += 1;
                }
            } else if self.pos + 1 >= self.chars.len() || !self.chars[self.pos + 1].is_alphabetic() {
                // bare trailing dot like 1. is a float
                is_float = true;
                self.pos += 1;
            }
        }
        // Exponent
        if self.pos < self.chars.len() && (self.chars[self.pos] == 'e' || self.chars[self.pos] == 'E') {
            is_float = true;
            self.pos += 1;
            if self.pos < self.chars.len() && (self.chars[self.pos] == '+' || self.chars[self.pos] == '-') {
                self.pos += 1;
            }
            while self.pos < self.chars.len() && (self.chars[self.pos].is_ascii_digit() || self.chars[self.pos] == '_') {
                self.pos += 1;
            }
        }
        // 'j' suffix for complex numbers — treat as float
        if self.pos < self.chars.len() && (self.chars[self.pos] == 'j' || self.chars[self.pos] == 'J') {
            is_float = true;
            self.pos += 1;
        }

        let s: String = self.chars[start..self.pos].iter().filter(|c| **c != '_').collect();
        if is_float {
            let s = s.trim_end_matches(|c| c == 'j' || c == 'J');
            let val: f64 = s.parse().map_err(|e| format!("line {}: bad float: {}", line, e))?;
            Ok(Token { kind: TokenKind::Float(val), line })
        } else {
            let val: i64 = s.parse().map_err(|e| format!("line {}: bad integer: {}", line, e))?;
            Ok(Token { kind: TokenKind::Int(val), line })
        }
    }

    // --- Strings ---

    fn is_string_prefix(&self) -> bool {
        let remaining: String = self.chars[self.pos..].iter().take(3).collect();
        let lower = remaining.to_lowercase();
        // Check for string prefix followed by quote
        for prefix in &["f\"", "f'", "r\"", "r'", "b\"", "b'", "u\"", "u'",
                         "rf\"", "rf'", "rb\"", "rb'", "br\"", "br'", "fr\"", "fr'"] {
            if lower.starts_with(prefix) { return true; }
        }
        false
    }

    fn lex_prefixed_string(&mut self) -> Result<Token, String> {
        let mut is_raw = false;
        let mut is_fstring = false;
        let mut is_bytes = false;

        // Consume prefix characters
        loop {
            let ch = self.peek_char().to_ascii_lowercase();
            match ch {
                'r' => { is_raw = true; self.advance_char(); }
                'f' => { is_fstring = true; self.advance_char(); }
                'b' => { is_bytes = true; self.advance_char(); }
                'u' => { self.advance_char(); } // u prefix, no special handling needed
                _ => break,
            }
        }
        if is_fstring {
            return self.lex_fstring();
        }
        if is_bytes {
            return self.lex_string(is_raw, true);
        }
        self.lex_string(is_raw, false)
    }

    fn lex_string(&mut self, is_raw: bool, is_bytes: bool) -> Result<Token, String> {
        let line = self.line;
        let quote = self.advance_char();
        let triple = self.pos + 1 < self.chars.len()
            && self.chars[self.pos] == quote
            && self.chars[self.pos + 1] == quote;

        if triple {
            self.pos += 2; // skip second and third quotes
            return self.lex_triple_string(quote, is_raw, is_bytes, line);
        }

        let mut s = String::new();
        while self.pos < self.chars.len() && self.chars[self.pos] != quote && self.chars[self.pos] != '\n' {
            if self.chars[self.pos] == '\\' && !is_raw {
                self.pos += 1;
                s.push(self.lex_escape_char()?);
            } else {
                s.push(self.chars[self.pos]);
                self.pos += 1;
            }
        }
        if self.pos >= self.chars.len() || self.chars[self.pos] != quote {
            return Err(format!("line {}: unterminated string literal", line));
        }
        self.pos += 1; // closing quote

        if is_bytes {
            Ok(Token { kind: TokenKind::ByteStr(s.into_bytes()), line })
        } else {
            Ok(Token { kind: TokenKind::Str(s), line })
        }
    }

    fn lex_triple_string(&mut self, quote: char, is_raw: bool, is_bytes: bool, start_line: u32) -> Result<Token, String> {
        let mut s = String::new();
        while self.pos < self.chars.len() {
            if self.chars[self.pos] == quote
                && self.pos + 2 < self.chars.len()
                && self.chars[self.pos + 1] == quote
                && self.chars[self.pos + 2] == quote
            {
                self.pos += 3;
                return if is_bytes {
                    Ok(Token { kind: TokenKind::ByteStr(s.into_bytes()), line: start_line })
                } else {
                    Ok(Token { kind: TokenKind::Str(s), line: start_line })
                };
            }
            if self.chars[self.pos] == '\\' && !is_raw {
                self.pos += 1;
                s.push(self.lex_escape_char()?);
            } else {
                if self.chars[self.pos] == '\n' { self.line += 1; }
                s.push(self.chars[self.pos]);
                self.pos += 1;
            }
        }
        Err(format!("line {}: unterminated triple-quoted string", start_line))
    }

    fn lex_fstring(&mut self) -> Result<Token, String> {
        let line = self.line;
        let quote = self.advance_char();
        let triple = self.pos + 1 < self.chars.len()
            && self.chars[self.pos] == quote
            && self.chars[self.pos + 1] == quote;
        if triple {
            self.pos += 2;
        }

        // Emit FStringStart, then lex text/expr segments, then FStringEnd
        self.pending.push_back(Token { kind: TokenKind::FStringStart, line });

        let mut text = String::new();
        let _end_pattern = if triple { 3 } else { 1 };

        while self.pos < self.chars.len() {
            // Check for end of f-string
            if self.chars[self.pos] == quote {
                if triple {
                    if self.pos + 2 < self.chars.len() && self.chars[self.pos + 1] == quote && self.chars[self.pos + 2] == quote {
                        self.pos += 3;
                        break;
                    }
                    text.push(self.chars[self.pos]);
                    self.pos += 1;
                    continue;
                } else {
                    self.pos += 1;
                    break;
                }
            }
            if self.chars[self.pos] == '\n' && !triple {
                return Err(format!("line {}: unterminated f-string", line));
            }
            if self.chars[self.pos] == '\\' && !triple {
                self.pos += 1;
                text.push(self.lex_escape_char()?);
                continue;
            }
            // {{ and }} are literal braces
            if self.chars[self.pos] == '{' && self.peek_ahead(1) == '{' {
                text.push('{');
                self.pos += 2;
                continue;
            }
            if self.chars[self.pos] == '}' && self.peek_ahead(1) == '}' {
                text.push('}');
                self.pos += 2;
                continue;
            }
            // Start of interpolation
            if self.chars[self.pos] == '{' {
                if !text.is_empty() {
                    self.pending.push_back(Token { kind: TokenKind::FStringText(std::mem::take(&mut text)), line: self.line });
                }
                self.pos += 1; // skip {
                // Lex expression tokens until matching }
                let saved_bracket = self.bracket_depth;
                self.bracket_depth = 0;
                let mut brace_depth = 1i32;
                let mut expr_tokens = Vec::new();
                loop {
                    if self.is_at_end() {
                        return Err(format!("line {}: unterminated f-string expression", line));
                    }
                    if self.chars[self.pos] == '}' && brace_depth == 1 {
                        self.pos += 1;
                        break;
                    }
                    // Track brace depth for nested dicts/sets
                    if self.chars[self.pos] == '{' { brace_depth += 1; }
                    if self.chars[self.pos] == '}' { brace_depth -= 1; }

                    // Skip colon format specs at top level: f"{x:.2f}"
                    if self.chars[self.pos] == ':' && brace_depth == 1 {
                        // Everything after : until } is format spec, skip it
                        self.pos += 1;
                        let mut fmt_depth = 1i32;
                        while self.pos < self.chars.len() {
                            if self.chars[self.pos] == '{' { fmt_depth += 1; }
                            if self.chars[self.pos] == '}' {
                                fmt_depth -= 1;
                                if fmt_depth == 0 { self.pos += 1; break; }
                            }
                            self.pos += 1;
                        }
                        break;
                    }
                    // Skip !r, !s, !a conversion specs
                    if self.chars[self.pos] == '!' && brace_depth == 1 {
                        let next = self.peek_ahead(1);
                        if next == 'r' || next == 's' || next == 'a' {
                            let after = self.peek_ahead(2);
                            if after == '}' || after == ':' {
                                self.pos += 2; // skip !x
                                continue;
                            }
                        }
                    }
                    let saved_line_start = self.at_line_start;
                    self.at_line_start = false;
                    let tok = self.lex_single_token()?;
                    self.at_line_start = saved_line_start;
                    expr_tokens.push(tok);
                }
                self.bracket_depth = saved_bracket;
                // Push expression tokens into pending
                for tok in expr_tokens {
                    self.pending.push_back(tok);
                }
                continue;
            }
            if self.chars[self.pos] == '\n' { self.line += 1; }
            text.push(self.chars[self.pos]);
            self.pos += 1;
        }

        if !text.is_empty() {
            self.pending.push_back(Token { kind: TokenKind::FStringText(text), line: self.line });
        }
        self.pending.push_back(Token { kind: TokenKind::FStringEnd, line: self.line });

        // Return first pending token
        Ok(self.pending.pop_front().unwrap())
    }

    fn lex_escape_char(&mut self) -> Result<char, String> {
        if self.pos >= self.chars.len() {
            return Err(format!("line {}: unexpected end of string escape", self.line));
        }
        let ch = self.chars[self.pos];
        self.pos += 1;
        match ch {
            'n' => Ok('\n'),
            't' => Ok('\t'),
            'r' => Ok('\r'),
            '\\' => Ok('\\'),
            '\'' => Ok('\''),
            '"' => Ok('"'),
            '0' => Ok('\0'),
            'a' => Ok('\x07'),
            'b' => Ok('\x08'),
            'f' => Ok('\x0C'),
            'v' => Ok('\x0B'),
            '\n' => { self.line += 1; Ok(' ') } // line continuation in string
            'x' => {
                let h = self.take_hex_digits(2)?;
                Ok(char::from_u32(h).unwrap_or('\u{FFFD}'))
            }
            'u' => {
                let h = self.take_hex_digits(4)?;
                Ok(char::from_u32(h).unwrap_or('\u{FFFD}'))
            }
            'U' => {
                let h = self.take_hex_digits(8)?;
                Ok(char::from_u32(h).unwrap_or('\u{FFFD}'))
            }
            c if c.is_ascii_digit() && c >= '0' && c <= '7' => {
                // Octal escape
                let mut val = (c as u32) - ('0' as u32);
                for _ in 0..2 {
                    if self.pos < self.chars.len() && self.chars[self.pos] >= '0' && self.chars[self.pos] <= '7' {
                        val = val * 8 + (self.chars[self.pos] as u32 - '0' as u32);
                        self.pos += 1;
                    } else {
                        break;
                    }
                }
                Ok(char::from_u32(val).unwrap_or('\u{FFFD}'))
            }
            // Unrecognized escape: keep both the backslash and char (Python behavior)
            other => Ok(other),
        }
    }

    fn take_hex_digits(&mut self, count: usize) -> Result<u32, String> {
        let mut val = 0u32;
        for _ in 0..count {
            if self.pos >= self.chars.len() || !self.chars[self.pos].is_ascii_hexdigit() {
                return Err(format!("line {}: expected hex digit in escape", self.line));
            }
            val = val * 16 + self.chars[self.pos].to_digit(16).unwrap();
            self.pos += 1;
        }
        Ok(val)
    }

    // --- Identifiers ---

    fn lex_identifier(&mut self) -> Result<Token, String> {
        let line = self.line;
        let start = self.pos;
        while self.pos < self.chars.len() && (self.chars[self.pos].is_alphanumeric() || self.chars[self.pos] == '_') {
            self.pos += 1;
        }
        let word: String = self.chars[start..self.pos].iter().collect();
        if let Some(kw) = TokenKind::from_keyword(&word) {
            Ok(Token { kind: kw, line })
        } else {
            Ok(Token { kind: TokenKind::Identifier(word), line })
        }
    }

    // --- Operators ---

    fn lex_operator(&mut self) -> Result<Token, String> {
        let line = self.line;
        let ch = self.advance_char();
        let next = self.peek_char();

        let kind = match ch {
            '(' => { self.bracket_depth += 1; TokenKind::LParen }
            ')' => { self.bracket_depth -= 1; TokenKind::RParen }
            '[' => { self.bracket_depth += 1; TokenKind::LBracket }
            ']' => { self.bracket_depth -= 1; TokenKind::RBracket }
            '{' => { self.bracket_depth += 1; TokenKind::LBrace }
            '}' => { self.bracket_depth -= 1; TokenKind::RBrace }
            ',' => TokenKind::Comma,
            ';' => TokenKind::Semicolon,
            '~' => TokenKind::Tilde,
            ':' if next == '=' => { self.advance_char(); TokenKind::ColonEq }
            ':' => TokenKind::Colon,
            '.' if next == '.' && self.peek_ahead(1) == '.' => { self.advance_char(); self.advance_char(); TokenKind::DotDotDot }
            '.' => TokenKind::Dot,
            '+' if next == '=' => { self.advance_char(); TokenKind::PlusEq }
            '+' => TokenKind::Plus,
            '-' if next == '>' => { self.advance_char(); TokenKind::Arrow }
            '-' if next == '=' => { self.advance_char(); TokenKind::MinusEq }
            '-' => TokenKind::Minus,
            '*' if next == '*' => {
                self.advance_char();
                if self.peek_char() == '=' { self.advance_char(); TokenKind::DoubleStarEq }
                else { TokenKind::DoubleStar }
            }
            '*' if next == '=' => { self.advance_char(); TokenKind::StarEq }
            '*' => TokenKind::Star,
            '/' if next == '/' => {
                self.advance_char();
                if self.peek_char() == '=' { self.advance_char(); TokenKind::DoubleSlashEq }
                else { TokenKind::DoubleSlash }
            }
            '/' if next == '=' => { self.advance_char(); TokenKind::SlashEq }
            '/' => TokenKind::Slash,
            '%' if next == '=' => { self.advance_char(); TokenKind::PercentEq }
            '%' => TokenKind::Percent,
            '@' if next == '=' => { self.advance_char(); TokenKind::AtEq }
            '@' => TokenKind::At,
            '|' if next == '=' => { self.advance_char(); TokenKind::PipeEq }
            '|' => TokenKind::Pipe,
            '&' if next == '=' => { self.advance_char(); TokenKind::AmpEq }
            '&' => TokenKind::Amp,
            '^' if next == '=' => { self.advance_char(); TokenKind::CaretEq }
            '^' => TokenKind::Caret,
            '<' if next == '<' => {
                self.advance_char();
                if self.peek_char() == '=' { self.advance_char(); TokenKind::LtLtEq }
                else { TokenKind::LtLt }
            }
            '<' if next == '=' => { self.advance_char(); TokenKind::LtEq }
            '<' => TokenKind::Lt,
            '>' if next == '>' => {
                self.advance_char();
                if self.peek_char() == '=' { self.advance_char(); TokenKind::GtGtEq }
                else { TokenKind::GtGt }
            }
            '>' if next == '=' => { self.advance_char(); TokenKind::GtEq }
            '>' => TokenKind::Gt,
            '=' if next == '=' => { self.advance_char(); TokenKind::EqEq }
            '=' => TokenKind::Eq,
            '!' if next == '=' => { self.advance_char(); TokenKind::BangEq }
            c => return Err(format!("line {}: unexpected character '{}'", line, c)),
        };
        Ok(Token { kind, line })
    }

    /// Lex a single token for use inside f-string expressions.
    /// Does NOT handle indentation or newlines — just raw token lexing.
    fn lex_single_token(&mut self) -> Result<Token, String> {
        self.skip_spaces_and_comments();

        if self.is_at_end() {
            return Ok(Token { kind: TokenKind::Eof, line: self.line });
        }

        let ch = self.peek_char();

        if ch == '\n' || ch == '\r' {
            // Inside f-string expression, newlines are just whitespace
            if ch == '\n' { self.pos += 1; self.line += 1; }
            else {
                self.pos += 1;
                if self.peek_char() == '\n' { self.pos += 1; }
                self.line += 1;
            }
            return self.lex_single_token();
        }

        if ch.is_ascii_digit() || (ch == '.' && self.peek_ahead(1).is_ascii_digit()) {
            return self.lex_number();
        }
        if ch == '\'' || ch == '"' {
            return self.lex_string(false, false);
        }
        if ch.is_alphabetic() || ch == '_' {
            return self.lex_identifier();
        }
        self.lex_operator()
    }

    // --- Helpers ---

    fn peek_char(&self) -> char {
        if self.pos < self.chars.len() { self.chars[self.pos] } else { '\0' }
    }

    fn peek_ahead(&self, n: usize) -> char {
        let idx = self.pos + n;
        if idx < self.chars.len() { self.chars[idx] } else { '\0' }
    }

    fn advance_char(&mut self) -> char {
        if self.pos < self.chars.len() {
            let ch = self.chars[self.pos];
            self.pos += 1;
            ch
        } else {
            '\0'
        }
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.chars.len()
    }
}
