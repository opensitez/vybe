//! Generic lexer engine — tokenizes source code based on a LexerSpec.
//!
//! Produces a flat Vec<Token> from source text. Language-agnostic:
//! the same code handles Pascal's `{comments}`, Python's `#comments`,
//! JS's `/*comments*/`, etc. — driven entirely by the grammar definition.

use crate::grammar::{LexerSpec, Terminator};

/// A token produced by the lexer.
#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub kind: TokenKind,
    pub text: String,
    pub line: u32,    // 0-based
    pub col: u32,     // 0-based
}

#[derive(Debug, Clone, PartialEq)]
pub enum TokenKind {
    Keyword,       // matches a keyword from the grammar
    Ident,         // identifier (not a keyword)
    IntLit,        // integer literal
    FloatLit,      // float literal
    StringLit,     // string literal (content without delimiters)
    CharLit,       // char literal
    Operator,      // operator or punctuation
    Newline,       // significant newline (for Python-style languages)
    Indent,        // indentation increase (for Python-style languages)
    Dedent,        // indentation decrease (for Python-style languages)
    Eof,
}

/// Tokenize source code using the given lexer spec.
pub fn tokenize(source: &str, spec: &LexerSpec, terminator: &Terminator, indentation_based: bool, case_sensitive: bool) -> Vec<Token> {
    let mut lexer = Lexer::new(source, spec, terminator, indentation_based, case_sensitive);
    let mut tokens = Vec::new();
    loop {
        let tok = lexer.next_token();
        let is_eof = tok.kind == TokenKind::Eof;
        tokens.push(tok);
        if is_eof { break; }
    }
    tokens
}

struct Lexer<'a> {
    chars: Vec<char>,
    pos: usize,
    line: u32,
    col: u32,
    spec: &'a LexerSpec,
    terminator: &'a Terminator,
    indentation_based: bool,
    case_sensitive: bool,
    // Indentation tracking (for Python-style)
    indent_stack: Vec<u32>,
    pending_dedents: u32,
    at_line_start: bool,
}

impl<'a> Lexer<'a> {
    fn new(source: &str, spec: &'a LexerSpec, terminator: &'a Terminator, indentation_based: bool, case_sensitive: bool) -> Self {
        // Sort operators longest-first for greedy matching
        Self {
            chars: source.chars().collect(),
            pos: 0,
            line: 0,
            col: 0,
            spec,
            terminator,
            indentation_based,
            case_sensitive,
            indent_stack: vec![0],
            pending_dedents: 0,
            at_line_start: true,
        }
    }

    fn peek(&self) -> Option<char> { self.chars.get(self.pos).copied() }

    fn advance(&mut self) -> Option<char> {
        let c = self.chars.get(self.pos).copied();
        if let Some(ch) = c {
            self.pos += 1;
            if ch == '\n' {
                self.line += 1;
                self.col = 0;
                self.at_line_start = true;
            } else {
                self.col += 1;
            }
        }
        c
    }

    #[allow(dead_code)]
    fn remaining(&self) -> &[char] { &self.chars[self.pos..] }

    fn starts_with(&self, s: &str) -> bool {
        let chars: Vec<char> = s.chars().collect();
        if self.pos + chars.len() > self.chars.len() { return false; }
        for (i, c) in chars.iter().enumerate() {
            let src = self.chars[self.pos + i];
            if self.case_sensitive {
                if src != *c { return false; }
            } else {
                if src.to_ascii_lowercase() != c.to_ascii_lowercase() { return false; }
            }
        }
        true
    }

    fn advance_n(&mut self, n: usize) -> String {
        let mut s = String::new();
        for _ in 0..n {
            if let Some(c) = self.advance() { s.push(c); }
        }
        s
    }

    fn make_token(&self, kind: TokenKind, text: String, line: u32, col: u32) -> Token {
        Token { kind, text, line, col }
    }

    fn next_token(&mut self) -> Token {
        // Emit pending DEDENT tokens (for indentation-based languages)
        if self.pending_dedents > 0 {
            self.pending_dedents -= 1;
            return self.make_token(TokenKind::Dedent, "DEDENT".into(), self.line, self.col);
        }

        // Handle indentation at the start of a line
        if self.indentation_based && self.at_line_start && self.peek().is_some() {
            return self.handle_indentation();
        }

        // Skip whitespace (but not newlines for newline-significant languages)
        self.skip_whitespace();

        // Skip comments
        while self.skip_comment() {
            self.skip_whitespace();
        }

        let line = self.line;
        let col = self.col;

        match self.peek() {
            None => {
                // At EOF, emit DEDENT for any remaining indentation
                if self.indentation_based && self.indent_stack.len() > 1 {
                    self.indent_stack.pop();
                    if self.indent_stack.len() > 1 {
                        self.pending_dedents = (self.indent_stack.len() - 1) as u32;
                        self.indent_stack.truncate(1);
                    }
                    return self.make_token(TokenKind::Dedent, "DEDENT".into(), line, col);
                }
                self.make_token(TokenKind::Eof, String::new(), line, col)
            }
            Some('\n') | Some('\r') => {
                // Newline
                if self.peek() == Some('\r') { self.advance(); }
                if self.peek() == Some('\n') { self.advance(); }
                if *self.terminator == Terminator::Newline || self.indentation_based {
                    self.at_line_start = true;
                    self.make_token(TokenKind::Newline, "\n".into(), line, col)
                } else {
                    // Newlines are just whitespace — skip and get next token
                    self.next_token()
                }
            }
            Some(c) if c == '\'' || c == '"' || c == '`' => {
                self.read_string(line, col)
            }
            Some(c) if c.is_ascii_digit() => {
                self.read_number(line, col)
            }
            Some(c) if c.is_alphabetic() || c == '_' => {
                self.read_ident_or_keyword(line, col)
            }
            Some('#') if self.spec.char_prefix.as_deref() == Some("#") => {
                // Pascal char literal: #65
                self.advance();
                let start = self.pos;
                while self.peek().map_or(false, |c| c.is_ascii_digit()) { self.advance(); }
                let text: String = self.chars[start..self.pos].iter().collect();
                let code: u32 = text.parse().unwrap_or(0);
                let ch = char::from_u32(code).unwrap_or('\0');
                self.make_token(TokenKind::CharLit, ch.to_string(), line, col)
            }
            Some('$') if self.spec.hex_prefix.as_deref() == Some("$") => {
                // Pascal hex literal: $FF
                self.advance();
                let start = self.pos;
                while self.peek().map_or(false, |c| c.is_ascii_hexdigit()) { self.advance(); }
                let text: String = self.chars[start..self.pos].iter().collect();
                let val = i64::from_str_radix(&text, 16).unwrap_or(0);
                self.make_token(TokenKind::IntLit, val.to_string(), line, col)
            }
            Some(_) => {
                // Try to match an operator (longest match first)
                if let Some(op) = self.try_match_operator() {
                    self.make_token(TokenKind::Operator, op, line, col)
                } else {
                    // Unknown character — skip it
                    let c = self.advance().unwrap();
                    self.make_token(TokenKind::Operator, c.to_string(), line, col)
                }
            }
        }
    }

    fn skip_whitespace(&mut self) {
        while let Some(c) = self.peek() {
            if c == ' ' || c == '\t' {
                self.advance();
            } else if (c == '\n' || c == '\r') && !self.indentation_based && *self.terminator != Terminator::Newline {
                self.advance();
            } else {
                break;
            }
        }
    }

    fn skip_comment(&mut self) -> bool {
        // Block comments
        for (open, close) in &self.spec.comment_block {
            if self.starts_with(open) {
                self.advance_n(open.len());
                let _close_chars: Vec<char> = close.chars().collect();
                while self.pos < self.chars.len() {
                    if self.starts_with(close) {
                        self.advance_n(close.len());
                        return true;
                    }
                    self.advance();
                }
                return true; // unterminated comment
            }
        }
        // Line comments
        for prefix in &self.spec.comment_line {
            if self.starts_with(prefix) {
                while self.peek().map_or(false, |c| c != '\n') { self.advance(); }
                return true;
            }
        }
        false
    }

    fn read_string(&mut self, line: u32, col: u32) -> Token {
        let delim = self.advance().unwrap();
        let delim_str = delim.to_string();

        // Check for triple-quoted strings
        let is_triple = self.spec.triple_string.iter().any(|t| {
            let chars: Vec<char> = t.chars().collect();
            chars.len() == 3 && chars[0] == delim
        }) && self.peek() == Some(delim) && self.chars.get(self.pos + 1) == Some(&delim);

        if is_triple {
            self.advance(); self.advance(); // consume the other two delimiters
            let mut text = String::new();
            loop {
                match self.peek() {
                    None => break,
                    Some(c) if c == delim && self.chars.get(self.pos + 1) == Some(&delim) && self.chars.get(self.pos + 2) == Some(&delim) => {
                        self.advance(); self.advance(); self.advance();
                        break;
                    }
                    _ => { text.push(self.advance().unwrap()); }
                }
            }
            return self.make_token(TokenKind::StringLit, text, line, col);
        }

        // Check if this delimiter is valid for strings
        if !self.spec.string_delimiters.contains(&delim_str) && self.spec.template_string.as_deref() != Some(&delim_str) {
            return self.make_token(TokenKind::Operator, delim_str, line, col);
        }

        let mut text = String::new();
        loop {
            match self.peek() {
                None | Some('\n') => break,
                Some(c) if c == delim => {
                    self.advance();
                    // Check for escaped delimiter (e.g. Pascal '' or general \')
                    if self.spec.string_escape.as_deref() == Some(&format!("{}{}", delim, delim)) && self.peek() == Some(delim) {
                        // Doubled delimiter escape (Pascal)
                        // Actually the escape IS the doubled char, and we already consumed the first delim.
                        // If next is also delim, it's an escaped delim.
                        // Wait — we already consumed the closing delim. Check if it was doubled:
                        if self.peek() == Some(delim) {
                            text.push(delim);
                            self.advance(); // consume the second delim
                            continue;
                        }
                    }
                    break; // end of string
                }
                Some('\\') if self.spec.string_escape.as_deref() == Some("\\") => {
                    self.advance(); // consume backslash
                    if let Some(c) = self.advance() {
                        match c {
                            'n' => text.push('\n'),
                            't' => text.push('\t'),
                            'r' => text.push('\r'),
                            '\\' => text.push('\\'),
                            _ => { text.push('\\'); text.push(c); }
                        }
                    }
                }
                _ => { text.push(self.advance().unwrap()); }
            }
        }
        self.make_token(TokenKind::StringLit, text, line, col)
    }

    fn read_number(&mut self, line: u32, col: u32) -> Token {
        let mut text = String::new();
        let mut is_float = false;

        // Check for 0x hex prefix
        if self.peek() == Some('0') && self.chars.get(self.pos + 1).map_or(false, |c| *c == 'x' || *c == 'X') {
            text.push(self.advance().unwrap()); // '0'
            text.push(self.advance().unwrap()); // 'x'
            while self.peek().map_or(false, |c| c.is_ascii_hexdigit() || c == '_') {
                let c = self.advance().unwrap();
                if c != '_' { text.push(c); }
            }
            return self.make_token(TokenKind::IntLit, text, line, col);
        }

        while let Some(c) = self.peek() {
            if c.is_ascii_digit() {
                text.push(self.advance().unwrap());
            } else if c == '.' && !is_float {
                // Check it's not '..' (range operator)
                if self.chars.get(self.pos + 1) == Some(&'.') { break; }
                // Check next char is digit (not method call like 1.toString)
                if self.chars.get(self.pos + 1).map_or(false, |c| c.is_ascii_digit()) {
                    is_float = true;
                    text.push(self.advance().unwrap());
                } else {
                    break;
                }
            } else if (c == 'e' || c == 'E') && !is_float {
                is_float = true;
                text.push(self.advance().unwrap());
                if self.peek() == Some('+') || self.peek() == Some('-') {
                    text.push(self.advance().unwrap());
                }
            } else if c == '_' {
                self.advance(); // skip numeric separators
            } else {
                break;
            }
        }

        if is_float {
            self.make_token(TokenKind::FloatLit, text, line, col)
        } else {
            self.make_token(TokenKind::IntLit, text, line, col)
        }
    }

    fn read_ident_or_keyword(&mut self, line: u32, col: u32) -> Token {
        let mut text = String::new();
        while let Some(c) = self.peek() {
            if c.is_alphanumeric() || c == '_' {
                text.push(self.advance().unwrap());
            } else {
                break;
            }
        }

        // Check if it's a keyword
        let is_kw = if self.case_sensitive {
            self.spec.keywords.iter().any(|k| k == &text)
        } else {
            let lower = text.to_lowercase();
            self.spec.keywords.iter().any(|k| k.to_lowercase() == lower)
        };

        if is_kw {
            // Normalize keyword case for case-insensitive languages
            let normalized = if !self.case_sensitive {
                self.spec.keywords.iter()
                    .find(|k| k.to_lowercase() == text.to_lowercase())
                    .cloned()
                    .unwrap_or(text)
            } else {
                text
            };
            self.make_token(TokenKind::Keyword, normalized, line, col)
        } else {
            self.make_token(TokenKind::Ident, text, line, col)
        }
    }

    fn try_match_operator(&mut self) -> Option<String> {
        // Try longest operators first (they should already be sorted)
        for op in &self.spec.operators {
            if self.starts_with(op) {
                return Some(self.advance_n(op.len()));
            }
        }
        None
    }

    fn handle_indentation(&mut self) -> Token {
        self.at_line_start = false;

        // Skip blank lines
        if self.peek() == Some('\n') || self.peek() == Some('\r') {
            return self.next_token();
        }

        // Count leading whitespace
        let mut indent = 0u32;
        let _start_pos = self.pos;
        while let Some(c) = self.peek() {
            match c {
                ' ' => { indent += 1; self.advance(); }
                '\t' => { indent += 4; self.advance(); } // tab = 4 spaces
                _ => break,
            }
        }

        // Skip comment-only lines
        if self.peek() == Some('#') || self.peek() == Some('\n') || self.peek().is_none() {
            return self.next_token();
        }

        let current_indent = *self.indent_stack.last().unwrap();
        let line = self.line;
        let col = 0;

        if indent > current_indent {
            self.indent_stack.push(indent);
            self.make_token(TokenKind::Indent, "INDENT".into(), line, col)
        } else if indent < current_indent {
            // May need multiple DEDENTs
            let mut count = 0u32;
            while self.indent_stack.last().map_or(false, |&top| top > indent) {
                self.indent_stack.pop();
                count += 1;
            }
            if count > 1 {
                self.pending_dedents = count - 1;
            }
            self.make_token(TokenKind::Dedent, "DEDENT".into(), line, col)
        } else {
            // Same indent — continue with next token
            self.next_token()
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn pascal_spec() -> LexerSpec {
        LexerSpec {
            comment_line: vec!["//".into()],
            comment_block: vec![("{".into(), "}".into()), ("(*".into(), "*)".into())],
            string_delimiters: vec!["'".into()],
            string_escape: Some("''".into()),
            triple_string: Vec::new(),
            string_prefixes: Vec::new(),
            interpolation: None,
            template_string: None,
            char_prefix: Some("#".into()),
            hex_prefix: Some("$".into()),
            keywords: vec!["program".into(), "begin".into(), "end".into(), "var".into(),
                "if".into(), "then".into(), "else".into(), "for".into(), "to".into(), "do".into(),
                "while".into(), "procedure".into(), "function".into(), "writeln".into(),
                "true".into(), "false".into(), "and".into(), "or".into(), "not".into(),
                "integer".into(), "string".into(), "boolean".into(), "real".into(),
            ],
            operators: vec![
                ":=".into(), "+=".into(), "-=".into(), "*=".into(), "/=".into(),
                "<>".into(), "<=".into(), ">=".into(), "..".into(),
                "+".into(), "-".into(), "*".into(), "/".into(),
                "=".into(), "<".into(), ">".into(),
                "(".into(), ")".into(), "[".into(), "]".into(),
                ".".into(), ",".into(), ";".into(), ":".into(), "^".into(), "@".into(),
            ],
        }
    }

    #[test]
    fn tokenize_pascal_hello() {
        let spec = pascal_spec();
        let tokens = tokenize(
            "program Test; begin WriteLn('hello'); end.",
            &spec, &Terminator::Char(';'), false, false,
        );
        let kinds: Vec<_> = tokens.iter().map(|t| (&t.kind, t.text.as_str())).collect();
        assert_eq!(kinds[0], (&TokenKind::Keyword, "program"));
        assert_eq!(kinds[1], (&TokenKind::Ident, "Test"));
        assert_eq!(kinds[2], (&TokenKind::Operator, ";"));
        assert_eq!(kinds[3], (&TokenKind::Keyword, "begin"));
        assert_eq!(kinds[4], (&TokenKind::Keyword, "writeln")); // case-insensitive match
        assert_eq!(kinds[5], (&TokenKind::Operator, "("));
        assert_eq!(kinds[6], (&TokenKind::StringLit, "hello"));
        assert_eq!(kinds[7], (&TokenKind::Operator, ")"));
    }

    #[test]
    fn tokenize_pascal_numbers() {
        let spec = pascal_spec();
        let tokens = tokenize("42 3.14 $FF #65", &spec, &Terminator::Char(';'), false, false);
        assert_eq!(tokens[0].kind, TokenKind::IntLit);
        assert_eq!(tokens[0].text, "42");
        assert_eq!(tokens[1].kind, TokenKind::FloatLit);
        assert_eq!(tokens[1].text, "3.14");
        assert_eq!(tokens[2].kind, TokenKind::IntLit);
        assert_eq!(tokens[2].text, "255"); // $FF
        assert_eq!(tokens[3].kind, TokenKind::CharLit);
        assert_eq!(tokens[3].text, "A"); // #65
    }

    #[test]
    fn tokenize_pascal_comments() {
        let spec = pascal_spec();
        let tokens = tokenize("x // comment\ny { block } z", &spec, &Terminator::Char(';'), false, false);
        let idents: Vec<_> = tokens.iter().filter(|t| t.kind == TokenKind::Ident).map(|t| t.text.as_str()).collect();
        assert_eq!(idents, &["x", "y", "z"]);
    }

    #[test]
    fn tokenize_pascal_case_insensitive() {
        let spec = pascal_spec();
        let tokens = tokenize("BEGIN End If", &spec, &Terminator::Char(';'), false, false);
        assert_eq!(tokens[0].kind, TokenKind::Keyword);
        assert_eq!(tokens[0].text, "begin"); // normalized
        assert_eq!(tokens[1].kind, TokenKind::Keyword);
        assert_eq!(tokens[1].text, "end");
        assert_eq!(tokens[2].kind, TokenKind::Keyword);
        assert_eq!(tokens[2].text, "if");
    }

    #[test]
    fn tokenize_pascal_string_escape() {
        let spec = pascal_spec();
        let tokens = tokenize("'it''s'", &spec, &Terminator::Char(';'), false, false);
        assert_eq!(tokens[0].kind, TokenKind::StringLit);
        assert_eq!(tokens[0].text, "it's");
    }

    #[test]
    fn tokenize_js_style() {
        let spec = LexerSpec {
            comment_line: vec!["//".into()],
            comment_block: vec![("/*".into(), "*/".into())],
            string_delimiters: vec!["'".into(), "\"".into()],
            string_escape: Some("\\".into()),
            triple_string: Vec::new(),
            string_prefixes: Vec::new(),
            interpolation: Some(("${".into(), "}".into())),
            template_string: Some("`".into()),
            char_prefix: None,
            hex_prefix: None,
            keywords: vec!["function".into(), "var".into(), "if".into(), "else".into(), "return".into(),
                "true".into(), "false".into(), "null".into()],
            operators: vec![
                "===".into(), "!==".into(), "==".into(), "!=".into(),
                "&&".into(), "||".into(), "=>".into(),
                "<=".into(), ">=".into(),
                "+".into(), "-".into(), "*".into(), "/".into(),
                "=".into(), "<".into(), ">".into(), "!".into(),
                "(".into(), ")".into(), "{".into(), "}".into(), "[".into(), "]".into(),
                ".".into(), ",".into(), ";".into(), ":".into(),
            ],
        };
        let tokens = tokenize(
            "function add(a, b) { return a + b; }",
            &spec, &Terminator::Char(';'), false, true,
        );
        assert_eq!(tokens[0].kind, TokenKind::Keyword);
        assert_eq!(tokens[0].text, "function");
        assert_eq!(tokens[1].kind, TokenKind::Ident);
        assert_eq!(tokens[1].text, "add");
    }

    #[test]
    fn tokenize_python_indentation() {
        let spec = LexerSpec {
            comment_line: vec!["#".into()],
            comment_block: Vec::new(),
            string_delimiters: vec!["'".into(), "\"".into()],
            string_escape: Some("\\".into()),
            triple_string: vec!["'''".into(), "\"\"\"".into()],
            string_prefixes: vec!["f".into(), "r".into()],
            interpolation: Some(("{".into(), "}".into())),
            template_string: None,
            char_prefix: None,
            hex_prefix: None,
            keywords: vec!["def".into(), "if".into(), "else".into(), "return".into(),
                "True".into(), "False".into(), "None".into(), "pass".into()],
            operators: vec![
                "==".into(), "!=".into(), "<=".into(), ">=".into(),
                "+".into(), "-".into(), "*".into(), "/".into(),
                "=".into(), "<".into(), ">".into(),
                "(".into(), ")".into(), "[".into(), "]".into(), "{".into(), "}".into(),
                ".".into(), ",".into(), ":".into(),
            ],
        };
        let tokens = tokenize(
            "def foo():\n    return 42\n",
            &spec, &Terminator::Newline, true, true,
        );
        let kinds: Vec<_> = tokens.iter().map(|t| (&t.kind, t.text.as_str())).collect();
        // Should produce: def foo ( ) : NEWLINE INDENT return 42 NEWLINE DEDENT EOF
        assert!(kinds.iter().any(|(k, _)| *k == &TokenKind::Indent));
        assert!(kinds.iter().any(|(k, _)| *k == &TokenKind::Dedent));
    }
}
