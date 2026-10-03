use crate::{Diagnostic, Span};

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum Kind {
    Ident(String),
    String(String),
    Character(char),
    Number(u32),
    Symbol(char),
    Range,
    Eof,
}

#[derive(Debug, Clone)]
pub(crate) struct Token {
    pub kind: Kind,
    pub span: Span,
}

pub(crate) fn lex(source: &str) -> Result<Vec<Token>, Diagnostic> {
    let mut lexer = Lexer { source, offset: 0 };
    let mut tokens = Vec::new();
    loop {
        lexer.skip_trivia()?;
        let start = lexer.offset;
        let Some(ch) = lexer.bump() else {
            tokens.push(Token {
                kind: Kind::Eof,
                span: Span { start, end: start },
            });
            return Ok(tokens);
        };
        let kind = match ch {
            '"' => Kind::String(lexer.quoted('"', start)?),
            '\'' => {
                let text = lexer.quoted('\'', start)?;
                let mut chars = text.chars();
                let first = chars.next();
                if first.is_none() || chars.next().is_some() {
                    return Err(lexer.error(
                        "G003",
                        "a character endpoint must contain exactly one Unicode scalar",
                        start,
                    ));
                }
                Kind::Character(first.unwrap())
            }
            '.' if lexer.rest().starts_with('.') => {
                lexer.offset += 1;
                Kind::Range
            }
            c if c.is_ascii_alphabetic() || c == '_' => {
                while lexer
                    .peek()
                    .is_some_and(|c| c.is_ascii_alphanumeric() || c == '_')
                {
                    lexer.bump();
                }
                Kind::Ident(source[start..lexer.offset].to_owned())
            }
            c if c.is_ascii_digit() => {
                while lexer.peek().is_some_and(|c| c.is_ascii_digit()) {
                    lexer.bump();
                }
                Kind::Number(source[start..lexer.offset].parse().map_err(|_| {
                    lexer.error("G004", "repetition/index integer exceeds u32", start)
                })?)
            }
            '=' | '{' | '}' | '(' | ')' | '[' | ']' | '~' | '|' | '?' | '*' | '+' | '&' | '!'
            | '@' | '$' | '^' | ',' | '-' | '#' => Kind::Symbol(ch),
            _ => {
                return Err(lexer.error(
                    "G001",
                    format!("unexpected grammar character {ch:?}"),
                    start,
                ));
            }
        };
        tokens.push(Token {
            kind,
            span: Span {
                start,
                end: lexer.offset,
            },
        });
    }
}

struct Lexer<'a> {
    source: &'a str,
    offset: usize,
}

impl Lexer<'_> {
    fn rest(&self) -> &str {
        &self.source[self.offset..]
    }
    fn peek(&self) -> Option<char> {
        self.rest().chars().next()
    }
    fn bump(&mut self) -> Option<char> {
        let ch = self.peek()?;
        self.offset += ch.len_utf8();
        Some(ch)
    }
    fn error(&self, code: &'static str, message: impl Into<String>, start: usize) -> Diagnostic {
        Diagnostic::new(
            code,
            message,
            Span {
                start,
                end: self.offset,
            },
        )
    }
    fn skip_trivia(&mut self) -> Result<(), Diagnostic> {
        loop {
            while self
                .peek()
                .is_some_and(|c| matches!(c, ' ' | '\t' | '\r' | '\n'))
            {
                self.bump();
            }
            if self.rest().starts_with("//") {
                while self.peek().is_some_and(|c| c != '\n') {
                    self.bump();
                }
            } else if self.rest().starts_with("/*") {
                let start = self.offset;
                self.offset += 2;
                let mut depth = 1usize;
                while depth > 0 {
                    if self.rest().starts_with("/*") {
                        self.offset += 2;
                        depth += 1;
                    } else if self.rest().starts_with("*/") {
                        self.offset += 2;
                        depth -= 1;
                    } else if self.bump().is_none() {
                        return Err(self.error("G002", "unterminated grammar comment", start));
                    }
                }
            } else {
                return Ok(());
            }
        }
    }
    fn quoted(&mut self, quote: char, start: usize) -> Result<String, Diagnostic> {
        let mut text = String::new();
        loop {
            let Some(ch) = self.bump() else {
                return Err(self.error("G002", "unterminated literal", start));
            };
            if ch == quote {
                return Ok(text);
            }
            if ch != '\\' {
                text.push(ch);
                continue;
            }
            let escape_start = self.offset - 1;
            let value = match self.bump() {
                Some('n') => '\n',
                Some('r') => '\r',
                Some('t') => '\t',
                Some('\\') => '\\',
                Some('"') => '"',
                Some('\'') => '\'',
                Some('x') => {
                    let value = self.hex_digits(2, escape_start)?;
                    if value > 0x7f {
                        return Err(self.error(
                            "G003",
                            "\\x escape must be ASCII (00..7f)",
                            escape_start,
                        ));
                    }
                    char::from_u32(value).unwrap()
                }
                Some('u') => {
                    if self.bump() != Some('{') {
                        return Err(self.error("G003", "expected '{' after \\u", escape_start));
                    }
                    let digits_start = self.offset;
                    while self.peek().is_some_and(|c| c.is_ascii_hexdigit()) {
                        self.bump();
                    }
                    let digits = &self.source[digits_start..self.offset];
                    if digits.is_empty() || digits.len() > 6 || self.bump() != Some('}') {
                        return Err(self.error(
                            "G003",
                            "expected 1..6 hex digits and '}' in Unicode escape",
                            escape_start,
                        ));
                    }
                    let value = u32::from_str_radix(digits, 16).unwrap();
                    char::from_u32(value).ok_or_else(|| {
                        self.error("G003", "escape is not a Unicode scalar", escape_start)
                    })?
                }
                _ => {
                    return Err(self.error(
                        "G003",
                        "unsupported or incomplete literal escape",
                        escape_start,
                    ));
                }
            };
            text.push(value);
        }
    }
    fn hex_digits(&mut self, count: usize, start: usize) -> Result<u32, Diagnostic> {
        let mut value = 0;
        for _ in 0..count {
            let Some(digit) = self
                .bump()
                .and_then(|c| c.to_digit(16).filter(|_| c.is_ascii()))
            else {
                return Err(self.error("G003", "expected hexadecimal escape digits", start));
            };
            value = value * 16 + digit;
        }
        Ok(value)
    }
}
