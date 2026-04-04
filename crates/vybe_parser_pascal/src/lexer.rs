use crate::token::Token;

pub struct Lexer {
    chars: Vec<char>,
    pos: usize,
}

impl Lexer {
    pub fn new(input: &str) -> Self {
        Self { chars: input.chars().collect(), pos: 0 }
    }

    fn peek(&self) -> Option<char> { self.chars.get(self.pos).copied() }
    fn peek2(&self) -> Option<char> { self.chars.get(self.pos + 1).copied() }
    fn advance(&mut self) -> Option<char> {
        let c = self.chars.get(self.pos).copied();
        if c.is_some() { self.pos += 1; }
        c
    }
    fn eat(&mut self, c: char) -> bool {
        if self.peek() == Some(c) { self.advance(); true } else { false }
    }

    fn skip_whitespace_and_comments(&mut self) {
        loop {
            while matches!(self.peek(), Some(c) if c.is_whitespace()) { self.advance(); }
            // { comment }
            if self.peek() == Some('{') {
                self.advance();
                while !matches!(self.peek(), None | Some('}')) { self.advance(); }
                self.advance();
                continue;
            }
            // (* comment *)
            if self.peek() == Some('(') && self.peek2() == Some('*') {
                self.advance(); self.advance();
                loop {
                    match self.advance() {
                        None => break,
                        Some('*') if self.peek() == Some(')') => { self.advance(); break; }
                        _ => {}
                    }
                }
                continue;
            }
            // // comment
            if self.peek() == Some('/') && self.peek2() == Some('/') {
                while !matches!(self.peek(), None | Some('\n')) { self.advance(); }
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
            '[' => Token::LBracket,
            ']' => Token::RBracket,
            ',' => Token::Comma,
            ';' => Token::Semicolon,
            '@' => Token::At,
            '^' => Token::Caret,
            '+' => { if self.eat('=') { Token::PlusAssign } else { Token::Plus } }
            '-' => { if self.eat('=') { Token::MinusAssign } else { Token::Minus } }
            '*' => { if self.eat('=') { Token::StarAssign } else { Token::Star } }
            '/' => { if self.eat('=') { Token::SlashAssign } else { Token::Slash } }
            '=' => Token::Eq,
            '<' => {
                if self.eat('>') { Token::NotEq }
                else if self.eat('=') { Token::Le }
                else { Token::Lt }
            }
            '>' => { if self.eat('=') { Token::Ge } else { Token::Gt } }
            ':' => { if self.eat('=') { Token::Assign } else { Token::Colon } }
            '.' => { if self.eat('.') { Token::DotDot } else { Token::Dot } }
            '\'' => self.read_string(),
            '#' => self.read_char_literal(),
            '$' => self.read_hex(),
            '0'..='9' => self.read_number(ch),
            'a'..='z' | 'A'..='Z' | '_' => self.read_ident(ch),
            _ => Token::EOF,
        }
    }

    fn read_string(&mut self) -> Token {
        let mut s = String::new();
        loop {
            match self.advance() {
                None => break,
                Some('\'') => {
                    if self.peek() == Some('\'') { self.advance(); s.push('\''); }
                    else { break; }
                }
                Some(c) => s.push(c),
            }
        }
        Token::StringLiteral(s)
    }

    fn read_char_literal(&mut self) -> Token {
        let mut n = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            n.push(self.advance().unwrap());
        }
        let code: u32 = n.parse().unwrap_or(0);
        Token::CharLiteral(char::from_u32(code).unwrap_or('\0'))
    }

    fn read_hex(&mut self) -> Token {
        let mut s = String::new();
        while matches!(self.peek(), Some(c) if c.is_ascii_hexdigit()) {
            s.push(self.advance().unwrap());
        }
        Token::IntLiteral(i64::from_str_radix(&s, 16).unwrap_or(0))
    }

    fn read_number(&mut self, first: char) -> Token {
        let mut s = String::from(first);
        let mut is_real = false;
        while let Some(c) = self.peek() {
            if c.is_ascii_digit() { s.push(self.advance().unwrap()); }
            else if c == '.' && !is_real && self.peek2() != Some('.') {
                is_real = true; s.push(self.advance().unwrap());
            } else if (c == 'e' || c == 'E') && !is_real {
                is_real = true; s.push(self.advance().unwrap());
                if matches!(self.peek(), Some('+') | Some('-')) { s.push(self.advance().unwrap()); }
            } else { break; }
        }
        if is_real { Token::RealLiteral(s.parse().unwrap_or(0.0)) }
        else { Token::IntLiteral(s.parse().unwrap_or(0)) }
    }

    fn read_ident(&mut self, first: char) -> Token {
        let mut s = String::from(first);
        while matches!(self.peek(), Some(c) if c.is_alphanumeric() || c == '_') {
            s.push(self.advance().unwrap());
        }
        keyword_or_ident(s)
    }
}

fn keyword_or_ident(s: String) -> Token {
    match s.to_lowercase().as_str() {
        "program" => Token::Program,
        "unit" => Token::Unit,
        "uses" => Token::Uses,
        "interface" => Token::Interface,
        "implementation" => Token::Implementation,
        "begin" => Token::Begin,
        "end" => Token::End,
        "var" => Token::Var,
        "const" => Token::Const,
        "type" => Token::Type,
        "procedure" => Token::Procedure,
        "function" => Token::Function,
        "forward" => Token::Forward,
        "if" => Token::If,
        "then" => Token::Then,
        "else" => Token::Else,
        "for" => Token::For,
        "to" => Token::To,
        "downto" => Token::Downto,
        "do" => Token::Do,
        "while" => Token::While,
        "repeat" => Token::Repeat,
        "until" => Token::Until,
        "case" => Token::Case,
        "of" => Token::Of,
        "otherwise" => Token::Otherwise,
        "record" => Token::Record,
        "array" => Token::Array,
        "set" => Token::Set,
        "file" => Token::File,
        "object" | "class" => Token::Class,
        "inherited" => Token::Inherited,
        "override" => Token::Override,
        "virtual" => Token::Virtual,
        "abstract" => Token::Abstract,
        "with" => Token::With,
        "in" => Token::In,
        "is" => Token::Is,
        "as" => Token::As,
        "try" => Token::Try,
        "except" => Token::Except,
        "finally" => Token::Finally,
        "raise" => Token::Raise,
        "on" => Token::On,
        "and" => Token::And,
        "or" => Token::Or,
        "not" => Token::Not,
        "xor" => Token::Xor,
        "div" => Token::Div,
        "mod" => Token::Mod,
        "shl" => Token::Shl,
        "shr" => Token::Shr,
        "nil" => Token::Nil,
        "true" => Token::True,
        "false" => Token::False,
        "string" => Token::String,
        "integer" | "longint" | "int64" | "cardinal" | "word" | "byte" | "shortint" => Token::Integer,
        "real" | "double" | "single" | "extended" => Token::Real,
        "boolean" => Token::Boolean,
        "char" => Token::Char,
        "pointer" => Token::Pointer,
        "exit" => Token::Exit,
        "break" => Token::Break,
        "continue" => Token::Continue,
        "halt" => Token::Halt,
        "new" => Token::New,
        "dispose" => Token::Dispose,
        "result" => Token::Result,
        "constructor" => Token::Constructor,
        "destructor" => Token::Destructor,
        "public" => Token::Public,
        "private" => Token::Private,
        "protected" => Token::Protected,
        "published" => Token::Published,
        "reintroduce" => Token::Identifier("reintroduce".into()),
        _ => Token::Identifier(s),
    }
}
