use crate::token::{Token, TokenKind};

pub struct Lexer {
    src: Vec<char>,
    pos: usize,
    line: u32,
}

impl Lexer {
    pub fn new(source: &str) -> Self {
        let src: Vec<char> = source.chars().collect();
        Lexer { src, pos: 0, line: 1 }
    }

    pub fn tokenize(&mut self) -> Result<Vec<Token>, String> {
        let mut tokens = Vec::new();
        loop {
            self.skip_whitespace_and_comments();
            if self.pos >= self.src.len() {
                tokens.push(Token { kind: TokenKind::Eof, line: self.line });
                break;
            }
            let tok = self.next_token()?;
            tokens.push(tok);
        }
        Ok(tokens)
    }

    fn skip_whitespace_and_comments(&mut self) {
        while self.pos < self.src.len() {
            let ch = self.src[self.pos];
            if ch == '\n' {
                self.line += 1;
                self.pos += 1;
            } else if ch == '\r' || ch == ' ' || ch == '\t' {
                self.pos += 1;
            } else if ch == '*' && self.at_line_start() {
                // Fixed-format comment (column 7 = *)
                while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else if self.peek_str("*>") {
                // Free-format inline comment
                while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else if self.peek_str(">>") {
                // Preprocessor directive — skip line
                while self.pos < self.src.len() && self.src[self.pos] != '\n' {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    fn at_line_start(&self) -> bool {
        if self.pos == 0 { return true; }
        // Check if previous non-whitespace was a newline
        let mut i = self.pos - 1;
        while i > 0 && (self.src[i] == ' ' || self.src[i] == '\t') {
            i -= 1;
        }
        i == 0 || self.src[i] == '\n'
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

    fn next_token(&mut self) -> Result<Token, String> {
        let line = self.line;
        let ch = self.advance();

        match ch {
            // String literals
            '"' => self.read_string('"', line),
            '\'' => self.read_string('\'', line),

            // Numbers
            '0'..='9' => self.read_number(ch, line),

            // Operators
            '+' => Ok(Token { kind: TokenKind::Plus, line }),
            '-' => Ok(Token { kind: TokenKind::Minus, line }),
            '*' => {
                if self.peek() == Some('*') { self.pos += 1; Ok(Token { kind: TokenKind::StarStar, line }) }
                else { Ok(Token { kind: TokenKind::Star, line }) }
            }
            '/' => Ok(Token { kind: TokenKind::Slash, line }),
            '=' => Ok(Token { kind: TokenKind::Eq, line }),
            '>' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::GtEq, line }) }
                else { Ok(Token { kind: TokenKind::Gt, line }) }
            }
            '<' => {
                if self.peek() == Some('=') { self.pos += 1; Ok(Token { kind: TokenKind::LtEq, line }) }
                else { Ok(Token { kind: TokenKind::Lt, line }) }
            }

            // Delimiters
            '(' => Ok(Token { kind: TokenKind::LParen, line }),
            ')' => Ok(Token { kind: TokenKind::RParen, line }),
            '.' => Ok(Token { kind: TokenKind::Period, line }),
            ',' => Ok(Token { kind: TokenKind::Comma, line }),
            ':' => Ok(Token { kind: TokenKind::Colon, line }),

            // Identifiers / keywords
            'a'..='z' | 'A'..='Z' | '_' => self.read_word(ch, line),

            _ => Err(format!("Unexpected character '{}' at line {}", ch, line)),
        }
    }

    fn read_string(&mut self, quote: char, line: u32) -> Result<Token, String> {
        let mut s = String::new();
        while self.pos < self.src.len() && self.src[self.pos] != quote {
            if self.src[self.pos] == '\n' { self.line += 1; }
            s.push(self.src[self.pos]);
            self.pos += 1;
        }
        if self.pos < self.src.len() { self.pos += 1; }
        Ok(Token { kind: TokenKind::Str(s), line })
    }

    fn read_number(&mut self, first: char, line: u32) -> Result<Token, String> {
        let mut s = String::new();
        s.push(first);
        let mut has_dot = false;
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c.is_ascii_digit() {
                s.push(c);
                self.pos += 1;
            } else if c == '.' && !has_dot && self.pos + 1 < self.src.len() && self.src[self.pos + 1].is_ascii_digit() {
                has_dot = true;
                s.push('.');
                self.pos += 1;
            } else {
                break;
            }
        }
        let val: f64 = s.parse().map_err(|_| format!("Invalid number at line {}", line))?;
        Ok(Token { kind: TokenKind::Number(val), line })
    }

    fn read_word(&mut self, first: char, line: u32) -> Result<Token, String> {
        let mut word = String::new();
        word.push(first);
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if c.is_alphanumeric() || c == '-' || c == '_' {
                word.push(c);
                self.pos += 1;
            } else {
                break;
            }
        }

        let upper = word.to_uppercase();
        let kind = match upper.as_str() {
            // Division keywords
            "IDENTIFICATION" => TokenKind::Identification,
            "DATA" => TokenKind::Data,
            "PROCEDURE" => TokenKind::Procedure,
            "ENVIRONMENT" => TokenKind::Environment,
            "DIVISION" => TokenKind::Division,
            "SECTION" => TokenKind::Section,

            // Identification
            "PROGRAM-ID" => TokenKind::ProgramId,
            "AUTHOR" => TokenKind::Author,
            "DATE-WRITTEN" => TokenKind::DateWritten,

            // Data division
            "WORKING-STORAGE" => TokenKind::WorkingStorage,
            "LOCAL-STORAGE" => TokenKind::LocalStorage,
            "FILE" => TokenKind::FileSection,
            "LINKAGE" => TokenKind::Linkage,
            "PIC" | "PICTURE" => TokenKind::Pic,
            "VALUE" => TokenKind::Value,
            "OCCURS" => TokenKind::Occurs,
            "TIMES" => TokenKind::Times,
            "REDEFINES" => TokenKind::Redefines,
            "INDEXED" => TokenKind::Indexed,
            "USAGE" => TokenKind::Usage,
            "BINARY" => TokenKind::Binary,
            "COMP" | "COMPUTATIONAL" => TokenKind::Comp,
            "COMP-3" => TokenKind::Comp3,
            "DISPLAY" => TokenKind::DisplayKw,
            "POINTER" => TokenKind::Pointer,
            "SPACES" | "SPACE" => TokenKind::Spaces,
            "ZEROS" | "ZERO" | "ZEROES" => TokenKind::Zeros,
            "LOW-VALUES" | "LOW-VALUE" => TokenKind::LowValues,
            "HIGH-VALUES" | "HIGH-VALUE" => TokenKind::HighValues,

            // Statements
            "MOVE" => TokenKind::Move,
            "TO" => TokenKind::To,
            "CORRESPONDING" => TokenKind::Corresponding,
            "CORR" => TokenKind::Corr,
            "ADD" => TokenKind::Add,
            "SUBTRACT" => TokenKind::Subtract,
            "FROM" => TokenKind::From,
            "GIVING" => TokenKind::Giving,
            "MULTIPLY" => TokenKind::Multiply,
            "BY" => TokenKind::By,
            "DIVIDE" => TokenKind::Divide,
            "REMAINDER" => TokenKind::Remainder,
            "COMPUTE" => TokenKind::Compute,
            "ACCEPT" => TokenKind::Accept,
            "IF" => TokenKind::If,
            "ELSE" => TokenKind::Else,
            "END-IF" => TokenKind::EndIf,
            "THEN" => TokenKind::Then,
            "EVALUATE" => TokenKind::Evaluate,
            "WHEN" => TokenKind::When,
            "OTHER" => TokenKind::Other,
            "END-EVALUATE" => TokenKind::EndEvaluate,
            "TRUE" => TokenKind::True,
            "FALSE" => TokenKind::False,
            "PERFORM" => TokenKind::Perform,
            "END-PERFORM" => TokenKind::EndPerform,
            "UNTIL" => TokenKind::Until,
            "VARYING" => TokenKind::Varying,
            "STRING" => TokenKind::String_,
            "DELIMITED" => TokenKind::Delimited,
            "SIZE" => TokenKind::Size,
            "INTO" => TokenKind::Into,
            "END-STRING" => TokenKind::EndString,
            "UNSTRING" => TokenKind::Unstring,
            "END-UNSTRING" => TokenKind::EndUnstring,
            "INSPECT" => TokenKind::Inspect,
            "TALLYING" => TokenKind::Tallying,
            "REPLACING" => TokenKind::Replacing,
            "ALL" => TokenKind::All,
            "LEADING" => TokenKind::Leading,
            "FIRST" => TokenKind::First,
            "FOR" => TokenKind::For,
            "SEARCH" => TokenKind::Search,
            "END-SEARCH" => TokenKind::EndSearch,
            "AT" => TokenKind::At,
            "CALL" => TokenKind::Call,
            "USING" => TokenKind::Using,
            "END-CALL" => TokenKind::EndCall,
            "GO" => TokenKind::Go,
            "STOP" => TokenKind::Ident("STOP".to_string()),
            "RUN" => TokenKind::Ident("RUN".to_string()),
            "GOBACK" => TokenKind::Goback,
            "INITIALIZE" => TokenKind::Initialize,
            "SET" => TokenKind::Set,
            "CONTINUE" => TokenKind::Continue,
            "RAISE" => TokenKind::Raise,
            "RESUME" => TokenKind::Resume,
            "EXCEPTION" => TokenKind::Exception,
            "REWRITE" => TokenKind::Rewrite,
            "END-REWRITE" => TokenKind::EndRewrite,
            "DELETE" => TokenKind::Delete,
            "END-DELETE" => TokenKind::EndDelete,
            "START" => TokenKind::Start,
            "END-START" => TokenKind::EndStart,
            "EXIT" => TokenKind::Exit,
            "PARAGRAPH" => TokenKind::Paragraph,
            "MERGE" => TokenKind::Merge,
            "COPY" => TokenKind::Copy,
            "CONVERTING" => TokenKind::Converting,
            "NUMERIC" => TokenKind::Numeric,
            "ALPHABETIC" => TokenKind::Alphabetic,
            "ALPHABETIC-LOWER" => TokenKind::AlphabeticLower,
            "ALPHABETIC-UPPER" => TokenKind::AlphabeticUpper,
            "POSITIVE" => TokenKind::Positive,
            "NEGATIVE" => TokenKind::Negative,
            "COUNT" => TokenKind::Count,
            "CLASS-ID" => TokenKind::ClassId,
            "METHOD-ID" => TokenKind::MethodId,
            "INVOKE" => TokenKind::Invoke,
            "END-CLASS" => TokenKind::EndClass,
            "END-METHOD" => TokenKind::EndMethod,
            "TYPEDEF" => TokenKind::Typedef,
            "VALIDATE" => TokenKind::Validate,
            "END-VALIDATE" => TokenKind::EndValidate,
            "FREE" => TokenKind::Free,
            "ALLOCATE" => TokenKind::Allocate,
            "BOOLEAN" => TokenKind::Boolean,
            "FLOAT-LONG" => TokenKind::FloatLong,
            "FLOAT-SHORT" => TokenKind::FloatShort,
            "NATIONAL" => TokenKind::National,
            "PROPERTY" => TokenKind::Property,
            "INHERITS" => TokenKind::Inherits,
            "IMPLEMENTS" => TokenKind::Implements,
            "INTERFACE-ID" => TokenKind::InterfaceId,
            "END-INTERFACE" => TokenKind::EndInterface,
            "FACTORY" => TokenKind::Factory,
            "OBJECT" => TokenKind::Object_,
            "END-FACTORY" => TokenKind::EndFactory,
            "END-OBJECT" => TokenKind::EndObject,
            "NEW" => TokenKind::New,
            "SELF" => TokenKind::Self_,
            "OVERRIDE" => TokenKind::Override,
            "GET" => TokenKind::Get,
            "ASYNC" => TokenKind::Async,
            "WAIT" => TokenKind::Wait,
            "RUN-UNIT" => TokenKind::RunUnit,
            "MONITOR" => TokenKind::Monitor,
            "LOCK" => TokenKind::Lock,
            "UNLOCK" => TokenKind::Unlock,
            "YIELD" => TokenKind::Yield_,
            "SUSPEND" => TokenKind::Suspend,
            "JSON" => TokenKind::Json,
            "GENERATE" => TokenKind::Generate,
            "PARSE" => TokenKind::Parse,
            "XML" => TokenKind::Xml,
            "NOT" => TokenKind::Not,
            "AND" => TokenKind::And,
            "OR" => TokenKind::Or,
            "THRU" | "THROUGH" => TokenKind::Thru,
            "WITH" => TokenKind::With,
            "TEST" => TokenKind::Test,
            "BEFORE" => TokenKind::Before,
            "AFTER" => TokenKind::After,
            "RETURNING" => TokenKind::Returning,
            "UP" => TokenKind::Up,
            "DOWN" => TokenKind::Down,
            "CHARACTERS" => TokenKind::Characters,
            "OPEN" => TokenKind::Open,
            "CLOSE" => TokenKind::Close,
            "READ" => TokenKind::Read,
            "WRITE" => TokenKind::Write,
            "END-READ" => TokenKind::EndRead,
            "END-WRITE" => TokenKind::EndWrite,
            "INPUT" => TokenKind::Input,
            "OUTPUT" => TokenKind::Output,
            "EXTEND" => TokenKind::Extend,
            "I-O" => TokenKind::IoMode,
            "SORT" => TokenKind::Sort,
            "END-SORT" => TokenKind::EndSort,
            "ON" => TokenKind::On,
            "ASCENDING" => TokenKind::Ascending,
            "DESCENDING" => TokenKind::Descending,
            "KEY" => TokenKind::Key,
            "RELEASE" => TokenKind::Release,
            "RETURN" => TokenKind::Return_,
            "END-RETURN" => TokenKind::EndReturn,
            "SELECT" => TokenKind::Select,
            "ASSIGN" => TokenKind::Assign,
            "FILE-STATUS" => TokenKind::FileStatus,
            "ORGANIZATION" => TokenKind::Organization,
            "SEQUENTIAL" => TokenKind::Sequential,
            "RELATIVE" => TokenKind::Relative,
            "LINE" => TokenKind::Line,
            "FILLER" => TokenKind::Filler,
            "BLANK" => TokenKind::Blank,
            "JUSTIFIED" => TokenKind::Justified,
            "RIGHT" => TokenKind::Right,
            "LEFT" => TokenKind::Left,
            "ALSO" => TokenKind::Also,
            "REFERENCE" => TokenKind::Reference,
            "CONTENT" => TokenKind::Content,
            "DEPENDING" => TokenKind::Ident("DEPENDING".to_string()),

            // Intrinsic functions
            "FUNCTION" => TokenKind::Function,
            "LENGTH" => TokenKind::Length,
            "UPPER-CASE" => TokenKind::UpperCase,
            "LOWER-CASE" => TokenKind::LowerCase,
            "TRIM" => TokenKind::Trim,
            "REVERSE" => TokenKind::Reverse,
            "CURRENT-DATE" => TokenKind::CurrentDate,
            "MAX" => TokenKind::Max,
            "MIN" => TokenKind::Min,
            "MOD" => TokenKind::Mod,
            "REM" => TokenKind::Rem,
            "NUMVAL" => TokenKind::Numval,
            "NUMVAL-C" => TokenKind::NumvalC,
            "ORD" => TokenKind::Ord,
            "CHAR" => TokenKind::Char,
            "SUBSTITUTE" => TokenKind::Substitute,
            "SQRT" => TokenKind::Sqrt,
            "SUM" => TokenKind::Sum,
            "INTEGER" => TokenKind::Integer,
            "ABS" => TokenKind::Abs,
            "LOG" => TokenKind::Log,
            "LOG10" => TokenKind::Log10,
            "EXP" => TokenKind::Exp,
            "SIN" => TokenKind::Sin,
            "COS" => TokenKind::Cos,
            "TAN" => TokenKind::Tan,
            "ASIN" => TokenKind::Asin,
            "ACOS" => TokenKind::Acos,
            "ATAN" => TokenKind::Atan,
            "CEILING" => TokenKind::Ceiling,
            "FLOOR" => TokenKind::Floor,
            "SIGN" => TokenKind::Sign,
            "RANDOM" => TokenKind::Random,
            "MEAN" => TokenKind::Mean,
            "MEDIAN" => TokenKind::Median,
            "VARIANCE" => TokenKind::Variance,
            "CONCATENATE" => TokenKind::Concatenate,
            "WHEN-COMPILED" => TokenKind::WhenCompiled,
            "FORMATTED-DATE" => TokenKind::FormattedDate,
            "FORMATTED-TIME" => TokenKind::FormattedTime,
            "DATE-OF-INTEGER" => TokenKind::DateOfInteger,
            "INTEGER-OF-DATE" => TokenKind::IntegerOfDate,
            "POWER" => TokenKind::Power,
            "ANNUITY" => TokenKind::Annuity,
            "PRESENT-VALUE" => TokenKind::PresentValue,
            "TEST-NUMVAL" => TokenKind::TestNumval,

            "OF" => TokenKind::Of,
            "IN" => TokenKind::In,
            "IS" => TokenKind::Ident("IS".to_string()),

            _ => {
                // Check if it's a level number (01-88)
                if word.len() == 2 && word.chars().all(|c| c.is_ascii_digit()) {
                    let level: u8 = word.parse().unwrap_or(0);
                    if matches!(level, 1..=49 | 66 | 77 | 88) {
                        return Ok(Token { kind: TokenKind::Level(level), line });
                    }
                }
                // PIC clause content: X(20), 9(5)V99, S9(5), etc.
                TokenKind::Ident(upper)
            }
        };
        Ok(Token { kind, line })
    }
}
