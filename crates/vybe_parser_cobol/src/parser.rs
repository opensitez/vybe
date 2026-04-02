use crate::token::{Token, TokenKind};
use crate::ast::*;
use crate::lexer::Lexer;

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(source: &str) -> Result<Self, String> {
        let mut lexer = Lexer::new(source);
        let tokens = lexer.tokenize()?;
        Ok(Parser { tokens, pos: 0 })
    }

    pub fn parse_program(&mut self) -> Result<Program, String> {
        let mut program_id = String::new();
        let mut author = None;
        let mut data_items = Vec::new();
        let mut paragraphs = Vec::new();
        let mut main_body = Vec::new();

        // Parse IDENTIFICATION DIVISION
        self.expect(&TokenKind::Identification)?;
        self.expect(&TokenKind::Division)?;
        self.expect(&TokenKind::Period)?;

        // PROGRAM-ID
        self.expect(&TokenKind::ProgramId)?;
        self.expect(&TokenKind::Period)?;
        program_id = self.expect_ident()?;
        self.expect(&TokenKind::Period)?;

        // Optional: AUTHOR, DATE-WRITTEN, etc.
        while matches!(self.current(), TokenKind::Author | TokenKind::DateWritten) {
            self.advance();
            self.expect(&TokenKind::Period)?;
            // Skip to next period
            while *self.current() != TokenKind::Period && *self.current() != TokenKind::Eof {
                self.advance();
            }
            if *self.current() == TokenKind::Period { self.advance(); }
        }

        // DATA DIVISION (optional)
        if *self.current() == TokenKind::Data {
            self.advance(); // DATA
            self.expect(&TokenKind::Division)?;
            self.expect(&TokenKind::Period)?;
            data_items = self.parse_data_division()?;
        }

        // PROCEDURE DIVISION
        if *self.current() == TokenKind::Procedure {
            self.advance();
            self.expect(&TokenKind::Division)?;
            // Optional USING clause
            if *self.current() == TokenKind::Using {
                self.advance();
                while *self.current() != TokenKind::Period && *self.current() != TokenKind::Eof {
                    self.advance();
                }
            }
            self.expect(&TokenKind::Period)?;
            // Parse main body + paragraphs
            let (body, paras) = self.parse_procedure_division()?;
            main_body = body;
            paragraphs = paras;
        }

        Ok(Program { program_id, author, data_items, paragraphs, main_body })
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    fn current(&self) -> &TokenKind {
        if self.pos < self.tokens.len() { &self.tokens[self.pos].kind } else { &TokenKind::Eof }
    }

    fn current_line(&self) -> u32 {
        if self.pos < self.tokens.len() { self.tokens[self.pos].line } else { 0 }
    }

    fn advance(&mut self) -> &TokenKind {
        let kind = &self.tokens[self.pos].kind;
        if self.pos < self.tokens.len() { self.pos += 1; }
        kind
    }

    fn expect(&mut self, expected: &TokenKind) -> Result<(), String> {
        if self.current() == expected {
            self.advance();
            Ok(())
        } else {
            Err(format!("Expected {:?}, got {:?} at line {}", expected, self.current(), self.current_line()))
        }
    }

    fn expect_ident(&mut self) -> Result<String, String> {
        match self.current().clone() {
            TokenKind::Ident(name) => { self.advance(); Ok(name) }
            // Many keywords can also be identifiers in COBOL
            _ => {
                let name = format!("{:?}", self.current());
                self.advance();
                Ok(name)
            }
        }
    }

    fn at_end(&self) -> bool {
        self.pos >= self.tokens.len() || *self.current() == TokenKind::Eof
    }

    fn match_token(&mut self, kind: &TokenKind) -> bool {
        if self.current() == kind {
            self.advance();
            true
        } else {
            false
        }
    }

    fn skip_periods(&mut self) {
        while *self.current() == TokenKind::Period { self.advance(); }
    }

    // ------------------------------------------------------------------
    // DATA DIVISION
    // ------------------------------------------------------------------

    fn parse_data_division(&mut self) -> Result<Vec<DataItem>, String> {
        let mut items = Vec::new();

        // WORKING-STORAGE SECTION / LOCAL-STORAGE SECTION / etc.
        while matches!(self.current(), TokenKind::WorkingStorage | TokenKind::LocalStorage | TokenKind::FileSection | TokenKind::Linkage) {
            self.advance(); // section type
            self.expect(&TokenKind::Section)?;
            self.expect(&TokenKind::Period)?;
            items.extend(self.parse_data_items()?);
        }
        Ok(items)
    }

    fn parse_data_items(&mut self) -> Result<Vec<DataItem>, String> {
        let mut items = Vec::new();
        while let TokenKind::Level(level) = self.current().clone() {
            let item = self.parse_data_item(level)?;
            items.push(item);
        }
        Ok(items)
    }

    fn parse_data_item(&mut self, level: u8) -> Result<DataItem, String> {
        self.advance(); // consume level number

        // 88-level condition
        if level == 88 {
            let name = self.expect_ident()?;
            let mut values = Vec::new();
            if self.match_token(&TokenKind::Value) {
                values.push(self.parse_literal()?);
                while self.match_token(&TokenKind::Comma) {
                    values.push(self.parse_literal()?);
                }
            }
            self.skip_to_period();
            return Ok(DataItem {
                level: 88, name: name.clone(), pic: None, value: None,
                occurs: None, redefines: None, usage: None,
                children: Vec::new(),
                conditions: vec![Condition88 { name, values }],
            });
        }

        let name = self.expect_ident()?;

        let mut pic = None;
        let mut value = None;
        let mut occurs = None;
        let mut redefines = None;
        let mut usage = None;

        // Parse clauses until period
        while *self.current() != TokenKind::Period && !self.at_end() {
            match self.current().clone() {
                TokenKind::Pic => {
                    self.advance();
                    // Optional IS
                    if let TokenKind::Ident(s) = self.current() { if s == "IS" { self.advance(); } }
                    pic = Some(self.parse_pic_string()?);
                }
                TokenKind::Value => {
                    self.advance();
                    // Optional IS
                    if let TokenKind::Ident(s) = self.current() { if s == "IS" { self.advance(); } }
                    value = Some(self.parse_literal()?);
                }
                TokenKind::Occurs => {
                    self.advance();
                    if let TokenKind::Number(n) = self.current().clone() {
                        self.advance();
                        occurs = Some(n as u32);
                    }
                    self.match_token(&TokenKind::Times);
                    // Skip INDEXED BY
                    if *self.current() == TokenKind::Indexed {
                        self.advance();
                        self.match_token(&TokenKind::By);
                        self.expect_ident()?;
                    }
                }
                TokenKind::Redefines => {
                    self.advance();
                    redefines = Some(self.expect_ident()?);
                }
                TokenKind::Usage => {
                    self.advance();
                    // Optional IS
                    if let TokenKind::Ident(s) = self.current() { if s == "IS" { self.advance(); } }
                    usage = Some(self.expect_ident()?);
                }
                _ => { self.advance(); } // skip unknown clauses
            }
        }
        self.match_token(&TokenKind::Period);

        // Parse child items (higher level numbers)
        let mut children = Vec::new();
        let mut conditions = Vec::new();
        while let TokenKind::Level(child_level) = self.current().clone() {
            if child_level <= level && child_level != 88 { break; }
            if child_level == 88 {
                let cond_item = self.parse_data_item(88)?;
                for c in cond_item.conditions { conditions.push(c); }
            } else {
                children.push(self.parse_data_item(child_level)?);
            }
        }

        Ok(DataItem { level, name, pic, value, occurs, redefines, usage, children, conditions })
    }

    fn parse_pic_string(&mut self) -> Result<String, String> {
        let mut pic = String::new();
        // PIC can be: X(20), 9(5)V99, S9(5), A(10), etc.
        while !matches!(self.current(), TokenKind::Period | TokenKind::Value | TokenKind::Occurs | TokenKind::Redefines | TokenKind::Usage | TokenKind::Eof) {
            match self.current().clone() {
                TokenKind::Ident(s) => { pic.push_str(&s); self.advance(); }
                TokenKind::Number(n) => { pic.push_str(&n.to_string()); self.advance(); }
                TokenKind::LParen => { pic.push('('); self.advance(); }
                TokenKind::RParen => { pic.push(')'); self.advance(); }
                _ => break,
            }
        }
        Ok(pic)
    }

    fn parse_literal(&mut self) -> Result<Literal, String> {
        match self.current().clone() {
            TokenKind::Number(n) => { self.advance(); Ok(Literal::Num(n)) }
            TokenKind::Str(s) => { self.advance(); Ok(Literal::Str(s)) }
            TokenKind::Spaces => { self.advance(); Ok(Literal::Spaces) }
            TokenKind::Zeros => { self.advance(); Ok(Literal::Zeros) }
            TokenKind::LowValues => { self.advance(); Ok(Literal::LowValues) }
            TokenKind::HighValues => { self.advance(); Ok(Literal::HighValues) }
            TokenKind::True => { self.advance(); Ok(Literal::True) }
            TokenKind::False => { self.advance(); Ok(Literal::False) }
            TokenKind::Minus => {
                self.advance();
                if let TokenKind::Number(n) = self.current().clone() {
                    self.advance();
                    Ok(Literal::Num(-n))
                } else {
                    Ok(Literal::Num(0.0))
                }
            }
            _ => Err(format!("Expected literal, got {:?} at line {}", self.current(), self.current_line())),
        }
    }

    fn skip_to_period(&mut self) {
        while *self.current() != TokenKind::Period && !self.at_end() {
            self.advance();
        }
        self.match_token(&TokenKind::Period);
    }

    // ------------------------------------------------------------------
    // PROCEDURE DIVISION
    // ------------------------------------------------------------------

    fn parse_procedure_division(&mut self) -> Result<(Vec<Statement>, Vec<Paragraph>), String> {
        let mut main_body = Vec::new();
        let mut paragraphs = Vec::new();

        // Parse statements until we hit a paragraph name (ident followed by period at start)
        // or end of input
        loop {
            if self.at_end() { break; }

            // Check for paragraph definition: IDENT.
            if self.is_paragraph_start() {
                break;
            }

            if let Some(stmt) = self.parse_statement()? {
                main_body.push(stmt);
            }
            self.skip_periods();
        }

        // Parse paragraphs
        while !self.at_end() {
            if self.is_paragraph_start() {
                let name = self.expect_ident()?;
                self.expect(&TokenKind::Period)?;
                let mut body = Vec::new();
                while !self.at_end() && !self.is_paragraph_start() {
                    if let Some(stmt) = self.parse_statement()? {
                        body.push(stmt);
                    }
                    self.skip_periods();
                }
                paragraphs.push(Paragraph { name, body });
            } else {
                break;
            }
        }

        Ok((main_body, paragraphs))
    }

    fn is_paragraph_start(&self) -> bool {
        if let TokenKind::Ident(_) = self.current() {
            if self.pos + 1 < self.tokens.len() && self.tokens[self.pos + 1].kind == TokenKind::Period {
                // Check that the next-next token is NOT a data keyword
                // (to avoid confusing "STOP. RUN." with a paragraph)
                if let TokenKind::Ident(name) = self.current() {
                    // Known paragraphs end with -PARA, -SECTION, etc.
                    // But really any IDENT followed by period at the start of a statement
                    // where the ident isn't a known keyword is a paragraph
                    return !matches!(name.as_str(),
                        "STOP" | "RUN" | "IS" | "DISPLAY" | "ACCEPT" | "MOVE" |
                        "ADD" | "SUBTRACT" | "MULTIPLY" | "DIVIDE" | "COMPUTE" |
                        "CALL" | "SET" | "INITIALIZE" | "INSPECT" | "STRING" | "UNSTRING"
                    );
                }
            }
        }
        false
    }

    fn parse_statement(&mut self) -> Result<Option<Statement>, String> {
        match self.current().clone() {
            TokenKind::DisplayKw => self.parse_display(),
            TokenKind::Accept => self.parse_accept(),
            TokenKind::Move => self.parse_move(),
            TokenKind::Add => self.parse_add(),
            TokenKind::Subtract => self.parse_subtract(),
            TokenKind::Multiply => self.parse_multiply(),
            TokenKind::Divide => self.parse_divide(),
            TokenKind::Compute => self.parse_compute(),
            TokenKind::If => self.parse_if(),
            TokenKind::Evaluate => self.parse_evaluate(),
            TokenKind::Perform => self.parse_perform(),
            TokenKind::String_ => self.parse_string_stmt(),
            TokenKind::Unstring => self.parse_unstring(),
            TokenKind::Inspect => self.parse_inspect(),
            TokenKind::Call => self.parse_call(),
            TokenKind::Initialize => self.parse_initialize(),
            TokenKind::Set => self.parse_set(),
            TokenKind::Continue => { self.advance(); Ok(Some(Statement::Continue)) }
            TokenKind::Goback => { self.advance(); Ok(Some(Statement::Goback)) }
            TokenKind::Go => self.parse_goto(),
            TokenKind::Raise => self.parse_raise(),
            TokenKind::Json => self.parse_json(),
            TokenKind::Open => self.parse_open(),
            TokenKind::Close => { self.advance(); let n = self.expect_ident()?; Ok(Some(Statement::Close(n))) }
            TokenKind::Read => self.parse_read(),
            TokenKind::Write => self.parse_write(),
            TokenKind::Sort => self.parse_sort(),
            TokenKind::Search => self.parse_search(),
            TokenKind::Ident(ref s) if s == "STOP" => {
                self.advance();
                if let TokenKind::Ident(s) = self.current() {
                    if s == "RUN" { self.advance(); }
                }
                Ok(Some(Statement::StopRun))
            }
            TokenKind::Period | TokenKind::Eof => { Ok(None) }
            _ => {
                self.advance();
                Ok(None)
            }
        }
    }

    // ------------------------------------------------------------------
    // Statement parsers
    // ------------------------------------------------------------------

    fn parse_display(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // DISPLAY
        let mut exprs = Vec::new();
        while !matches!(self.current(), TokenKind::Period | TokenKind::Eof | TokenKind::EndIf | TokenKind::Else | TokenKind::EndPerform | TokenKind::EndEvaluate | TokenKind::When | TokenKind::Other) {
            exprs.push(self.parse_expr()?);
        }
        Ok(Some(Statement::Display(exprs)))
    }

    fn parse_accept(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // ACCEPT
        let name = self.expect_ident()?;
        // Check for FROM DATE/TIME/DAY
        if self.match_token(&TokenKind::From) {
            let source = match self.current() {
                TokenKind::Ident(s) if s == "DATE" => { self.advance(); AcceptSource::Date }
                TokenKind::Ident(s) if s == "TIME" => { self.advance(); AcceptSource::Time }
                TokenKind::Ident(s) if s == "DAY" => { self.advance(); AcceptSource::Day }
                TokenKind::Ident(s) if s == "DAY-OF-WEEK" => { self.advance(); AcceptSource::DayOfWeek }
                _ => AcceptSource::Console,
            };
            return Ok(Some(Statement::AcceptFrom { var: name, source }));
        }
        Ok(Some(Statement::Accept(name)))
    }

    fn parse_move(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // MOVE
        if self.match_token(&TokenKind::Corresponding) || self.match_token(&TokenKind::Corr) {
            let src = self.expect_ident()?;
            self.expect(&TokenKind::To)?;
            let dst = self.expect_ident()?;
            return Ok(Some(Statement::MoveCorresponding { src, dst }));
        }
        let src = self.parse_expr()?;
        self.expect(&TokenKind::To)?;
        let mut dsts = vec![self.expect_ident()?];
        while let TokenKind::Ident(_) = self.current() {
            if matches!(self.current(), TokenKind::Period) { break; }
            dsts.push(self.expect_ident()?);
        }
        Ok(Some(Statement::Move { src, dsts }))
    }

    fn parse_add(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // ADD
        let mut srcs = vec![self.parse_expr()?];
        while !matches!(self.current(), TokenKind::To | TokenKind::Giving | TokenKind::Period | TokenKind::Eof) {
            srcs.push(self.parse_expr()?);
        }
        self.expect(&TokenKind::To)?;
        let to = self.expect_ident()?;
        let giving = if self.match_token(&TokenKind::Giving) {
            Some(self.expect_ident()?)
        } else { None };
        Ok(Some(Statement::Add { srcs, to, giving }))
    }

    fn parse_subtract(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // SUBTRACT
        let src = self.parse_expr()?;
        self.expect(&TokenKind::From)?;
        let from = self.expect_ident()?;
        let giving = if self.match_token(&TokenKind::Giving) {
            Some(self.expect_ident()?)
        } else { None };
        Ok(Some(Statement::Subtract { src, from, giving }))
    }

    fn parse_multiply(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // MULTIPLY
        let src = self.parse_expr()?;
        self.expect(&TokenKind::By)?;
        let by = self.expect_ident()?;
        let giving = if self.match_token(&TokenKind::Giving) {
            Some(self.expect_ident()?)
        } else { None };
        Ok(Some(Statement::Multiply { src, by, giving }))
    }

    fn parse_divide(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // DIVIDE
        let src = self.parse_expr()?;
        self.expect(&TokenKind::By)?;
        let by = self.parse_expr()?;
        self.expect(&TokenKind::Giving)?;
        let giving = self.expect_ident()?;
        let remainder = if self.match_token(&TokenKind::Remainder) {
            Some(self.expect_ident()?)
        } else { None };
        Ok(Some(Statement::Divide { src, by, giving, remainder }))
    }

    fn parse_compute(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // COMPUTE
        let dst = self.expect_ident()?;
        self.expect(&TokenKind::Eq)?;
        let expr = self.parse_expr()?;
        Ok(Some(Statement::Compute { dst, expr }))
    }

    fn parse_if(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // IF
        let test = self.parse_condition()?;
        self.match_token(&TokenKind::Then);
        let mut body = Vec::new();
        while !matches!(self.current(), TokenKind::Else | TokenKind::EndIf | TokenKind::Period | TokenKind::Eof) {
            if let Some(stmt) = self.parse_statement()? {
                body.push(stmt);
            }
            self.skip_periods();
        }
        let else_body = if self.match_token(&TokenKind::Else) {
            let mut stmts = Vec::new();
            while !matches!(self.current(), TokenKind::EndIf | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? {
                    stmts.push(stmt);
                }
                self.skip_periods();
            }
            Some(stmts)
        } else { None };
        self.match_token(&TokenKind::EndIf);
        Ok(Some(Statement::If { test, body, else_body }))
    }

    fn parse_evaluate(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // EVALUATE
        let subject = self.parse_expr()?;
        let mut whens = Vec::new();
        let mut other = None;

        while self.match_token(&TokenKind::When) {
            if self.match_token(&TokenKind::Other) {
                let mut body = Vec::new();
                while !matches!(self.current(), TokenKind::When | TokenKind::EndEvaluate | TokenKind::Period | TokenKind::Eof) {
                    if let Some(stmt) = self.parse_statement()? {
                        body.push(stmt);
                    }
                    self.skip_periods();
                }
                other = Some(body);
                break;
            }
            let mut values = vec![self.parse_expr()?];
            // Handle WHEN val1 ALSO val2 or multiple values
            while self.match_token(&TokenKind::Comma) {
                values.push(self.parse_expr()?);
            }
            let mut body = Vec::new();
            while !matches!(self.current(), TokenKind::When | TokenKind::EndEvaluate | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? {
                    body.push(stmt);
                }
                self.skip_periods();
            }
            whens.push(WhenClause { values, body });
        }
        self.match_token(&TokenKind::EndEvaluate);
        Ok(Some(Statement::Evaluate { subject, whens, other }))
    }

    fn parse_perform(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // PERFORM

        // PERFORM VARYING
        if self.match_token(&TokenKind::Varying) {
            let var = self.expect_ident()?;
            self.expect(&TokenKind::From)?;
            let from = self.parse_expr()?;
            self.expect(&TokenKind::By)?;
            let by = self.parse_expr()?;
            self.expect(&TokenKind::Until)?;
            let until = self.parse_condition()?;
            let mut body = Vec::new();
            while !matches!(self.current(), TokenKind::EndPerform | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? {
                    body.push(stmt);
                }
                self.skip_periods();
            }
            self.match_token(&TokenKind::EndPerform);
            return Ok(Some(Statement::PerformVarying { var, from, by, until, body }));
        }

        // PERFORM UNTIL
        if self.match_token(&TokenKind::Until) {
            let test = self.parse_condition()?;
            let mut body = Vec::new();
            while !matches!(self.current(), TokenKind::EndPerform | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? {
                    body.push(stmt);
                }
                self.skip_periods();
            }
            self.match_token(&TokenKind::EndPerform);
            return Ok(Some(Statement::PerformUntil { test, body }));
        }

        // PERFORM n TIMES
        if let TokenKind::Number(_) = self.current() {
            let count = self.parse_expr()?;
            self.expect(&TokenKind::Times)?;
            let mut body = Vec::new();
            while !matches!(self.current(), TokenKind::EndPerform | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? {
                    body.push(stmt);
                }
                self.skip_periods();
            }
            self.match_token(&TokenKind::EndPerform);
            return Ok(Some(Statement::PerformTimes { count, body }));
        }

        // PERFORM paragraph-name [THRU paragraph-name]
        if let TokenKind::Ident(name) = self.current().clone() {
            self.advance();
            if self.match_token(&TokenKind::Thru) {
                let thru = self.expect_ident()?;
                return Ok(Some(Statement::PerformThru { from: name, thru }));
            }
            return Ok(Some(Statement::PerformParagraph(name)));
        }

        Ok(None)
    }

    fn parse_string_stmt(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // STRING
        let mut sources = Vec::new();
        while !matches!(self.current(), TokenKind::Into | TokenKind::Period | TokenKind::Eof) {
            let value = self.parse_expr()?;
            let delimited_by = if self.match_token(&TokenKind::Delimited) {
                self.expect(&TokenKind::By)?;
                if self.match_token(&TokenKind::Size) {
                    DelimitedBy::Size
                } else if let TokenKind::Str(s) = self.current().clone() {
                    self.advance();
                    DelimitedBy::Value(s)
                } else {
                    let name = self.expect_ident()?;
                    DelimitedBy::Value(name)
                }
            } else {
                DelimitedBy::Size
            };
            sources.push(StringSource { value, delimited_by });
        }
        self.expect(&TokenKind::Into)?;
        let into = self.expect_ident()?;
        self.match_token(&TokenKind::EndString);
        Ok(Some(Statement::StringConcat { sources, into }))
    }

    fn parse_unstring(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // UNSTRING
        let src = self.expect_ident()?;
        let mut delimiters = Vec::new();
        if self.match_token(&TokenKind::Delimited) {
            self.expect(&TokenKind::By)?;
            if let TokenKind::Str(s) = self.current().clone() {
                self.advance();
                delimiters.push(s);
            }
        }
        self.expect(&TokenKind::Into)?;
        let mut into = Vec::new();
        while !matches!(self.current(), TokenKind::EndUnstring | TokenKind::Period | TokenKind::Eof) {
            into.push(self.expect_ident()?);
        }
        self.match_token(&TokenKind::EndUnstring);
        Ok(Some(Statement::Unstring { src, delimiters, into }))
    }

    fn parse_inspect(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // INSPECT
        let var = self.expect_ident()?;
        if self.match_token(&TokenKind::Tallying) {
            let counter = self.expect_ident()?;
            self.expect(&TokenKind::For)?;
            let mode = self.parse_inspect_mode()?;
            let target = if let TokenKind::Str(s) = self.current().clone() {
                self.advance(); s
            } else { self.expect_ident()? };
            return Ok(Some(Statement::InspectTallying { var, counter, mode, target }));
        }
        if self.match_token(&TokenKind::Replacing) {
            let mode = self.parse_inspect_mode()?;
            let old = if let TokenKind::Str(s) = self.current().clone() {
                self.advance(); s
            } else { self.expect_ident()? };
            self.expect(&TokenKind::By)?;
            let new = if let TokenKind::Str(s) = self.current().clone() {
                self.advance(); s
            } else { self.expect_ident()? };
            return Ok(Some(Statement::InspectReplacing { var, mode, old, new }));
        }
        Ok(None)
    }

    fn parse_inspect_mode(&mut self) -> Result<InspectMode, String> {
        if self.match_token(&TokenKind::All) { Ok(InspectMode::All) }
        else if self.match_token(&TokenKind::Leading) { Ok(InspectMode::Leading) }
        else if self.match_token(&TokenKind::First) { Ok(InspectMode::First) }
        else { Ok(InspectMode::All) }
    }

    fn parse_call(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // CALL
        let name = if let TokenKind::Str(s) = self.current().clone() {
            self.advance(); s
        } else { self.expect_ident()? };
        let mut args = Vec::new();
        if self.match_token(&TokenKind::Using) {
            while !matches!(self.current(), TokenKind::EndCall | TokenKind::Period | TokenKind::Eof) {
                args.push(self.expect_ident()?);
            }
        }
        self.match_token(&TokenKind::EndCall);
        Ok(Some(Statement::Call { name, args }))
    }

    fn parse_initialize(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // INITIALIZE
        let name = self.expect_ident()?;
        Ok(Some(Statement::Initialize(name)))
    }

    fn parse_set(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // SET
        let target = self.expect_ident()?;
        self.expect(&TokenKind::To)?;
        let value = self.match_token(&TokenKind::True);
        if !value { self.match_token(&TokenKind::False); }
        Ok(Some(Statement::Set { target, value }))
    }

    fn parse_goto(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // GO
        self.match_token(&TokenKind::To);
        let name = self.expect_ident()?;
        Ok(Some(Statement::GoTo(name)))
    }

    fn parse_raise(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // RAISE
        self.match_token(&TokenKind::Exception);
        let msg = self.parse_expr()?;
        Ok(Some(Statement::Raise(msg)))
    }

    fn parse_json(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // JSON
        if self.match_token(&TokenKind::Generate) {
            let dst = self.expect_ident()?;
            self.expect(&TokenKind::From)?;
            let src = self.expect_ident()?;
            return Ok(Some(Statement::JsonGenerate { dst, src }));
        }
        if self.match_token(&TokenKind::Parse) {
            let src = self.expect_ident()?;
            self.expect(&TokenKind::Into)?;
            let dst = self.expect_ident()?;
            return Ok(Some(Statement::JsonParse { src, dst }));
        }
        Ok(None)
    }

    fn parse_open(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // OPEN
        let mode = match self.current() {
            TokenKind::Input => { self.advance(); FileMode::Input }
            TokenKind::Output => { self.advance(); FileMode::Output }
            TokenKind::Extend => { self.advance(); FileMode::Extend }
            TokenKind::IoMode => { self.advance(); FileMode::IoMode }
            _ => FileMode::Input,
        };
        let file = self.expect_ident()?;
        Ok(Some(Statement::Open { mode, file }))
    }

    fn parse_read(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // READ
        let file = self.expect_ident()?;
        let into = if self.match_token(&TokenKind::Into) {
            Some(self.expect_ident()?)
        } else { None };
        self.match_token(&TokenKind::EndRead);
        Ok(Some(Statement::ReadFile { file, into }))
    }

    fn parse_write(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // WRITE
        let record = self.expect_ident()?;
        let from = if self.match_token(&TokenKind::From) {
            Some(self.expect_ident()?)
        } else { None };
        self.match_token(&TokenKind::EndWrite);
        Ok(Some(Statement::WriteFile { record, from }))
    }

    fn parse_sort(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // SORT
        let file = self.expect_ident()?;
        self.match_token(&TokenKind::On);
        let ascending = if self.match_token(&TokenKind::Descending) { false }
        else { self.match_token(&TokenKind::Ascending); true };
        self.match_token(&TokenKind::Key);
        let key = self.expect_ident()?;
        Ok(Some(Statement::Sort { file, ascending, key }))
    }

    fn parse_search(&mut self) -> Result<Option<Statement>, String> {
        self.advance(); // SEARCH
        let table = self.expect_ident()?;
        let mut at_end = Vec::new();
        let mut when_cond = Expr::Bool(true);
        let mut when_body = Vec::new();

        // AT END
        if self.match_token(&TokenKind::At) {
            if let TokenKind::Ident(s) = self.current().clone() {
                if s == "END" { self.advance(); }
            }
            while !matches!(self.current(), TokenKind::When | TokenKind::EndSearch | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? { at_end.push(stmt); }
                self.skip_periods();
            }
        }

        // WHEN condition
        if self.match_token(&TokenKind::When) {
            when_cond = self.parse_condition()?;
            while !matches!(self.current(), TokenKind::EndSearch | TokenKind::Period | TokenKind::Eof) {
                if let Some(stmt) = self.parse_statement()? { when_body.push(stmt); }
                self.skip_periods();
            }
        }
        self.match_token(&TokenKind::EndSearch);
        Ok(Some(Statement::SearchTable { table, at_end, when_cond, when_body }))
    }

    // ------------------------------------------------------------------
    // Expression / Condition parsing
    // ------------------------------------------------------------------

    fn parse_condition(&mut self) -> Result<Expr, String> {
        self.parse_or_expr()
    }

    fn parse_or_expr(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_and_expr()?;
        while self.match_token(&TokenKind::Or) {
            let right = self.parse_and_expr()?;
            left = Expr::Logic { op: LogicOp::Or, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_and_expr(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_not_expr()?;
        while self.match_token(&TokenKind::And) {
            let right = self.parse_not_expr()?;
            left = Expr::Logic { op: LogicOp::And, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_not_expr(&mut self) -> Result<Expr, String> {
        if self.match_token(&TokenKind::Not) {
            let expr = self.parse_comparison()?;
            return Ok(Expr::Not(Box::new(expr)));
        }
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Expr, String> {
        let left = self.parse_expr()?;
        let op = match self.current() {
            TokenKind::Eq => Some(CmpOp::Eq),
            TokenKind::Gt => Some(CmpOp::Gt),
            TokenKind::Lt => Some(CmpOp::Lt),
            TokenKind::GtEq => Some(CmpOp::Ge),
            TokenKind::LtEq => Some(CmpOp::Le),
            TokenKind::Not => {
                // NOT = (not equal)
                if self.pos + 1 < self.tokens.len() && self.tokens[self.pos + 1].kind == TokenKind::Eq {
                    self.advance(); // NOT
                    Some(CmpOp::Ne)
                } else {
                    None
                }
            }
            _ => None,
        };
        if let Some(op) = op {
            self.advance();
            let right = self.parse_expr()?;
            Ok(Expr::Compare { op, left: Box::new(left), right: Box::new(right) })
        } else {
            Ok(left)
        }
    }

    fn parse_expr(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_term()?;
        while matches!(self.current(), TokenKind::Plus | TokenKind::Minus) {
            let op = if self.match_token(&TokenKind::Plus) { BinOp::Add } else { self.advance(); BinOp::Sub };
            let right = self.parse_term()?;
            left = Expr::BinOp { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_term(&mut self) -> Result<Expr, String> {
        let mut left = self.parse_power()?;
        while matches!(self.current(), TokenKind::Star | TokenKind::Slash) {
            let op = if self.match_token(&TokenKind::Star) { BinOp::Mul } else { self.advance(); BinOp::Div };
            let right = self.parse_power()?;
            left = Expr::BinOp { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<Expr, String> {
        let left = self.parse_atom()?;
        if self.match_token(&TokenKind::StarStar) {
            let right = self.parse_power()?;
            Ok(Expr::BinOp { op: BinOp::Pow, left: Box::new(left), right: Box::new(right) })
        } else {
            Ok(left)
        }
    }

    fn parse_atom(&mut self) -> Result<Expr, String> {
        match self.current().clone() {
            TokenKind::Number(n) => { self.advance(); Ok(Expr::Lit(Literal::Num(n))) }
            TokenKind::Str(s) => { self.advance(); Ok(Expr::Lit(Literal::Str(s))) }
            TokenKind::Spaces => { self.advance(); Ok(Expr::Lit(Literal::Spaces)) }
            TokenKind::Zeros => { self.advance(); Ok(Expr::Lit(Literal::Zeros)) }
            TokenKind::True => { self.advance(); Ok(Expr::Bool(true)) }
            TokenKind::False => { self.advance(); Ok(Expr::Bool(false)) }
            TokenKind::LParen => {
                self.advance();
                let expr = self.parse_expr()?;
                self.expect(&TokenKind::RParen)?;
                Ok(expr)
            }
            TokenKind::Function => {
                self.advance();
                self.parse_function_call()
            }
            TokenKind::Minus => {
                self.advance();
                let expr = self.parse_atom()?;
                Ok(Expr::BinOp { op: BinOp::Sub, left: Box::new(Expr::Lit(Literal::Num(0.0))), right: Box::new(expr) })
            }
            TokenKind::Ident(name) => {
                self.advance();
                // Check for subscript or reference modification: WS-ITEM(1) or WS-ITEM(1:5)
                if *self.current() == TokenKind::LParen {
                    self.advance();
                    let first = self.parse_expr()?;
                    // Check for : (reference modification)
                    if *self.current() == TokenKind::Colon {
                        self.advance();
                        let length = if *self.current() != TokenKind::RParen {
                            Some(Box::new(self.parse_expr()?))
                        } else {
                            None
                        };
                        self.expect(&TokenKind::RParen)?;
                        return Ok(Expr::RefMod { name, start: Box::new(first), length });
                    }
                    self.expect(&TokenKind::RParen)?;
                    Ok(Expr::Subscript(name, Box::new(first)))
                }
                // Check for qualified name: X OF Y
                else if *self.current() == TokenKind::Of || *self.current() == TokenKind::In {
                    self.advance();
                    let parent = self.expect_ident()?;
                    Ok(Expr::Qualified(name, parent))
                }
                else {
                    Ok(Expr::Ident(name))
                }
            }
            _ => Err(format!("Expected expression, got {:?} at line {}", self.current(), self.current_line())),
        }
    }

    fn parse_function_call(&mut self) -> Result<Expr, String> {
        let name = match self.current().clone() {
            TokenKind::Length => { self.advance(); "LENGTH".to_string() }
            TokenKind::UpperCase => { self.advance(); "UPPER-CASE".to_string() }
            TokenKind::LowerCase => { self.advance(); "LOWER-CASE".to_string() }
            TokenKind::Trim => { self.advance(); "TRIM".to_string() }
            TokenKind::Reverse => { self.advance(); "REVERSE".to_string() }
            TokenKind::CurrentDate => { self.advance(); "CURRENT-DATE".to_string() }
            TokenKind::Max => { self.advance(); "MAX".to_string() }
            TokenKind::Min => { self.advance(); "MIN".to_string() }
            TokenKind::Mod => { self.advance(); "MOD".to_string() }
            TokenKind::Rem => { self.advance(); "REM".to_string() }
            TokenKind::Numval => { self.advance(); "NUMVAL".to_string() }
            TokenKind::Substitute => { self.advance(); "SUBSTITUTE".to_string() }
            TokenKind::Sqrt => { self.advance(); "SQRT".to_string() }
            TokenKind::Sum => { self.advance(); "SUM".to_string() }
            TokenKind::Integer => { self.advance(); "INTEGER".to_string() }
            TokenKind::Abs => { self.advance(); "ABS".to_string() }
            TokenKind::Ord => { self.advance(); "ORD".to_string() }
            TokenKind::Char => { self.advance(); "CHAR".to_string() }
            TokenKind::Ident(name) => { self.advance(); name }
            _ => { return Err(format!("Expected function name at line {}", self.current_line())); }
        };

        // Parse args (may or may not have parens)
        let mut args = Vec::new();
        if self.match_token(&TokenKind::LParen) {
            while *self.current() != TokenKind::RParen && !self.at_end() {
                args.push(self.parse_expr()?);
                self.match_token(&TokenKind::Comma);
            }
            self.expect(&TokenKind::RParen)?;
        } else {
            // Args without parens (space-separated until non-expr token)
            while matches!(self.current(), TokenKind::Number(_) | TokenKind::Str(_) | TokenKind::Ident(_) | TokenKind::LParen) {
                args.push(self.parse_atom()?);
            }
        }

        Ok(Expr::FunctionCall { name, args })
    }
}
