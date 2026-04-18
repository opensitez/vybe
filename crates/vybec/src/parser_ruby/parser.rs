use super::token::{Token, TokenKind};
use super::ast::*;
use super::lexer::Lexer;

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
        let mut body = Vec::new();
        self.skip_newlines();
        while !self.at_end() {
            let stmt = self.parse_statement()?;
            body.push(stmt);
            self.skip_newlines();
        }
        Ok(Program { body })
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    fn current(&self) -> &TokenKind {
        &self.tokens[self.pos].kind
    }

    fn current_line(&self) -> u32 {
        self.tokens[self.pos].line
    }

    fn at_end(&self) -> bool {
        self.pos >= self.tokens.len() || *self.current() == TokenKind::Eof
    }

    fn advance(&mut self) -> &TokenKind {
        let kind = &self.tokens[self.pos].kind;
        self.pos += 1;
        kind
    }

    fn peek(&self) -> &TokenKind {
        if self.pos + 1 < self.tokens.len() {
            &self.tokens[self.pos + 1].kind
        } else {
            &TokenKind::Eof
        }
    }

    fn expect(&mut self, expected: &TokenKind) -> Result<(), String> {
        if self.current() == expected {
            self.advance();
            Ok(())
        } else {
            Err(format!("Expected {:?}, got {:?} at line {}", expected, self.current(), self.current_line()))
        }
    }

    fn skip_newlines(&mut self) {
        while !self.at_end() && *self.current() == TokenKind::Newline {
            self.advance();
        }
    }

    fn at_statement_end(&self) -> bool {
        matches!(self.current(),
            TokenKind::Newline | TokenKind::Semicolon | TokenKind::Eof |
            TokenKind::End | TokenKind::Else | TokenKind::Elsif |
            TokenKind::Rescue | TokenKind::Ensure | TokenKind::When
        )
    }

    fn consume_terminator(&mut self) {
        if matches!(self.current(), TokenKind::Newline | TokenKind::Semicolon) {
            self.advance();
        }
    }

    fn match_token(&mut self, kind: &TokenKind) -> bool {
        if self.current() == kind {
            self.advance();
            true
        } else {
            false
        }
    }

    // ------------------------------------------------------------------
    // Statements
    // ------------------------------------------------------------------

    fn parse_statement(&mut self) -> Result<Statement, String> {
        self.skip_newlines();
        if self.at_end() { return Ok(Statement::Empty); }

        let stmt = match self.current().clone() {
            TokenKind::Puts => {
                self.advance();
                let args = self.parse_expr_list()?;
                // Check for modifier if/unless/while/until
                let stmt = Statement::Puts(args);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Print => {
                self.advance();
                let args = self.parse_expr_list()?;
                let stmt = Statement::Print(args);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::P => {
                self.advance();
                let args = self.parse_expr_list()?;
                let stmt = Statement::P(args);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Def => self.parse_def()?,
            TokenKind::Class => self.parse_class()?,
            TokenKind::Module => self.parse_module()?,
            TokenKind::If => self.parse_if()?,
            TokenKind::Unless => self.parse_unless()?,
            TokenKind::While => self.parse_while()?,
            TokenKind::Until => self.parse_until()?,
            TokenKind::For => self.parse_for()?,
            TokenKind::Case => self.parse_case()?,
            TokenKind::Begin => self.parse_begin()?,
            TokenKind::Return => {
                self.advance();
                let val = if self.at_statement_end() { None } else {
                    if matches!(self.current(), TokenKind::If | TokenKind::Unless) { None }
                    else { Some(self.parse_expression()?) }
                };
                let stmt = Statement::Return(val);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Break => {
                self.advance();
                let val = if self.at_statement_end() { None } else {
                    // Don't parse if/unless as expression — they are modifiers
                    if matches!(self.current(), TokenKind::If | TokenKind::Unless) {
                        None
                    } else {
                        Some(self.parse_expression()?)
                    }
                };
                let stmt = Statement::Break(val);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Next => {
                self.advance();
                let val = if self.at_statement_end() { None } else {
                    if matches!(self.current(), TokenKind::If | TokenKind::Unless) {
                        None
                    } else {
                        Some(self.parse_expression()?)
                    }
                };
                let stmt = Statement::Next(val);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Raise => {
                self.advance();
                let val = if self.at_statement_end() { None } else { Some(self.parse_expression()?) };
                Statement::Raise(val)
            }
            TokenKind::Require => {
                self.advance();
                if let TokenKind::Str(s) = self.current().clone() {
                    self.advance();
                    Statement::Require(s)
                } else {
                    let expr = self.parse_expression()?;
                    Statement::Expression(expr)
                }
            }
            TokenKind::Retry => {
                self.advance();
                Statement::Retry
            }
            TokenKind::Alias => {
                self.advance();
                let new_name = self.parse_method_name()?;
                let old_name = self.parse_method_name()?;
                Statement::Alias { new_name, old_name }
            }
            TokenKind::Private => {
                self.advance();
                Statement::AccessModifier(AccessLevel::Private)
            }
            TokenKind::Protected => {
                self.advance();
                Statement::AccessModifier(AccessLevel::Protected)
            }
            TokenKind::Public => {
                self.advance();
                Statement::AccessModifier(AccessLevel::Public)
            }
            TokenKind::Loop => {
                self.advance();
                if self.match_token(&TokenKind::Do) || self.match_token(&TokenKind::LBrace) {
                    let close = if self.tokens[self.pos - 1].kind == TokenKind::LBrace {
                        TokenKind::RBrace
                    } else {
                        TokenKind::End
                    };
                    self.skip_newlines();
                    let body = self.parse_body_until(&[close.clone()])?;
                    self.expect(&close)?;
                    Statement::Loop(body)
                } else {
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::End])?;
                    self.expect(&TokenKind::End)?;
                    Statement::Loop(body)
                }
            }
            TokenKind::Redo => {
                self.advance();
                let stmt = Statement::Redo;
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::AtExit => {
                self.advance();
                if *self.current() == TokenKind::LBrace {
                    self.advance();
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::RBrace])?;
                    self.expect(&TokenKind::RBrace)?;
                    Statement::AtExit(body)
                } else if *self.current() == TokenKind::Do {
                    self.advance();
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::End])?;
                    self.expect(&TokenKind::End)?;
                    Statement::AtExit(body)
                } else {
                    Statement::AtExit(Vec::new())
                }
            }
            TokenKind::Pp => {
                self.advance();
                let args = self.parse_expr_list()?;
                let stmt = Statement::P(args);
                self.parse_stmt_modifier(stmt)?
            }
            TokenKind::Catch => {
                self.advance();
                self.expect(&TokenKind::LParen)?;
                let tag = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                if self.match_token(&TokenKind::Do) || self.match_token(&TokenKind::LBrace) {
                    let close = if self.tokens[self.pos - 1].kind == TokenKind::LBrace {
                        TokenKind::RBrace
                    } else {
                        TokenKind::End
                    };
                    self.skip_newlines();
                    let body = self.parse_body_until(&[close.clone()])?;
                    self.expect(&close)?;
                    Statement::CatchThrow { tag, body }
                } else {
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::End])?;
                    self.expect(&TokenKind::End)?;
                    Statement::CatchThrow { tag, body }
                }
            }
            _ => {
                let expr = self.parse_expression()?;
                // Check for trailing if/unless/while/until modifiers
                let stmt = self.parse_modifier(expr)?;
                stmt
            }
        };

        self.consume_terminator();
        Ok(stmt)
    }

    /// Wrap any statement in a modifier if/unless/while/until
    fn parse_stmt_modifier(&mut self, stmt: Statement) -> Result<Statement, String> {
        match self.current() {
            TokenKind::If => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::If { test, body: vec![stmt], elsifs: Vec::new(), else_body: None })
            }
            TokenKind::Unless => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::Unless { test, body: vec![stmt], else_body: None })
            }
            TokenKind::While => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::While { test, body: vec![stmt] })
            }
            TokenKind::Until => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::Until { test, body: vec![stmt] })
            }
            _ => Ok(stmt),
        }
    }

    fn parse_modifier(&mut self, expr: Expression) -> Result<Statement, String> {
        // Check for comma (multiple assignment: a, b = 1, 2)
        if *self.current() == TokenKind::Comma {
            let mut targets = vec![expr.clone()];
            let mut splat_index = None;
            while self.match_token(&TokenKind::Comma) {
                if *self.current() == TokenKind::Star {
                    self.advance();
                    splat_index = Some(targets.len());
                }
                targets.push(self.parse_expression()?);
            }
            if self.match_token(&TokenKind::Eq) {
                let mut values = vec![self.parse_expression()?];
                while self.match_token(&TokenKind::Comma) {
                    values.push(self.parse_expression()?);
                }
                return Ok(Statement::MultiAssign { targets, splat_index, values });
            }
        }
        // Check for assignment
        if let Some(op) = self.try_assign_op() {
            // For chained assignment (a = b = c = 1), parse the rhs
            // which may itself be an assignment
            let value = self.parse_assign_value()?;
            let stmt = Statement::Assignment { target: expr, op, value };
            return self.parse_stmt_modifier(stmt);
        }

        // Check for trailing modifiers: expr if cond / expr unless cond / etc
        match self.current() {
            TokenKind::If => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::If {
                    test,
                    body: vec![Statement::Expression(expr)],
                    elsifs: Vec::new(),
                    else_body: None,
                })
            }
            TokenKind::Unless => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::Unless {
                    test,
                    body: vec![Statement::Expression(expr)],
                    else_body: None,
                })
            }
            TokenKind::While => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::While {
                    test,
                    body: vec![Statement::Expression(expr)],
                })
            }
            TokenKind::Until => {
                self.advance();
                let test = self.parse_expression()?;
                Ok(Statement::Until {
                    test,
                    body: vec![Statement::Expression(expr)],
                })
            }
            _ => Ok(Statement::Expression(expr)),
        }
    }

    /// Parse assignment value, handling chained assignments like a = b = c = 1
    fn parse_assign_value(&mut self) -> Result<Expression, String> {
        let expr = self.parse_expression()?;
        // Check if followed by = (chained assignment)
        if *self.current() == TokenKind::Eq {
            self.advance();
            let inner_value = self.parse_assign_value()?;
            // Return the inner value — the chained target was already an expression
            // We need to emit assignment to the intermediate. Wrap as ChainedAssign.
            return Ok(Expression::ChainedAssign {
                targets: vec![expr],
                value: Box::new(inner_value),
            });
        }
        Ok(expr)
    }

    fn try_assign_op(&mut self) -> Option<AssignOp> {
        let op = match self.current() {
            TokenKind::Eq => AssignOp::Assign,
            TokenKind::PlusEq => AssignOp::AddAssign,
            TokenKind::MinusEq => AssignOp::SubAssign,
            TokenKind::StarEq => AssignOp::MulAssign,
            TokenKind::SlashEq => AssignOp::DivAssign,
            TokenKind::PercentEq => AssignOp::ModAssign,
            TokenKind::AmpAmpEq => AssignOp::AndAssign,
            TokenKind::PipePipeEq => AssignOp::OrAssign,
            TokenKind::AmpEq => AssignOp::BitAndAssign,
            TokenKind::PipeEq => AssignOp::BitOrAssign,
            TokenKind::CaretEq => AssignOp::BitXorAssign,
            TokenKind::LtLtEq => AssignOp::ShlAssign,
            TokenKind::GtGtEq => AssignOp::ShrAssign,
            _ => return None,
        };
        self.advance();
        Some(op)
    }

    fn parse_body_until(&mut self, terminators: &[TokenKind]) -> Result<Vec<Statement>, String> {
        let mut stmts = Vec::new();
        self.skip_newlines();
        while !self.at_end() && !terminators.contains(self.current()) {
            let stmt = self.parse_statement()?;
            stmts.push(stmt);
            self.skip_newlines();
        }
        Ok(stmts)
    }

    // ------------------------------------------------------------------
    // Control flow
    // ------------------------------------------------------------------

    fn parse_if(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::If)?;
        let test = self.parse_expression()?;
        self.match_token(&TokenKind::Then);
        self.skip_newlines();

        let body = self.parse_body_until(&[TokenKind::Elsif, TokenKind::Else, TokenKind::End])?;

        let mut elsifs = Vec::new();
        while *self.current() == TokenKind::Elsif {
            self.advance();
            let etest = self.parse_expression()?;
            self.match_token(&TokenKind::Then);
            self.skip_newlines();
            let ebody = self.parse_body_until(&[TokenKind::Elsif, TokenKind::Else, TokenKind::End])?;
            elsifs.push(ElsIf { test: etest, body: ebody });
        }

        let else_body = if *self.current() == TokenKind::Else {
            self.advance();
            self.skip_newlines();
            Some(self.parse_body_until(&[TokenKind::End])?)
        } else {
            None
        };

        self.expect(&TokenKind::End)?;
        Ok(Statement::If { test, body, elsifs, else_body })
    }

    fn parse_unless(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Unless)?;
        let test = self.parse_expression()?;
        self.match_token(&TokenKind::Then);
        self.skip_newlines();

        let body = self.parse_body_until(&[TokenKind::Else, TokenKind::End])?;
        let else_body = if *self.current() == TokenKind::Else {
            self.advance();
            self.skip_newlines();
            Some(self.parse_body_until(&[TokenKind::End])?)
        } else {
            None
        };

        self.expect(&TokenKind::End)?;
        Ok(Statement::Unless { test, body, else_body })
    }

    fn parse_while(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::While)?;
        let test = self.parse_expression()?;
        self.match_token(&TokenKind::Do);
        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;
        Ok(Statement::While { test, body })
    }

    fn parse_until(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Until)?;
        let test = self.parse_expression()?;
        self.match_token(&TokenKind::Do);
        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;
        Ok(Statement::Until { test, body })
    }

    fn parse_for(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::For)?;
        let var = match self.current().clone() {
            TokenKind::Identifier(name) => { self.advance(); name }
            _ => return Err(format!("Expected identifier in for loop at line {}", self.current_line())),
        };
        self.expect(&TokenKind::In)?;
        let iterable = self.parse_expression()?;
        self.match_token(&TokenKind::Do);
        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;
        Ok(Statement::For { var, iterable, body })
    }

    fn parse_case(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Case)?;
        let subject = if self.at_statement_end() { None } else { Some(self.parse_expression()?) };
        self.skip_newlines();

        let mut whens = Vec::new();
        while *self.current() == TokenKind::When {
            self.advance();
            let mut conditions = vec![self.parse_expression()?];
            while self.match_token(&TokenKind::Comma) {
                conditions.push(self.parse_expression()?);
            }
            self.match_token(&TokenKind::Then);
            self.skip_newlines();
            let body = self.parse_body_until(&[TokenKind::When, TokenKind::Else, TokenKind::End])?;
            whens.push(WhenClause { conditions, body });
        }

        let else_body = if *self.current() == TokenKind::Else {
            self.advance();
            self.skip_newlines();
            Some(self.parse_body_until(&[TokenKind::End])?)
        } else {
            None
        };

        self.expect(&TokenKind::End)?;
        Ok(Statement::Case { subject, whens, else_body })
    }

    fn parse_begin(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Begin)?;
        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::Rescue, TokenKind::Else, TokenKind::Ensure, TokenKind::End])?;

        let mut rescues = Vec::new();
        while *self.current() == TokenKind::Rescue {
            self.advance();
            let mut types = Vec::new();
            let mut var = None;

            // rescue ExceptionType[, Type2] [=> var]
            if let TokenKind::Constant(name) = self.current().clone() {
                self.advance();
                types.push(name);
                while self.match_token(&TokenKind::Comma) {
                    if let TokenKind::Constant(name) = self.current().clone() {
                        self.advance();
                        types.push(name);
                    }
                }
            }
            if self.match_token(&TokenKind::FatArrow) {
                if let TokenKind::Identifier(name) = self.current().clone() {
                    self.advance();
                    var = Some(name);
                }
            }
            self.skip_newlines();
            let rbody = self.parse_body_until(&[TokenKind::Rescue, TokenKind::Else, TokenKind::Ensure, TokenKind::End])?;
            rescues.push(RescueClause { types, var, body: rbody });
        }

        let else_body = if *self.current() == TokenKind::Else {
            self.advance();
            self.skip_newlines();
            Some(self.parse_body_until(&[TokenKind::Ensure, TokenKind::End])?)
        } else {
            None
        };

        let ensure = if *self.current() == TokenKind::Ensure {
            self.advance();
            self.skip_newlines();
            Some(self.parse_body_until(&[TokenKind::End])?)
        } else {
            None
        };

        self.expect(&TokenKind::End)?;
        Ok(Statement::Begin { body, rescues, else_body, ensure })
    }

    // ------------------------------------------------------------------
    // Definitions
    // ------------------------------------------------------------------

    fn parse_def(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Def)?;

        // Check for self.method_name (class method)
        let (name, is_self) = if *self.current() == TokenKind::Self_ {
            self.advance();
            self.expect(&TokenKind::Dot)?;
            let name = self.parse_method_name()?;
            (name, true)
        } else {
            let name = self.parse_method_name()?;
            (name, false)
        };

        let params = if self.match_token(&TokenKind::LParen) {
            let params = self.parse_params()?;
            self.expect(&TokenKind::RParen)?;
            params
        } else if !self.at_statement_end() && !matches!(self.current(), TokenKind::Newline | TokenKind::Semicolon | TokenKind::End) {
            // Ruby allows params without parens: def foo x, y
            self.parse_params()?
        } else {
            Vec::new()
        };

        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;
        Ok(Statement::MethodDef(MethodDecl { name, params, body, is_self }))
    }

    fn parse_method_name(&mut self) -> Result<String, String> {
        match self.current().clone() {
            TokenKind::Identifier(name) => { self.advance(); Ok(name) }
            TokenKind::Constant(name) => { self.advance(); Ok(name) }
            // Operator methods
            TokenKind::Plus => { self.advance(); Ok("+".to_string()) }
            TokenKind::Minus => { self.advance(); Ok("-".to_string()) }
            TokenKind::Star => { self.advance(); Ok("*".to_string()) }
            TokenKind::Slash => { self.advance(); Ok("/".to_string()) }
            TokenKind::EqEq => { self.advance(); Ok("==".to_string()) }
            TokenKind::LtEq => { self.advance(); Ok("<=".to_string()) }
            TokenKind::GtEq => { self.advance(); Ok(">=".to_string()) }
            TokenKind::Spaceship => { self.advance(); Ok("<=>".to_string()) }
            TokenKind::LBracket => {
                self.advance();
                if self.match_token(&TokenKind::RBracket) {
                    if self.match_token(&TokenKind::Eq) {
                        Ok("[]=".to_string())
                    } else {
                        Ok("[]".to_string())
                    }
                } else {
                    Err("Expected ] after [".to_string())
                }
            }
            // Keywords that can also be method names after .
            TokenKind::Class => { self.advance(); Ok("class".to_string()) }
            TokenKind::Freeze => { self.advance(); Ok("freeze".to_string()) }
            TokenKind::Frozen => { self.advance(); Ok("frozen?".to_string()) }
            TokenKind::Include => { self.advance(); Ok("include?".to_string()) }
            TokenKind::Extend => { self.advance(); Ok("extend".to_string()) }
            TokenKind::Nil => {
                self.advance();
                // .nil? — consume trailing ? if present
                self.match_token(&TokenKind::Question);
                Ok("nil?".to_string())
            }
            TokenKind::Defined => { self.advance(); Ok("defined?".to_string()) }
            _ => Err(format!("Expected method name, got {:?} at line {}", self.current(), self.current_line())),
        }
    }

    fn parse_params(&mut self) -> Result<Vec<Param>, String> {
        let mut params = Vec::new();
        if matches!(self.current(), TokenKind::RParen | TokenKind::Newline | TokenKind::Semicolon | TokenKind::Pipe | TokenKind::Eof) {
            return Ok(params);
        }
        loop {
            let mut splat = false;
            let mut double_splat = false;
            let mut block = false;

            if self.match_token(&TokenKind::Star) {
                splat = true;
            } else if self.match_token(&TokenKind::StarStar) {
                double_splat = true;
            } else if self.match_token(&TokenKind::Amp) {
                block = true;
            }

            let name = match self.current().clone() {
                TokenKind::Identifier(name) => { self.advance(); name }
                _ => return Err(format!("Expected parameter name at line {}", self.current_line())),
            };

            // Check for keyword argument: name:
            let keyword = if self.match_token(&TokenKind::Colon) {
                true
            } else {
                false
            };

            let default = if keyword && !matches!(self.current(), TokenKind::Comma | TokenKind::RParen | TokenKind::Pipe | TokenKind::Newline | TokenKind::Eof) {
                // keyword arg with default: name: default_val
                Some(self.parse_expression()?)
            } else if !keyword && self.match_token(&TokenKind::Eq) {
                Some(self.parse_expression()?)
            } else {
                None
            };

            params.push(Param { name, default, splat, double_splat, block, keyword });

            if !self.match_token(&TokenKind::Comma) { break; }
        }
        Ok(params)
    }

    fn parse_class(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Class)?;
        let name = match self.current().clone() {
            TokenKind::Constant(name) => { self.advance(); name }
            _ => return Err(format!("Expected class name at line {}", self.current_line())),
        };

        let parent = if self.match_token(&TokenKind::Lt) {
            match self.current().clone() {
                TokenKind::Constant(name) => { self.advance(); Some(name) }
                _ => return Err(format!("Expected parent class name at line {}", self.current_line())),
            }
        } else {
            None
        };

        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;

        Ok(Statement::ClassDef(ClassDecl { name, parent, body }))
    }

    fn parse_module(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::Module)?;
        let name = match self.current().clone() {
            TokenKind::Constant(name) => { self.advance(); name }
            _ => return Err(format!("Expected module name at line {}", self.current_line())),
        };
        self.skip_newlines();
        let body = self.parse_body_until(&[TokenKind::End])?;
        self.expect(&TokenKind::End)?;
        Ok(Statement::ModuleDef(ModuleDecl { name, body }))
    }

    // ------------------------------------------------------------------
    // Expressions — Pratt parser
    // ------------------------------------------------------------------

    fn parse_expression(&mut self) -> Result<Expression, String> {
        let expr = self.parse_ternary()?;
        // Inline rescue: expr rescue default_value
        if *self.current() == TokenKind::Rescue {
            self.advance();
            let rescue_val = self.parse_ternary()?;
            return Ok(Expression::InlineRescue {
                expr: Box::new(expr),
                rescue_val: Box::new(rescue_val),
            });
        }
        Ok(expr)
    }

    fn parse_ternary(&mut self) -> Result<Expression, String> {
        let expr = self.parse_or()?;
        if self.match_token(&TokenKind::Question) {
            let consequent = self.parse_expression()?;
            self.expect(&TokenKind::Colon)?;
            let alternate = self.parse_expression()?;
            Ok(Expression::Ternary {
                test: Box::new(expr),
                consequent: Box::new(consequent),
                alternate: Box::new(alternate),
            })
        } else {
            Ok(expr)
        }
    }

    fn parse_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_and()?;
        while matches!(self.current(), TokenKind::PipePipe | TokenKind::Or) {
            self.advance();
            let right = self.parse_and()?;
            left = Expression::Binary {
                op: BinaryOp::Or,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_not()?;
        while matches!(self.current(), TokenKind::AmpAmp | TokenKind::And) {
            self.advance();
            let right = self.parse_not()?;
            left = Expression::Binary {
                op: BinaryOp::And,
                left: Box::new(left),
                right: Box::new(right),
            };
        }
        Ok(left)
    }

    fn parse_not(&mut self) -> Result<Expression, String> {
        if matches!(self.current(), TokenKind::Bang | TokenKind::Not) {
            self.advance();
            let expr = self.parse_not()?;
            return Ok(Expression::Unary { op: UnaryOp::Not, expr: Box::new(expr) });
        }
        self.parse_comparison()
    }

    fn parse_comparison(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitor()?;
        loop {
            let op = match self.current() {
                TokenKind::EqEq => BinaryOp::Eq,
                TokenKind::BangEq => BinaryOp::Ne,
                TokenKind::Lt => BinaryOp::Lt,
                TokenKind::Gt => BinaryOp::Gt,
                TokenKind::LtEq => BinaryOp::Le,
                TokenKind::GtEq => BinaryOp::Ge,
                TokenKind::Spaceship => BinaryOp::Spaceship,
                TokenKind::EqEqEq => BinaryOp::Eq, // case equality
                TokenKind::EqTilde => BinaryOp::Eq, // regex match (treated as equality for now)
                _ => break,
            };
            self.advance();
            let right = self.parse_bitor()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitxor()?;
        while *self.current() == TokenKind::Pipe {
            self.advance();
            let right = self.parse_bitxor()?;
            left = Expression::Binary { op: BinaryOp::BitOr, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitxor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitand()?;
        while *self.current() == TokenKind::Caret {
            self.advance();
            let right = self.parse_bitand()?;
            left = Expression::Binary { op: BinaryOp::BitXor, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitand(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_shift()?;
        while *self.current() == TokenKind::Amp {
            self.advance();
            let right = self.parse_shift()?;
            left = Expression::Binary { op: BinaryOp::BitAnd, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_shift(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_range()?;
        loop {
            let op = match self.current() {
                TokenKind::LtLt => BinaryOp::Shl,
                TokenKind::GtGt => BinaryOp::Shr,
                _ => break,
            };
            self.advance();
            let right = self.parse_range()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_range(&mut self) -> Result<Expression, String> {
        let left = self.parse_additive()?;
        if matches!(self.current(), TokenKind::DotDot | TokenKind::DotDotDot) {
            let exclusive = *self.current() == TokenKind::DotDotDot;
            self.advance();
            let right = self.parse_additive()?;
            Ok(Expression::Range {
                start: Box::new(left),
                end: Box::new(right),
                exclusive,
            })
        } else {
            Ok(left)
        }
    }

    fn parse_additive(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_multiplicative()?;
        loop {
            let op = match self.current() {
                TokenKind::Plus => BinaryOp::Add,
                TokenKind::Minus => BinaryOp::Sub,
                _ => break,
            };
            self.advance();
            let right = self.parse_multiplicative()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_power()?;
        loop {
            let op = match self.current() {
                TokenKind::Star => BinaryOp::Mul,
                TokenKind::Slash => BinaryOp::Div,
                TokenKind::Percent => BinaryOp::Mod,
                _ => break,
            };
            self.advance();
            let right = self.parse_power()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<Expression, String> {
        let left = self.parse_unary()?;
        if *self.current() == TokenKind::StarStar {
            self.advance();
            let right = self.parse_power()?; // right-associative
            Ok(Expression::Binary { op: BinaryOp::Pow, left: Box::new(left), right: Box::new(right) })
        } else {
            Ok(left)
        }
    }

    fn parse_unary(&mut self) -> Result<Expression, String> {
        match self.current().clone() {
            TokenKind::Minus => {
                self.advance();
                let expr = self.parse_postfix()?;
                Ok(Expression::Unary { op: UnaryOp::Neg, expr: Box::new(expr) })
            }
            TokenKind::Plus => {
                self.advance();
                let expr = self.parse_postfix()?;
                Ok(Expression::Unary { op: UnaryOp::Pos, expr: Box::new(expr) })
            }
            TokenKind::Tilde => {
                self.advance();
                let expr = self.parse_postfix()?;
                Ok(Expression::Unary { op: UnaryOp::BitNot, expr: Box::new(expr) })
            }
            TokenKind::Star => {
                self.advance();
                let expr = self.parse_postfix()?;
                Ok(Expression::Splat(Box::new(expr)))
            }
            TokenKind::Amp => {
                self.advance();
                // &:method_name → symbol-to-proc
                if let TokenKind::Symbol(name) = self.current().clone() {
                    self.advance();
                    return Ok(Expression::SymbolProc(name));
                }
                // &block_var
                let expr = self.parse_postfix()?;
                Ok(expr)
            }
            _ => self.parse_postfix(),
        }
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            match self.current() {
                TokenKind::Dot => {
                    self.advance();
                    let method = self.parse_method_name()?;
                    let (args, block) = self.parse_call_args_and_block()?;
                    expr = Expression::MethodCall {
                        receiver: Some(Box::new(expr)),
                        method,
                        args,
                        block,
                    };
                }
                TokenKind::AmpDot => {
                    self.advance();
                    let method = self.parse_method_name()?;
                    let (args, block) = self.parse_call_args_and_block()?;
                    expr = Expression::SafeNav {
                        receiver: Box::new(expr),
                        method,
                        args,
                        block,
                    };
                }
                TokenKind::ColonColon => {
                    self.advance();
                    match self.current().clone() {
                        TokenKind::Constant(name) | TokenKind::Identifier(name) => {
                            self.advance();
                            // Check if it's a method call
                            if *self.current() == TokenKind::LParen {
                                let (args, block) = self.parse_call_args_and_block()?;
                                expr = Expression::MethodCall {
                                    receiver: Some(Box::new(expr)),
                                    method: name,
                                    args,
                                    block,
                                };
                            } else {
                                expr = Expression::ScopeResolution {
                                    left: Box::new(expr),
                                    name,
                                };
                            }
                        }
                        _ => return Err(format!("Expected name after :: at line {}", self.current_line())),
                    }
                }
                TokenKind::LBracket => {
                    self.advance();
                    let index = self.parse_expression()?;
                    self.expect(&TokenKind::RBracket)?;
                    expr = Expression::IndexAccess {
                        object: Box::new(expr),
                        index: Box::new(index),
                    };
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_call_args_and_block(&mut self) -> Result<(Vec<Expression>, Option<Box<BlockArg>>), String> {
        let mut args = Vec::new();
        if self.match_token(&TokenKind::LParen) {
            if *self.current() != TokenKind::RParen {
                loop {
                    // Check for keyword arg: ident: value → becomes Hash entry
                    if let TokenKind::Identifier(_) = self.current().clone() {
                        if self.peek() == &TokenKind::Colon {
                            // keyword arg — build inline hash
                            let mut pairs = Vec::new();
                            loop {
                                if let TokenKind::Identifier(name) = self.current().clone() {
                                    if self.peek() == &TokenKind::Colon {
                                        self.advance(); // ident
                                        self.advance(); // colon
                                        let val = self.parse_expression()?;
                                        pairs.push((Expression::Symbol(name), val));
                                    } else {
                                        break;
                                    }
                                } else {
                                    break;
                                }
                                if !self.match_token(&TokenKind::Comma) { break; }
                            }
                            if !pairs.is_empty() {
                                args.push(Expression::Hash(pairs));
                                break;
                            }
                        }
                    }
                    // Check for &:symbol
                    args.push(self.parse_expression()?);
                    if !self.match_token(&TokenKind::Comma) { break; }
                }
            }
            self.expect(&TokenKind::RParen)?;
        }
        // Check for block { |params| body } or do |params| body end
        let block = self.try_parse_block()?;
        Ok((args, block))
    }

    fn try_parse_block(&mut self) -> Result<Option<Box<BlockArg>>, String> {
        if *self.current() == TokenKind::LBrace {
            self.advance();
            let params = self.parse_block_params()?;
            self.skip_newlines();
            let body = self.parse_body_until(&[TokenKind::RBrace])?;
            self.expect(&TokenKind::RBrace)?;
            Ok(Some(Box::new(BlockArg { params, body })))
        } else if *self.current() == TokenKind::Do {
            self.advance();
            let params = self.parse_block_params()?;
            self.skip_newlines();
            let body = self.parse_body_until(&[TokenKind::End])?;
            self.expect(&TokenKind::End)?;
            Ok(Some(Box::new(BlockArg { params, body })))
        } else {
            Ok(None)
        }
    }

    fn parse_block_params(&mut self) -> Result<Vec<String>, String> {
        if !self.match_token(&TokenKind::Pipe) { return Ok(Vec::new()); }
        let mut params = Vec::new();
        while *self.current() != TokenKind::Pipe {
            match self.current().clone() {
                TokenKind::Identifier(name) => { self.advance(); params.push(name); }
                _ => return Err(format!("Expected block param at line {}", self.current_line())),
            }
            self.match_token(&TokenKind::Comma);
        }
        self.expect(&TokenKind::Pipe)?;
        Ok(params)
    }

    fn parse_expr_list(&mut self) -> Result<Vec<Expression>, String> {
        let mut exprs = Vec::new();
        if self.at_statement_end() { return Ok(exprs); }
        exprs.push(self.parse_expression()?);
        while self.match_token(&TokenKind::Comma) {
            exprs.push(self.parse_expression()?);
        }
        Ok(exprs)
    }

    // ------------------------------------------------------------------
    // Primary expressions
    // ------------------------------------------------------------------

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.current().clone() {
            TokenKind::Number(n) => {
                self.advance();
                Ok(Expression::Number(n))
            }
            // Handle %w[] / %i[] / %q{} string that was encoded by lexer (before general Str arm)
            TokenKind::Str(ref s) if s.starts_with('\x03') => {
                let s = s.clone();
                self.advance();
                if s.starts_with("\x03w") {
                    let content = &s[2..];
                    let words: Vec<Expression> = if content.is_empty() {
                        Vec::new()
                    } else {
                        content.split('\x04').map(|w| Expression::Str(w.to_string())).collect()
                    };
                    Ok(Expression::Array(words))
                } else if s.starts_with("\x03i") {
                    let content = &s[2..];
                    let syms: Vec<Expression> = if content.is_empty() {
                        Vec::new()
                    } else {
                        content.split('\x04').map(|w| Expression::Symbol(w.to_string())).collect()
                    };
                    Ok(Expression::Array(syms))
                } else {
                    Ok(Expression::Str(s[2..].to_string()))
                }
            }

            TokenKind::Str(s) => {
                self.advance();
                // Check for string interpolation
                if s.starts_with('\x01') {
                    self.parse_interpolated_string(&s)
                } else {
                    Ok(Expression::Str(s))
                }
            }
            TokenKind::Symbol(s) => {
                self.advance();
                Ok(Expression::Symbol(s))
            }
            TokenKind::True => { self.advance(); Ok(Expression::Bool(true)) }
            TokenKind::False => { self.advance(); Ok(Expression::Bool(false)) }
            TokenKind::Nil => { self.advance(); Ok(Expression::Nil) }
            TokenKind::Self_ => { self.advance(); Ok(Expression::SelfExpr) }

            // Magic constants (must be before general Identifier arm)
            TokenKind::Identifier(ref name) if name == "__FILE__" => {
                self.advance();
                Ok(Expression::MagicConstant(MagicConst::File))
            }
            TokenKind::Identifier(ref name) if name == "__LINE__" => {
                let line = self.current_line();
                self.advance();
                Ok(Expression::Number(line as f64))
            }
            TokenKind::Identifier(ref name) if name == "__dir__" => {
                self.advance();
                Ok(Expression::MagicConstant(MagicConst::Dir))
            }
            TokenKind::Identifier(ref name) if name == "__method__" => {
                self.advance();
                Ok(Expression::MagicConstant(MagicConst::Method))
            }

            TokenKind::Identifier(name) => {
                self.advance();
                // Check if it's a method call with parens
                if *self.current() == TokenKind::LParen {
                    let (args, block) = self.parse_call_args_and_block()?;
                    Ok(Expression::MethodCall {
                        receiver: None,
                        method: name,
                        args,
                        block,
                    })
                } else {
                    Ok(Expression::Identifier(name))
                }
            }

            TokenKind::InstanceVar(name) => { self.advance(); Ok(Expression::InstanceVar(name)) }
            TokenKind::ClassVar(name) => { self.advance(); Ok(Expression::ClassVar(name)) }
            TokenKind::GlobalVar(name) => { self.advance(); Ok(Expression::GlobalVar(name)) }

            TokenKind::Constant(name) => {
                self.advance();
                // Check for method call: ClassName.new() or ClassName::method()
                if *self.current() == TokenKind::Dot || *self.current() == TokenKind::ColonColon {
                    // Return as ConstantRef, postfix will handle the rest
                    Ok(Expression::ConstantRef(name))
                } else if *self.current() == TokenKind::LParen {
                    // ConstantName(args) — function-style call
                    let (args, block) = self.parse_call_args_and_block()?;
                    Ok(Expression::MethodCall {
                        receiver: None,
                        method: name,
                        args,
                        block,
                    })
                } else {
                    Ok(Expression::ConstantRef(name))
                }
            }

            // Array literal
            TokenKind::LBracket => {
                self.advance();
                let mut elements = Vec::new();
                self.skip_newlines();
                if *self.current() != TokenKind::RBracket {
                    loop {
                        self.skip_newlines();
                        elements.push(self.parse_expression()?);
                        self.skip_newlines();
                        if !self.match_token(&TokenKind::Comma) { break; }
                    }
                }
                self.skip_newlines();
                self.expect(&TokenKind::RBracket)?;
                Ok(Expression::Array(elements))
            }

            // Hash literal
            TokenKind::LBrace => {
                self.advance();
                let mut pairs = Vec::new();
                self.skip_newlines();
                if *self.current() != TokenKind::RBrace {
                    loop {
                        self.skip_newlines();
                        // key: value (symbol shorthand) or key => value
                        let key = if let TokenKind::Identifier(name) = self.current().clone() {
                            if self.peek() == &TokenKind::Colon {
                                self.advance(); // identifier
                                self.advance(); // colon
                                Expression::Symbol(name)
                            } else {
                                self.parse_expression()?
                            }
                        } else {
                            self.parse_expression()?
                        };

                        if *self.current() == TokenKind::FatArrow {
                            self.advance();
                        }
                        // value may already be parsed (for symbol shorthand the key consumed the colon)
                        let value = self.parse_expression()?;
                        pairs.push((key, value));
                        self.skip_newlines();
                        if !self.match_token(&TokenKind::Comma) { break; }
                    }
                }
                self.skip_newlines();
                self.expect(&TokenKind::RBrace)?;
                Ok(Expression::Hash(pairs))
            }

            // Parenthesized expression
            TokenKind::LParen => {
                self.advance();
                let expr = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                Ok(expr)
            }

            // Lambda: -> (params) { body }
            TokenKind::Arrow => {
                self.advance();
                let params = if self.match_token(&TokenKind::LParen) {
                    let p = self.parse_params()?;
                    self.expect(&TokenKind::RParen)?;
                    p
                } else {
                    Vec::new()
                };
                self.expect(&TokenKind::LBrace)?;
                self.skip_newlines();
                let body = self.parse_body_until(&[TokenKind::RBrace])?;
                self.expect(&TokenKind::RBrace)?;
                Ok(Expression::Lambda { params, body })
            }

            // lambda keyword
            TokenKind::Lambda => {
                self.advance();
                if *self.current() == TokenKind::LBrace {
                    self.advance();
                    let params = self.parse_block_params()?;
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::RBrace])?;
                    self.expect(&TokenKind::RBrace)?;
                    Ok(Expression::Lambda {
                        params: params.iter().map(|n| Param { name: n.clone(), default: None, splat: false, double_splat: false, block: false, keyword: false }).collect(),
                        body,
                    })
                } else {
                    Ok(Expression::Lambda { params: Vec::new(), body: Vec::new() })
                }
            }

            // yield
            TokenKind::Yield => {
                self.advance();
                let args = if self.match_token(&TokenKind::LParen) {
                    let mut a = Vec::new();
                    if *self.current() != TokenKind::RParen {
                        loop {
                            a.push(self.parse_expression()?);
                            if !self.match_token(&TokenKind::Comma) { break; }
                        }
                    }
                    self.expect(&TokenKind::RParen)?;
                    a
                } else if !self.at_statement_end() {
                    self.parse_expr_list()?
                } else {
                    Vec::new()
                };
                Ok(Expression::Yield(args))
            }

            TokenKind::Block_given => {
                self.advance();
                Ok(Expression::BlockGiven)
            }

            // super
            TokenKind::Super => {
                self.advance();
                let args = if self.match_token(&TokenKind::LParen) {
                    let mut a = Vec::new();
                    if *self.current() != TokenKind::RParen {
                        loop {
                            a.push(self.parse_expression()?);
                            if !self.match_token(&TokenKind::Comma) { break; }
                        }
                    }
                    self.expect(&TokenKind::RParen)?;
                    a
                } else {
                    Vec::new()
                };
                Ok(Expression::Super(args))
            }

            // attr_reader / attr_writer / attr_accessor
            TokenKind::Attr_reader | TokenKind::Attr_writer | TokenKind::Attr_accessor => {
                let kind = match self.current() {
                    TokenKind::Attr_reader => AttrKind::Reader,
                    TokenKind::Attr_writer => AttrKind::Writer,
                    TokenKind::Attr_accessor => AttrKind::Accessor,
                    _ => unreachable!(),
                };
                self.advance();
                let mut names = Vec::new();
                loop {
                    match self.current().clone() {
                        TokenKind::Symbol(name) => { self.advance(); names.push(name); }
                        TokenKind::Colon => {
                            self.advance();
                            if let TokenKind::Identifier(name) = self.current().clone() {
                                self.advance();
                                names.push(name);
                            }
                        }
                        _ => break,
                    }
                    if !self.match_token(&TokenKind::Comma) { break; }
                }
                Ok(Expression::AttrDecl { kind, names })
            }

            // include / extend
            TokenKind::Include => {
                self.advance();
                match self.current().clone() {
                    TokenKind::Constant(name) => { self.advance(); Ok(Expression::Include(name)) }
                    _ => Err(format!("Expected module name after include at line {}", self.current_line())),
                }
            }
            TokenKind::Extend => {
                self.advance();
                match self.current().clone() {
                    TokenKind::Constant(name) => { self.advance(); Ok(Expression::Extend(name)) }
                    _ => Err(format!("Expected module name after extend at line {}", self.current_line())),
                }
            }

            // Regex literal
            TokenKind::Regex(pattern) => {
                self.advance();
                Ok(Expression::Regex(pattern))
            }

            // defined?(expr)
            TokenKind::Defined => {
                self.advance();
                if self.match_token(&TokenKind::LParen) {
                    let expr = self.parse_expression()?;
                    self.expect(&TokenKind::RParen)?;
                    Ok(Expression::Defined(Box::new(expr)))
                } else {
                    let expr = self.parse_expression()?;
                    Ok(Expression::Defined(Box::new(expr)))
                }
            }

            // proc { |x| body }
            TokenKind::Proc => {
                self.advance();
                if *self.current() == TokenKind::LBrace {
                    self.advance();
                    let params = self.parse_block_params()?;
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::RBrace])?;
                    self.expect(&TokenKind::RBrace)?;
                    Ok(Expression::ProcLiteral { params, body })
                } else if *self.current() == TokenKind::Do {
                    self.advance();
                    let params = self.parse_block_params()?;
                    self.skip_newlines();
                    let body = self.parse_body_until(&[TokenKind::End])?;
                    self.expect(&TokenKind::End)?;
                    Ok(Expression::ProcLiteral { params, body })
                } else {
                    // Proc.new { }
                    Ok(Expression::ProcLiteral { params: Vec::new(), body: Vec::new() })
                }
            }

            // if/unless/begin/case as expression
            TokenKind::If => {
                self.advance();
                let test = self.parse_expression()?;
                self.match_token(&TokenKind::Then);
                self.skip_newlines();
                let body = self.parse_body_until(&[TokenKind::Elsif, TokenKind::Else, TokenKind::End])?;
                let mut elsifs = Vec::new();
                while *self.current() == TokenKind::Elsif {
                    self.advance();
                    let etest = self.parse_expression()?;
                    self.match_token(&TokenKind::Then);
                    self.skip_newlines();
                    let ebody = self.parse_body_until(&[TokenKind::Elsif, TokenKind::Else, TokenKind::End])?;
                    elsifs.push(ElsIf { test: etest, body: ebody });
                }
                let else_body = if *self.current() == TokenKind::Else {
                    self.advance();
                    self.skip_newlines();
                    Some(self.parse_body_until(&[TokenKind::End])?)
                } else { None };
                self.expect(&TokenKind::End)?;
                Ok(Expression::IfExpr { test: Box::new(test), body, elsifs, else_body })
            }
            TokenKind::Unless => {
                self.advance();
                let test = self.parse_expression()?;
                self.match_token(&TokenKind::Then);
                self.skip_newlines();
                let body = self.parse_body_until(&[TokenKind::Else, TokenKind::End])?;
                let else_body = if *self.current() == TokenKind::Else {
                    self.advance(); self.skip_newlines();
                    Some(self.parse_body_until(&[TokenKind::End])?)
                } else { None };
                self.expect(&TokenKind::End)?;
                Ok(Expression::UnlessExpr { test: Box::new(test), body, else_body })
            }
            TokenKind::Begin => {
                self.advance();
                self.skip_newlines();
                let body = self.parse_body_until(&[TokenKind::Rescue, TokenKind::Else, TokenKind::Ensure, TokenKind::End])?;
                let mut rescues = Vec::new();
                while *self.current() == TokenKind::Rescue {
                    self.advance();
                    let mut types = Vec::new();
                    let mut var = None;
                    if let TokenKind::Constant(name) = self.current().clone() {
                        self.advance(); types.push(name);
                        while self.match_token(&TokenKind::Comma) {
                            if let TokenKind::Constant(name) = self.current().clone() { self.advance(); types.push(name); }
                        }
                    }
                    if self.match_token(&TokenKind::FatArrow) {
                        if let TokenKind::Identifier(name) = self.current().clone() { self.advance(); var = Some(name); }
                    }
                    self.skip_newlines();
                    let rbody = self.parse_body_until(&[TokenKind::Rescue, TokenKind::Else, TokenKind::Ensure, TokenKind::End])?;
                    rescues.push(RescueClause { types, var, body: rbody });
                }
                let else_body = if *self.current() == TokenKind::Else {
                    self.advance(); self.skip_newlines();
                    Some(self.parse_body_until(&[TokenKind::Ensure, TokenKind::End])?)
                } else { None };
                let ensure = if *self.current() == TokenKind::Ensure {
                    self.advance(); self.skip_newlines();
                    Some(self.parse_body_until(&[TokenKind::End])?)
                } else { None };
                self.expect(&TokenKind::End)?;
                Ok(Expression::BeginExpr { body, rescues, else_body, ensure })
            }

            // Backtick shell command
            TokenKind::Backtick(cmd) => {
                self.advance();
                Ok(Expression::Backtick(cmd))
            }

            // pp (pretty print)
            TokenKind::Pp => {
                self.advance();
                let args = self.parse_expr_list()?;
                Ok(Expression::MethodCall {
                    receiver: None,
                    method: "pp".to_string(),
                    args,
                    block: None,
                })
            }

            // sprintf / format
            TokenKind::Sprintf | TokenKind::Format => {
                self.advance();
                if self.match_token(&TokenKind::LParen) {
                    let mut args = Vec::new();
                    if *self.current() != TokenKind::RParen {
                        loop {
                            args.push(self.parse_expression()?);
                            if !self.match_token(&TokenKind::Comma) { break; }
                        }
                    }
                    self.expect(&TokenKind::RParen)?;
                    Ok(Expression::MethodCall {
                        receiver: None,
                        method: "sprintf".to_string(),
                        args,
                        block: None,
                    })
                } else {
                    let args = self.parse_expr_list()?;
                    Ok(Expression::MethodCall {
                        receiver: None,
                        method: "sprintf".to_string(),
                        args,
                        block: None,
                    })
                }
            }

            // throw :tag [, value]
            TokenKind::Throw => {
                self.advance();
                let tag = self.parse_expression()?;
                let value = if self.match_token(&TokenKind::Comma) {
                    Some(Box::new(self.parse_expression()?))
                } else {
                    None
                };
                Ok(Expression::Throw { tag: Box::new(tag), value })
            }

            _ => Err(format!("Unexpected token {:?} at line {}", self.current(), self.current_line())),
        }
    }

    fn parse_interpolated_string(&self, encoded: &str) -> Result<Expression, String> {
        let mut parts = Vec::new();
        let mut current = String::new();
        let chars: Vec<char> = encoded.chars().collect();
        let mut i = 0;
        // Skip initial \x01
        if !chars.is_empty() && chars[0] == '\x01' { i = 1; }
        while i < chars.len() {
            if chars[i] == '\x02' {
                if !current.is_empty() {
                    parts.push(InterpolPart::Lit(std::mem::take(&mut current)));
                }
                i += 1;
                let mut expr_str = String::new();
                while i < chars.len() && chars[i] != '\x01' {
                    expr_str.push(chars[i]);
                    i += 1;
                }
                if i < chars.len() { i += 1; } // skip \x01
                // Parse the expression
                let mut sub_parser = Parser::new(&expr_str)?;
                let expr = sub_parser.parse_expression()?;
                parts.push(InterpolPart::Expr(expr));
            } else {
                current.push(chars[i]);
                i += 1;
            }
        }
        if !current.is_empty() {
            parts.push(InterpolPart::Lit(current));
        }
        Ok(Expression::Interpolated(parts))
    }
}
