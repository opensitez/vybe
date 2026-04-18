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
        Ok(Self { tokens, pos: 0 })
    }

    pub fn parse_module(&mut self) -> Result<Module, String> {
        // Skip leading newlines
        while self.check(&TokenKind::Newline) {
            self.advance();
        }
        let mut body = Vec::new();
        while !self.is_at_end() {
            if self.check(&TokenKind::Newline) {
                self.advance();
                continue;
            }
            body.push(self.parse_statement()?);
        }
        Ok(Module { body })
    }

    // ── Statement parsing ────────────────────────────────────────────

    fn parse_statement(&mut self) -> Result<Statement, String> {
        match self.current_kind() {
            TokenKind::Def => self.parse_function_def(false, Vec::new()),
            TokenKind::Async => self.parse_async_stmt(),
            TokenKind::Class => self.parse_class_def(Vec::new()),
            TokenKind::If => self.parse_if_stmt(),
            TokenKind::While => self.parse_while_stmt(),
            TokenKind::For => self.parse_for_stmt(false),
            TokenKind::Try => self.parse_try_stmt(),
            TokenKind::With => self.parse_with_stmt(false),
            TokenKind::Return => self.parse_return_stmt(),
            TokenKind::Raise => self.parse_raise_stmt(),
            TokenKind::Import => self.parse_import_stmt(),
            TokenKind::From => self.parse_import_from_stmt(),
            TokenKind::Global => self.parse_global_stmt(),
            TokenKind::Nonlocal => self.parse_nonlocal_stmt(),
            TokenKind::Break => { self.advance(); self.expect_newline()?; Ok(Statement::Break) }
            TokenKind::Continue => { self.advance(); self.expect_newline()?; Ok(Statement::Continue) }
            TokenKind::Pass => { self.advance(); self.expect_newline()?; Ok(Statement::Pass) }
            TokenKind::Del => self.parse_del_stmt(),
            TokenKind::Assert => self.parse_assert_stmt(),
            TokenKind::Match => self.parse_match_stmt(),
            TokenKind::At => self.parse_decorated(),
            _ => self.parse_expr_or_assign_stmt(),
        }
    }

    fn parse_async_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'async'
        match self.current_kind() {
            TokenKind::Def => self.parse_function_def(true, Vec::new()),
            TokenKind::For => self.parse_for_stmt(true),
            TokenKind::With => self.parse_with_stmt(true),
            _ => Err(self.error("expected 'def', 'for', or 'with' after 'async'")),
        }
    }

    fn parse_function_def(&mut self, is_async: bool, decorators: Vec<Expression>) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::Def)?;
        let name = self.expect_identifier()?;
        self.expect_kind(&TokenKind::LParen)?;
        let params = self.parse_parameters()?;
        self.expect_kind(&TokenKind::RParen)?;
        let returns = if self.eat(&TokenKind::Arrow) {
            Some(self.parse_expression()?)
        } else {
            None
        };
        let body = self.parse_block()?;
        Ok(Statement::FunctionDef { name, params, body, decorators, returns, is_async })
    }

    fn parse_class_def(&mut self, decorators: Vec<Expression>) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::Class)?;
        let name = self.expect_identifier()?;
        let mut bases = Vec::new();
        let mut keywords = Vec::new();
        if self.eat(&TokenKind::LParen) {
            while !self.check(&TokenKind::RParen) && !self.is_at_end() {
                // keyword arg: name=value or **kwargs
                if self.check_identifier() && self.peek_kind(1) == Some(&TokenKind::Eq) {
                    let kw_name = self.expect_identifier()?;
                    self.advance(); // skip =
                    let val = self.parse_expression()?;
                    keywords.push(Keyword { name: Some(kw_name), value: val });
                } else if self.check(&TokenKind::DoubleStar) {
                    self.advance();
                    let val = self.parse_expression()?;
                    keywords.push(Keyword { name: None, value: val });
                } else {
                    bases.push(self.parse_expression()?);
                }
                if !self.eat(&TokenKind::Comma) { break; }
            }
            self.expect_kind(&TokenKind::RParen)?;
        }
        let body = self.parse_block()?;
        Ok(Statement::ClassDef { name, bases, keywords, body, decorators })
    }

    fn parse_if_stmt(&mut self) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::If)?;
        let test = self.parse_expression()?;
        // Single-line if: if cond: stmt
        if self.eat(&TokenKind::Colon) && !self.check(&TokenKind::Newline) {
            let body = vec![self.parse_simple_stmt()?];
            self.eat(&TokenKind::Newline);
            return Ok(Statement::If { test, body, elif_clauses: Vec::new(), else_body: None });
        }
        // Already ate colon in eat() above, or need block
        let body = if self.check(&TokenKind::Newline) || self.check(&TokenKind::Indent) {
            self.parse_block_after_colon()?
        } else {
            // Colon was already consumed by eat()
            self.parse_block_after_colon()?
        };
        let mut elif_clauses = Vec::new();
        while self.check(&TokenKind::Elif) {
            self.advance();
            let elif_test = self.parse_expression()?;
            let elif_body = self.parse_block()?;
            elif_clauses.push((elif_test, elif_body));
        }
        let else_body = if self.check(&TokenKind::Else) {
            self.advance();
            Some(self.parse_block()?)
        } else {
            None
        };
        Ok(Statement::If { test, body, elif_clauses, else_body })
    }

    fn parse_while_stmt(&mut self) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::While)?;
        let test = self.parse_expression()?;
        let body = self.parse_block()?;
        let else_body = if self.check(&TokenKind::Else) {
            self.advance();
            Some(self.parse_block()?)
        } else {
            None
        };
        Ok(Statement::While { test, body, else_body })
    }

    fn parse_for_stmt(&mut self, is_async: bool) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::For)?;
        let target = self.parse_target_list()?;
        self.expect_kind(&TokenKind::In)?;
        let iter = self.parse_expression_list()?;
        let body = self.parse_block()?;
        let else_body = if self.check(&TokenKind::Else) {
            self.advance();
            Some(self.parse_block()?)
        } else {
            None
        };
        Ok(Statement::For { target, iter, body, else_body, is_async })
    }

    fn parse_try_stmt(&mut self) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::Try)?;
        let body = self.parse_block()?;
        let mut handlers = Vec::new();
        while self.check(&TokenKind::Except) {
            self.advance();
            let (exc_type, name) = if self.check(&TokenKind::Colon) {
                (None, None)
            } else {
                let exc = self.parse_expression()?;
                let name = if self.eat(&TokenKind::As) {
                    Some(self.expect_identifier()?)
                } else {
                    None
                };
                (Some(exc), name)
            };
            let handler_body = self.parse_block()?;
            handlers.push(ExceptHandler { exc_type, name, body: handler_body });
        }
        let else_body = if self.check(&TokenKind::Else) {
            self.advance();
            Some(self.parse_block()?)
        } else {
            None
        };
        let finally_body = if self.check(&TokenKind::Finally) {
            self.advance();
            Some(self.parse_block()?)
        } else {
            None
        };
        Ok(Statement::Try { body, handlers, else_body, finally_body })
    }

    fn parse_with_stmt(&mut self, is_async: bool) -> Result<Statement, String> {
        self.expect_kind(&TokenKind::With)?;
        let mut items = Vec::new();
        loop {
            let context_expr = self.parse_expression()?;
            let optional_vars = if self.eat(&TokenKind::As) {
                Some(self.parse_target_list()?)
            } else {
                None
            };
            items.push(WithItem { context_expr, optional_vars });
            if !self.eat(&TokenKind::Comma) { break; }
        }
        let body = self.parse_block()?;
        Ok(Statement::With { items, body, is_async })
    }

    fn parse_return_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'return'
        if self.check(&TokenKind::Newline) || self.check(&TokenKind::Eof) {
            self.eat(&TokenKind::Newline);
            return Ok(Statement::Return(None));
        }
        let expr = self.parse_expression_list()?;
        self.expect_newline()?;
        Ok(Statement::Return(Some(expr)))
    }

    fn parse_raise_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'raise'
        if self.check(&TokenKind::Newline) || self.check(&TokenKind::Eof) {
            self.eat(&TokenKind::Newline);
            return Ok(Statement::Raise { exc: None, cause: None });
        }
        let exc = self.parse_expression()?;
        let cause = if self.eat(&TokenKind::From) {
            Some(self.parse_expression()?)
        } else {
            None
        };
        self.expect_newline()?;
        Ok(Statement::Raise { exc: Some(exc), cause })
    }

    fn parse_import_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'import'
        let mut names = Vec::new();
        loop {
            let name = self.parse_dotted_name()?;
            let asname = if self.eat(&TokenKind::As) {
                Some(self.expect_identifier()?)
            } else {
                None
            };
            names.push(Alias { name, asname });
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect_newline()?;
        Ok(Statement::Import { names })
    }

    fn parse_import_from_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'from'
        let mut level = 0;
        while self.check(&TokenKind::Dot) || self.check(&TokenKind::DotDotDot) {
            if self.eat(&TokenKind::DotDotDot) { level += 3; }
            else { self.advance(); level += 1; }
        }
        let module = if !self.check(&TokenKind::Import) {
            Some(self.parse_dotted_name()?)
        } else {
            None
        };
        self.expect_kind(&TokenKind::Import)?;
        let names = if self.eat(&TokenKind::Star) {
            vec![Alias { name: "*".to_string(), asname: None }]
        } else {
            let paren = self.eat(&TokenKind::LParen);
            let mut names = Vec::new();
            loop {
                if paren && self.check(&TokenKind::RParen) { break; }
                let name = self.expect_identifier()?;
                let asname = if self.eat(&TokenKind::As) {
                    Some(self.expect_identifier()?)
                } else {
                    None
                };
                names.push(Alias { name, asname });
                if !self.eat(&TokenKind::Comma) { break; }
            }
            if paren { self.expect_kind(&TokenKind::RParen)?; }
            names
        };
        self.expect_newline()?;
        Ok(Statement::ImportFrom { module, names, level })
    }

    fn parse_global_stmt(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut names = Vec::new();
        loop {
            names.push(self.expect_identifier()?);
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect_newline()?;
        Ok(Statement::Global(names))
    }

    fn parse_nonlocal_stmt(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut names = Vec::new();
        loop {
            names.push(self.expect_identifier()?);
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect_newline()?;
        Ok(Statement::Nonlocal(names))
    }

    fn parse_del_stmt(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut targets = Vec::new();
        loop {
            targets.push(self.parse_expression()?);
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect_newline()?;
        Ok(Statement::Delete(targets))
    }

    fn parse_assert_stmt(&mut self) -> Result<Statement, String> {
        self.advance();
        let test = self.parse_expression()?;
        let msg = if self.eat(&TokenKind::Comma) {
            Some(self.parse_expression()?)
        } else {
            None
        };
        self.expect_newline()?;
        Ok(Statement::Assert { test, msg })
    }

    fn parse_match_stmt(&mut self) -> Result<Statement, String> {
        self.advance(); // skip 'match'
        let subject = self.parse_expression_list()?;
        let mut cases = Vec::new();
        self.expect_kind(&TokenKind::Colon)?;
        self.expect_kind(&TokenKind::Newline)?;
        self.expect_kind(&TokenKind::Indent)?;
        while self.check(&TokenKind::Case) {
            self.advance();
            let pattern = self.parse_pattern()?;
            let guard = if self.eat(&TokenKind::If) {
                Some(self.parse_expression()?)
            } else {
                None
            };
            let body = self.parse_block()?;
            cases.push(MatchCase { pattern, guard, body });
        }
        self.expect_kind(&TokenKind::Dedent)?;
        Ok(Statement::Match { subject, cases })
    }

    fn parse_pattern(&mut self) -> Result<Pattern, String> {
        let pat = self.parse_single_pattern()?;
        if self.check(&TokenKind::Pipe) {
            let mut patterns = vec![pat];
            while self.eat(&TokenKind::Pipe) {
                patterns.push(self.parse_single_pattern()?);
            }
            Ok(Pattern::Or(patterns))
        } else {
            Ok(pat)
        }
    }

    fn parse_single_pattern(&mut self) -> Result<Pattern, String> {
        let pat = match self.current_kind() {
            TokenKind::Star => {
                self.advance();
                if self.check_identifier() {
                    let name = self.expect_identifier()?;
                    Pattern::Star(Some(name))
                } else {
                    Pattern::Star(None)
                }
            }
            TokenKind::Identifier(name) if name == "_" => {
                self.advance();
                Pattern::Wildcard
            }
            TokenKind::None | TokenKind::True | TokenKind::False => {
                let expr = self.parse_primary()?;
                Pattern::Singleton(expr)
            }
            TokenKind::LBracket => {
                self.advance();
                let mut pats = Vec::new();
                while !self.check(&TokenKind::RBracket) && !self.is_at_end() {
                    pats.push(self.parse_pattern()?);
                    if !self.eat(&TokenKind::Comma) { break; }
                }
                self.expect_kind(&TokenKind::RBracket)?;
                Pattern::Sequence(pats)
            }
            TokenKind::LBrace => {
                self.advance();
                let mut pairs = Vec::new();
                while !self.check(&TokenKind::RBrace) && !self.is_at_end() {
                    let key = self.parse_expression()?;
                    self.expect_kind(&TokenKind::Colon)?;
                    let val = self.parse_pattern()?;
                    pairs.push((key, val));
                    if !self.eat(&TokenKind::Comma) { break; }
                }
                self.expect_kind(&TokenKind::RBrace)?;
                Pattern::Mapping(pairs)
            }
            TokenKind::Identifier(name) if !name.starts_with('_') => {
                // Could be a name capture (case x:) or a value (case MyConst:)
                // If followed by 'if' or ':', treat as name capture for guard support
                let name = name.clone();
                self.advance();
                if self.check(&TokenKind::If) || self.check(&TokenKind::Colon) {
                    // Name capture pattern
                    Pattern::As { pattern: None, name: Some(name) }
                } else if self.check(&TokenKind::LParen) {
                    // Class pattern: ClassName(...)
                    self.advance();
                    let mut pats = Vec::new();
                    while !self.check(&TokenKind::RParen) && !self.is_at_end() {
                        pats.push(self.parse_pattern()?);
                        if !self.eat(&TokenKind::Comma) { break; }
                    }
                    self.expect_kind(&TokenKind::RParen)?;
                    Pattern::Class { cls: Expression::Name(name), patterns: pats, kw_patterns: Vec::new() }
                } else if self.check(&TokenKind::Dot) {
                    // Dotted name like module.CONST — parse as value
                    let mut expr = Expression::Name(name);
                    while self.eat(&TokenKind::Dot) {
                        let attr = self.expect_identifier()?;
                        expr = Expression::Attribute { value: Box::new(expr), attr };
                    }
                    Pattern::Value(expr)
                } else {
                    // Treat as name capture
                    Pattern::As { pattern: None, name: Some(name) }
                }
            }
            _ => {
                let expr = self.parse_expression()?;
                Pattern::Value(expr)
            }
        };

        // 'as' binding
        if self.eat(&TokenKind::As) {
            let name = self.expect_identifier()?;
            Ok(Pattern::As { pattern: Some(Box::new(pat)), name: Some(name) })
        } else {
            Ok(pat)
        }
    }

    fn parse_decorated(&mut self) -> Result<Statement, String> {
        let mut decorators = Vec::new();
        while self.check(&TokenKind::At) {
            self.advance();
            decorators.push(self.parse_expression()?);
            self.expect_kind(&TokenKind::Newline)?;
        }
        match self.current_kind() {
            TokenKind::Def => self.parse_function_def(false, decorators),
            TokenKind::Async => {
                self.advance();
                self.parse_function_def(true, decorators)
            }
            TokenKind::Class => self.parse_class_def(decorators),
            _ => Err(self.error("expected 'def', 'async def', or 'class' after decorator")),
        }
    }

    fn parse_expr_or_assign_stmt(&mut self) -> Result<Statement, String> {
        let expr = self.parse_expression_list()?;

        // Augmented assignment: +=, -=, etc.
        if let Some(op) = self.try_aug_op() {
            self.advance();
            let value = self.parse_expression_list()?;
            self.expect_newline()?;
            return Ok(Statement::AugAssign { target: expr, op, value });
        }

        // Regular assignment: a = b (= c ...)
        if self.check(&TokenKind::Eq) {
            let mut targets = vec![expr];
            while self.eat(&TokenKind::Eq) {
                targets.push(self.parse_expression_list()?);
            }
            let value = targets.pop().unwrap();
            self.expect_newline()?;
            return Ok(Statement::Assign { targets, value });
        }

        // Annotated assignment: x: type = value
        if self.check(&TokenKind::Colon) && !self.at_block_colon() {
            self.advance();
            let annotation = self.parse_expression()?;
            let value = if self.eat(&TokenKind::Eq) {
                Some(self.parse_expression_list()?)
            } else {
                None
            };
            self.expect_newline()?;
            return Ok(Statement::AnnAssign { target: expr, annotation, value });
        }

        self.expect_newline()?;
        Ok(Statement::Expression(expr))
    }

    fn try_aug_op(&self) -> Option<AugOp> {
        match self.current_kind() {
            TokenKind::PlusEq => Some(AugOp::Add),
            TokenKind::MinusEq => Some(AugOp::Sub),
            TokenKind::StarEq => Some(AugOp::Mul),
            TokenKind::SlashEq => Some(AugOp::Div),
            TokenKind::DoubleSlashEq => Some(AugOp::FloorDiv),
            TokenKind::PercentEq => Some(AugOp::Mod),
            TokenKind::DoubleStarEq => Some(AugOp::Pow),
            TokenKind::LtLtEq => Some(AugOp::LShift),
            TokenKind::GtGtEq => Some(AugOp::RShift),
            TokenKind::PipeEq => Some(AugOp::BitOr),
            TokenKind::AmpEq => Some(AugOp::BitAnd),
            TokenKind::CaretEq => Some(AugOp::BitXor),
            TokenKind::AtEq => Some(AugOp::MatMul),
            _ => None,
        }
    }

    // Check if current colon is a block-starting colon (followed by newline/indent)
    fn at_block_colon(&self) -> bool {
        if !self.check(&TokenKind::Colon) { return false; }
        matches!(self.peek_kind(1), Some(TokenKind::Newline) | Some(TokenKind::Eof) | None)
    }

    fn parse_simple_stmt(&mut self) -> Result<Statement, String> {
        // Parse a single simple statement (for single-line if/while)
        match self.current_kind() {
            TokenKind::Return => self.parse_return_stmt_inline(),
            TokenKind::Break => { self.advance(); Ok(Statement::Break) }
            TokenKind::Continue => { self.advance(); Ok(Statement::Continue) }
            TokenKind::Pass => { self.advance(); Ok(Statement::Pass) }
            TokenKind::Raise => self.parse_raise_stmt_inline(),
            _ => {
                let expr = self.parse_expression_list()?;
                if let Some(op) = self.try_aug_op() {
                    self.advance();
                    let value = self.parse_expression_list()?;
                    return Ok(Statement::AugAssign { target: expr, op, value });
                }
                if self.check(&TokenKind::Eq) {
                    let mut targets = vec![expr];
                    while self.eat(&TokenKind::Eq) {
                        targets.push(self.parse_expression_list()?);
                    }
                    let value = targets.pop().unwrap();
                    return Ok(Statement::Assign { targets, value });
                }
                Ok(Statement::Expression(expr))
            }
        }
    }

    fn parse_return_stmt_inline(&mut self) -> Result<Statement, String> {
        self.advance();
        if self.check(&TokenKind::Newline) || self.check(&TokenKind::Eof) || self.check(&TokenKind::Semicolon) {
            return Ok(Statement::Return(None));
        }
        let expr = self.parse_expression_list()?;
        Ok(Statement::Return(Some(expr)))
    }

    fn parse_raise_stmt_inline(&mut self) -> Result<Statement, String> {
        self.advance();
        if self.check(&TokenKind::Newline) || self.check(&TokenKind::Eof) || self.check(&TokenKind::Semicolon) {
            return Ok(Statement::Raise { exc: None, cause: None });
        }
        let exc = self.parse_expression()?;
        let cause = if self.eat(&TokenKind::From) { Some(self.parse_expression()?) } else { None };
        Ok(Statement::Raise { exc: Some(exc), cause })
    }

    // ── Block parsing ────────────────────────────────────────────────

    fn parse_block(&mut self) -> Result<Vec<Statement>, String> {
        self.expect_kind(&TokenKind::Colon)?;

        // Single-line block: `if cond: stmt`
        if !self.check(&TokenKind::Newline) {
            let stmt = self.parse_simple_stmt()?;
            self.eat(&TokenKind::Newline);
            return Ok(vec![stmt]);
        }

        self.parse_block_after_colon()
    }

    fn parse_block_after_colon(&mut self) -> Result<Vec<Statement>, String> {
        self.expect_kind(&TokenKind::Newline)?;
        self.expect_kind(&TokenKind::Indent)?;
        let mut stmts = Vec::new();
        while !self.check(&TokenKind::Dedent) && !self.is_at_end() {
            if self.check(&TokenKind::Newline) {
                self.advance();
                continue;
            }
            stmts.push(self.parse_statement()?);
        }
        self.expect_kind(&TokenKind::Dedent)?;
        Ok(stmts)
    }

    // ── Parameters ───────────────────────────────────────────────────

    fn parse_parameters(&mut self) -> Result<Parameters, String> {
        let mut params = Parameters::default();
        let mut seen_star = false;
        let mut positional_defaults = Vec::new();

        while !self.check(&TokenKind::RParen) && !self.is_at_end() {
            // **kwargs
            if self.eat(&TokenKind::DoubleStar) {
                let name = self.expect_identifier()?;
                let annotation = if self.eat(&TokenKind::Colon) {
                    Some(Box::new(self.parse_expression()?))
                } else {
                    None
                };
                params.kwarg = Some(Param { name, annotation });
                self.eat(&TokenKind::Comma);
                break;
            }

            // *args or bare * (keyword-only separator)
            if self.eat(&TokenKind::Star) {
                seen_star = true;
                if self.check_identifier() {
                    let name = self.expect_identifier()?;
                    let annotation = if self.eat(&TokenKind::Colon) {
                        Some(Box::new(self.parse_expression()?))
                    } else {
                        None
                    };
                    params.vararg = Some(Param { name, annotation });
                }
                if !self.eat(&TokenKind::Comma) { break; }
                continue;
            }

            let name = self.expect_identifier()?;
            let annotation = if self.eat(&TokenKind::Colon) {
                Some(Box::new(self.parse_expression()?))
            } else {
                None
            };
            let default = if self.eat(&TokenKind::Eq) {
                Some(self.parse_expression()?)
            } else {
                None
            };

            if seen_star {
                params.kwonly_args.push(Param { name, annotation });
                params.kw_defaults.push(default);
            } else {
                params.args.push(Param { name, annotation });
                if let Some(d) = default {
                    positional_defaults.push(d);
                }
            }

            if !self.eat(&TokenKind::Comma) { break; }
        }

        params.defaults = positional_defaults;
        Ok(params)
    }

    // ── Expression parsing (precedence climbing) ─────────────────────

    fn parse_expression(&mut self) -> Result<Expression, String> {
        // Check for lambda
        if self.check(&TokenKind::Lambda) {
            return self.parse_lambda();
        }
        self.parse_named_expr()
    }

    fn parse_expression_list(&mut self) -> Result<Expression, String> {
        let first = self.parse_expression()?;
        if self.check(&TokenKind::Comma) && !self.at_block_colon_or_assign() {
            let mut elements = vec![first];
            while self.eat(&TokenKind::Comma) {
                if self.is_expr_end() { break; }
                elements.push(self.parse_expression()?);
            }
            if elements.len() == 1 {
                return Ok(elements.into_iter().next().unwrap());
            }
            Ok(Expression::Tuple(elements))
        } else {
            Ok(first)
        }
    }

    fn at_block_colon_or_assign(&self) -> bool {
        // Don't consume comma if we're about to hit = or : (annotation context)
        false // let the caller handle this
    }

    fn is_expr_end(&self) -> bool {
        matches!(self.current_kind(),
            TokenKind::Newline | TokenKind::Eof | TokenKind::Colon
            | TokenKind::RParen | TokenKind::RBracket | TokenKind::RBrace
            | TokenKind::Semicolon | TokenKind::Eq
            | TokenKind::PlusEq | TokenKind::MinusEq | TokenKind::StarEq
            | TokenKind::SlashEq | TokenKind::DoubleSlashEq | TokenKind::PercentEq
            | TokenKind::DoubleStarEq | TokenKind::AtEq | TokenKind::PipeEq
            | TokenKind::AmpEq | TokenKind::CaretEq | TokenKind::LtLtEq | TokenKind::GtGtEq
        )
    }

    fn parse_lambda(&mut self) -> Result<Expression, String> {
        self.advance(); // skip 'lambda'
        let params = if self.check(&TokenKind::Colon) {
            Parameters::default()
        } else {
            self.parse_lambda_params()?
        };
        self.expect_kind(&TokenKind::Colon)?;
        let body = self.parse_expression()?;
        Ok(Expression::Lambda { params, body: Box::new(body) })
    }

    /// Parse lambda parameters (no parens, terminated by colon).
    /// Lambda params don't support annotations.
    fn parse_lambda_params(&mut self) -> Result<Parameters, String> {
        let mut params = Parameters::default();
        let mut seen_star = false;
        let mut positional_defaults = Vec::new();

        while !self.check(&TokenKind::Colon) && !self.is_at_end() {
            if self.eat(&TokenKind::DoubleStar) {
                let name = self.expect_identifier()?;
                params.kwarg = Some(Param { name, annotation: None });
                self.eat(&TokenKind::Comma);
                break;
            }
            if self.eat(&TokenKind::Star) {
                seen_star = true;
                if self.check_identifier() {
                    let name = self.expect_identifier()?;
                    params.vararg = Some(Param { name, annotation: None });
                }
                if !self.eat(&TokenKind::Comma) { break; }
                continue;
            }
            let name = self.expect_identifier()?;
            let default = if self.eat(&TokenKind::Eq) {
                Some(self.parse_expression()?)
            } else {
                None
            };
            if seen_star {
                params.kwonly_args.push(Param { name, annotation: None });
                params.kw_defaults.push(default);
            } else {
                params.args.push(Param { name, annotation: None });
                if let Some(d) = default { positional_defaults.push(d); }
            }
            if !self.eat(&TokenKind::Comma) { break; }
        }
        params.defaults = positional_defaults;
        Ok(params)
    }

    fn parse_named_expr(&mut self) -> Result<Expression, String> {
        let expr = self.parse_ternary()?;
        if self.check(&TokenKind::ColonEq) {
            self.advance();
            let value = self.parse_named_expr()?;
            Ok(Expression::NamedExpr { target: Box::new(expr), value: Box::new(value) })
        } else {
            Ok(expr)
        }
    }

    fn parse_ternary(&mut self) -> Result<Expression, String> {
        let body = self.parse_or()?;
        if self.check(&TokenKind::If) {
            self.advance();
            let test = self.parse_or()?;
            self.expect_kind(&TokenKind::Else)?;
            let orelse = self.parse_ternary()?;
            Ok(Expression::IfExp { test: Box::new(test), body: Box::new(body), orelse: Box::new(orelse) })
        } else {
            Ok(body)
        }
    }

    fn parse_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_and()?;
        while self.check(&TokenKind::Or) {
            self.advance();
            let right = self.parse_and()?;
            left = Expression::BoolOp { op: BoolOp::Or, values: flatten_bool_op(BoolOp::Or, left, right) };
        }
        Ok(left)
    }

    fn parse_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_not()?;
        while self.check(&TokenKind::And) {
            self.advance();
            let right = self.parse_not()?;
            left = Expression::BoolOp { op: BoolOp::And, values: flatten_bool_op(BoolOp::And, left, right) };
        }
        Ok(left)
    }

    fn parse_not(&mut self) -> Result<Expression, String> {
        if self.check(&TokenKind::Not) {
            self.advance();
            let operand = self.parse_not()?;
            Ok(Expression::UnaryOp { op: UnaryOp::Not, operand: Box::new(operand) })
        } else {
            self.parse_comparison()
        }
    }

    fn parse_comparison(&mut self) -> Result<Expression, String> {
        let left = self.parse_bitor()?;
        let mut ops = Vec::new();
        let mut comparators = Vec::new();
        loop {
            let op = match self.current_kind() {
                TokenKind::Lt => { self.advance(); CmpOp::Lt }
                TokenKind::Gt => { self.advance(); CmpOp::Gt }
                TokenKind::LtEq => { self.advance(); CmpOp::LtE }
                TokenKind::GtEq => { self.advance(); CmpOp::GtE }
                TokenKind::EqEq => { self.advance(); CmpOp::Eq }
                TokenKind::BangEq => { self.advance(); CmpOp::NotEq }
                TokenKind::In => { self.advance(); CmpOp::In }
                TokenKind::Not if matches!(self.peek_kind(1), Some(TokenKind::In)) => {
                    self.advance(); self.advance(); CmpOp::NotIn
                }
                TokenKind::Is => {
                    self.advance();
                    if self.eat(&TokenKind::Not) { CmpOp::IsNot } else { CmpOp::Is }
                }
                _ => break,
            };
            ops.push(op);
            comparators.push(self.parse_bitor()?);
        }
        if ops.is_empty() {
            Ok(left)
        } else {
            Ok(Expression::Compare { left: Box::new(left), ops, comparators })
        }
    }

    fn parse_bitor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitxor()?;
        while self.check(&TokenKind::Pipe) {
            self.advance();
            let right = self.parse_bitxor()?;
            left = Expression::BinOp { op: BinOp::BitOr, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitxor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitand()?;
        while self.check(&TokenKind::Caret) {
            self.advance();
            let right = self.parse_bitand()?;
            left = Expression::BinOp { op: BinOp::BitXor, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitand(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_shift()?;
        while self.check(&TokenKind::Amp) {
            self.advance();
            let right = self.parse_shift()?;
            left = Expression::BinOp { op: BinOp::BitAnd, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_shift(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_additive()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::LtLt => BinOp::LShift,
                TokenKind::GtGt => BinOp::RShift,
                _ => break,
            };
            self.advance();
            let right = self.parse_additive()?;
            left = Expression::BinOp { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_additive(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_multiplicative()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::Plus => BinOp::Add,
                TokenKind::Minus => BinOp::Sub,
                _ => break,
            };
            self.advance();
            let right = self.parse_multiplicative()?;
            left = Expression::BinOp { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_unary()?;
        loop {
            let op = match self.current_kind() {
                TokenKind::Star => BinOp::Mul,
                TokenKind::Slash => BinOp::Div,
                TokenKind::DoubleSlash => BinOp::FloorDiv,
                TokenKind::Percent => BinOp::Mod,
                TokenKind::At => BinOp::MatMul,
                _ => break,
            };
            self.advance();
            let right = self.parse_unary()?;
            left = Expression::BinOp { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<Expression, String> {
        match self.current_kind() {
            TokenKind::Minus => {
                self.advance();
                let operand = self.parse_unary()?;
                Ok(Expression::UnaryOp { op: UnaryOp::USub, operand: Box::new(operand) })
            }
            TokenKind::Plus => {
                self.advance();
                let operand = self.parse_unary()?;
                Ok(Expression::UnaryOp { op: UnaryOp::UAdd, operand: Box::new(operand) })
            }
            TokenKind::Tilde => {
                self.advance();
                let operand = self.parse_unary()?;
                Ok(Expression::UnaryOp { op: UnaryOp::Invert, operand: Box::new(operand) })
            }
            _ => self.parse_power(),
        }
    }

    fn parse_power(&mut self) -> Result<Expression, String> {
        let base = self.parse_await_expr()?;
        if self.check(&TokenKind::DoubleStar) {
            self.advance();
            // Right-associative
            let exp = self.parse_unary()?;
            Ok(Expression::BinOp { op: BinOp::Pow, left: Box::new(base), right: Box::new(exp) })
        } else {
            Ok(base)
        }
    }

    fn parse_await_expr(&mut self) -> Result<Expression, String> {
        if self.check(&TokenKind::Await) {
            self.advance();
            let expr = self.parse_unary()?;
            Ok(Expression::Await(Box::new(expr)))
        } else {
            self.parse_postfix()
        }
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            match self.current_kind() {
                TokenKind::Dot => {
                    self.advance();
                    let attr = self.expect_identifier()?;
                    expr = Expression::Attribute { value: Box::new(expr), attr };
                }
                TokenKind::LParen => {
                    self.advance();
                    let (args, keywords) = self.parse_call_args()?;
                    self.expect_kind(&TokenKind::RParen)?;
                    expr = Expression::Call { func: Box::new(expr), args, keywords };
                }
                TokenKind::LBracket => {
                    self.advance();
                    let slice = self.parse_subscript()?;
                    self.expect_kind(&TokenKind::RBracket)?;
                    expr = Expression::Subscript { value: Box::new(expr), slice: Box::new(slice) };
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_call_args(&mut self) -> Result<(Vec<Expression>, Vec<Keyword>), String> {
        let mut args = Vec::new();
        let mut keywords = Vec::new();
        while !self.check(&TokenKind::RParen) && !self.is_at_end() {
            // **kwargs
            if self.check(&TokenKind::DoubleStar) {
                self.advance();
                let val = self.parse_expression()?;
                keywords.push(Keyword { name: None, value: val });
            }
            // *args
            else if self.check(&TokenKind::Star) {
                self.advance();
                let val = self.parse_expression()?;
                args.push(Expression::Starred(Box::new(val)));
            }
            // keyword=value
            else if self.check_identifier() && self.peek_kind(1) == Some(&TokenKind::Eq) {
                let name = self.expect_identifier()?;
                self.advance(); // skip =
                let val = self.parse_expression()?;
                keywords.push(Keyword { name: Some(name), value: val });
            }
            // Generator expression: f(x for x in y)
            else {
                let expr = self.parse_expression()?;
                if self.check(&TokenKind::For) && args.is_empty() && keywords.is_empty() {
                    let generators = self.parse_comp_clauses()?;
                    args.push(Expression::GeneratorExp { element: Box::new(expr), generators });
                } else {
                    args.push(expr);
                }
            }
            if !self.eat(&TokenKind::Comma) { break; }
        }
        Ok((args, keywords))
    }

    fn parse_subscript(&mut self) -> Result<Expression, String> {
        // Could be a simple index, a slice, or a tuple of slices
        if self.check(&TokenKind::Colon) {
            return self.parse_slice(None);
        }
        let first = self.parse_expression()?;
        if self.check(&TokenKind::Colon) {
            return self.parse_slice(Some(first));
        }
        // Tuple subscript: a[1, 2]
        if self.check(&TokenKind::Comma) {
            let mut elems = vec![first];
            while self.eat(&TokenKind::Comma) {
                if self.check(&TokenKind::RBracket) { break; }
                elems.push(self.parse_subscript_item()?);
            }
            return Ok(Expression::Tuple(elems));
        }
        Ok(first)
    }

    fn parse_subscript_item(&mut self) -> Result<Expression, String> {
        if self.check(&TokenKind::Colon) {
            return self.parse_slice(None);
        }
        let expr = self.parse_expression()?;
        if self.check(&TokenKind::Colon) {
            return self.parse_slice(Some(expr));
        }
        Ok(expr)
    }

    fn parse_slice(&mut self, lower: Option<Expression>) -> Result<Expression, String> {
        self.expect_kind(&TokenKind::Colon)?;
        let upper = if self.check(&TokenKind::Colon) || self.check(&TokenKind::RBracket) || self.check(&TokenKind::Comma) {
            None
        } else {
            Some(Box::new(self.parse_expression()?))
        };
        let step = if self.eat(&TokenKind::Colon) {
            if self.check(&TokenKind::RBracket) || self.check(&TokenKind::Comma) {
                None
            } else {
                Some(Box::new(self.parse_expression()?))
            }
        } else {
            None
        };
        Ok(Expression::Slice { lower: lower.map(Box::new), upper, step })
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.current_kind() {
            TokenKind::Int(n) => { let v = n; self.advance(); Ok(Expression::Int(v)) }
            TokenKind::Float(n) => { let v = n; self.advance(); Ok(Expression::Float(v)) }
            TokenKind::Str(ref s) => { let v = s.clone(); self.advance(); Ok(self.concat_strings(v)?) }
            TokenKind::FStringStart => self.parse_fstring_expr(),
            TokenKind::ByteStr(ref b) => { let v = b.clone(); self.advance(); Ok(Expression::Str(String::from_utf8_lossy(&v).into_owned())) }
            TokenKind::True => { self.advance(); Ok(Expression::Bool(true)) }
            TokenKind::False => { self.advance(); Ok(Expression::Bool(false)) }
            TokenKind::None => { self.advance(); Ok(Expression::None) }
            TokenKind::DotDotDot => { self.advance(); Ok(Expression::Ellipsis) }
            TokenKind::Identifier(ref name) => { let n = name.clone(); self.advance(); Ok(Expression::Name(n)) }

            // Yield expression
            TokenKind::Yield => {
                self.advance();
                if self.eat(&TokenKind::From) {
                    let expr = self.parse_expression()?;
                    Ok(Expression::YieldFrom(Box::new(expr)))
                } else if self.is_expr_end() {
                    Ok(Expression::Yield(None))
                } else {
                    let expr = self.parse_expression_list()?;
                    Ok(Expression::Yield(Some(Box::new(expr))))
                }
            }

            // Star expression (for unpacking)
            TokenKind::Star => {
                self.advance();
                let expr = self.parse_expression()?;
                Ok(Expression::Starred(Box::new(expr)))
            }

            // Parenthesized expression, tuple, or generator
            TokenKind::LParen => {
                self.advance();
                if self.check(&TokenKind::RParen) {
                    self.advance();
                    return Ok(Expression::Tuple(Vec::new()));
                }
                let first = self.parse_expression()?;
                // Generator expression
                if self.check(&TokenKind::For) {
                    let generators = self.parse_comp_clauses()?;
                    self.expect_kind(&TokenKind::RParen)?;
                    return Ok(Expression::GeneratorExp { element: Box::new(first), generators });
                }
                // Tuple
                if self.eat(&TokenKind::Comma) {
                    let mut elements = vec![first];
                    while !self.check(&TokenKind::RParen) && !self.is_at_end() {
                        elements.push(self.parse_expression()?);
                        if !self.eat(&TokenKind::Comma) { break; }
                    }
                    self.expect_kind(&TokenKind::RParen)?;
                    return Ok(Expression::Tuple(elements));
                }
                self.expect_kind(&TokenKind::RParen)?;
                Ok(first)
            }

            // List or list comprehension
            TokenKind::LBracket => {
                self.advance();
                if self.check(&TokenKind::RBracket) {
                    self.advance();
                    return Ok(Expression::List(Vec::new()));
                }
                let first = self.parse_expression()?;
                // List comprehension
                if self.check(&TokenKind::For) {
                    let generators = self.parse_comp_clauses()?;
                    self.expect_kind(&TokenKind::RBracket)?;
                    return Ok(Expression::ListComp { element: Box::new(first), generators });
                }
                // Normal list
                let mut elements = vec![first];
                while self.eat(&TokenKind::Comma) {
                    if self.check(&TokenKind::RBracket) { break; }
                    elements.push(self.parse_expression()?);
                }
                self.expect_kind(&TokenKind::RBracket)?;
                Ok(Expression::List(elements))
            }

            // Dict, set, or comprehension
            TokenKind::LBrace => {
                self.advance();
                if self.check(&TokenKind::RBrace) {
                    self.advance();
                    return Ok(Expression::Dict { keys: Vec::new(), values: Vec::new() });
                }
                // ** unpacking in dict
                if self.check(&TokenKind::DoubleStar) {
                    return self.parse_dict_rest(Vec::new(), Vec::new());
                }
                let first = self.parse_expression()?;
                // Dict: {key: value, ...}
                if self.check(&TokenKind::Colon) {
                    self.advance();
                    let value = self.parse_expression()?;
                    // Dict comprehension
                    if self.check(&TokenKind::For) {
                        let generators = self.parse_comp_clauses()?;
                        self.expect_kind(&TokenKind::RBrace)?;
                        return Ok(Expression::DictComp { key: Box::new(first), value: Box::new(value), generators });
                    }
                    let keys = vec![Some(first)];
                    let values = vec![value];
                    if self.eat(&TokenKind::Comma) {
                        return self.parse_dict_rest(keys, values);
                    }
                    self.expect_kind(&TokenKind::RBrace)?;
                    return Ok(Expression::Dict { keys, values });
                }
                // Set comprehension
                if self.check(&TokenKind::For) {
                    let generators = self.parse_comp_clauses()?;
                    self.expect_kind(&TokenKind::RBrace)?;
                    return Ok(Expression::SetComp { element: Box::new(first), generators });
                }
                // Set literal
                let mut elements = vec![first];
                while self.eat(&TokenKind::Comma) {
                    if self.check(&TokenKind::RBrace) { break; }
                    elements.push(self.parse_expression()?);
                }
                self.expect_kind(&TokenKind::RBrace)?;
                Ok(Expression::Set(elements))
            }

            _ => Err(self.error(&format!("unexpected token {:?}", self.current_kind()))),
        }
    }

    fn parse_dict_rest(&mut self, mut keys: Vec<Option<Expression>>, mut values: Vec<Expression>) -> Result<Expression, String> {
        while !self.check(&TokenKind::RBrace) && !self.is_at_end() {
            if self.check(&TokenKind::DoubleStar) {
                self.advance();
                keys.push(None);
                values.push(self.parse_expression()?);
            } else {
                let key = self.parse_expression()?;
                self.expect_kind(&TokenKind::Colon)?;
                let value = self.parse_expression()?;
                keys.push(Some(key));
                values.push(value);
            }
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect_kind(&TokenKind::RBrace)?;
        Ok(Expression::Dict { keys, values })
    }

    fn parse_fstring_expr(&mut self) -> Result<Expression, String> {
        self.advance(); // skip FStringStart
        let mut parts = Vec::new();
        while !self.check(&TokenKind::FStringEnd) && !self.is_at_end() {
            match self.current_kind() {
                TokenKind::FStringText(ref s) => {
                    let text = s.clone();
                    self.advance();
                    parts.push(FStringPart::Literal(text));
                }
                _ => {
                    // Expression tokens
                    let expr = self.parse_expression()?;
                    // Check for format spec after expression
                    if let TokenKind::FStringFormatSpec(ref spec) = self.current_kind() {
                        let spec = spec.clone();
                        self.advance();
                        parts.push(FStringPart::FormattedExpr(expr, spec));
                    } else {
                        parts.push(FStringPart::Expr(expr));
                    }
                }
            }
        }
        self.eat_kind(&TokenKind::FStringEnd);
        Ok(Expression::FString { parts })
    }

    // Implicit string concatenation: "hello" " world"
    fn concat_strings(&mut self, first: String) -> Result<Expression, String> {
        let mut s = first;
        loop {
            match self.current_kind() {
                TokenKind::Str(ref next) => {
                    s.push_str(next);
                    self.advance();
                }
                TokenKind::FStringStart => {
                    // Mixed concat with f-string: not supported yet, just return what we have
                    break;
                }
                _ => break,
            }
        }
        Ok(Expression::Str(s))
    }

    // ── Comprehension clauses ────────────────────────────────────────

    fn parse_comp_clauses(&mut self) -> Result<Vec<Comprehension>, String> {
        let mut generators = Vec::new();
        while self.check(&TokenKind::For) || self.check(&TokenKind::Async) {
            let is_async = self.eat(&TokenKind::Async);
            self.expect_kind(&TokenKind::For)?;
            let target = self.parse_target_list()?;
            self.expect_kind(&TokenKind::In)?;
            let iter = self.parse_or()?;
            let mut ifs = Vec::new();
            while self.check(&TokenKind::If) {
                self.advance();
                ifs.push(self.parse_or()?);
            }
            generators.push(Comprehension { target, iter, ifs, is_async });
        }
        Ok(generators)
    }

    // ── Target list (for assignments, for loops) ─────────────────────

    fn parse_target_list(&mut self) -> Result<Expression, String> {
        let first = self.parse_target()?;
        if self.check(&TokenKind::Comma) && !self.check_ahead_is(&[TokenKind::In, TokenKind::Eq]) {
            let mut targets = vec![first];
            while self.eat(&TokenKind::Comma) {
                if matches!(self.current_kind(), TokenKind::In | TokenKind::Eq | TokenKind::Colon | TokenKind::Newline | TokenKind::RParen) {
                    break;
                }
                targets.push(self.parse_target()?);
            }
            if targets.len() == 1 {
                Ok(targets.into_iter().next().unwrap())
            } else {
                Ok(Expression::Tuple(targets))
            }
        } else {
            Ok(first)
        }
    }

    fn parse_target(&mut self) -> Result<Expression, String> {
        if self.check(&TokenKind::Star) {
            self.advance();
            let expr = self.parse_postfix()?;
            Ok(Expression::Starred(Box::new(expr)))
        } else if self.check(&TokenKind::LParen) {
            self.advance();
            let target = self.parse_target_list()?;
            self.expect_kind(&TokenKind::RParen)?;
            Ok(target)
        } else if self.check(&TokenKind::LBracket) {
            self.advance();
            let mut targets = Vec::new();
            while !self.check(&TokenKind::RBracket) && !self.is_at_end() {
                targets.push(self.parse_target()?);
                if !self.eat(&TokenKind::Comma) { break; }
            }
            self.expect_kind(&TokenKind::RBracket)?;
            Ok(Expression::List(targets))
        } else {
            self.parse_postfix()
        }
    }

    // ── Dotted name (for imports) ────────────────────────────────────

    fn parse_dotted_name(&mut self) -> Result<String, String> {
        let mut name = self.expect_identifier()?;
        while self.eat(&TokenKind::Dot) {
            name.push('.');
            name.push_str(&self.expect_identifier()?);
        }
        Ok(name)
    }

    // ── Helpers ──────────────────────────────────────────────────────

    fn current_kind(&self) -> TokenKind {
        if self.pos < self.tokens.len() {
            self.tokens[self.pos].kind.clone()
        } else {
            TokenKind::Eof
        }
    }

    fn current_line(&self) -> u32 {
        if self.pos < self.tokens.len() { self.tokens[self.pos].line } else { 0 }
    }

    fn check(&self, kind: &TokenKind) -> bool {
        std::mem::discriminant(&self.current_kind()) == std::mem::discriminant(kind)
    }

    fn check_identifier(&self) -> bool {
        matches!(self.current_kind(), TokenKind::Identifier(_))
    }

    fn peek_kind(&self, offset: usize) -> Option<&TokenKind> {
        self.tokens.get(self.pos + offset).map(|t| &t.kind)
    }

    fn check_ahead_is(&self, kinds: &[TokenKind]) -> bool {
        if let Some(next) = self.peek_kind(1) {
            kinds.iter().any(|k| std::mem::discriminant(next) == std::mem::discriminant(k))
        } else {
            false
        }
    }

    fn advance(&mut self) -> &Token {
        let tok = &self.tokens[self.pos.min(self.tokens.len() - 1)];
        if self.pos < self.tokens.len() { self.pos += 1; }
        tok
    }

    fn eat(&mut self, kind: &TokenKind) -> bool {
        if self.check(kind) {
            self.advance();
            true
        } else {
            false
        }
    }

    fn eat_kind(&mut self, kind: &TokenKind) -> bool {
        self.eat(kind)
    }

    fn expect_kind(&mut self, kind: &TokenKind) -> Result<(), String> {
        if self.check(kind) {
            self.advance();
            Ok(())
        } else {
            Err(self.error(&format!("expected {:?}, got {:?}", kind, self.current_kind())))
        }
    }

    fn expect_identifier(&mut self) -> Result<String, String> {
        match self.current_kind() {
            TokenKind::Identifier(name) => { self.advance(); Ok(name) }
            // Allow soft keywords as identifiers
            TokenKind::Match => { self.advance(); Ok("match".to_string()) }
            TokenKind::Case => { self.advance(); Ok("case".to_string()) }
            _ => Err(self.error(&format!("expected identifier, got {:?}", self.current_kind()))),
        }
    }

    fn expect_newline(&mut self) -> Result<(), String> {
        if self.check(&TokenKind::Newline) || self.check(&TokenKind::Eof) || self.check(&TokenKind::Semicolon) {
            self.eat(&TokenKind::Newline);
            self.eat(&TokenKind::Semicolon);
            Ok(())
        } else {
            Err(self.error(&format!("expected newline, got {:?}", self.current_kind())))
        }
    }

    fn is_at_end(&self) -> bool {
        self.pos >= self.tokens.len() || self.current_kind() == TokenKind::Eof
    }

    fn error(&self, msg: &str) -> String {
        format!("line {}: {}", self.current_line(), msg)
    }
}

// Flatten chained BoolOps: (a and b) and c → BoolOp { And, [a, b, c] }
fn flatten_bool_op(op: BoolOp, left: Expression, right: Expression) -> Vec<Expression> {
    let mut values = match left {
        Expression::BoolOp { op: lop, values } if lop == op => values,
        other => vec![other],
    };
    values.push(right);
    values
}
