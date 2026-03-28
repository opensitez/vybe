use crate::ast::*;
use crate::token::{Token, TokenKind};

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(tokens: Vec<Token>) -> Self {
        Parser { tokens, pos: 0 }
    }

    pub fn parse(&mut self) -> Result<CompilationUnit, String> {
        let mut usings = Vec::new();
        let mut namespace = None;
        let mut members = Vec::new();
        let mut top_level_statements = Vec::new();

        while !self.at_end() {
            match self.current() {
                TokenKind::Using => {
                    self.advance();
                    let mut parts = vec![self.expect_ident()?];
                    while self.eat(TokenKind::Dot) {
                        parts.push(self.expect_ident()?);
                    }
                    self.expect(TokenKind::Semicolon)?;
                    usings.push(parts.join("."));
                }
                TokenKind::Namespace => {
                    self.advance();
                    let mut parts = vec![self.expect_ident()?];
                    while self.eat(TokenKind::Dot) {
                        parts.push(self.expect_ident()?);
                    }
                    namespace = Some(parts.join("."));
                    if self.eat(TokenKind::LBrace) {
                        while !self.check(TokenKind::RBrace) && !self.at_end() {
                            members.push(self.parse_type_decl()?);
                        }
                        self.expect(TokenKind::RBrace)?;
                    } else {
                        self.expect(TokenKind::Semicolon)?;
                        while !self.at_end() {
                            members.push(self.parse_type_decl()?);
                        }
                    }
                }
                TokenKind::LBracket => {
                    // Attribute — skip [...]
                    self.advance();
                    let mut depth = 1;
                    while depth > 0 && !self.at_end() {
                        if self.check(TokenKind::LBracket) { depth += 1; }
                        if self.check(TokenKind::RBracket) { depth -= 1; }
                        self.advance();
                    }
                }
                TokenKind::Public | TokenKind::Private | TokenKind::Protected | TokenKind::Internal
                | TokenKind::Static | TokenKind::Abstract | TokenKind::Partial | TokenKind::Sealed
                | TokenKind::Class | TokenKind::Struct | TokenKind::Interface | TokenKind::Enum => {
                    members.push(self.parse_type_decl()?);
                }
                _ => {
                    // Top-level statement (C# 9+)
                    top_level_statements.push(self.parse_statement()?);
                }
            }
        }

        Ok(CompilationUnit { usings, namespace, members, top_level_statements })
    }

    // ================================================================
    // Type declarations
    // ================================================================

    fn parse_type_decl(&mut self) -> Result<TypeDecl, String> {
        let mut access = Access::Internal;
        let mut is_static = false;
        let mut is_abstract = false;
        let mut is_partial = false;
        let mut is_sealed = false;

        // Parse modifiers
        loop {
            match self.current() {
                TokenKind::Public => { access = Access::Public; self.advance(); }
                TokenKind::Private => { access = Access::Private; self.advance(); }
                TokenKind::Protected => { access = Access::Protected; self.advance(); }
                TokenKind::Internal => { access = Access::Internal; self.advance(); }
                TokenKind::Static => { is_static = true; self.advance(); }
                TokenKind::Abstract => { is_abstract = true; self.advance(); }
                TokenKind::Partial => { is_partial = true; self.advance(); }
                TokenKind::Sealed => { is_sealed = true; self.advance(); }
                _ => break,
            }
        }

        match self.current() {
            TokenKind::Class => {
                self.advance();
                let name = self.expect_ident()?;
                // Skip generic params <T>
                if self.eat(TokenKind::Lt) {
                    let mut depth = 1;
                    while depth > 0 && !self.at_end() {
                        if self.check(TokenKind::Lt) { depth += 1; }
                        if self.check(TokenKind::Gt) { depth -= 1; }
                        self.advance();
                    }
                }
                let mut base_type = None;
                let mut interfaces = Vec::new();
                if self.eat(TokenKind::Colon) {
                    let first = self.parse_type_name()?;
                    // First could be base class or interface
                    base_type = Some(first);
                    while self.eat(TokenKind::Comma) {
                        interfaces.push(self.parse_type_name()?);
                    }
                }
                self.expect(TokenKind::LBrace)?;
                let members = self.parse_class_body()?;
                self.expect(TokenKind::RBrace)?;
                Ok(TypeDecl::Class(ClassDecl {
                    name, is_partial, is_static, is_abstract,
                    base_type, interfaces, members,
                }))
            }
            TokenKind::Struct => {
                self.advance();
                let name = self.expect_ident()?;
                self.expect(TokenKind::LBrace)?;
                let members = self.parse_class_body()?;
                self.expect(TokenKind::RBrace)?;
                Ok(TypeDecl::Struct(StructDecl { name, members }))
            }
            TokenKind::Enum => {
                self.advance();
                let name = self.expect_ident()?;
                self.expect(TokenKind::LBrace)?;
                let mut members = Vec::new();
                while !self.check(TokenKind::RBrace) && !self.at_end() {
                    let member_name = self.expect_ident()?;
                    let value = if self.eat(TokenKind::Assign) {
                        Some(self.parse_expression()?)
                    } else { None };
                    members.push((member_name, value));
                    self.eat(TokenKind::Comma);
                }
                self.expect(TokenKind::RBrace)?;
                Ok(TypeDecl::Enum(EnumDecl { name, members }))
            }
            TokenKind::Interface => {
                self.advance();
                let name = self.expect_ident()?;
                self.expect(TokenKind::LBrace)?;
                let members = self.parse_class_body()?;
                self.expect(TokenKind::RBrace)?;
                Ok(TypeDecl::Interface(InterfaceDecl { name, members }))
            }
            _ => Err(format!("Expected class/struct/enum/interface, got {:?} at line {}", self.current(), self.line())),
        }
    }

    fn parse_class_body(&mut self) -> Result<Vec<MemberDecl>, String> {
        let mut members = Vec::new();
        while !self.check(TokenKind::RBrace) && !self.at_end() {
            // Skip attributes
            if self.check(TokenKind::LBracket) {
                self.advance();
                let mut depth = 1;
                while depth > 0 && !self.at_end() {
                    if self.check(TokenKind::LBracket) { depth += 1; }
                    if self.check(TokenKind::RBracket) { depth -= 1; }
                    self.advance();
                }
                continue;
            }
            members.push(self.parse_member_decl()?);
        }
        Ok(members)
    }

    fn parse_member_decl(&mut self) -> Result<MemberDecl, String> {
        let mut access = Access::Private;
        let mut is_static = false;
        let mut is_override = false;
        let mut is_virtual = false;
        let mut is_abstract = false;
        let mut is_async = false;
        let mut is_readonly = false;
        let mut is_const = false;
        let mut is_event = false;

        loop {
            match self.current() {
                TokenKind::Public => { access = Access::Public; self.advance(); }
                TokenKind::Private => { access = Access::Private; self.advance(); }
                TokenKind::Protected => { access = Access::Protected; self.advance(); }
                TokenKind::Internal => { access = Access::Internal; self.advance(); }
                TokenKind::Static => { is_static = true; self.advance(); }
                TokenKind::Override => { is_override = true; self.advance(); }
                TokenKind::Virtual => { is_virtual = true; self.advance(); }
                TokenKind::Abstract => { is_abstract = true; self.advance(); }
                TokenKind::Async => { is_async = true; self.advance(); }
                TokenKind::Readonly => { is_readonly = true; self.advance(); }
                TokenKind::Const => { is_const = true; self.advance(); }
                TokenKind::Event => { is_event = true; self.advance(); }
                TokenKind::Sealed => { self.advance(); }
                TokenKind::New => {
                    // 'new' modifier (hides base member) — skip
                    if self.peek_is_type_or_ident() { self.advance(); }
                    else { break; }
                }
                _ => break,
            }
        }

        // Constructor: ClassName(params) { body }
        // Check if current token is the class name followed by (
        if let TokenKind::Identifier(ref name) = self.current() {
            let name = name.clone();
            if self.peek_at(1) == Some(&TokenKind::LParen) && !self.is_type_name_next() {
                self.advance(); // skip name
                let params = self.parse_params()?;
                let base_args = if self.eat(TokenKind::Colon) {
                    if self.eat(TokenKind::Base) || self.eat(TokenKind::This) {
                        self.expect(TokenKind::LParen)?;
                        let args = self.parse_args()?;
                        self.expect(TokenKind::RParen)?;
                        Some(args)
                    } else { None }
                } else { None };
                let body = self.parse_block()?;
                return Ok(MemberDecl::Constructor(ConstructorDecl { params, body, base_args, access }));
            }
        }

        // Event declaration
        if is_event {
            let type_name = self.parse_type_name()?;
            let name = self.expect_ident()?;
            self.expect(TokenKind::Semicolon)?;
            return Ok(MemberDecl::Event { name, type_name, access });
        }

        // Type + name — could be field, method, or property
        let return_type = self.parse_type_name()?;
        let name = self.expect_ident()?;

        // Method: name(params) { body }
        if self.check(TokenKind::LParen) {
            let params = self.parse_params()?;
            let body = if is_abstract || self.eat(TokenKind::Semicolon) {
                Vec::new()
            } else {
                self.parse_block()?
            };
            return Ok(MemberDecl::Method(MethodDecl {
                name, return_type: Some(return_type), params, body,
                is_static, is_override, is_virtual, is_abstract, is_async, access,
            }));
        }

        // Property: name { get; set; } or name { get { ... } set { ... } }
        if self.check(TokenKind::LBrace) {
            self.advance();
            let mut getter = None;
            let mut setter = None;
            let mut is_auto = false;
            while !self.check(TokenKind::RBrace) && !self.at_end() {
                // Skip access modifiers on get/set
                if matches!(self.current(), TokenKind::Public | TokenKind::Private | TokenKind::Protected) {
                    self.advance();
                }
                if let TokenKind::Identifier(ref kw) = self.current() {
                    match kw.as_str() {
                        "get" => {
                            self.advance();
                            if self.eat(TokenKind::Semicolon) {
                                is_auto = true;
                                getter = Some(Vec::new());
                            } else {
                                getter = Some(self.parse_block()?);
                            }
                        }
                        "set" => {
                            self.advance();
                            if self.eat(TokenKind::Semicolon) {
                                is_auto = true;
                                setter = Some(("value".into(), Vec::new()));
                            } else {
                                let body = self.parse_block()?;
                                setter = Some(("value".into(), body));
                            }
                        }
                        _ => { self.advance(); }
                    }
                } else {
                    self.advance();
                }
            }
            self.expect(TokenKind::RBrace)?;
            // Auto-property initializer: = value;
            if self.eat(TokenKind::Assign) {
                let _init = self.parse_expression()?;
                self.expect(TokenKind::Semicolon)?;
            }
            return Ok(MemberDecl::Property(PropertyDecl {
                name, type_name: Some(return_type), getter, setter, is_auto, access,
            }));
        }

        // Field: type name = init;
        let initializer = if self.eat(TokenKind::Assign) {
            Some(self.parse_expression()?)
        } else { None };
        self.expect(TokenKind::Semicolon)?;
        Ok(MemberDecl::Field {
            name, type_name: Some(return_type), initializer,
            is_static, access,
        })
    }

    // ================================================================
    // Statements
    // ================================================================

    fn parse_statement(&mut self) -> Result<Statement, String> {
        match self.current() {
            TokenKind::LBrace => {
                let body = self.parse_block()?;
                Ok(Statement::Block(body))
            }
            TokenKind::If => self.parse_if(),
            TokenKind::For => self.parse_for(),
            TokenKind::ForEach => self.parse_foreach(),
            TokenKind::While => self.parse_while(),
            TokenKind::Do => self.parse_do_while(),
            TokenKind::Switch => self.parse_switch(),
            TokenKind::Return => {
                self.advance();
                let value = if !self.check(TokenKind::Semicolon) {
                    Some(self.parse_expression()?)
                } else { None };
                self.expect(TokenKind::Semicolon)?;
                Ok(Statement::Return(value))
            }
            TokenKind::Break => { self.advance(); self.expect(TokenKind::Semicolon)?; Ok(Statement::Break) }
            TokenKind::Continue => { self.advance(); self.expect(TokenKind::Semicolon)?; Ok(Statement::Continue) }
            TokenKind::Throw => {
                self.advance();
                let expr = self.parse_expression()?;
                self.expect(TokenKind::Semicolon)?;
                Ok(Statement::Throw(expr))
            }
            TokenKind::Try => self.parse_try(),
            TokenKind::Var => self.parse_local_var_decl(),
            TokenKind::Semicolon => { self.advance(); Ok(Statement::Empty) }
            // Type keyword followed by identifier — local declaration
            TokenKind::Int | TokenKind::String_ | TokenKind::Double | TokenKind::Float
            | TokenKind::Bool | TokenKind::Char | TokenKind::Long | TokenKind::Byte | TokenKind::Object => {
                self.parse_typed_local_decl()
            }
            _ => {
                // Could be: expression statement, or typed local decl (e.g. MyClass x = ...)
                // Try expression first
                let expr = self.parse_expression()?;
                // Check for compound assignment
                let stmt = match self.current() {
                    TokenKind::Assign => {
                        self.advance();
                        let value = self.parse_expression()?;
                        Statement::Assignment { target: expr, value }
                    }
                    TokenKind::PlusAssign => { self.advance(); let v = self.parse_expression()?; Statement::CompoundAssignment { target: expr, op: CompoundOp::AddAssign, value: v } }
                    TokenKind::MinusAssign => { self.advance(); let v = self.parse_expression()?; Statement::CompoundAssignment { target: expr, op: CompoundOp::SubAssign, value: v } }
                    TokenKind::StarAssign => { self.advance(); let v = self.parse_expression()?; Statement::CompoundAssignment { target: expr, op: CompoundOp::MulAssign, value: v } }
                    TokenKind::SlashAssign => { self.advance(); let v = self.parse_expression()?; Statement::CompoundAssignment { target: expr, op: CompoundOp::DivAssign, value: v } }
                    _ => Statement::Expression(expr),
                };
                self.expect(TokenKind::Semicolon)?;
                Ok(stmt)
            }
        }
    }

    fn parse_block(&mut self) -> Result<Vec<Statement>, String> {
        self.expect(TokenKind::LBrace)?;
        let mut stmts = Vec::new();
        while !self.check(TokenKind::RBrace) && !self.at_end() {
            stmts.push(self.parse_statement()?);
        }
        self.expect(TokenKind::RBrace)?;
        Ok(stmts)
    }

    fn parse_if(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::If)?;
        self.expect(TokenKind::LParen)?;
        let condition = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        let then_body = if self.check(TokenKind::LBrace) {
            self.parse_block()?
        } else {
            vec![self.parse_statement()?]
        };
        let mut else_if = Vec::new();
        let mut else_body = None;
        while self.eat(TokenKind::Else) {
            if self.check(TokenKind::If) {
                self.advance();
                self.expect(TokenKind::LParen)?;
                let cond = self.parse_expression()?;
                self.expect(TokenKind::RParen)?;
                let body = if self.check(TokenKind::LBrace) { self.parse_block()? } else { vec![self.parse_statement()?] };
                else_if.push((cond, body));
            } else {
                else_body = Some(if self.check(TokenKind::LBrace) { self.parse_block()? } else { vec![self.parse_statement()?] });
                break;
            }
        }
        Ok(Statement::If { condition, then_body, else_if, else_body })
    }

    fn parse_for(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::For)?;
        self.expect(TokenKind::LParen)?;
        let init = if !self.check(TokenKind::Semicolon) {
            Some(Box::new(if self.check(TokenKind::Var) || self.is_type_keyword() {
                self.parse_local_var_decl_no_semi()?
            } else {
                let expr = self.parse_expression()?;
                if self.eat(TokenKind::Assign) {
                    let val = self.parse_expression()?;
                    Statement::Assignment { target: expr, value: val }
                } else {
                    Statement::Expression(expr)
                }
            }))
        } else { None };
        self.expect(TokenKind::Semicolon)?;
        let condition = if !self.check(TokenKind::Semicolon) { Some(self.parse_expression()?) } else { None };
        self.expect(TokenKind::Semicolon)?;
        let update = if !self.check(TokenKind::RParen) {
            let expr = self.parse_expression()?;
            Some(Box::new(Statement::Expression(expr)))
        } else { None };
        self.expect(TokenKind::RParen)?;
        let body = if self.check(TokenKind::LBrace) { self.parse_block()? } else { vec![self.parse_statement()?] };
        Ok(Statement::For { init, condition, update, body })
    }

    fn parse_foreach(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::ForEach)?;
        self.expect(TokenKind::LParen)?;
        // var x in collection  OR  Type x in collection
        if self.check(TokenKind::Var) { self.advance(); }
        else { self.parse_type_name()?; } // skip type
        let var_name = self.expect_ident()?;
        self.expect(TokenKind::In)?;
        let iterable = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        let body = if self.check(TokenKind::LBrace) { self.parse_block()? } else { vec![self.parse_statement()?] };
        Ok(Statement::ForEach { var_name, iterable, body })
    }

    fn parse_while(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::While)?;
        self.expect(TokenKind::LParen)?;
        let condition = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        let body = if self.check(TokenKind::LBrace) { self.parse_block()? } else { vec![self.parse_statement()?] };
        Ok(Statement::While { condition, body })
    }

    fn parse_do_while(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Do)?;
        let body = self.parse_block()?;
        self.expect(TokenKind::While)?;
        self.expect(TokenKind::LParen)?;
        let condition = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        self.expect(TokenKind::Semicolon)?;
        Ok(Statement::DoWhile { body, condition })
    }

    fn parse_switch(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Switch)?;
        self.expect(TokenKind::LParen)?;
        let expr = self.parse_expression()?;
        self.expect(TokenKind::RParen)?;
        self.expect(TokenKind::LBrace)?;
        let mut cases = Vec::new();
        while !self.check(TokenKind::RBrace) && !self.at_end() {
            let mut labels = Vec::new();
            while self.check(TokenKind::Case) || self.check(TokenKind::Default) {
                if self.eat(TokenKind::Case) {
                    labels.push(Some(self.parse_expression()?));
                    self.expect(TokenKind::Colon)?;
                } else {
                    self.advance(); // default
                    self.expect(TokenKind::Colon)?;
                    labels.push(None);
                }
            }
            let mut body = Vec::new();
            while !self.check(TokenKind::Case) && !self.check(TokenKind::Default) && !self.check(TokenKind::RBrace) && !self.at_end() {
                body.push(self.parse_statement()?);
            }
            cases.push(SwitchCase { labels, body });
        }
        self.expect(TokenKind::RBrace)?;
        Ok(Statement::Switch { expr, cases })
    }

    fn parse_try(&mut self) -> Result<Statement, String> {
        self.expect(TokenKind::Try)?;
        let try_body = self.parse_block()?;
        let mut catches = Vec::new();
        while self.eat(TokenKind::Catch) {
            let (type_name, var_name) = if self.eat(TokenKind::LParen) {
                let tn = self.parse_type_name()?;
                let vn = if !self.check(TokenKind::RParen) { Some(self.expect_ident()?) } else { None };
                self.expect(TokenKind::RParen)?;
                (Some(tn), vn)
            } else { (None, None) };
            let body = self.parse_block()?;
            catches.push(CatchClause { type_name, var_name, body });
        }
        let finally_body = if self.eat(TokenKind::Finally) { Some(self.parse_block()?) } else { None };
        Ok(Statement::TryCatchFinally { try_body, catches, finally_body })
    }

    fn parse_local_var_decl(&mut self) -> Result<Statement, String> {
        let stmt = self.parse_local_var_decl_no_semi()?;
        self.expect(TokenKind::Semicolon)?;
        Ok(stmt)
    }

    fn parse_local_var_decl_no_semi(&mut self) -> Result<Statement, String> {
        let is_var = self.eat(TokenKind::Var);
        let type_name = if !is_var { Some(self.parse_type_name()?) } else { None };
        let name = self.expect_ident()?;
        let initializer = if self.eat(TokenKind::Assign) { Some(self.parse_expression()?) } else { None };
        Ok(Statement::LocalDecl { name, type_name, initializer, is_var })
    }

    fn parse_typed_local_decl(&mut self) -> Result<Statement, String> {
        let type_name = self.parse_type_name()?;
        let name = self.expect_ident()?;
        let initializer = if self.eat(TokenKind::Assign) { Some(self.parse_expression()?) } else { None };
        self.expect(TokenKind::Semicolon)?;
        Ok(Statement::LocalDecl { name, type_name: Some(type_name), initializer, is_var: false })
    }

    // ================================================================
    // Expressions (Pratt parser)
    // ================================================================

    fn parse_expression(&mut self) -> Result<Expression, String> {
        self.parse_ternary()
    }

    fn parse_ternary(&mut self) -> Result<Expression, String> {
        let expr = self.parse_null_coalescing()?;
        if self.eat(TokenKind::Question) {
            let then_expr = self.parse_expression()?;
            self.expect(TokenKind::Colon)?;
            let else_expr = self.parse_expression()?;
            Ok(Expression::Conditional(Box::new(expr), Box::new(then_expr), Box::new(else_expr)))
        } else {
            Ok(expr)
        }
    }

    fn parse_null_coalescing(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_or()?;
        while self.eat(TokenKind::QuestionQuestion) {
            let right = self.parse_or()?;
            expr = Expression::NullCoalescing(Box::new(expr), Box::new(right));
        }
        Ok(expr)
    }

    fn parse_or(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_and()?;
        while self.eat(TokenKind::Or) {
            let right = self.parse_and()?;
            expr = Expression::Binary(BinaryOp::Or, Box::new(expr), Box::new(right));
        }
        Ok(expr)
    }

    fn parse_and(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_equality()?;
        while self.eat(TokenKind::And) {
            let right = self.parse_equality()?;
            expr = Expression::Binary(BinaryOp::And, Box::new(expr), Box::new(right));
        }
        Ok(expr)
    }

    fn parse_equality(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_relational()?;
        loop {
            if self.eat(TokenKind::Eq) {
                let r = self.parse_relational()?;
                expr = Expression::Binary(BinaryOp::Eq, Box::new(expr), Box::new(r));
            } else if self.eat(TokenKind::Neq) {
                let r = self.parse_relational()?;
                expr = Expression::Binary(BinaryOp::Neq, Box::new(expr), Box::new(r));
            } else { break; }
        }
        Ok(expr)
    }

    fn parse_relational(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_additive()?;
        loop {
            if self.eat(TokenKind::Lt) { let r = self.parse_additive()?; expr = Expression::Binary(BinaryOp::Lt, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Gt) { let r = self.parse_additive()?; expr = Expression::Binary(BinaryOp::Gt, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Le) { let r = self.parse_additive()?; expr = Expression::Binary(BinaryOp::Le, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Ge) { let r = self.parse_additive()?; expr = Expression::Binary(BinaryOp::Ge, Box::new(expr), Box::new(r)); }
            else if self.check(TokenKind::Is) { self.advance(); let tn = self.parse_type_name()?; expr = Expression::Is(Box::new(expr), tn); }
            else if self.check(TokenKind::As) { self.advance(); let tn = self.parse_type_name()?; expr = Expression::As(Box::new(expr), tn); }
            else { break; }
        }
        Ok(expr)
    }

    fn parse_additive(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_multiplicative()?;
        loop {
            if self.eat(TokenKind::Plus) { let r = self.parse_multiplicative()?; expr = Expression::Binary(BinaryOp::Add, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Minus) { let r = self.parse_multiplicative()?; expr = Expression::Binary(BinaryOp::Sub, Box::new(expr), Box::new(r)); }
            else { break; }
        }
        Ok(expr)
    }

    fn parse_multiplicative(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_unary()?;
        loop {
            if self.eat(TokenKind::Star) { let r = self.parse_unary()?; expr = Expression::Binary(BinaryOp::Mul, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Slash) { let r = self.parse_unary()?; expr = Expression::Binary(BinaryOp::Div, Box::new(expr), Box::new(r)); }
            else if self.eat(TokenKind::Percent) { let r = self.parse_unary()?; expr = Expression::Binary(BinaryOp::Mod, Box::new(expr), Box::new(r)); }
            else { break; }
        }
        Ok(expr)
    }

    fn parse_unary(&mut self) -> Result<Expression, String> {
        if self.eat(TokenKind::Minus) { let e = self.parse_unary()?; return Ok(Expression::Unary(UnaryOp::Neg, Box::new(e))); }
        if self.eat(TokenKind::Not) { let e = self.parse_unary()?; return Ok(Expression::Unary(UnaryOp::Not, Box::new(e))); }
        if self.eat(TokenKind::Tilde) { let e = self.parse_unary()?; return Ok(Expression::Unary(UnaryOp::BitNot, Box::new(e))); }
        if self.eat(TokenKind::Increment) { let e = self.parse_unary()?; return Ok(Expression::PreIncrement(Box::new(e))); }
        if self.eat(TokenKind::Decrement) { let e = self.parse_unary()?; return Ok(Expression::PreDecrement(Box::new(e))); }
        if self.check(TokenKind::Await) { self.advance(); let e = self.parse_unary()?; return Ok(Expression::Await(Box::new(e))); }
        self.parse_postfix()
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            if self.eat(TokenKind::Dot) {
                let member = self.expect_ident()?;
                if self.check(TokenKind::LParen) {
                    // Method call
                    self.advance();
                    let args = self.parse_args()?;
                    self.expect(TokenKind::RParen)?;
                    expr = Expression::Call(
                        Box::new(Expression::MemberAccess(Box::new(expr), member)),
                        args,
                    );
                } else {
                    expr = Expression::MemberAccess(Box::new(expr), member);
                }
            } else if self.eat(TokenKind::QuestionDot) {
                let member = self.expect_ident()?;
                expr = Expression::NullConditionalAccess(Box::new(expr), member);
            } else if self.eat(TokenKind::LBracket) {
                let index = self.parse_expression()?;
                self.expect(TokenKind::RBracket)?;
                expr = Expression::Index(Box::new(expr), Box::new(index));
            } else if self.check(TokenKind::LParen) && self.is_callable(&expr) {
                self.advance();
                let args = self.parse_args()?;
                self.expect(TokenKind::RParen)?;
                expr = Expression::Call(Box::new(expr), args);
            } else if self.eat(TokenKind::Increment) {
                expr = Expression::PostIncrement(Box::new(expr));
            } else if self.eat(TokenKind::Decrement) {
                expr = Expression::PostDecrement(Box::new(expr));
            } else {
                break;
            }
        }
        Ok(expr)
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.current() {
            TokenKind::IntLit(n) => { let n = n; self.advance(); Ok(Expression::IntLiteral(n)) }
            TokenKind::DoubleLit(n) => { let n = n; self.advance(); Ok(Expression::DoubleLiteral(n)) }
            TokenKind::StringLit(ref s) => { let s = s.clone(); self.advance(); Ok(Expression::StringLiteral(s)) }
            TokenKind::CharLit(c) => { let c = c; self.advance(); Ok(Expression::CharLiteral(c)) }
            TokenKind::True => { self.advance(); Ok(Expression::BoolLiteral(true)) }
            TokenKind::False => { self.advance(); Ok(Expression::BoolLiteral(false)) }
            TokenKind::Null => { self.advance(); Ok(Expression::NullLiteral) }
            TokenKind::This => { self.advance(); Ok(Expression::This) }
            TokenKind::Base => { self.advance(); Ok(Expression::Base) }
            TokenKind::New => {
                self.advance();
                let type_name = self.parse_type_name()?;
                if self.eat(TokenKind::LBracket) {
                    // new Type[size]
                    let size = self.parse_expression()?;
                    self.expect(TokenKind::RBracket)?;
                    Ok(Expression::NewArray(type_name, Box::new(size)))
                } else if self.eat(TokenKind::LParen) {
                    let args = self.parse_args()?;
                    self.expect(TokenKind::RParen)?;
                    // Check for object initializer { prop = val, ... }
                    if self.check(TokenKind::LBrace) {
                        self.advance();
                        let mut inits = Vec::new();
                        while !self.check(TokenKind::RBrace) && !self.at_end() {
                            let prop = self.expect_ident()?;
                            self.expect(TokenKind::Assign)?;
                            let val = self.parse_expression()?;
                            inits.push((prop, val));
                            self.eat(TokenKind::Comma);
                        }
                        self.expect(TokenKind::RBrace)?;
                        Ok(Expression::ObjectInit(
                            Box::new(Expression::New(type_name, args)),
                            inits,
                        ))
                    } else {
                        Ok(Expression::New(type_name, args))
                    }
                } else if self.check(TokenKind::LBrace) {
                    // new Type { items }
                    self.advance();
                    let mut items = Vec::new();
                    while !self.check(TokenKind::RBrace) && !self.at_end() {
                        items.push(self.parse_expression()?);
                        self.eat(TokenKind::Comma);
                    }
                    self.expect(TokenKind::RBrace)?;
                    Ok(Expression::ArrayInit(items))
                } else {
                    Ok(Expression::New(type_name, Vec::new()))
                }
            }
            TokenKind::TypeOf => {
                self.advance();
                self.expect(TokenKind::LParen)?;
                let tn = self.parse_type_name()?;
                self.expect(TokenKind::RParen)?;
                Ok(Expression::TypeOf(tn))
            }
            TokenKind::NameOf => {
                self.advance();
                self.expect(TokenKind::LParen)?;
                let name = self.expect_ident()?;
                self.expect(TokenKind::RParen)?;
                Ok(Expression::NameOf(name))
            }
            TokenKind::Default => {
                self.advance();
                if self.eat(TokenKind::LParen) {
                    let tn = self.parse_type_name()?;
                    self.expect(TokenKind::RParen)?;
                    Ok(Expression::Default(Some(tn)))
                } else {
                    Ok(Expression::Default(None))
                }
            }
            TokenKind::LParen => {
                self.advance();
                // Could be cast: (Type)expr or grouping: (expr)
                // Try grouping
                let expr = self.parse_expression()?;
                self.expect(TokenKind::RParen)?;
                Ok(expr)
            }
            TokenKind::Identifier(ref name) => {
                let name = name.clone();
                self.advance();
                // Lambda: (x, y) => expr  or  x => expr
                if self.check(TokenKind::Arrow) {
                    self.advance();
                    if self.check(TokenKind::LBrace) {
                        let body = self.parse_block()?;
                        return Ok(Expression::LambdaBlock(vec![name], body));
                    } else {
                        let expr = self.parse_expression()?;
                        return Ok(Expression::Lambda(vec![name], Box::new(expr)));
                    }
                }
                Ok(Expression::Identifier(name))
            }
            _ => Err(format!("Unexpected token {:?} at line {}", self.current(), self.line())),
        }
    }

    // ================================================================
    // Helpers
    // ================================================================

    fn parse_params(&mut self) -> Result<Vec<Parameter>, String> {
        self.expect(TokenKind::LParen)?;
        let mut params = Vec::new();
        while !self.check(TokenKind::RParen) && !self.at_end() {
            let mut is_params = false;
            let mut is_ref = false;
            let mut is_out = false;
            if self.eat(TokenKind::Params) { is_params = true; }
            if self.eat(TokenKind::Ref) { is_ref = true; }
            if self.eat(TokenKind::Out) { is_out = true; }
            let type_name = Some(self.parse_type_name()?);
            let name = self.expect_ident()?;
            let default = if self.eat(TokenKind::Assign) { Some(self.parse_expression()?) } else { None };
            params.push(Parameter { name, type_name, default, is_params, is_ref, is_out });
            self.eat(TokenKind::Comma);
        }
        self.expect(TokenKind::RParen)?;
        Ok(params)
    }

    fn parse_args(&mut self) -> Result<Vec<Expression>, String> {
        let mut args = Vec::new();
        while !self.check(TokenKind::RParen) && !self.at_end() {
            // Skip ref/out keywords in args
            self.eat(TokenKind::Ref);
            self.eat(TokenKind::Out);
            args.push(self.parse_expression()?);
            self.eat(TokenKind::Comma);
        }
        Ok(args)
    }

    fn parse_type_name(&mut self) -> Result<String, String> {
        let mut name = match self.current() {
            TokenKind::Int => { self.advance(); "int".into() }
            TokenKind::String_ => { self.advance(); "string".into() }
            TokenKind::Double => { self.advance(); "double".into() }
            TokenKind::Float => { self.advance(); "float".into() }
            TokenKind::Bool => { self.advance(); "bool".into() }
            TokenKind::Char => { self.advance(); "char".into() }
            TokenKind::Long => { self.advance(); "long".into() }
            TokenKind::Byte => { self.advance(); "byte".into() }
            TokenKind::Object => { self.advance(); "object".into() }
            TokenKind::Void => { self.advance(); "void".into() }
            TokenKind::Identifier(ref s) => { let s = s.clone(); self.advance(); s }
            _ => return Err(format!("Expected type name, got {:?} at line {}", self.current(), self.line())),
        };
        // Dotted names: System.Windows.Forms.Button
        while self.eat(TokenKind::Dot) {
            name.push('.');
            name.push_str(&self.expect_ident()?);
        }
        // Generic args: List<string>
        if self.eat(TokenKind::Lt) {
            name.push('<');
            name.push_str(&self.parse_type_name()?);
            while self.eat(TokenKind::Comma) {
                name.push(',');
                name.push_str(&self.parse_type_name()?);
            }
            self.expect(TokenKind::Gt)?;
            name.push('>');
        }
        // Array: string[]
        if self.eat(TokenKind::LBracket) {
            self.expect(TokenKind::RBracket)?;
            name.push_str("[]");
        }
        // Nullable: int?
        if self.eat(TokenKind::Question) {
            name.push('?');
        }
        Ok(name)
    }

    // ================================================================
    // Token helpers
    // ================================================================

    fn current(&self) -> TokenKind {
        self.tokens.get(self.pos).map(|t| t.kind.clone()).unwrap_or(TokenKind::Eof)
    }

    fn line(&self) -> u32 {
        self.tokens.get(self.pos).map(|t| t.line).unwrap_or(0)
    }

    fn advance(&mut self) { if self.pos < self.tokens.len() { self.pos += 1; } }

    fn at_end(&self) -> bool { matches!(self.current(), TokenKind::Eof) }

    fn check(&self, kind: TokenKind) -> bool {
        std::mem::discriminant(&self.current()) == std::mem::discriminant(&kind)
    }

    fn eat(&mut self, kind: TokenKind) -> bool {
        if self.check(kind) { self.advance(); true } else { false }
    }

    fn expect(&mut self, kind: TokenKind) -> Result<(), String> {
        if self.check(kind.clone()) { self.advance(); Ok(()) }
        else { Err(format!("Expected {:?}, got {:?} at line {}", kind, self.current(), self.line())) }
    }

    fn expect_ident(&mut self) -> Result<String, String> {
        if let TokenKind::Identifier(name) = self.current() {
            self.advance();
            Ok(name)
        } else {
            Err(format!("Expected identifier, got {:?} at line {}", self.current(), self.line()))
        }
    }

    fn peek_at(&self, offset: usize) -> Option<&TokenKind> {
        self.tokens.get(self.pos + offset).map(|t| &t.kind)
    }

    fn peek_is_type_or_ident(&self) -> bool {
        matches!(self.peek_at(1), Some(TokenKind::Identifier(_) | TokenKind::Int | TokenKind::String_
            | TokenKind::Double | TokenKind::Bool | TokenKind::Void | TokenKind::Object))
    }

    fn is_type_name_next(&self) -> bool {
        // Check if the next token after ident is a type keyword or another ident (type name)
        // This distinguishes Constructor() from ReturnType Method()
        matches!(self.peek_at(2), Some(TokenKind::Identifier(_)))
    }

    fn is_type_keyword(&self) -> bool {
        matches!(self.current(), TokenKind::Int | TokenKind::String_ | TokenKind::Double
            | TokenKind::Float | TokenKind::Bool | TokenKind::Char | TokenKind::Long
            | TokenKind::Byte | TokenKind::Object)
    }

    fn is_callable(&self, expr: &Expression) -> bool {
        matches!(expr, Expression::Identifier(_) | Expression::MemberAccess(_, _))
    }
}
