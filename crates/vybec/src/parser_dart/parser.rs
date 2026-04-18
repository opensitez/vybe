use super::ast::*;
use super::lexer::Lexer;
use super::token::Token;

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(input: &str) -> Self {
        let mut lexer = Lexer::new(input);
        let mut tokens = Vec::new();
        loop {
            let t = lexer.next_token();
            let done = t == Token::EOF;
            tokens.push(t);
            if done { break; }
        }
        Self { tokens, pos: 0 }
    }

    fn peek(&self) -> &Token { self.tokens.get(self.pos).unwrap_or(&Token::EOF) }
    fn peek2(&self) -> &Token { self.tokens.get(self.pos + 1).unwrap_or(&Token::EOF) }

    fn advance(&mut self) -> Token {
        let t = self.tokens.get(self.pos).cloned().unwrap_or(Token::EOF);
        if self.pos < self.tokens.len() { self.pos += 1; }
        t
    }

    fn eat(&mut self, t: &Token) -> bool {
        if self.peek() == t { self.advance(); true } else { false }
    }

    fn expect(&mut self, t: Token) -> Result<(), String> {
        if self.peek() == &t { self.advance(); Ok(()) }
        else { Err(format!("Expected {:?}, got {:?}", t, self.peek())) }
    }

    /// Parse operator name after the `operator` keyword.
    /// Returns the operator symbol as a string: "+", "-", "*", "[]", "[]=", "==", "<", etc.
    fn parse_operator_name(&mut self) -> Result<String, String> {
        let tok = self.advance();
        let name = match tok {
            Token::Plus => "+",
            Token::Minus => "-",
            Token::Star => "*",
            Token::Slash => "/",
            Token::Percent => "%",
            Token::TildeSlash => "~/",
            Token::EqEq => "==",
            Token::BangEq => "!=",
            Token::Less => "<",
            Token::Greater => ">",
            Token::LessEq => "<=",
            Token::GreaterEq => ">=",
            Token::Amp => "&",
            Token::Bar => "|",
            Token::Caret => "^",
            Token::Tilde => "~",
            Token::LessLess => "<<",
            Token::GreaterGreater => ">>",
            Token::GreaterGreaterGreater => ">>>",
            Token::LBracket => {
                // operator[] or operator[]=
                self.expect(Token::RBracket)?;
                if self.eat(&Token::Eq) { return Ok("[]=".to_string()); }
                return Ok("[]".to_string());
            }
            _ => return Err(format!("Expected operator symbol after 'operator', got {:?}", tok)),
        };
        Ok(name.to_string())
    }

    fn expect_ident(&mut self) -> Result<String, String> {
        let t = self.advance();
        if let Some(s) = t.to_ident_str() {
            return Ok(s);
        }
        Err(format!("Expected identifier, got {:?}", t))
    }

    // ── Top level ─────────────────────────────────────────────────────────────

    pub fn parse_program(&mut self) -> Result<Program, String> {
        let mut body = Vec::new();
        while self.peek() != &Token::EOF {
            body.push(self.parse_top_level()?);
        }
        Ok(Program { body })
    }

    fn parse_top_level(&mut self) -> Result<TopLevel, String> {
        // Skip annotations
        while self.peek() == &Token::At {
            self.advance();
            self.expect_ident()?;
            if self.peek() == &Token::LParen { self.skip_balanced(Token::LParen, Token::RParen)?; }
        }
        match self.peek() {
            Token::Import => { self.advance(); Ok(TopLevel::Import(self.parse_import()?)) }
            Token::Class | Token::Abstract | Token::Mixin => Ok(TopLevel::Class(self.parse_class()?)),
            Token::Extension => Ok(TopLevel::Extension(self.parse_extension()?)),
            Token::Enum => { self.advance(); Ok(TopLevel::Enum(self.parse_enum()?)) }
            Token::Typedef => { self.advance(); Ok(TopLevel::Typedef(self.parse_typedef()?)) }
            _ => {
                // Try function or variable
                let stmt = self.parse_statement()?;
                match stmt {
                    Statement::FunctionDecl(f) => Ok(TopLevel::Function(f)),
                    Statement::VarDecl(v) => Ok(TopLevel::Variable(v)),
                    s => Ok(TopLevel::Statement(s)),
                }
            }
        }
    }

    fn parse_import(&mut self) -> Result<ImportDecl, String> {
        let uri = match self.advance() {
            Token::StringLiteral(s) => s,
            t => return Err(format!("Expected string after import, got {:?}", t)),
        };
        let mut prefix = None;
        let mut show = Vec::new();
        let mut hide = Vec::new();
        loop {
            match self.peek() {
                Token::As => { self.advance(); prefix = Some(self.expect_ident()?); }
                Token::Show => { self.advance(); show = self.parse_ident_list()?; }
                Token::Hide => { self.advance(); hide = self.parse_ident_list()?; }
                Token::Semicolon => { self.advance(); break; }
                _ => break,
            }
        }
        Ok(ImportDecl { uri, prefix, show, hide })
    }

    fn parse_typedef(&mut self) -> Result<TypedefDecl, String> {
        let name = self.expect_ident()?;
        self.expect(Token::Eq)?;
        let type_ann = self.parse_type_annotation()?;
        self.eat(&Token::Semicolon);
        Ok(TypedefDecl { name, type_ann })
    }

    fn parse_ident_list(&mut self) -> Result<Vec<String>, String> {
        let mut list = vec![self.expect_ident()?];
        while self.eat(&Token::Comma) { list.push(self.expect_ident()?); }
        Ok(list)
    }

    // ── Class ─────────────────────────────────────────────────────────────────

    fn parse_enum(&mut self) -> Result<EnumDecl, String> {
        let name = self.expect_ident()?;
        self.expect(Token::LBrace)?;
        let mut values = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            values.push(self.expect_ident()?);
            if !self.eat(&Token::Comma) { break; }
        }
        self.expect(Token::RBrace)?;
        Ok(EnumDecl { name, values })
    }

    fn parse_class(&mut self) -> Result<ClassDecl, String> {
        let is_abstract = self.eat(&Token::Abstract);
        // Accept both `class` and `mixin` (mixin compiles like a class)
        if !self.eat(&Token::Class) { self.eat(&Token::Mixin); }
        let name = self.expect_ident()?;
        let type_params = self.parse_type_params()?;
        let mut extends = None;
        let mut implements = Vec::new();
        let mut mixins = Vec::new();
        if self.eat(&Token::Extends) { extends = Some(self.expect_ident()?); self.skip_type_args(); }
        if self.eat(&Token::With) {
            loop { mixins.push(self.expect_ident()?); self.skip_type_args(); if !self.eat(&Token::Comma) { break; } }
        }
        if self.eat(&Token::Implements) {
            loop { implements.push(self.expect_ident()?); self.skip_type_args(); if !self.eat(&Token::Comma) { break; } }
        }
        self.expect(Token::LBrace)?;
        let mut members = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            // skip annotations
            while self.peek() == &Token::At {
                self.advance(); self.expect_ident()?;
                if self.peek() == &Token::LParen { self.skip_balanced(Token::LParen, Token::RParen)?; }
            }
            if self.peek() == &Token::RBrace { break; }
            members.extend(self.parse_class_member(&name)?);
        }
        self.expect(Token::RBrace)?;
        Ok(ClassDecl { name, type_params, extends, implements, mixins, is_abstract, members })
    }

    fn parse_extension(&mut self) -> Result<ExtensionDecl, String> {
        self.expect(Token::Extension)?;
        let name = if let Token::Identifier(_) = self.peek() {
            Some(self.expect_ident()?)
        } else {
            None
        };
        self.expect(Token::On)?;
        let on_type = self.parse_type_annotation()?;
        self.expect(Token::LBrace)?;
        let mut members = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            members.extend(self.parse_class_member(name.as_deref().unwrap_or("Extension"))?);
        }
        self.expect(Token::RBrace)?;
        Ok(ExtensionDecl { name, on_type, members })
    }


    fn parse_class_member(&mut self, class_name: &str) -> Result<Vec<ClassMember>, String> {
        let is_static = self.eat(&Token::Static);
        let is_abstract = self.eat(&Token::Abstract);
        let is_override = self.eat(&Token::Override);
        let is_const = self.eat(&Token::Const);
        let is_factory = self.eat(&Token::Factory);
        let is_final = self.eat(&Token::Final);
        let is_late = self.eat(&Token::Late);

        // getter/setter keyword
        let kind = if self.peek() == &Token::Get { self.advance(); MethodKind::Getter }
                   else if self.peek() == &Token::Set { self.advance(); MethodKind::Setter }
                   else { MethodKind::Method };

        // Constructor detection: ClassName( or ClassName.named(
        let is_ctor = match self.peek() {
            Token::Identifier(n) if n == class_name => {
                matches!(self.peek2(), Token::LParen | Token::Dot)
            }
            _ => false,
        };

        if is_ctor || is_factory {
            return Ok(vec![self.parse_constructor(class_name, is_const, is_factory)?]);
        }

        // Operator overloading: `operator +(Other o) { ... }` or `ReturnType operator +(Other o) { ... }`
        // Must check BEFORE type annotation parsing to avoid `operator` being consumed as a type.
        if self.peek() == &Token::Operator {
            self.advance();
            let op_name = self.parse_operator_name()?;
            let name = format!("operator{}", op_name);
            let params = self.parse_params()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params: vec![], params, return_type: None, body, is_async, is_generator: false };
            return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
        }

        // Getters: `get name { ... }` or `get name => expr;` — name immediately after get
        // Setters: `set name(param) { ... }` — name then params
        if kind == MethodKind::Getter {
            let return_type = self.try_parse_type_annotation();
            let name = self.expect_ident()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params: vec![], params: Params { positional: vec![], optional_pos: vec![], named: vec![] }, return_type, body, is_async, is_generator: false };
            return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
        }
        if kind == MethodKind::Setter {
            let return_type = self.try_parse_type_annotation();
            let name = self.expect_ident()?;
            let params = self.parse_params()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params: vec![], params, return_type, body, is_async, is_generator: false };
            return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
        }

        // Operator overloading: `ReturnType operator +(Other o) { ... }`
        // The `operator` keyword may appear after the return type annotation.
        // We check for it before or after type annotation.
        if self.peek() == &Token::Operator {
            self.advance(); // consume 'operator'
            let op_name = self.parse_operator_name()?;
            let name = format!("operator{}", op_name);
            let params = self.parse_params()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params: vec![], params, return_type: None, body, is_async, is_generator: false };
            return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
        }

        // Check for `ReturnType operator <op>(...)` pattern — peek is type name, peek2 is `operator`
        if self.peek2() == &Token::Operator {
            if let Token::Identifier(_) | Token::Void | Token::Dynamic = self.peek() {
                // Force-consume the return type identifier
                let saved = self.pos;
                let return_type = if let Ok(t) = self.parse_type_annotation() { Some(t) } else { self.pos = saved; None };
                if self.peek() == &Token::Operator {
                    self.advance();
                    let op_name = self.parse_operator_name()?;
                    let name = format!("operator{}", op_name);
                    let params = self.parse_params()?;
                    let is_async = self.eat(&Token::Async);
                    let body = self.parse_function_body()?;
                    let decl = FunctionDecl { name, type_params: vec![], params, return_type, body, is_async, is_generator: false };
                    return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
                }
                self.pos = saved; // rollback if not operator after all
            }
        }

        // Return type or field/method type
        let type_ann = self.try_parse_type_annotation();

        // Check for operator after return type (for cases try_parse_type_annotation consumed it)
        if self.peek() == &Token::Operator {
            self.advance();
            let op_name = self.parse_operator_name()?;
            let name = format!("operator{}", op_name);
            let params = self.parse_params()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params: vec![], params, return_type: type_ann, body, is_async, is_generator: false };
            return Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }]);
        }

        let name = self.expect_ident()?;

        if self.peek() == &Token::LParen || self.peek() == &Token::Less {
            // Method
            let type_params = self.parse_type_params()?;
            let params = self.parse_params()?;
            let is_async = self.eat(&Token::Async);
            let body = self.parse_function_body()?;
            let decl = FunctionDecl { name, type_params, params, return_type: type_ann, body, is_async, is_generator: false };
            Ok(vec![ClassMember::Method { is_static, is_abstract, is_override, kind, decl }])
        } else {
            // Field(s)
            let mut fields = Vec::new();
            let initializer = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
            fields.push(ClassMember::Field { is_static, is_final, is_late, type_ann: type_ann.clone(), name, initializer });
            
            while self.eat(&Token::Comma) {
                let name = self.expect_ident()?;
                let initializer = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
                fields.push(ClassMember::Field { is_static, is_final, is_late, type_ann: type_ann.clone(), name, initializer });
            }
            
            self.eat(&Token::Semicolon);
            Ok(fields)
        }
    }

    fn parse_constructor(&mut self, _class_name: &str, is_const: bool, is_factory: bool) -> Result<ClassMember, String> {
        self.expect_ident()?; // consume class name
        let ctor_name = if self.eat(&Token::Dot) { Some(self.expect_ident()?) } else { None };
        let params = self.parse_params()?;
        // initializer list
        let mut initializers = Vec::new();
        if self.eat(&Token::Colon) {
            loop {
                if self.peek() == &Token::Super {
                    self.advance();
                    let _name = if self.eat(&Token::Dot) { Some(self.expect_ident()?) } else { None };
                    let args = self.parse_args()?;
                    initializers.push(CtorInitializer::SuperCall(args));
                } else if self.peek() == &Token::This {
                    self.advance();
                    let name = if self.eat(&Token::Dot) { Some(self.expect_ident()?) } else { None };
                    let args = self.parse_args()?;
                    initializers.push(CtorInitializer::RedirectingCall(name, args));
                } else if let Token::Identifier(_) = self.peek() {
                    let field = self.expect_ident()?;
                    self.expect(Token::Eq)?;
                    let val = self.parse_expr()?;
                    initializers.push(CtorInitializer::FieldInit(field, val));
                } else { break; }
                if !self.eat(&Token::Comma) { break; }
            }
        }
        let body = if self.peek() == &Token::LBrace {
            self.advance();
            let stmts = self.parse_block_body()?;
            self.expect(Token::RBrace)?;
            Some(stmts)
        } else { self.eat(&Token::Semicolon); None };
        Ok(ClassMember::Constructor { name: ctor_name, params, initializers, body, is_const, is_factory })
    }

    // ── Statements ────────────────────────────────────────────────────────────

    fn parse_statement(&mut self) -> Result<Statement, String> {
        // skip annotations
        while self.peek() == &Token::At {
            self.advance(); self.expect_ident()?;
            if self.peek() == &Token::LParen { self.skip_balanced(Token::LParen, Token::RParen)?; }
        }
        match self.peek().clone() {
            Token::LBrace => {
                self.advance();
                let stmts = self.parse_block_body()?;
                self.expect(Token::RBrace)?;
                Ok(Statement::Block(stmts))
            }
            Token::If => self.parse_if(),
            Token::While => self.parse_while(),
            Token::Do => self.parse_do_while(),
            Token::For => self.parse_for(),
            Token::Switch => self.parse_switch(),
            Token::Return => {
                self.advance();
                let val = if self.peek() == &Token::Semicolon { None } else { Some(self.parse_expr()?) };
                self.eat(&Token::Semicolon);
                Ok(Statement::Return(val))
            }
            Token::Break => {
                self.advance();
                let label = if let Token::Identifier(l) = self.peek().clone() { self.advance(); Some(l) } else { None };
                self.eat(&Token::Semicolon);
                Ok(Statement::Break(label))
            }
            Token::Continue => {
                self.advance();
                let label = if let Token::Identifier(l) = self.peek().clone() { self.advance(); Some(l) } else { None };
                self.eat(&Token::Semicolon);
                Ok(Statement::Continue(label))
            }
            Token::Throw => {
                self.advance();
                let e = self.parse_expr()?;
                self.eat(&Token::Semicolon);
                Ok(Statement::Throw(e))
            }
            Token::Rethrow => {
                self.advance(); self.eat(&Token::Semicolon);
                Ok(Statement::Throw(Expression::Identifier("__rethrow__".into())))
            }
            Token::Try => self.parse_try(),
            Token::Assert => {
                self.advance();
                self.expect(Token::LParen)?;
                let cond = self.parse_expr()?;
                let msg = if self.eat(&Token::Comma) { Some(self.parse_expr()?) } else { None };
                self.expect(Token::RParen)?;
                self.eat(&Token::Semicolon);
                Ok(Statement::Assert(cond, msg))
            }
            Token::Semicolon => { self.advance(); Ok(Statement::Empty) }
            Token::Var | Token::Final | Token::Const | Token::Late => self.parse_var_decl_stmt(),
            Token::Void => self.parse_func_or_expr_stmt(),
            Token::Async => {
                self.advance();
                self.eat(&Token::Star);
                self.parse_func_or_expr_stmt_async(true)
            }
            Token::Identifier(_) => self.parse_ambiguous_stmt(),
            _ => {
                let e = self.parse_expr()?;
                self.eat(&Token::Semicolon);
                Ok(Statement::Expression(e))
            }
        }
    }

    fn parse_block_body(&mut self) -> Result<Vec<Statement>, String> {
        let mut stmts = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            stmts.push(self.parse_statement()?);
        }
        Ok(stmts)
    }

    fn parse_if(&mut self) -> Result<Statement, String> {
        self.advance();
        self.expect(Token::LParen)?;
        let cond = self.parse_expr()?;
        self.expect(Token::RParen)?;
        let then = Box::new(self.parse_statement()?);
        let else_ = if self.eat(&Token::Else) { Some(Box::new(self.parse_statement()?)) } else { None };
        Ok(Statement::If { condition: cond, then_branch: then, else_branch: else_ })
    }

    fn parse_while(&mut self) -> Result<Statement, String> {
        self.advance();
        self.expect(Token::LParen)?;
        let cond = self.parse_expr()?;
        self.expect(Token::RParen)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::While { condition: cond, body })
    }

    fn parse_do_while(&mut self) -> Result<Statement, String> {
        self.advance();
        let body = Box::new(self.parse_statement()?);
        self.expect(Token::While)?;
        self.expect(Token::LParen)?;
        let cond = self.parse_expr()?;
        self.expect(Token::RParen)?;
        self.eat(&Token::Semicolon);
        Ok(Statement::DoWhile { body, condition: cond })
    }

    fn parse_for(&mut self) -> Result<Statement, String> {
        self.advance();
        self.expect(Token::LParen)?;

        // Check for for-in: (var/final/type name in expr)
        let saved = self.pos;
        let is_for_in = self.try_parse_for_in_header();
        if let Some((is_final, var_type, var_name)) = is_for_in {
            self.expect(Token::In)?;
            let iterable = self.parse_expr()?;
            self.expect(Token::RParen)?;
            let body = Box::new(self.parse_statement()?);
            return Ok(Statement::ForIn { is_final, var_type, var_name, iterable, body });
        }
        self.pos = saved;

        // Classic for
        let init = if self.peek() == &Token::Semicolon {
            self.advance(); None
        } else {
            let s = self.parse_for_init()?;
            Some(s)
        };
        let condition = if self.peek() == &Token::Semicolon { self.advance(); None }
                        else { let e = self.parse_expr()?; self.expect(Token::Semicolon)?; Some(e) };
        let mut update = Vec::new();
        while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
            update.push(self.parse_expr()?);
            if !self.eat(&Token::Comma) { break; }
        }
        self.expect(Token::RParen)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::For(ForStatement { init, condition, update, body }))
    }

    fn try_parse_for_in_header(&mut self) -> Option<(bool, Option<String>, String)> {
        let is_final = self.eat(&Token::Final) || self.eat(&Token::Var);
        let type_or_name = match self.peek().clone() {
            Token::Identifier(s) => { self.advance(); s }
            _ => return None,
        };
        // If next is `in`, type_or_name is the variable name with no type
        if self.peek() == &Token::In {
            return Some((is_final, None, type_or_name));
        }
        // If next is an identifier, type_or_name is the type and next is the var name
        if let Token::Identifier(var) = self.peek().clone() {
            self.advance();
            if self.peek() == &Token::In {
                return Some((is_final, Some(type_or_name), var));
            }
        }
        None
    }

    fn parse_for_init(&mut self) -> Result<ForInit, String> {
        match self.peek().clone() {
            Token::Var | Token::Final | Token::Const | Token::Late => {
                let v = self.parse_var_decl()?;
                Ok(ForInit::VarDecl(v))
            }
            Token::Identifier(_) if matches!(self.peek2(), Token::Identifier(_)) => {
                let v = self.parse_var_decl()?;
                Ok(ForInit::VarDecl(v))
            }
            _ => {
                let e = self.parse_expr()?;
                self.expect(Token::Semicolon)?;
                Ok(ForInit::Expression(e))
            }
        }
    }

    fn parse_switch(&mut self) -> Result<Statement, String> {
        self.advance();
        self.expect(Token::LParen)?;
        let expr = self.parse_expr()?;
        self.expect(Token::RParen)?;
        self.expect(Token::LBrace)?;
        let mut cases = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            // Support modern patterns or legacy case/default
            let label = if self.eat(&Token::Case) {
                // Here we should ideally handle patterns in Switch statements too
                Some(self.parse_expr()?)
            } else if self.eat(&Token::Default) {
                None
            } else {
                break;
            };
            self.expect(Token::Colon)?;
            let mut body = Vec::new();
            while !matches!(self.peek(), Token::Case | Token::Default | Token::RBrace | Token::EOF) {
                body.push(self.parse_statement()?);
            }
            cases.push(SwitchCase { label, body });
        }
        self.expect(Token::RBrace)?;
        Ok(Statement::Switch { expr, cases })
    }

    fn parse_try(&mut self) -> Result<Statement, String> {
        self.advance();
        self.expect(Token::LBrace)?;
        let body = self.parse_block_body()?;
        self.expect(Token::RBrace)?;
        let mut catches = Vec::new();
        while matches!(self.peek(), Token::On | Token::Catch) {
            let on_type = if self.eat(&Token::On) { Some(self.expect_ident()?) } else { None };
            let (var_name, stack_name) = if self.eat(&Token::Catch) {
                self.expect(Token::LParen)?;
                let v = self.expect_ident()?;
                let s = if self.eat(&Token::Comma) { Some(self.expect_ident()?) } else { None };
                self.expect(Token::RParen)?;
                (Some(v), s)
            } else { (None, None) };
            self.expect(Token::LBrace)?;
            let cb = self.parse_block_body()?;
            self.expect(Token::RBrace)?;
            catches.push(CatchClause { on_type, var_name, stack_name, body: cb });
        }
        let finally = if self.eat(&Token::Finally) {
            self.expect(Token::LBrace)?;
            let f = self.parse_block_body()?;
            self.expect(Token::RBrace)?;
            Some(f)
        } else { None };
        Ok(Statement::Try { body, catches, finally })
    }

    fn parse_var_decl_stmt(&mut self) -> Result<Statement, String> {
        let is_late = self.eat(&Token::Late);
        let is_final = self.eat(&Token::Final);
        let is_const = self.eat(&Token::Const);
        if !is_final && !is_const { self.eat(&Token::Late); } // late after var
        let is_var = self.eat(&Token::Var);
        
        let type_name = if !is_var {
            match self.peek().clone() {
                Token::Identifier(t) if matches!(self.peek2(), Token::Identifier(_)) => {
                    self.advance(); Some(t)
                }
                Token::Void | Token::Dynamic => {
                    let t = format!("{:?}", self.advance()).to_lowercase();
                    Some(t)
                }
                _ => None,
            }
        } else { None };

        let mut decls = Vec::new();
        loop {
            let name = self.expect_ident()?;
            let initializer = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
            decls.push(Statement::VarDecl(VarDecl { is_final, is_const, is_late, type_name: type_name.clone(), name, initializer }));
            if !self.eat(&Token::Comma) { break; }
        }
        self.eat(&Token::Semicolon);

        if decls.len() == 1 {
            Ok(decls.pop().unwrap())
        } else {
            Ok(Statement::Block(decls))
        }
    }

    fn parse_var_decl(&mut self) -> Result<VarDecl, String> {
        let is_late = self.eat(&Token::Late);
        let is_final = self.eat(&Token::Final);
        let is_const = self.eat(&Token::Const);
        if !is_final && !is_const { self.eat(&Token::Late); } // late after var
        let is_var = self.eat(&Token::Var);
        let type_name = if !is_var {
            match self.peek().clone() {
                Token::Identifier(t) if matches!(self.peek2(), Token::Identifier(_)) => {
                    self.advance(); Some(t)
                }
                Token::Void | Token::Dynamic => {
                    let t = format!("{:?}", self.advance()).to_lowercase();
                    Some(t)
                }
                _ => None,
            }
        } else { None };
        let name = self.expect_ident()?;
        let initializer = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
        self.eat(&Token::Semicolon);
        Ok(VarDecl { is_final, is_const, is_late, type_name, name, initializer })
    }

    fn parse_func_or_expr_stmt(&mut self) -> Result<Statement, String> {
        self.parse_func_or_expr_stmt_async(false)
    }

    fn parse_func_or_expr_stmt_async(&mut self, is_async: bool) -> Result<Statement, String> {
        // void/type name(...) { } — function declaration
        let type_ann = self.try_parse_type_annotation();
        if let Token::Identifier(_) = self.peek() {
            let name = self.expect_ident()?;
            if self.peek() == &Token::LParen || self.peek() == &Token::Less {
                let type_params = self.parse_type_params()?;
                let params = self.parse_params()?;
                let is_async2 = is_async || self.eat(&Token::Async);
                let body = self.parse_function_body()?;
                return Ok(Statement::FunctionDecl(FunctionDecl {
                    name, type_params, params, return_type: type_ann, body, is_async: is_async2, is_generator: false,
                }));
            }
            // Variable declaration: Type name = ...
            let initializer = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
            self.eat(&Token::Semicolon);
            return Ok(Statement::VarDecl(VarDecl {
                is_final: false, is_const: false, is_late: false,
                type_name: type_ann.map(|t| t.name), name, initializer,
            }));
        }
        let e = self.parse_expr()?;
        self.eat(&Token::Semicolon);
        Ok(Statement::Expression(e))
    }

    /// Ambiguous: starts with Identifier — could be type+name, assignment, call, etc.
    fn parse_ambiguous_stmt(&mut self) -> Result<Statement, String> {
        // Heuristic: Identifier followed by Identifier → type declaration or function
        if let (Token::Identifier(_), Token::Identifier(_)) = (self.peek().clone(), self.peek2().clone()) {
            return self.parse_func_or_expr_stmt();
        }
        let e = self.parse_expr()?;
        self.eat(&Token::Semicolon);
        Ok(Statement::Expression(e))
    }

    // ── Functions ─────────────────────────────────────────────────────────────

    fn parse_function_body(&mut self) -> Result<FunctionBody, String> {
        if self.eat(&Token::Arrow) {
            let e = self.parse_expr()?;
            self.eat(&Token::Semicolon);
            Ok(FunctionBody::Expression(e))
        } else if self.peek() == &Token::LBrace {
            self.advance();
            let stmts = self.parse_block_body()?;
            self.expect(Token::RBrace)?;
            Ok(FunctionBody::Block(stmts))
        } else if self.eat(&Token::Semicolon) {
            Ok(FunctionBody::Empty)
        } else {
            Err(format!("Expected function body, got {:?}", self.peek()))
        }
    }

    fn parse_params(&mut self) -> Result<Params, String> {
        self.expect(Token::LParen)?;
        let mut positional = Vec::new();
        let mut optional_pos = Vec::new();
        let mut named = Vec::new();

        if self.peek() == &Token::RParen { self.advance(); return Ok(Params { positional, optional_pos, named }); }

        if self.eat(&Token::LBracket) {
            // [optional positional]
            while self.peek() != &Token::RBracket && self.peek() != &Token::EOF {
                optional_pos.push(self.parse_param(false)?);
                if !self.eat(&Token::Comma) { break; }
            }
            self.expect(Token::RBracket)?;
        } else if self.eat(&Token::LBrace) {
            // {named}
            while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
                named.push(self.parse_param(true)?);
                if !self.eat(&Token::Comma) { break; }
            }
            self.expect(Token::RBrace)?;
        } else {
            loop {
                if self.peek() == &Token::RParen { break; }
                if self.peek() == &Token::LBracket {
                    self.advance();
                    while self.peek() != &Token::RBracket && self.peek() != &Token::EOF {
                        optional_pos.push(self.parse_param(false)?);
                        if !self.eat(&Token::Comma) { break; }
                    }
                    self.expect(Token::RBracket)?;
                    break;
                }
                if self.peek() == &Token::LBrace {
                    self.advance();
                    while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
                        named.push(self.parse_param(true)?);
                        if !self.eat(&Token::Comma) { break; }
                    }
                    self.expect(Token::RBrace)?;
                    break;
                }
                positional.push(self.parse_param(false)?);
                if !self.eat(&Token::Comma) { break; }
            }
        }
        self.expect(Token::RParen)?;
        Ok(Params { positional, optional_pos, named })
    }

    fn parse_param(&mut self, is_named: bool) -> Result<Param, String> {
        let is_required = self.eat(&Token::Required);
        self.eat(&Token::Covariant);
        let is_final = self.eat(&Token::Final);
        let _ = is_final;
        // this.field shorthand
        let is_this = if self.peek() == &Token::This {
            self.advance(); self.expect(Token::Dot)?; true
        } else { false };
        // type annotation (optional)
        let type_ann = if matches!(self.peek(), Token::Identifier(_)) && matches!(self.peek2(), Token::Identifier(_)) {
            Some(self.parse_type_annotation()?)
        } else { None };
        let name = self.expect_ident()?;
        let default_value = if self.eat(&Token::Eq) || self.eat(&Token::Colon) {
            Some(self.parse_expr()?)
        } else { None };
        Ok(Param { name, type_ann, default_value, is_required: is_required || !is_named, is_this })
    }

    // ── Expressions ───────────────────────────────────────────────────────────

    pub fn parse_expr(&mut self) -> Result<Expression, String> {
        self.parse_assign()
    }

    fn parse_assign(&mut self) -> Result<Expression, String> {
        let left = self.parse_ternary()?;
        let op = match self.peek() {
            Token::Eq => AssignOp::Assign,
            Token::PlusEq => AssignOp::AddAssign,
            Token::MinusEq => AssignOp::SubAssign,
            Token::StarEq => AssignOp::MulAssign,
            Token::SlashEq => AssignOp::DivAssign,
            Token::PercentEq => AssignOp::ModAssign,
            Token::TildeSlashEq => AssignOp::IntDivAssign,
            Token::AmpEq => AssignOp::BitAndAssign,
            Token::BarEq => AssignOp::BitOrAssign,
            Token::CaretEq => AssignOp::BitXorAssign,
            Token::LessLessEq => AssignOp::ShlAssign,
            Token::GreaterGreaterEq => AssignOp::ShrAssign,
            Token::QuestionQuestionEq => AssignOp::NullAssign,
            _ => return Ok(left),
        };
        self.advance();
        let right = self.parse_assign()?;
        Ok(Expression::Assign { op, left: Box::new(left), right: Box::new(right) })
    }

    fn parse_ternary(&mut self) -> Result<Expression, String> {
        let cond = self.parse_null_coalesce()?;
        if self.eat(&Token::Question) {
            let then = self.parse_expr()?;
            self.expect(Token::Colon)?;
            let else_ = self.parse_expr()?;
            Ok(Expression::Ternary { cond: Box::new(cond), then: Box::new(then), else_: Box::new(else_) })
        } else { Ok(cond) }
    }

    fn parse_null_coalesce(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_or()?;
        while self.eat(&Token::QuestionQuestion) {
            let right = self.parse_or()?;
            left = Expression::NullCoalesce { left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_and()?;
        while self.eat(&Token::BarBar) {
            let right = self.parse_and()?;
            left = Expression::Binary { op: BinOp::Or, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitor()?;
        while self.eat(&Token::AmpAmp) {
            let right = self.parse_bitor()?;
            left = Expression::Binary { op: BinOp::And, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitxor()?;
        while self.eat(&Token::Bar) {
            let right = self.parse_bitxor()?;
            left = Expression::Binary { op: BinOp::BitOr, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitxor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitand()?;
        while self.eat(&Token::Caret) {
            let right = self.parse_bitand()?;
            left = Expression::Binary { op: BinOp::BitXor, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitand(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_equality()?;
        while self.eat(&Token::Amp) {
            let right = self.parse_equality()?;
            left = Expression::Binary { op: BinOp::BitAnd, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_equality(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_relational()?;
        loop {
            let op = match self.peek() {
                Token::EqEq => BinOp::Eq,
                Token::BangEq => BinOp::NotEq,
                _ => break,
            };
            self.advance();
            let right = self.parse_relational()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_relational(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_shift()?;
        loop {
            match self.peek().clone() {
                Token::Less => { self.advance(); let r = self.parse_shift()?; left = Expression::Binary { op: BinOp::Lt, left: Box::new(left), right: Box::new(r) }; }
                Token::Greater => { self.advance(); let r = self.parse_shift()?; left = Expression::Binary { op: BinOp::Gt, left: Box::new(left), right: Box::new(r) }; }
                Token::LessEq => { self.advance(); let r = self.parse_shift()?; left = Expression::Binary { op: BinOp::Le, left: Box::new(left), right: Box::new(r) }; }
                Token::GreaterEq => { self.advance(); let r = self.parse_shift()?; left = Expression::Binary { op: BinOp::Ge, left: Box::new(left), right: Box::new(r) }; }
                Token::Is => {
                    self.advance();
                    let negated = self.eat(&Token::Bang);
                    let type_ann = self.parse_type_annotation()?;
                    left = Expression::Is { expr: Box::new(left), type_ann, negated };
                }
                Token::As => {
                    self.advance();
                    let type_ann = self.parse_type_annotation()?;
                    left = Expression::As { expr: Box::new(left), type_ann };
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_shift(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_additive()?;
        loop {
            let op = match self.peek() {
                Token::LessLess => BinOp::Shl,
                Token::GreaterGreater => BinOp::Shr,
                Token::GreaterGreaterGreater => BinOp::UShr,
                _ => break,
            };
            self.advance();
            let right = self.parse_additive()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_additive(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_multiplicative()?;
        loop {
            let op = match self.peek() {
                Token::Plus => BinOp::Add,
                Token::Minus => BinOp::Sub,
                _ => break,
            };
            self.advance();
            let right = self.parse_multiplicative()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_multiplicative(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_unary()?;
        loop {
            let op = match self.peek() {
                Token::Star => BinOp::Mul,
                Token::Slash => BinOp::Div,
                Token::Percent => BinOp::Mod,
                Token::TildeSlash => BinOp::IntDiv,
                _ => break,
            };
            self.advance();
            let right = self.parse_unary()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            Token::Bang => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::Not, expr: Box::new(e) }) }
            Token::Minus => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::Neg, expr: Box::new(e) }) }
            Token::Tilde => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::BitNot, expr: Box::new(e) }) }
            Token::PlusPlus => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::PreInc, expr: Box::new(e) }) }
            Token::MinusMinus => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::PreDec, expr: Box::new(e) }) }
            Token::Await => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Await(Box::new(e))) }
            _ => self.parse_postfix(),
        }
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            match self.peek().clone() {
                Token::PlusPlus => { self.advance(); expr = Expression::PostfixUnary { op: PostfixOp::PostInc, expr: Box::new(expr) }; }
                Token::MinusMinus => { self.advance(); expr = Expression::PostfixUnary { op: PostfixOp::PostDec, expr: Box::new(expr) }; }
                Token::Dot => {
                    self.advance();
                    let member = self.expect_ident()?;
                    if self.peek() == &Token::LParen {
                        let args = self.parse_args()?;
                        expr = Expression::Call { callee: Box::new(Expression::Member { object: Box::new(expr), member, null_safe: false }), type_args: vec![], args, null_safe: false };
                    } else {
                        expr = Expression::Member { object: Box::new(expr), member, null_safe: false };
                    }
                }
                Token::QuestionDot => {
                    self.advance();
                    let member = self.expect_ident()?;
                    if self.peek() == &Token::LParen {
                        let args = self.parse_args()?;
                        expr = Expression::Call { callee: Box::new(Expression::Member { object: Box::new(expr), member, null_safe: true }), type_args: vec![], args, null_safe: true };
                    } else {
                        expr = Expression::Member { object: Box::new(expr), member, null_safe: true };
                    }
                }
                Token::LBracket => {
                    self.advance();
                    let idx = self.parse_expr()?;
                    self.expect(Token::RBracket)?;
                    expr = Expression::Index { object: Box::new(expr), index: Box::new(idx) };
                }
                Token::LParen => {
                    let args = self.parse_args()?;
                    expr = Expression::Call { callee: Box::new(expr), type_args: vec![], args, null_safe: false };
                }
                Token::DotDot | Token::QuestionDotDot => {
                    // Cascade
                    let is_null_safe = self.peek() == &Token::QuestionDotDot;
                    self.advance();
                    let mut ops = Vec::new();
                    loop {
                        let member = self.expect_ident()?;
                        if self.peek() == &Token::LParen {
                            let args = self.parse_args()?;
                            ops.push(CascadeOp::Method(member, args));
                        } else if self.eat(&Token::Eq) {
                            let val = self.parse_expr()?;
                            ops.push(CascadeOp::Assign(member, val));
                        } else {
                            ops.push(CascadeOp::Field(member));
                        }
                        
                        // Check for another cascade op
                        if self.peek() == &Token::DotDot {
                            self.advance();
                        } else {
                            break;
                        }
                    }
                    expr = Expression::Cascade { object: Box::new(expr), ops, null_safe: is_null_safe };
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            Token::IntLiteral(n) => { self.advance(); Ok(Expression::Int(n)) }
            Token::DoubleLiteral(n) => { self.advance(); Ok(Expression::Double(n)) }
            Token::True => { self.advance(); Ok(Expression::Bool(true)) }
            Token::False => { self.advance(); Ok(Expression::Bool(false)) }
            Token::Null => { self.advance(); Ok(Expression::Null) }
            Token::This => { self.advance(); Ok(Expression::This) }
            Token::Super => { self.advance(); Ok(Expression::Super) }
            Token::StringLiteral(s) => {
                self.advance();
                if s.starts_with("\x00INTERP\x00") {
                    let parts = decode_interp_parts(&s[8..]);
                    Ok(Expression::String(StringExpr::Interpolated(parts)))
                } else {
                    Ok(Expression::String(StringExpr::Simple(s)))
                }
            }
            Token::LBracket => {
                self.advance();
                let mut elems = Vec::new();
                while self.peek() != &Token::RBracket && self.peek() != &Token::EOF {
                    if self.eat(&Token::DotDotDot) {
                        elems.push(Expression::Spread(Box::new(self.parse_expr()?)));
                    } else {
                        elems.push(self.parse_expr()?);
                    }
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RBracket)?;
                Ok(Expression::List { type_arg: None, elements: elems })
            }
            Token::LBrace => {
                self.advance();
                if self.eat(&Token::RBrace) {
                    return Ok(Expression::Map { type_args: None, entries: Vec::new() });
                }
                
                // Distinguish between Set and Map
                let first = self.parse_expr()?;
                if self.peek() == &Token::Colon {
                    self.advance(); // :
                    let val = self.parse_expr()?;
                    let mut entries = vec![(first, val)];
                    if self.eat(&Token::Comma) {
                        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
                            let k = self.parse_expr()?;
                            self.expect(Token::Colon)?;
                            let v = self.parse_expr()?;
                            entries.push((k, v));
                            if !self.eat(&Token::Comma) { break; }
                        }
                    }
                    self.expect(Token::RBrace)?;
                    Ok(Expression::Map { type_args: None, entries })
                } else {
                    // Set
                    let mut elements = vec![first];
                    if self.eat(&Token::Comma) {
                        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
                            elements.push(self.parse_expr()?);
                            if !self.eat(&Token::Comma) { break; }
                        }
                    }
                    self.expect(Token::RBrace)?;
                    Ok(Expression::Set { type_arg: None, elements })
                }
            }
            Token::New => {
                self.advance();
                let class = self.expect_ident()?;
                let constructor = if self.eat(&Token::Dot) { Some(self.expect_ident()?) } else { None };
                self.skip_type_args();
                let args = self.parse_args()?;
                Ok(Expression::New { class, constructor, type_args: vec![], args })
            }
            Token::Const => {
                self.advance();
                let class = self.expect_ident()?;
                let constructor = if self.eat(&Token::Dot) { Some(self.expect_ident()?) } else { None };
                self.skip_type_args();
                let args = self.parse_args()?;
                Ok(Expression::Const { class, constructor, args })
            }
            Token::Switch => {
                self.parse_switch_expression()
            }
            Token::Identifier(name) => {
                self.advance();
                // Check for lambda: name => expr  or  (params) => expr
                if self.peek() == &Token::Arrow {
                    self.advance();
                    let body_expr = self.parse_expr()?;
                    let params = Params {
                        positional: vec![Param { name: name.clone(), type_ann: None, default_value: None, is_required: true, is_this: false }],
                        optional_pos: vec![], named: vec![],
                    };
                    return Ok(Expression::Lambda { params, body: Box::new(FunctionBody::Expression(body_expr)), is_async: false });
                }
                // Constructor call without `new`: ClassName(...)
                if self.peek() == &Token::LParen && name.chars().next().map(|c| c.is_uppercase()).unwrap_or(false) {
                    // Could be a function call or constructor — treat as call
                }
                Ok(Expression::Identifier(name))
            }
            Token::LParen => {
                // Could be lambda params
                let saved = self.pos;
                if let Ok(params) = self.try_parse_lambda_params() {
                    if self.eat(&Token::Arrow) {
                        let body_expr = self.parse_expr()?;
                        return Ok(Expression::Lambda { params, body: Box::new(FunctionBody::Expression(body_expr)), is_async: false });
                    } else if self.peek() == &Token::LBrace {
                        self.advance();
                        let body_block = self.parse_block_body()?;
                        self.expect(Token::RBrace)?;
                        return Ok(Expression::Lambda { params, body: Box::new(FunctionBody::Block(body_block)), is_async: false });
                    }
                }
                self.pos = saved;
                self.advance();
                
                // Peek ahead for record: (1,) or (a: 1) or (1, 2)
                // If it's just (expr), it's a parenthesized expression.
                let mut elements = Vec::new();
                let mut is_record = false;
                
                while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
                    let label = if let (Token::Identifier(id), Token::Colon) = (self.peek().clone(), self.peek2().clone()) {
                        self.advance(); // id
                        self.advance(); // :
                        is_record = true;
                        Some(id)
                    } else {
                        None
                    };
                    let value = self.parse_expr()?;
                    elements.push(Argument { label, value });
                    if self.peek() == &Token::Comma {
                        self.advance();
                        is_record = true; // (x,) is a record
                    } else {
                        break;
                    }
                }
                
                self.expect(Token::RParen)?;
                if is_record || elements.is_empty() {
                    Ok(Expression::Record { elements })
                } else {
                    // Just a parenthesized expression
                    Ok(elements[0].value.clone())
                }
            }
            t => Err(format!("Unexpected token in expression: {:?}", t)),
        }
    }

    fn try_parse_lambda_params(&mut self) -> Result<Params, String> {
        self.expect(Token::LParen)?;
        let mut positional = Vec::new();
        while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
            let name = self.expect_ident()?;
            positional.push(Param { name, type_ann: None, default_value: None, is_required: true, is_this: false });
            if !self.eat(&Token::Comma) { break; }
        }
        self.expect(Token::RParen)?;
        Ok(Params { positional, optional_pos: vec![], named: vec![] })
    }

    fn parse_args(&mut self) -> Result<Vec<Argument>, String> {
        self.expect(Token::LParen)?;
        let mut args = Vec::new();
        while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
            // named arg: name: value
            let label = if let Token::Identifier(n) = self.peek().clone() {
                if self.peek2() == &Token::Colon {
                    let n = n.clone();
                    self.advance(); self.advance();
                    Some(n)
                } else { None }
            } else { None };
            let value = if self.eat(&Token::DotDotDot) {
                Expression::Spread(Box::new(self.parse_expr()?))
            } else {
                self.parse_expr()?
            };
            args.push(Argument { label, value });
            if !self.eat(&Token::Comma) { break; }
        }
        self.expect(Token::RParen)?;
        Ok(args)
    }

    // ── Types ─────────────────────────────────────────────────────────────────

    fn parse_type_annotation(&mut self) -> Result<TypeAnnotation, String> {
        let name = match self.advance() {
            Token::Identifier(s) => s,
            Token::Void => "void".into(),
            Token::Dynamic => "dynamic".into(),
            t => return Err(format!("Expected type name, got {:?}", t)),
        };
        let args = if self.peek() == &Token::Less { self.parse_type_arg_list()? } else { vec![] };
        let nullable = self.eat(&Token::Question);
        Ok(TypeAnnotation { name, args, nullable })
    }

    fn try_parse_type_annotation(&mut self) -> Option<TypeAnnotation> {
        match self.peek() {
            Token::Void | Token::Dynamic => {
                let name = format!("{:?}", self.advance()).to_lowercase();
                Some(TypeAnnotation { name, args: vec![], nullable: false })
            }
            Token::Identifier(_) if matches!(self.peek2(), Token::Identifier(_) | Token::Less | Token::Question) => {
                let saved = self.pos;
                if let Ok(t) = self.parse_type_annotation() { Some(t) } else { self.pos = saved; None }
            }
            _ => None,
        }
    }

    fn parse_type_arg_list(&mut self) -> Result<Vec<TypeAnnotation>, String> {
        self.expect(Token::Less)?;
        let mut args = vec![self.parse_type_annotation()?];
        while self.eat(&Token::Comma) { args.push(self.parse_type_annotation()?); }
        self.expect(Token::Greater)?;
        Ok(args)
    }

    fn parse_type_params(&mut self) -> Result<Vec<String>, String> {
        if self.peek() != &Token::Less { return Ok(vec![]); }
        self.advance();
        let mut params = vec![self.expect_ident()?];
        while self.eat(&Token::Comma) { params.push(self.expect_ident()?); }
        self.expect(Token::Greater)?;
        Ok(params)
    }

    fn skip_type_args(&mut self) {
        if self.peek() == &Token::Less {
            let _ = self.skip_balanced(Token::Less, Token::Greater);
        }
    }

    fn skip_balanced(&mut self, open: Token, _close: Token) -> Result<(), String> {
        self.expect(open)?;
        let mut depth = 1;
        while depth > 0 && self.peek() != &Token::EOF {
            let t = self.advance();
            if t == Token::Less || t == Token::LParen || t == Token::LBrace || t == Token::LBracket { depth += 1; }
            else if t == Token::Greater || t == Token::RParen || t == Token::RBrace || t == Token::RBracket { depth -= 1; }
        }
        Ok(())
    }

    fn parse_switch_expression(&mut self) -> Result<Expression, String> {
        self.advance(); // switch
        self.expect(Token::LParen)?;
        let expr = self.parse_expr()?;
        self.expect(Token::RParen)?;
        self.expect(Token::LBrace)?;
        
        let mut cases = Vec::new();
        while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
            let pattern = self.parse_pattern()?;
            
            let mut guard = None;
            if let Token::Identifier(id) = self.peek() {
                if id == "when" {
                    self.advance();
                    guard = Some(self.parse_expr()?);
                }
            }
            
            self.expect(Token::Arrow)?;
            let result = self.parse_expr()?;
            
            cases.push(SwitchExpressionCase { pattern, guard, result });
            if !self.eat(&Token::Comma) { break; }
        }
        
        self.expect(Token::RBrace)?;
        Ok(Expression::Switch { expr: Box::new(expr), cases })
    }

    fn parse_pattern(&mut self) -> Result<Pattern, String> {
        let mut left = self.parse_primary_pattern()?;
        
        // Handle logical patterns: &&, ||
        loop {
            if self.peek() == &Token::BarBar {
                self.advance();
                let right = self.parse_primary_pattern()?;
                left = Pattern::Logical(Box::new(left), Box::new(right), true);
            } else if self.peek() == &Token::AmpAmp {
                self.advance();
                let right = self.parse_primary_pattern()?;
                left = Pattern::Logical(Box::new(left), Box::new(right), false);
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_primary_pattern(&mut self) -> Result<Pattern, String> {
        match self.peek().clone() {
            Token::Identifier(id) if id == "_" => {
                self.advance();
                Ok(Pattern::Wildcard)
            }
            // Relational patterns: > 5, <= 10
            Token::Greater | Token::Less | Token::GreaterEq | Token::LessEq | Token::EqEq | Token::BangEq => {
                let op = match self.advance() {
                    Token::Greater => ">",
                    Token::Less => "<",
                    Token::GreaterEq => ">=",
                    Token::LessEq => "<=",
                    Token::EqEq => "==",
                    Token::BangEq => "!=",
                    _ => unreachable!(),
                };
                let val = self.parse_expr()?;
                Ok(Pattern::Relational { op: op.to_string(), val })
            }
            // List pattern: [1, 2, _]
            Token::LBracket => {
                self.advance();
                let mut patterns = Vec::new();
                while self.peek() != &Token::RBracket && self.peek() != &Token::EOF {
                    patterns.push(self.parse_pattern()?);
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RBracket)?;
                Ok(Pattern::List(patterns))
            }
            // Map pattern: {'a': 1, 'b': _}
            Token::LBrace => {
                self.advance();
                let mut entries = Vec::new();
                while self.peek() != &Token::RBrace && self.peek() != &Token::EOF {
                    let key = self.parse_expr()?;
                    self.expect(Token::Colon)?;
                    let pat = self.parse_pattern()?;
                    entries.push((key, pat));
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RBrace)?;
                Ok(Pattern::Map(entries))
            }
            // Record or parenthesized pattern: (1, 2) or (a: 1)
            Token::LParen => {
                self.advance();
                let mut elements = Vec::new();
                while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
                    let label = if let (Token::Identifier(id), Token::Colon) = (self.peek().clone(), self.peek2().clone()) {
                        self.advance(); self.advance();
                        Some(id)
                    } else { None };
                    let pattern = self.parse_pattern()?;
                    elements.push(ArgumentPattern { label, pattern });
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RParen)?;
                Ok(Pattern::Record(elements))
            }
            Token::Identifier(id) => {
                let id = id.clone();
                self.advance();
                // Object pattern: ClassName(field: pattern)
                if self.peek() == &Token::LParen {
                    self.advance();
                    let mut fields = Vec::new();
                    while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
                        let field_name = self.expect_ident()?;
                        self.expect(Token::Colon)?;
                        let pat = self.parse_pattern()?;
                        fields.push((field_name, pat));
                        if !self.eat(&Token::Comma) { break; }
                    }
                    self.expect(Token::RParen)?;
                    Ok(Pattern::Object { class_name: id, fields })
                } else if let Token::Identifier(var_name) = self.peek() {
                    // Type varName pattern: int x
                    let var_name = var_name.clone();
                    self.advance();
                    Ok(Pattern::Variable(var_name))
                } else {
                    // Constant: id
                    Ok(Pattern::Constant(Expression::Identifier(id)))
                }
            }
            Token::Null => { self.advance(); Ok(Pattern::Constant(Expression::Null)) }
            Token::True => { self.advance(); Ok(Pattern::Constant(Expression::Bool(true))) }
            Token::False => { self.advance(); Ok(Pattern::Constant(Expression::Bool(false))) }
            _ => {
                let e = self.parse_primary()?;
                Ok(Pattern::Constant(e))
            }
        }
    }
}

/// Decode interpolation parts from the encoded string.
fn decode_interp_parts(encoded: &str) -> Vec<StringPart> {
    // Simple: split on \x01, each part starts with L or E
    encoded.split('\x01').filter_map(|p| {
        if p.starts_with('L') { Some(StringPart::Literal(p[1..].to_string())) }
        else if p.starts_with('E') {
            // Re-parse the expression
            let inner = &p[1..];
            // Strip debug repr markers if any
            Some(StringPart::Expr(parse_expr_str(inner)))
        } else { None }
    }).collect()
}

/// Parse a single expression from a string (used for interpolation).
pub fn parse_expr_str(s: &str) -> Expression {
    let mut p = Parser::new(s);
    p.parse_expr().unwrap_or(Expression::String(StringExpr::Simple(s.to_string())))
}

// (PartialEq is derived on Token in token.rs)

// close impl Parser and any unclosed block
