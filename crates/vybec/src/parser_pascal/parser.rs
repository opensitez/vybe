use super::ast::*;
use super::lexer::Lexer;
use super::token::Token;

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(input: &str) -> Self {
        let mut lex = Lexer::new(input);
        let mut tokens = Vec::new();
        loop {
            let t = lex.next_token();
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
    fn expect_ident(&mut self) -> Result<String, String> {
        match self.advance() {
            Token::Identifier(s) => Ok(s),
            Token::Result => Ok("Result".into()),
            t => Err(format!("Expected identifier, got {:?}", t)),
        }
    }
    #[allow(dead_code)]
    fn eat_semicolons(&mut self) { while self.eat(&Token::Semicolon) {} }

    // ── Program ───────────────────────────────────────────────────────────────

    pub fn parse_program(&mut self) -> Result<Program, String> {
        let name = if self.eat(&Token::Program) || self.eat(&Token::Unit) {
            let n = self.expect_ident()?;
            self.eat(&Token::Semicolon);
            n
        } else { "main".into() };

        let uses = if self.eat(&Token::Uses) {
            let mut u = vec![self.expect_ident()?];
            while self.eat(&Token::Comma) { u.push(self.expect_ident()?); }
            self.eat(&Token::Semicolon);
            u
        } else { Vec::new() };

        self.eat(&Token::Interface);
        self.eat(&Token::Implementation);

        let decls = self.parse_decl_section()?;

        let body = if self.peek() == &Token::Begin {
            self.parse_compound()?
        } else { Vec::new() };
        self.eat(&Token::Dot);

        Ok(Program { name, uses, decls, body })
    }

    // ── Declarations ──────────────────────────────────────────────────────────

    fn parse_decl_section(&mut self) -> Result<Vec<Decl>, String> {
        let mut decls = Vec::new();
        loop {
            match self.peek() {
                Token::Var => { self.advance(); decls.push(Decl::Var(self.parse_var_section()?)); }
                Token::Const => { self.advance(); decls.push(Decl::Const(self.parse_const_section()?)); }
                Token::Type => { self.advance(); decls.push(Decl::Type(self.parse_type_section()?)); }
                Token::Procedure => {
                    self.advance();
                    if self.is_method_impl() {
                        decls.push(Decl::Method(self.parse_method_impl(MethodKind::Procedure)?));
                    } else {
                        decls.push(Decl::Procedure(self.parse_procedure()?));
                    }
                }
                Token::Function => {
                    self.advance();
                    if self.is_method_impl() {
                        decls.push(Decl::Method(self.parse_method_impl(MethodKind::Function)?));
                    } else {
                        decls.push(Decl::Function(self.parse_function()?));
                    }
                }
                Token::Constructor => {
                    self.advance();
                    decls.push(Decl::Method(self.parse_method_impl(MethodKind::Constructor)?));
                }
                Token::Destructor => {
                    self.advance();
                    decls.push(Decl::Method(self.parse_method_impl(MethodKind::Destructor)?));
                }
                _ => break,
            }
        }
        Ok(decls)
    }

    fn is_method_impl(&self) -> bool {
        matches!(self.peek(), Token::Identifier(_)) && self.peek2() == &Token::Dot
    }

    fn parse_method_impl(&mut self, kind: MethodKind) -> Result<MethodImpl, String> {
        let class_name = self.expect_ident()?;
        self.expect(Token::Dot)?;
        let method_name = self.expect_ident()?;
        let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
        let return_type = if kind == MethodKind::Function {
            self.expect(Token::Colon)?;
            Some(self.parse_type_ref()?)
        } else { None };
        self.eat(&Token::Semicolon);
        let decls = self.parse_decl_section()?;
        let body = self.parse_compound()?;
        self.eat(&Token::Semicolon);
        Ok(MethodImpl { kind, class_name, method_name, params, return_type, decls, body })
    }

    fn parse_var_section(&mut self) -> Result<Vec<VarDecl>, String> {
        let mut vars = Vec::new();
        while let Token::Identifier(_) = self.peek() {
            let mut names = vec![self.expect_ident()?];
            while self.eat(&Token::Comma) { names.push(self.expect_ident()?); }
            self.expect(Token::Colon)?;
            let type_name = self.parse_type_ref()?;
            let init = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
            self.eat(&Token::Semicolon);
            vars.push(VarDecl { names, type_name, init });
        }
        Ok(vars)
    }

    fn parse_const_section(&mut self) -> Result<Vec<ConstDecl>, String> {
        let mut consts = Vec::new();
        while let Token::Identifier(_) = self.peek() {
            let name = self.expect_ident()?;
            // Typed constant: `N: Integer = 42;`
            let type_name = if self.eat(&Token::Colon) {
                let t = self.parse_type_ref()?;
                Some(t)
            } else { None };
            self.expect(Token::Eq)?;
            let value = self.parse_expr()?;
            self.eat(&Token::Semicolon);
            consts.push(ConstDecl { name, type_name, value });
        }
        Ok(consts)
    }

    fn parse_type_section(&mut self) -> Result<Vec<TypeDecl>, String> {
        let mut types = Vec::new();
        while let Token::Identifier(_) = self.peek() {
            let name = self.expect_ident()?;
            self.expect(Token::Eq)?;
            let def = self.parse_type_def()?;
            self.eat(&Token::Semicolon);
            types.push(TypeDecl { name, def });
        }
        Ok(types)
    }

    fn parse_type_def(&mut self) -> Result<TypeDef, String> {
        match self.peek().clone() {
            Token::Record => {
                self.advance();
                self.parse_record_def()
            }
            Token::Class => {
                self.advance();
                self.parse_class_def()
            }
            Token::Interface => {
                self.advance();
                self.parse_interface_def()
            }
            Token::Array => {
                self.advance();
                let index = if self.eat(&Token::LBracket) {
                    let lo = self.parse_expr()?;
                    self.expect(Token::DotDot)?;
                    let hi = self.parse_expr()?;
                    self.expect(Token::RBracket)?;
                    Some((lo, hi))
                } else { None };
                self.expect(Token::Of)?;
                let elem = self.parse_type_ref()?;
                Ok(TypeDef::Array { index, element: Box::new(elem) })
            }
            Token::Caret => {
                self.advance();
                let t = self.parse_type_ref()?;
                Ok(TypeDef::Pointer(Box::new(t)))
            }
            Token::LParen => {
                // Enum: (Red, Green, Blue) or (Red = 0, Green = 1, ...)
                self.advance();
                let mut values = Vec::new();
                while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
                    let name = self.expect_ident()?;
                    let value = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
                    values.push(EnumValue { name, value });
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RParen)?;
                Ok(TypeDef::Enum(values))
            }
            _ => Ok(TypeDef::Alias(self.parse_type_ref()?)),
        }
    }

    fn parse_record_def(&mut self) -> Result<TypeDef, String> {
        let mut fields = Vec::new();
        let mut methods = Vec::new();
        loop {
            // Skip visibility
            match self.peek() {
                Token::Public | Token::Private | Token::Protected => { self.advance(); continue; }
                _ => {}
            }
            if self.peek() == &Token::End || self.peek() == &Token::EOF { break; }
            match self.peek().clone() {
                Token::Constructor | Token::Destructor | Token::Procedure | Token::Function => {
                    let kind = match self.advance() {
                        Token::Constructor => MethodKind::Constructor,
                        Token::Destructor => MethodKind::Destructor,
                        Token::Procedure => MethodKind::Procedure,
                        Token::Function => MethodKind::Function,
                        _ => unreachable!(),
                    };
                    methods.push(self.parse_method_sig(kind)?);
                }
                Token::Class => {
                    self.advance();
                    let kind = match self.advance() {
                        Token::Procedure => MethodKind::Procedure,
                        Token::Function => MethodKind::Function,
                        _ => break,
                    };
                    let mut sig = self.parse_method_sig(kind)?;
                    sig.is_static = true;
                    methods.push(sig);
                }
                Token::Identifier(_) => {
                    if let Token::Identifier(name) = self.peek().clone() {
                        if name.to_lowercase() == "operator" {
                            self.advance();
                            let mut sig = self.parse_method_sig(MethodKind::Function)?;
                            sig.is_operator = true;
                            methods.push(sig);
                            continue;
                        }
                    }
                    fields.push(self.parse_field_decl()?);
                }
                _ => break,
            }
        }
        self.expect(Token::End)?;
        Ok(TypeDef::Record(RecordDef { fields, methods }))
    }

    fn parse_class_def(&mut self) -> Result<TypeDef, String> {
        // Optional parent + interfaces: class(TParent, IFoo, IBar)
        let mut parent = None;
        let mut interfaces = Vec::new();
        if self.eat(&Token::LParen) {
            let first = self.expect_ident()?;
            // Check if first is an interface (starts with I) or a class
            parent = Some(first);
            while self.eat(&Token::Comma) {
                interfaces.push(self.expect_ident()?);
            }
            self.expect(Token::RParen)?;
        }

        // Forward declaration: `TFoo = class;`
        if self.peek() == &Token::Semicolon {
            return Ok(TypeDef::Class(ClassDef { parent, interfaces, members: Vec::new() }));
        }

        let mut members = Vec::new();
        loop {
            // Skip visibility keywords
            match self.peek() {
                Token::Public | Token::Private | Token::Protected | Token::Published => {
                    self.advance();
                    continue;
                }
                _ => {}
            }
            if self.peek() == &Token::End || self.peek() == &Token::EOF { break; }

            match self.peek().clone() {
                Token::Constructor => {
                    self.advance();
                    members.push(ClassMember::MethodDecl(self.parse_method_sig(MethodKind::Constructor)?));
                }
                Token::Destructor => {
                    self.advance();
                    members.push(ClassMember::MethodDecl(self.parse_method_sig(MethodKind::Destructor)?));
                }
                Token::Procedure => {
                    self.advance();
                    members.push(ClassMember::MethodDecl(self.parse_method_sig_with_static(MethodKind::Procedure, false)?));
                }
                Token::Function => {
                    self.advance();
                    members.push(ClassMember::MethodDecl(self.parse_method_sig_with_static(MethodKind::Function, false)?));
                }
                Token::Class => {
                    self.advance();
                    match self.peek().clone() {
                        Token::Procedure => {
                            self.advance();
                            members.push(ClassMember::MethodDecl(self.parse_method_sig_with_static(MethodKind::Procedure, true)?));
                        }
                        Token::Function => {
                            self.advance();
                            members.push(ClassMember::MethodDecl(self.parse_method_sig_with_static(MethodKind::Function, true)?));
                        }
                        Token::Var => {
                            // class var FCount: Integer;
                            self.advance();
                            let field = self.parse_field_decl()?;
                            members.push(ClassMember::ClassVar(field));
                        }
                        _ => break,
                    }
                }
                Token::Identifier(_) => {
                    if let Token::Identifier(name) = self.peek().clone() {
                        if name.to_lowercase() == "property" {
                            self.advance();
                            members.push(ClassMember::PropertyDecl(self.parse_property_def()?));
                            continue;
                        }
                    }
                    let field = self.parse_field_decl()?;
                    members.push(ClassMember::Field(field));
                }
                _ => break,
            }
        }
        self.expect(Token::End)?;
        Ok(TypeDef::Class(ClassDef { parent, interfaces, members }))
    }

    fn parse_interface_def(&mut self) -> Result<TypeDef, String> {
        let parent = if self.eat(&Token::LParen) {
            let p = self.expect_ident()?;
            self.expect(Token::RParen)?;
            Some(p)
        } else { None };

        // Forward declaration
        if self.peek() == &Token::Semicolon {
            return Ok(TypeDef::InterfaceDef(InterfaceDecl { parent, methods: Vec::new(), properties: Vec::new() }));
        }

        let mut methods = Vec::new();
        let mut properties = Vec::new();
        loop {
            if self.peek() == &Token::End || self.peek() == &Token::EOF { break; }
            match self.peek().clone() {
                Token::Procedure => {
                    self.advance();
                    methods.push(self.parse_method_sig(MethodKind::Procedure)?);
                }
                Token::Function => {
                    self.advance();
                    methods.push(self.parse_method_sig(MethodKind::Function)?);
                }
                Token::Identifier(_) => {
                    if let Token::Identifier(name) = self.peek().clone() {
                        if name.to_lowercase() == "property" {
                            self.advance();
                            properties.push(self.parse_property_def()?);
                            continue;
                        }
                    }
                    break;
                }
                _ => break,
            }
        }
        self.expect(Token::End)?;
        Ok(TypeDef::InterfaceDef(InterfaceDecl { parent, methods, properties }))
    }

    fn parse_method_sig(&mut self, kind: MethodKind) -> Result<MethodSig, String> {
        self.parse_method_sig_with_static(kind, false)
    }

    fn parse_method_sig_with_static(&mut self, kind: MethodKind, is_static: bool) -> Result<MethodSig, String> {
        let name = self.expect_ident()?;
        let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
        let return_type = if kind == MethodKind::Function {
            self.expect(Token::Colon)?;
            Some(self.parse_type_ref()?)
        } else { None };
        self.eat(&Token::Semicolon);

        let mut directives = Vec::new();
        loop {
            match self.peek() {
                Token::Virtual => { self.advance(); directives.push(MethodDirective::Virtual); self.eat(&Token::Semicolon); }
                Token::Override => { self.advance(); directives.push(MethodDirective::Override); self.eat(&Token::Semicolon); }
                Token::Abstract => { self.advance(); directives.push(MethodDirective::Abstract); self.eat(&Token::Semicolon); }
                _ => {
                    if let Token::Identifier(s) = self.peek() {
                        match s.to_lowercase().as_str() {
                            "reintroduce" => {
                                self.advance();
                                directives.push(MethodDirective::Reintroduce);
                                self.eat(&Token::Semicolon);
                                continue;
                            }
                            "overload" | "inline" | "cdecl" | "stdcall" | "register" | "dynamic" => {
                                self.advance();
                                self.eat(&Token::Semicolon);
                                continue;
                            }
                            _ => {}
                        }
                    }
                    break;
                }
            }
        }

        Ok(MethodSig { kind, name, params, return_type, directives, is_static, is_operator: false })
    }

    fn parse_field_decl(&mut self) -> Result<VarDecl, String> {
        let mut names = vec![self.expect_ident()?];
        while self.eat(&Token::Comma) { names.push(self.expect_ident()?); }
        self.expect(Token::Colon)?;
        let type_name = self.parse_type_ref()?;
        self.eat(&Token::Semicolon);
        Ok(VarDecl { names, type_name, init: None })
    }

    fn parse_property_def(&mut self) -> Result<PropertyDef, String> {
        let name = self.expect_ident()?;
        // Optional index: property Items[Index: Integer]: String
        let index_type = if self.eat(&Token::LBracket) {
            // skip param name
            self.expect_ident()?;
            self.expect(Token::Colon)?;
            let t = self.parse_type_ref()?;
            self.expect(Token::RBracket)?;
            Some(t)
        } else { None };
        self.expect(Token::Colon)?;
        let type_name = self.parse_type_ref()?;
        let mut reader = None;
        let mut writer = None;
        let mut default = false;
        while let Token::Identifier(kw) = self.peek() {
            match kw.to_lowercase().as_str() {
                "read" => { self.advance(); reader = Some(self.expect_ident()?); }
                "write" => { self.advance(); writer = Some(self.expect_ident()?); }
                "default" => {
                    self.advance();
                    // `default;` at end means default array property
                    if self.peek() == &Token::Semicolon {
                        default = true;
                    } else {
                        self.parse_expr()?; // skip default value expr
                    }
                    break;
                }
                "stored" | "nodefault" => { self.advance(); self.parse_expr().ok(); break; }
                _ => break,
            }
        }
        self.eat(&Token::Semicolon);
        Ok(PropertyDef { name, type_name, reader, writer, default, index_type })
    }

    fn parse_type_ref(&mut self) -> Result<TypeRef, String> {
        let name = match self.peek().clone() {
            Token::String => { self.advance(); "String".into() }
            Token::Integer => { self.advance(); "Integer".into() }
            Token::Real => { self.advance(); "Real".into() }
            Token::Boolean => { self.advance(); "Boolean".into() }
            Token::Char => { self.advance(); "Char".into() }
            Token::Pointer => { self.advance(); "Pointer".into() }
            // Procedural types: procedure or function(...): Type
            Token::Procedure => {
                self.advance();
                // Skip optional params
                if self.peek() == &Token::LParen {
                    let mut depth = 1;
                    self.advance();
                    while depth > 0 && self.peek() != &Token::EOF {
                        match self.advance() {
                            Token::LParen => depth += 1,
                            Token::RParen => depth -= 1,
                            _ => {}
                        }
                    }
                }
                "Procedure".into()
            }
            Token::Function => {
                self.advance();
                if self.peek() == &Token::LParen {
                    let mut depth = 1;
                    self.advance();
                    while depth > 0 && self.peek() != &Token::EOF {
                        match self.advance() {
                            Token::LParen => depth += 1,
                            Token::RParen => depth -= 1,
                            _ => {}
                        }
                    }
                }
                if self.eat(&Token::Colon) { self.parse_type_ref()?; }
                "Function".into()
            }
            Token::Array => {
                self.advance();
                if self.eat(&Token::LBracket) {
                    while self.peek() != &Token::RBracket && self.peek() != &Token::EOF { self.advance(); }
                    self.eat(&Token::RBracket);
                }
                self.expect(Token::Of)?;
                let elem = self.parse_type_ref()?;
                return Ok(TypeRef { name: format!("Array<{}>", elem.name), generic: None });
            }
            Token::Identifier(s) => {
                self.advance();
                // Check for generic: TList<Integer>
                if self.peek() == &Token::Lt {
                    self.advance();
                    let generic = self.parse_type_ref()?;
                    self.expect(Token::Gt)?;
                    return Ok(TypeRef { name: s, generic: Some(Box::new(generic)) });
                }
                s
            }
            t => return Err(format!("Expected type, got {:?}", t)),
        };
        Ok(TypeRef::simple(&name))
    }

    fn parse_procedure(&mut self) -> Result<ProcDecl, String> {
        let name = self.expect_ident()?;
        let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
        self.eat(&Token::Semicolon);
        let is_forward = self.eat(&Token::Forward);
        if is_forward { self.eat(&Token::Semicolon); return Ok(ProcDecl { name, params, decls: vec![], body: vec![], is_forward: true }); }
        let decls = self.parse_decl_section()?;
        let body = self.parse_compound()?;
        self.eat(&Token::Semicolon);
        Ok(ProcDecl { name, params, decls, body, is_forward: false })
    }

    fn parse_function(&mut self) -> Result<FuncDecl, String> {
        let name = self.expect_ident()?;
        let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
        self.expect(Token::Colon)?;
        let return_type = self.parse_type_ref()?;
        self.eat(&Token::Semicolon);
        let is_forward = self.eat(&Token::Forward);
        if is_forward { self.eat(&Token::Semicolon); return Ok(FuncDecl { name, params, return_type, decls: vec![], body: vec![], is_forward: true }); }
        let decls = self.parse_decl_section()?;
        let body = self.parse_compound()?;
        self.eat(&Token::Semicolon);
        Ok(FuncDecl { name, params, return_type, decls, body, is_forward: false })
    }

    fn parse_params(&mut self) -> Result<Vec<Param>, String> {
        self.expect(Token::LParen)?;
        let mut params = Vec::new();
        while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
            let pass_by = match self.peek() {
                Token::Var => { self.advance(); PassBy::Var }
                Token::Const => { self.advance(); PassBy::Const }
                _ => {
                    if let Token::Identifier(s) = self.peek() {
                        if s.to_lowercase() == "out" { self.advance(); PassBy::Out }
                        else { PassBy::Value }
                    } else { PassBy::Value }
                }
            };
            let mut names = vec![self.expect_ident()?];
            while self.eat(&Token::Comma) {
                if let Token::Identifier(_) = self.peek() { names.push(self.expect_ident()?); }
                else { break; }
            }
            self.expect(Token::Colon)?;
            let type_name = self.parse_type_ref()?;
            let default = if self.eat(&Token::Eq) { Some(self.parse_expr()?) } else { None };
            params.push(Param { names, type_name, pass_by, default });
            if !self.eat(&Token::Semicolon) { break; }
        }
        self.expect(Token::RParen)?;
        Ok(params)
    }

    // ── Statements ────────────────────────────────────────────────────────────

    fn parse_compound(&mut self) -> Result<Vec<Statement>, String> {
        self.expect(Token::Begin)?;
        let mut stmts = Vec::new();
        while self.peek() != &Token::End && self.peek() != &Token::EOF {
            stmts.push(self.parse_statement()?);
            self.eat(&Token::Semicolon);
        }
        self.expect(Token::End)?;
        Ok(stmts)
    }

    fn parse_statement(&mut self) -> Result<Statement, String> {
        match self.peek().clone() {
            Token::Begin => Ok(Statement::Block(self.parse_compound()?)),
            Token::If => self.parse_if(),
            Token::For => self.parse_for(),
            Token::While => self.parse_while(),
            Token::Repeat => self.parse_repeat(),
            Token::Case => self.parse_case(),
            Token::With => self.parse_with(),
            Token::Try => self.parse_try(),
            Token::Raise => {
                self.advance();
                let e = if !matches!(self.peek(), Token::Semicolon | Token::End | Token::Else) {
                    Some(self.parse_expr()?)
                } else { None };
                Ok(Statement::Raise(e))
            }
            Token::Exit => {
                self.advance();
                let e = if self.eat(&Token::LParen) {
                    let v = self.parse_expr()?;
                    self.expect(Token::RParen)?;
                    Some(v)
                } else { None };
                Ok(Statement::Exit(e))
            }
            Token::Break => { self.advance(); Ok(Statement::Break) }
            Token::Continue => { self.advance(); Ok(Statement::Continue) }
            Token::Halt => {
                self.advance();
                if self.eat(&Token::LParen) { self.parse_expr()?; self.expect(Token::RParen)?; }
                Ok(Statement::Exit(None))
            }
            Token::Inherited => {
                self.advance();
                if let Token::Identifier(_) = self.peek() {
                    let method = self.expect_ident()?;
                    let args = if self.peek() == &Token::LParen { self.parse_arg_list()? } else { Vec::new() };
                    Ok(Statement::Call {
                        name: Expression::Inherited { method: Some(method), args: args.clone() },
                        args: Vec::new(),
                    })
                } else {
                    Ok(Statement::Call {
                        name: Expression::Inherited { method: None, args: Vec::new() },
                        args: Vec::new(),
                    })
                }
            }
            Token::Semicolon | Token::End | Token::Else | Token::Until | Token::EOF => Ok(Statement::Empty),
            _ => self.parse_assign_or_call(),
        }
    }

    fn parse_assign_or_call(&mut self) -> Result<Statement, String> {
        let expr = self.parse_expr()?;
        match self.peek() {
            Token::Assign => {
                self.advance();
                let value = self.parse_expr()?;
                Ok(Statement::Assign { target: expr, value })
            }
            Token::PlusAssign => {
                self.advance();
                let value = self.parse_expr()?;
                Ok(Statement::CompoundAssign { target: expr, op: CompoundOp::Add, value })
            }
            Token::MinusAssign => {
                self.advance();
                let value = self.parse_expr()?;
                Ok(Statement::CompoundAssign { target: expr, op: CompoundOp::Sub, value })
            }
            Token::StarAssign => {
                self.advance();
                let value = self.parse_expr()?;
                Ok(Statement::CompoundAssign { target: expr, op: CompoundOp::Mul, value })
            }
            Token::SlashAssign => {
                self.advance();
                let value = self.parse_expr()?;
                Ok(Statement::CompoundAssign { target: expr, op: CompoundOp::Div, value })
            }
            _ => {
                // Expression statement (procedure call)
                Ok(Statement::Call { name: expr, args: Vec::new() })
            }
        }
    }

    fn parse_if(&mut self) -> Result<Statement, String> {
        self.advance();
        let cond = self.parse_expr()?;
        self.expect(Token::Then)?;
        let then = Box::new(self.parse_statement()?);
        let else_ = if self.eat(&Token::Else) { Some(Box::new(self.parse_statement()?)) } else { None };
        Ok(Statement::If { cond, then, else_ })
    }

    fn parse_for(&mut self) -> Result<Statement, String> {
        self.advance();
        let var = self.expect_ident()?;

        // for..in: `for item in collection do`
        if self.eat(&Token::In) {
            let collection = self.parse_expr()?;
            self.expect(Token::Do)?;
            let body = Box::new(self.parse_statement()?);
            return Ok(Statement::ForIn { var, collection, body });
        }

        self.expect(Token::Assign)?;
        let from = self.parse_expr()?;
        let downto = if self.eat(&Token::Downto) { true } else { self.expect(Token::To)?; false };
        let to = self.parse_expr()?;
        self.expect(Token::Do)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::For { var, from, to, downto, body })
    }

    fn parse_while(&mut self) -> Result<Statement, String> {
        self.advance();
        let cond = self.parse_expr()?;
        self.expect(Token::Do)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::While { cond, body })
    }

    fn parse_repeat(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut body = Vec::new();
        while self.peek() != &Token::Until && self.peek() != &Token::EOF {
            body.push(self.parse_statement()?);
            self.eat(&Token::Semicolon);
        }
        self.expect(Token::Until)?;
        let until = self.parse_expr()?;
        Ok(Statement::Repeat { body, until })
    }

    fn parse_case(&mut self) -> Result<Statement, String> {
        self.advance();
        let expr = self.parse_expr()?;
        self.expect(Token::Of)?;
        let mut arms = Vec::new();
        let mut else_ = None;
        loop {
            match self.peek() {
                Token::End | Token::EOF => break,
                Token::Else | Token::Otherwise => {
                    self.advance(); self.eat(&Token::Colon);
                    let mut stmts = Vec::new();
                    while !matches!(self.peek(), Token::End | Token::EOF) {
                        stmts.push(self.parse_statement()?);
                        self.eat(&Token::Semicolon);
                    }
                    else_ = Some(stmts);
                    break;
                }
                _ => {
                    let mut values = Vec::new();
                    loop {
                        let lo = self.parse_expr()?;
                        let v = if self.eat(&Token::DotDot) {
                            let hi = self.parse_expr()?;
                            CaseValue::Range(lo, hi)
                        } else { CaseValue::Single(lo) };
                        values.push(v);
                        if !self.eat(&Token::Comma) { break; }
                    }
                    self.expect(Token::Colon)?;
                    let mut body = Vec::new();
                    if self.peek() == &Token::Begin {
                        body = self.parse_compound()?;
                    } else {
                        body.push(self.parse_statement()?);
                    }
                    self.eat(&Token::Semicolon);
                    arms.push(CaseArm { values, body });
                }
            }
        }
        self.expect(Token::End)?;
        Ok(Statement::Case { expr, arms, else_ })
    }

    fn parse_with(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut vars = vec![self.parse_expr()?];
        while self.eat(&Token::Comma) { vars.push(self.parse_expr()?); }
        self.expect(Token::Do)?;
        let body = Box::new(self.parse_statement()?);
        Ok(Statement::With { vars, body })
    }

    fn parse_try(&mut self) -> Result<Statement, String> {
        self.advance();
        let mut body = Vec::new();
        while !matches!(self.peek(), Token::Except | Token::Finally | Token::EOF) {
            body.push(self.parse_statement()?);
            self.eat(&Token::Semicolon);
        }
        let handler = if self.eat(&Token::Except) {
            let mut clauses = Vec::new();
            let mut else_stmts = None;
            while self.peek() == &Token::On {
                self.advance();
                let var_name = if let Token::Identifier(_) = self.peek() {
                    if self.peek2() == &Token::Colon {
                        let v = self.expect_ident()?;
                        self.advance();
                        Some(v)
                    } else { None }
                } else { None };
                let on_type = Some(self.expect_ident()?);
                self.expect(Token::Do)?;
                let cb = vec![self.parse_statement()?];
                self.eat(&Token::Semicolon);
                clauses.push(ExceptClause { on_type, var_name, body: cb });
            }
            if clauses.is_empty() {
                let mut stmts = Vec::new();
                while !matches!(self.peek(), Token::End | Token::EOF) {
                    stmts.push(self.parse_statement()?);
                    self.eat(&Token::Semicolon);
                }
                else_stmts = Some(stmts);
            }
            self.expect(Token::End)?;
            TryHandler::Except(clauses, else_stmts)
        } else {
            self.expect(Token::Finally)?;
            let mut stmts = Vec::new();
            while !matches!(self.peek(), Token::End | Token::EOF) {
                stmts.push(self.parse_statement()?);
                self.eat(&Token::Semicolon);
            }
            self.expect(Token::End)?;
            TryHandler::Finally(stmts)
        };
        Ok(Statement::Try { body, handler })
    }

    // ── Expressions ───────────────────────────────────────────────────────────

    fn parse_expr(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_relational()?;
        // is / as operators (lower precedence than relational)
        loop {
            match self.peek() {
                Token::Is => {
                    self.advance();
                    let type_name = self.expect_ident()?;
                    left = Expression::IsCheck { expr: Box::new(left), type_name };
                }
                Token::As => {
                    self.advance();
                    let type_name = self.expect_ident()?;
                    left = Expression::AsCast { expr: Box::new(left), type_name };
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_relational(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_additive()?;
        loop {
            let op = match self.peek() {
                Token::Eq => BinOp::Eq,
                Token::NotEq => BinOp::NotEq,
                Token::Lt => BinOp::Lt,
                Token::Gt => BinOp::Gt,
                Token::Le => BinOp::Le,
                Token::Ge => BinOp::Ge,
                Token::In => BinOp::In,
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
                Token::Or => BinOp::Or,
                Token::Xor => BinOp::Xor,
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
                Token::Div => BinOp::IDiv,
                Token::Mod => BinOp::Mod,
                Token::And => BinOp::And,
                Token::Shl => BinOp::Shl,
                Token::Shr => BinOp::Shr,
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
            Token::Minus => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::Neg, expr: Box::new(e) }) }
            Token::Not => { self.advance(); let e = self.parse_unary()?; Ok(Expression::Unary { op: UnaryOp::Not, expr: Box::new(e) }) }
            Token::At => { self.advance(); let e = self.parse_unary()?; Ok(Expression::AddrOf(Box::new(e))) }
            _ => self.parse_postfix(),
        }
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            match self.peek() {
                Token::Dot => {
                    self.advance();
                    let field = self.expect_ident()?;
                    if self.peek() == &Token::LParen {
                        let args = self.parse_arg_list()?;
                        expr = Expression::Call {
                            callee: Box::new(Expression::Field { record: Box::new(expr), field }),
                            args,
                        };
                    } else {
                        expr = Expression::Field { record: Box::new(expr), field };
                    }
                }
                Token::LBracket => {
                    self.advance();
                    let idx = self.parse_expr()?;
                    self.expect(Token::RBracket)?;
                    expr = Expression::Index { array: Box::new(expr), index: Box::new(idx) };
                }
                Token::Caret => {
                    self.advance();
                    expr = Expression::Deref(Box::new(expr));
                }
                Token::LParen => {
                    let args = self.parse_arg_list()?;
                    expr = Expression::Call { callee: Box::new(expr), args };
                }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            Token::IntLiteral(n) => { self.advance(); Ok(Expression::Int(n)) }
            Token::RealLiteral(n) => { self.advance(); Ok(Expression::Real(n)) }
            Token::True => { self.advance(); Ok(Expression::Bool(true)) }
            Token::False => { self.advance(); Ok(Expression::Bool(false)) }
            Token::Nil => { self.advance(); Ok(Expression::Nil) }
            Token::StringLiteral(s) => { self.advance(); Ok(Expression::Str(s)) }
            Token::CharLiteral(c) => { self.advance(); Ok(Expression::Char(c)) }
            Token::LParen => {
                self.advance();
                let e = self.parse_expr()?;
                self.expect(Token::RParen)?;
                Ok(e)
            }
            Token::LBracket => {
                self.advance();
                let mut elems = Vec::new();
                while self.peek() != &Token::RBracket && self.peek() != &Token::EOF {
                    elems.push(self.parse_expr()?);
                    if !self.eat(&Token::Comma) { break; }
                }
                self.expect(Token::RBracket)?;
                Ok(Expression::SetLiteral(elems))
            }
            Token::New => {
                self.advance();
                let name = self.expect_ident()?;
                let args = if self.peek() == &Token::LParen { self.parse_arg_list()? } else { Vec::new() };
                Ok(Expression::Call { callee: Box::new(Expression::Identifier(name)), args })
            }
            // Anonymous procedure: `procedure(params) begin ... end`
            Token::Procedure => {
                self.advance();
                let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
                let body = self.parse_compound()?;
                Ok(Expression::Lambda { params, return_type: None, body })
            }
            // Anonymous function: `function(params): Type begin ... end`
            Token::Function => {
                self.advance();
                let params = if self.peek() == &Token::LParen { self.parse_params()? } else { Vec::new() };
                let return_type = if self.eat(&Token::Colon) {
                    Some(self.parse_type_ref()?)
                } else { None };
                let body = self.parse_compound()?;
                Ok(Expression::Lambda { params, return_type, body })
            }
            Token::Identifier(name) => {
                self.advance();
                Ok(Expression::Identifier(name))
            }
            Token::Result => { self.advance(); Ok(Expression::Identifier("Result".into())) }
            // Type cast: Integer(x), String(x), etc.
            Token::Integer | Token::Real | Token::Boolean | Token::Char | Token::String => {
                let type_name = format!("{:?}", self.advance());
                if self.peek() == &Token::LParen {
                    self.advance();
                    let e = self.parse_expr()?;
                    self.expect(Token::RParen)?;
                    Ok(Expression::Cast { type_name, expr: Box::new(e) })
                } else {
                    Ok(Expression::Identifier(type_name))
                }
            }
            t => Err(format!("Unexpected token in expression: {:?}", t)),
        }
    }

    fn parse_arg_list(&mut self) -> Result<Vec<Expression>, String> {
        self.expect(Token::LParen)?;
        let mut args = Vec::new();
        while self.peek() != &Token::RParen && self.peek() != &Token::EOF {
            args.push(self.parse_expr()?);
            if !self.eat(&Token::Comma) { break; }
        }
        self.expect(Token::RParen)?;
        Ok(args)
    }
}
