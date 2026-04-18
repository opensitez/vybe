use super::ast::*;
use super::token::{Token, TokenKind};
use super::lexer::Lexer;

pub struct Parser {
    tokens: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(source: &str) -> Result<Self, String> {
        let tokens = Lexer::new(source).tokenize()?;
        Ok(Parser { tokens, pos: 0 })
    }

    // ------------------------------------------------------------------
    // Helpers
    // ------------------------------------------------------------------

    fn peek(&self) -> &TokenKind {
        &self.tokens[self.pos].kind
    }

    fn peek2(&self) -> &TokenKind {
        if self.pos + 1 < self.tokens.len() { &self.tokens[self.pos + 1].kind }
        else { &TokenKind::Eof }
    }

    fn line(&self) -> u32 {
        self.tokens[self.pos].line
    }

    fn advance(&mut self) -> &TokenKind {
        let k = &self.tokens[self.pos].kind;
        if self.pos + 1 < self.tokens.len() { self.pos += 1; }
        k
    }

    #[allow(dead_code)]
    fn check(&self, kind: &TokenKind) -> bool {
        self.peek() == kind
    }

    fn eat(&mut self, kind: &TokenKind) -> bool {
        if self.peek() == kind { self.advance(); true } else { false }
    }

    fn expect(&mut self, kind: &TokenKind) -> Result<(), String> {
        if self.peek() == kind {
            self.advance();
            Ok(())
        } else {
            Err(format!("line {}: expected {:?}, got {:?}", self.line(), kind, self.peek()))
        }
    }

    fn expect_ident(&mut self) -> Result<String, String> {
        match self.peek().clone() {
            TokenKind::Identifier(s) => { self.advance(); Ok(s) }
            // Also accept keywords used as identifiers in some positions
            k => {
                if let Some(s) = self.keyword_as_ident(&k) {
                    self.advance();
                    Ok(s)
                } else {
                    Err(format!("line {}: expected identifier, got {:?}", self.line(), k))
                }
            }
        }
    }

    fn keyword_as_ident(&self, k: &TokenKind) -> Option<String> {
        match k {
            TokenKind::Identifier(s) => Some(s.clone()),
            TokenKind::Static => Some("static".into()),
            TokenKind::Abstract => Some("abstract".into()),
            TokenKind::Final => Some("final".into()),
            TokenKind::Match => Some("match".into()),
            TokenKind::List => Some("list".into()),
            TokenKind::Fn => Some("fn".into()),
            _ => None,
        }
    }

    fn expect_var(&mut self) -> Result<String, String> {
        match self.peek().clone() {
            TokenKind::Variable(s) => { self.advance(); Ok(s) }
            k => Err(format!("line {}: expected variable, got {:?}", self.line(), k)),
        }
    }

    fn is_at_end(&self) -> bool {
        matches!(self.peek(), TokenKind::Eof)
    }

    // ------------------------------------------------------------------
    // Program
    // ------------------------------------------------------------------

    pub fn parse_program(&mut self) -> Result<Program, String> {
        let mut body = Vec::new();
        while !self.is_at_end() {
            // Skip stray namespace declarations (just ignore for now)
            if matches!(self.peek(), TokenKind::Namespace) {
                self.advance();
                // consume namespace path
                while matches!(self.peek(), TokenKind::Identifier(_) | TokenKind::Backslash) {
                    self.advance();
                }
                self.eat(&TokenKind::Semicolon);
                continue;
            }
            // Skip use namespace imports
            if matches!(self.peek(), TokenKind::Use) {
                self.advance();
                while !matches!(self.peek(), TokenKind::Semicolon | TokenKind::Eof) {
                    self.advance();
                }
                self.eat(&TokenKind::Semicolon);
                continue;
            }
            if let Some(s) = self.parse_statement()? {
                body.push(s);
            }
        }
        Ok(Program { body })
    }

    // ------------------------------------------------------------------
    // Statements
    // ------------------------------------------------------------------

    fn parse_statement(&mut self) -> Result<Option<Statement>, String> {
        match self.peek().clone() {
            TokenKind::Semicolon => { self.advance(); Ok(Some(Statement::Empty)) }
            TokenKind::LBrace => {
                self.advance();
                let block = self.parse_block_body()?;
                Ok(Some(Statement::Block(block)))
            }
            TokenKind::Echo => {
                self.advance();
                let mut exprs = vec![self.parse_expression()?];
                while self.eat(&TokenKind::Comma) {
                    exprs.push(self.parse_expression()?);
                }
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Echo(exprs)))
            }
            TokenKind::Print => {
                self.advance();
                let expr = self.parse_expression()?;
                self.eat(&TokenKind::Semicolon);
                Ok(Some(Statement::Echo(vec![expr])))
            }
            TokenKind::Return => {
                self.advance();
                if self.eat(&TokenKind::Semicolon) {
                    return Ok(Some(Statement::Return(None)));
                }
                let expr = self.parse_expression()?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Return(Some(expr))))
            }
            TokenKind::Break => {
                self.advance();
                let val = if !matches!(self.peek(), TokenKind::Semicolon) {
                    Some(self.parse_expression()?)
                } else { None };
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Break(val)))
            }
            TokenKind::Continue => {
                self.advance();
                let val = if !matches!(self.peek(), TokenKind::Semicolon) {
                    Some(self.parse_expression()?)
                } else { None };
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Continue(val)))
            }
            TokenKind::If => {
                self.advance();
                Ok(Some(self.parse_if()?))
            }
            TokenKind::While => {
                self.advance();
                self.expect(&TokenKind::LParen)?;
                let test = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                let body = Box::new(self.parse_statement_required()?);
                Ok(Some(Statement::While { test, body }))
            }
            TokenKind::Do => {
                self.advance();
                let body = Box::new(self.parse_statement_required()?);
                self.expect(&TokenKind::While)?;
                self.expect(&TokenKind::LParen)?;
                let test = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::DoWhile { body, test }))
            }
            TokenKind::For => {
                self.advance();
                Ok(Some(self.parse_for()?))
            }
            TokenKind::ForEach => {
                self.advance();
                Ok(Some(self.parse_foreach()?))
            }
            TokenKind::Switch => {
                self.advance();
                Ok(Some(self.parse_switch()?))
            }
            TokenKind::Throw => {
                self.advance();
                let expr = self.parse_expression()?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Throw(expr)))
            }
            TokenKind::Try => {
                self.advance();
                Ok(Some(self.parse_try()?))
            }
            TokenKind::Function => {
                self.advance();
                let decl = self.parse_function_decl(Visibility::None, false)?;
                Ok(Some(Statement::FunctionDeclaration(decl)))
            }
            TokenKind::Class | TokenKind::Abstract | TokenKind::Final | TokenKind::Readonly => {
                // Consume modifiers: abstract, final, readonly
                while matches!(self.peek(), TokenKind::Abstract | TokenKind::Final | TokenKind::Readonly) {
                    self.advance();
                }
                self.expect(&TokenKind::Class)?;
                let decl = self.parse_class_decl()?;
                Ok(Some(Statement::ClassDeclaration(decl)))
            }
            TokenKind::Interface => {
                // Parse interface as a class with no body for now
                self.advance();
                let name = self.expect_ident()?;
                // skip extends list
                if self.eat(&TokenKind::Extends) {
                    self.expect_ident()?;
                    while self.eat(&TokenKind::Comma) { self.expect_ident()?; }
                }
                self.expect(&TokenKind::LBrace)?;
                // consume interface body
                let mut depth = 1;
                while depth > 0 && !self.is_at_end() {
                    match self.advance() {
                        TokenKind::LBrace => depth += 1,
                        TokenKind::RBrace => depth -= 1,
                        _ => {}
                    }
                }
                Ok(Some(Statement::ClassDeclaration(ClassDecl {
                    name,
                    parent: None,
                    interfaces: Vec::new(),
                    traits: Vec::new(),
                    members: Vec::new(),
                })))
            }
            TokenKind::Enum => {
                // enum Color { case Red; case Green; }
                // enum Suit: string { case Hearts = 'H'; }
                // Compiled as a class with static constants
                self.advance();
                let name = self.expect_ident()?;
                // Optional backing type: enum Suit: string
                if self.eat(&TokenKind::Colon) {
                    let _ = self.expect_ident()?; // skip type name (string, int)
                }
                // Optional implements
                let mut interfaces = Vec::new();
                if self.eat(&TokenKind::Implements) {
                    interfaces.push(self.expect_ident()?);
                    while self.eat(&TokenKind::Comma) { interfaces.push(self.expect_ident()?); }
                }
                self.expect(&TokenKind::LBrace)?;
                let mut members = Vec::new();
                while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
                    if self.eat(&TokenKind::Semicolon) { continue; }
                    if matches!(self.peek(), TokenKind::Case) {
                        self.advance();
                        let case_name = self.expect_ident()?;
                        let value = if self.eat(&TokenKind::Eq) {
                            Some(self.parse_expression()?)
                        } else {
                            // Auto-assign: use case name as string value
                            Some(Expression::Str(case_name.clone()))
                        };
                        self.eat(&TokenKind::Semicolon);
                        members.push(ClassMember::Constant { name: case_name, value: value.unwrap() });
                    } else if let Some(m) = self.parse_class_member()? {
                        members.push(m);
                    }
                }
                self.expect(&TokenKind::RBrace)?;
                Ok(Some(Statement::ClassDeclaration(ClassDecl {
                    name, parent: None, interfaces, traits: Vec::new(), members,
                })))
            }
            TokenKind::Trait => {
                // Parse trait as a class — same structure (methods + properties)
                // Compiled as a regular class; `use TraitName` in another class
                // will be handled by inheriting its methods.
                self.advance();
                let decl = self.parse_class_decl()?;
                Ok(Some(Statement::ClassDeclaration(decl)))
            }
            TokenKind::Const => {
                self.advance();
                let name = self.expect_ident()?;
                self.expect(&TokenKind::Eq)?;
                let value = self.parse_expression()?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::ConstDeclaration { name, value }))
            }
            TokenKind::Global => {
                self.advance();
                let mut vars = vec![self.expect_var()?];
                while self.eat(&TokenKind::Comma) {
                    vars.push(self.expect_var()?);
                }
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Global(vars)))
            }
            // Visibility modifiers at statement level = class member in global context
            // We just handle it by falling through to expression statement
            _ => {
                let expr = self.parse_expression()?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(Statement::Expression(expr)))
            }
        }
    }

    fn parse_statement_required(&mut self) -> Result<Statement, String> {
        match self.parse_statement()? {
            Some(s) => Ok(s),
            None => Err(format!("line {}: expected statement", self.line())),
        }
    }

    fn parse_block_body(&mut self) -> Result<Vec<Statement>, String> {
        let mut stmts = Vec::new();
        while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
            if let Some(s) = self.parse_statement()? {
                stmts.push(s);
            }
        }
        self.expect(&TokenKind::RBrace)?;
        Ok(stmts)
    }

    fn parse_if(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::LParen)?;
        let test = self.parse_expression()?;
        self.expect(&TokenKind::RParen)?;
        let consequent = Box::new(self.parse_statement_required()?);
        let mut alternates = Vec::new();
        let mut alternate = None;
        loop {
            if matches!(self.peek(), TokenKind::ElseIf) {
                self.advance();
                self.expect(&TokenKind::LParen)?;
                let t = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                let body = Box::new(self.parse_statement_required()?);
                alternates.push(ElseIf { test: t, body });
            } else if matches!(self.peek(), TokenKind::Else) {
                self.advance();
                // Handle `else if` (two tokens)
                if matches!(self.peek(), TokenKind::If) {
                    self.advance();
                    self.expect(&TokenKind::LParen)?;
                    let t = self.parse_expression()?;
                    self.expect(&TokenKind::RParen)?;
                    let body = Box::new(self.parse_statement_required()?);
                    alternates.push(ElseIf { test: t, body });
                } else {
                    alternate = Some(Box::new(self.parse_statement_required()?));
                    break;
                }
            } else {
                break;
            }
        }
        Ok(Statement::If { test, consequent, alternates, alternate })
    }

    fn parse_for(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::LParen)?;
        let mut init = Vec::new();
        if !self.eat(&TokenKind::Semicolon) {
            init.push(self.parse_expression()?);
            while self.eat(&TokenKind::Comma) {
                init.push(self.parse_expression()?);
            }
            self.expect(&TokenKind::Semicolon)?;
        }
        let test = if self.eat(&TokenKind::Semicolon) {
            None
        } else {
            let e = self.parse_expression()?;
            self.expect(&TokenKind::Semicolon)?;
            Some(e)
        };
        let mut update = Vec::new();
        if !matches!(self.peek(), TokenKind::RParen) {
            update.push(self.parse_expression()?);
            while self.eat(&TokenKind::Comma) {
                update.push(self.parse_expression()?);
            }
        }
        self.expect(&TokenKind::RParen)?;
        let body = Box::new(self.parse_statement_required()?);
        Ok(Statement::For { init, test, update, body })
    }

    fn parse_foreach(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::LParen)?;
        let array = self.parse_expression()?;
        self.expect(&TokenKind::As)?;
        // foreach ($arr as $val) or foreach ($arr as $key => $val)
        let first_var = self.expect_var()?;
        let (key, value) = if self.eat(&TokenKind::FatArrow) {
            let val = self.expect_var()?;
            (Some(first_var), val)
        } else {
            (None, first_var)
        };
        self.expect(&TokenKind::RParen)?;
        let body = Box::new(self.parse_statement_required()?);
        Ok(Statement::ForEach { array, key, value, body })
    }

    fn parse_switch(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::LParen)?;
        let discriminant = self.parse_expression()?;
        self.expect(&TokenKind::RParen)?;
        self.expect(&TokenKind::LBrace)?;
        let mut cases = Vec::new();
        while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
            let test = if self.eat(&TokenKind::Case) {
                let e = self.parse_expression()?;
                self.expect(&TokenKind::Colon)?;
                Some(e)
            } else if self.eat(&TokenKind::Default) {
                self.expect(&TokenKind::Colon)?;
                None
            } else {
                break;
            };
            let mut body = Vec::new();
            while !matches!(self.peek(), TokenKind::Case | TokenKind::Default | TokenKind::RBrace | TokenKind::Eof) {
                if let Some(s) = self.parse_statement()? {
                    body.push(s);
                }
            }
            cases.push(SwitchCase { test, body });
        }
        self.expect(&TokenKind::RBrace)?;
        Ok(Statement::Switch { discriminant, cases })
    }

    fn parse_try(&mut self) -> Result<Statement, String> {
        self.expect(&TokenKind::LBrace)?;
        let block = self.parse_block_body()?;
        let mut catches = Vec::new();
        while matches!(self.peek(), TokenKind::Catch) {
            self.advance();
            self.expect(&TokenKind::LParen)?;
            let mut types = Vec::new();
            types.push(self.expect_ident()?);
            while self.eat(&TokenKind::Pipe) {
                types.push(self.expect_ident()?);
            }
            let var = if matches!(self.peek(), TokenKind::Variable(_)) {
                Some(self.expect_var()?)
            } else { None };
            self.expect(&TokenKind::RParen)?;
            self.expect(&TokenKind::LBrace)?;
            let body = self.parse_block_body()?;
            catches.push(CatchClause { types, var, body });
        }
        let finalizer = if self.eat(&TokenKind::Finally) {
            self.expect(&TokenKind::LBrace)?;
            Some(self.parse_block_body()?)
        } else { None };
        Ok(Statement::Try { block, catches, finalizer })
    }

    fn parse_function_decl(&mut self, visibility: Visibility, is_static: bool) -> Result<FunctionDecl, String> {
        let return_by_ref = self.eat(&TokenKind::Amp);
        let name = self.expect_ident()?;
        let params = self.parse_params()?;
        // optional return type hint
        if self.eat(&TokenKind::Colon) {
            self.skip_type_hint();
        }
        // Abstract methods have no body — just a semicolon
        if self.eat(&TokenKind::Semicolon) {
            return Ok(FunctionDecl { name, params, body: Vec::new(), is_static, visibility, return_by_ref });
        }
        self.expect(&TokenKind::LBrace)?;
        let body = self.parse_block_body()?;
        Ok(FunctionDecl { name, params, body, is_static, visibility, return_by_ref })
    }

    fn parse_params(&mut self) -> Result<Vec<Param>, String> {
        self.expect(&TokenKind::LParen)?;
        let mut params = Vec::new();
        while !matches!(self.peek(), TokenKind::RParen | TokenKind::Eof) {
            // Skip constructor promotion modifiers: public/protected/private/readonly
            while matches!(self.peek(), TokenKind::Public | TokenKind::Protected | TokenKind::Private | TokenKind::Readonly) {
                self.advance();
            }
            // optional type hint BEFORE ... and & — handle int ...$nums, ?Type &$ref
            let type_hint = if self.has_type_hint_ahead() {
                let start = self.pos;
                self.skip_type_hint();
                // Reconstruct type name from consumed tokens (approximation for AST)
                let end = self.pos;
                let th: String = (start..end).map(|i| {
                    match &self.tokens[i].kind {
                        TokenKind::Identifier(s) => s.clone(),
                        TokenKind::Question => "?".into(),
                        TokenKind::Pipe => "|".into(),
                        TokenKind::Amp => "&".into(),
                        TokenKind::Backslash => "\\".into(),
                        TokenKind::LParen => "(".into(),
                        TokenKind::RParen => ")".into(),
                        TokenKind::Null => "null".into(),
                        TokenKind::True => "true".into(),
                        TokenKind::False => "false".into(),
                        TokenKind::Static => "static".into(),
                        _ => String::new(),
                    }
                }).collect();
                if th.is_empty() { None } else { Some(th) }
            } else { None };
            // ... and & come after type hint: int ...$nums, ?string &$ref
            let variadic = self.eat(&TokenKind::Ellipsis);
            let by_ref = self.eat(&TokenKind::Amp);
            let name = self.expect_var()?;
            let default = if self.eat(&TokenKind::Eq) {
                Some(self.parse_expression()?)
            } else { None };
            params.push(Param { name, default, by_ref, variadic, type_hint });
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect(&TokenKind::RParen)?;
        Ok(params)
    }

    /// Check if the current position has a type hint followed by a $variable
    /// Check if current position is `( ... )` — first-class callable syntax
    fn is_first_class_callable(&self) -> bool {
        self.pos + 2 < self.tokens.len()
            && self.tokens[self.pos].kind == TokenKind::LParen
            && self.tokens[self.pos + 1].kind == TokenKind::Ellipsis
            && self.tokens[self.pos + 2].kind == TokenKind::RParen
    }

    fn has_type_hint_ahead(&self) -> bool {
        // Look ahead: if we see Identifier/? followed eventually by Variable, it's a type hint
        let mut i = self.pos;
        // skip ?
        if i < self.tokens.len() && self.tokens[i].kind == TokenKind::Question { i += 1; }
        // Handle DNF parenthesized groups: (A&B)|C
        if i < self.tokens.len() && self.tokens[i].kind == TokenKind::LParen {
            let mut depth = 1;
            i += 1;
            while i < self.tokens.len() && depth > 0 {
                match &self.tokens[i].kind {
                    TokenKind::LParen => depth += 1,
                    TokenKind::RParen => depth -= 1,
                    _ => {}
                }
                i += 1;
            }
            // After closing ), skip |Type chains
            while i < self.tokens.len() {
                match &self.tokens[i].kind {
                    TokenKind::Pipe | TokenKind::Amp => { i += 1; }
                    TokenKind::Identifier(_) | TokenKind::Static | TokenKind::Null => { i += 1; }
                    TokenKind::LParen => {
                        let mut d2 = 1; i += 1;
                        while i < self.tokens.len() && d2 > 0 {
                            match &self.tokens[i].kind { TokenKind::LParen => d2 += 1, TokenKind::RParen => d2 -= 1, _ => {} }
                            i += 1;
                        }
                    }
                    _ => break,
                }
            }
            if i < self.tokens.len() {
                return matches!(&self.tokens[i].kind, TokenKind::Variable(_) | TokenKind::Amp | TokenKind::Ellipsis);
            }
            return false;
        }
        // need at least one identifier
        if i >= self.tokens.len() { return false; }
        match &self.tokens[i].kind {
            TokenKind::Identifier(_) | TokenKind::Static | TokenKind::Null => {}
            _ => return false,
        }
        i += 1;
        // skip |Type, &Type, \Namespace chains, (groups)
        while i < self.tokens.len() {
            match &self.tokens[i].kind {
                TokenKind::Pipe | TokenKind::Amp | TokenKind::Backslash => { i += 1; }
                TokenKind::Identifier(_) | TokenKind::Static | TokenKind::Null => { i += 1; }
                TokenKind::LParen => {
                    let mut d2 = 1; i += 1;
                    while i < self.tokens.len() && d2 > 0 {
                        match &self.tokens[i].kind { TokenKind::LParen => d2 += 1, TokenKind::RParen => d2 -= 1, _ => {} }
                        i += 1;
                    }
                }
                _ => break,
            }
        }
        // Must be followed by $var, &$var, or ...$var
        if i < self.tokens.len() {
            matches!(&self.tokens[i].kind, TokenKind::Variable(_) | TokenKind::Amp | TokenKind::Ellipsis)
        } else { false }
    }

    fn skip_type_hint(&mut self) {
        // consume nullable marker
        self.eat(&TokenKind::Question);
        // consume type names, | separators, & separators, and (group) for DNF types
        loop {
            match self.peek() {
                TokenKind::Identifier(_) | TokenKind::Static | TokenKind::Null | TokenKind::True | TokenKind::False => {
                    self.advance();
                }
                TokenKind::Backslash => { self.advance(); }
                TokenKind::Pipe | TokenKind::Amp => { self.advance(); }
                TokenKind::LParen => {
                    // DNF type group: (A&B)|C
                    self.advance();
                    self.skip_type_hint(); // recurse for inner type
                    self.eat(&TokenKind::RParen);
                }
                _ => break,
            }
        }
    }

    /// Parse anonymous class body: { members } with optional extends/implements
    fn parse_class_decl_body(&mut self, name: String) -> Result<ClassDecl, String> {
        let parent = if self.eat(&TokenKind::Extends) {
            Some(self.expect_ident()?)
        } else { None };
        let mut interfaces = Vec::new();
        if self.eat(&TokenKind::Implements) {
            interfaces.push(self.expect_ident()?);
            while self.eat(&TokenKind::Comma) { interfaces.push(self.expect_ident()?); }
        }
        self.expect(&TokenKind::LBrace)?;
        let mut members = Vec::new();
        let mut traits = Vec::new();
        while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
            if self.eat(&TokenKind::Semicolon) { continue; }
            if matches!(self.peek(), TokenKind::Use) {
                self.advance();
                if let Ok(n) = self.expect_ident() { traits.push(n); while self.eat(&TokenKind::Comma) { if let Ok(n) = self.expect_ident() { traits.push(n); } } }
                if self.eat(&TokenKind::LBrace) { let mut d=1; while d>0 && !self.is_at_end() { match self.advance() { TokenKind::LBrace=>d+=1, TokenKind::RBrace=>d-=1, _=>{} } } } else { self.eat(&TokenKind::Semicolon); }
            } else if let Some(m) = self.parse_class_member()? { members.push(m); }
        }
        self.expect(&TokenKind::RBrace)?;
        Ok(ClassDecl { name, parent, interfaces, traits, members })
    }

    fn parse_class_decl(&mut self) -> Result<ClassDecl, String> {
        let name = self.expect_ident()?;
        let parent = if self.eat(&TokenKind::Extends) {
            Some(self.expect_ident()?)
        } else { None };
        let mut interfaces = Vec::new();
        if self.eat(&TokenKind::Implements) {
            interfaces.push(self.expect_ident()?);
            while self.eat(&TokenKind::Comma) {
                interfaces.push(self.expect_ident()?);
            }
        }
        self.expect(&TokenKind::LBrace)?;
        let mut members = Vec::new();
        let mut traits = Vec::new();
        while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
            match self.peek().clone() {
                TokenKind::Semicolon => { self.advance(); }
                TokenKind::Use => {
                    // `use TraitName, TraitName2;` — collect trait names
                    self.advance();
                    if let Ok(name) = self.expect_ident() {
                        traits.push(name);
                        while self.eat(&TokenKind::Comma) {
                            if let Ok(name) = self.expect_ident() {
                                traits.push(name);
                            }
                        }
                    }
                    // Handle `use Trait { ... }` conflict resolution block
                    if self.eat(&TokenKind::LBrace) {
                        let mut d = 1;
                        while d > 0 && !self.is_at_end() {
                            match self.advance() {
                                TokenKind::LBrace => d += 1,
                                TokenKind::RBrace => d -= 1,
                                _ => {}
                            }
                        }
                    } else {
                        self.eat(&TokenKind::Semicolon);
                    }
                }
                _ => {
                    if let Some(m) = self.parse_class_member()? {
                        members.push(m);
                    }
                }
            }
        }
        self.expect(&TokenKind::RBrace)?;
        Ok(ClassDecl { name, parent, interfaces, traits, members })
    }

    fn parse_class_member(&mut self) -> Result<Option<ClassMember>, String> {
        // Attributes (#[...]) are skipped at the lexer level
        // Consume modifiers
        let mut visibility = Visibility::None;
        let mut is_static = false;
        loop {
            match self.peek() {
                TokenKind::Public => { visibility = Visibility::Public; self.advance(); }
                TokenKind::Private => { visibility = Visibility::Private; self.advance(); }
                TokenKind::Protected => { visibility = Visibility::Protected; self.advance(); }
                TokenKind::Static => { is_static = true; self.advance(); }
                TokenKind::Abstract | TokenKind::Final | TokenKind::Readonly => { self.advance(); }
                _ => break,
            }
        }
        match self.peek().clone() {
            TokenKind::Function => {
                self.advance();
                let decl = self.parse_function_decl(visibility, is_static)?;
                Ok(Some(ClassMember::Method(decl)))
            }
            TokenKind::Const => {
                self.advance();
                // PHP 8.3: optional type hint before constant name: const string NAME = 'val'
                // Peek: if next is Identifier and next-next is also Identifier or Eq, skip type
                if matches!(self.peek(), TokenKind::Identifier(_)) && matches!(self.peek2(), TokenKind::Identifier(_)) {
                    self.advance(); // skip type hint
                }
                let name = self.expect_ident()?;
                self.expect(&TokenKind::Eq)?;
                let value = self.parse_expression()?;
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(ClassMember::Constant { name, value }))
            }
            TokenKind::Variable(name) => {
                let name = name.clone();
                self.advance();
                let default = if self.eat(&TokenKind::Eq) {
                    Some(self.parse_expression()?)
                } else { None };
                self.expect(&TokenKind::Semicolon)?;
                Ok(Some(ClassMember::Property { name, visibility, is_static, default }))
            }
            TokenKind::Identifier(_) => {
                // Could be a typed property: TypeHint $var;
                self.skip_type_hint();
                if let TokenKind::Variable(name) = self.peek().clone() {
                    let name = name.clone();
                    self.advance();
                    let default = if self.eat(&TokenKind::Eq) {
                        Some(self.parse_expression()?)
                    } else { None };
                    self.expect(&TokenKind::Semicolon)?;
                    Ok(Some(ClassMember::Property { name, visibility, is_static, default }))
                } else {
                    // Skip unknown member
                    while !matches!(self.peek(), TokenKind::Semicolon | TokenKind::RBrace | TokenKind::Eof) {
                        self.advance();
                    }
                    self.eat(&TokenKind::Semicolon);
                    Ok(None)
                }
            }
            // skip unknown
            _ => {
                self.advance();
                Ok(None)
            }
        }
    }

    // ------------------------------------------------------------------
    // Expressions — Pratt-style precedence climbing
    // ------------------------------------------------------------------

    pub fn parse_expression(&mut self) -> Result<Expression, String> {
        // yield [expr]  or  yield from expr
        if matches!(self.peek(), TokenKind::Yield) {
            self.advance();
            // yield from expr
            if matches!(self.peek(), TokenKind::Identifier(s) if s == "from") {
                self.advance();
                let expr = self.parse_assign()?;
                return Ok(Expression::YieldFrom(Box::new(expr)));
            }
            // yield expr  or  bare yield
            if matches!(self.peek(), TokenKind::Semicolon | TokenKind::RParen | TokenKind::RBracket | TokenKind::RBrace | TokenKind::Comma) {
                return Ok(Expression::Yield(None));
            }
            let expr = self.parse_assign()?;
            return Ok(Expression::Yield(Some(Box::new(expr))));
        }
        self.parse_assign()
    }

    fn parse_assign(&mut self) -> Result<Expression, String> {
        let left = self.parse_ternary()?;

        let op = match self.peek() {
            TokenKind::Eq => AssignOp::Assign,
            TokenKind::PlusEq => AssignOp::AddAssign,
            TokenKind::MinusEq => AssignOp::SubAssign,
            TokenKind::StarEq => AssignOp::MulAssign,
            TokenKind::SlashEq => AssignOp::DivAssign,
            TokenKind::PercentEq => AssignOp::ModAssign,
            TokenKind::DotEq => AssignOp::ConcatAssign,
            TokenKind::AmpAmpEq => AssignOp::AndAssign,
            TokenKind::PipePipeEq => AssignOp::OrAssign,
            TokenKind::QuestionQuestionEq => AssignOp::NullCoalesceAssign,
            _ => return Ok(left),
        };
        self.advance();
        let right = self.parse_assign()?;
        Ok(Expression::Assign { op, left: Box::new(left), right: Box::new(right) })
    }

    fn parse_ternary(&mut self) -> Result<Expression, String> {
        let expr = self.parse_null_coalesce()?;
        if self.eat(&TokenKind::Question) {
            if self.eat(&TokenKind::Colon) {
                // Elvis: $a ?: $b
                let alt = self.parse_ternary()?;
                return Ok(Expression::Ternary {
                    test: Box::new(expr),
                    consequent: None,
                    alternate: Box::new(alt),
                });
            }
            let cons = self.parse_expression()?;
            self.expect(&TokenKind::Colon)?;
            let alt = self.parse_ternary()?;
            return Ok(Expression::Ternary {
                test: Box::new(expr),
                consequent: Some(Box::new(cons)),
                alternate: Box::new(alt),
            });
        }
        Ok(expr)
    }

    fn parse_null_coalesce(&mut self) -> Result<Expression, String> {
        let left = self.parse_logical_or()?;
        if self.eat(&TokenKind::QuestionQuestion) {
            let right = self.parse_null_coalesce()?;
            return Ok(Expression::NullCoalesce { left: Box::new(left), right: Box::new(right) });
        }
        Ok(left)
    }

    fn parse_logical_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_logical_and()?;
        while matches!(self.peek(), TokenKind::PipePipe) {
            self.advance();
            let right = self.parse_logical_and()?;
            left = Expression::Binary { op: BinaryOp::Or, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_logical_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_or()?;
        while matches!(self.peek(), TokenKind::AmpAmp) {
            self.advance();
            let right = self.parse_bitwise_or()?;
            left = Expression::Binary { op: BinaryOp::And, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitwise_or(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_xor()?;
        while matches!(self.peek(), TokenKind::Pipe) {
            self.advance();
            let right = self.parse_bitwise_xor()?;
            left = Expression::Binary { op: BinaryOp::BitOr, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitwise_xor(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_bitwise_and()?;
        while matches!(self.peek(), TokenKind::Caret) {
            self.advance();
            let right = self.parse_bitwise_and()?;
            left = Expression::Binary { op: BinaryOp::BitXor, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_bitwise_and(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_equality()?;
        while matches!(self.peek(), TokenKind::Amp) {
            self.advance();
            let right = self.parse_equality()?;
            left = Expression::Binary { op: BinaryOp::BitAnd, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_equality(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_comparison()?;
        loop {
            let op = match self.peek() {
                TokenKind::EqEq => BinaryOp::Eq,
                TokenKind::BangEq => BinaryOp::Ne,
                TokenKind::EqEqEq => BinaryOp::SEq,
                TokenKind::BangEqEq => BinaryOp::SNe,
                TokenKind::InstanceOf => BinaryOp::InstanceOf,
                _ => break,
            };
            self.advance();
            let right = self.parse_comparison()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_comparison(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_shift()?;
        loop {
            let op = match self.peek() {
                TokenKind::Lt => BinaryOp::Lt,
                TokenKind::Gt => BinaryOp::Gt,
                TokenKind::LtEq => BinaryOp::Le,
                TokenKind::GtEq => BinaryOp::Ge,
                TokenKind::Spaceship => BinaryOp::Spaceship,
                _ => break,
            };
            self.advance();
            let right = self.parse_shift()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_shift(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_addition()?;
        loop {
            let op = match self.peek() {
                TokenKind::LtLt => BinaryOp::Shl,
                TokenKind::GtGt => BinaryOp::Shr,
                _ => break,
            };
            self.advance();
            let right = self.parse_addition()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_addition(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_multiplication()?;
        loop {
            let op = match self.peek() {
                TokenKind::Plus => BinaryOp::Add,
                TokenKind::Minus => BinaryOp::Sub,
                TokenKind::Dot => BinaryOp::Concat,
                _ => break,
            };
            self.advance();
            let right = self.parse_multiplication()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_multiplication(&mut self) -> Result<Expression, String> {
        let mut left = self.parse_instanceof()?;
        loop {
            let op = match self.peek() {
                TokenKind::Star => BinaryOp::Mul,
                TokenKind::Slash => BinaryOp::Div,
                TokenKind::Percent => BinaryOp::Mod,
                _ => break,
            };
            self.advance();
            let right = self.parse_instanceof()?;
            left = Expression::Binary { op, left: Box::new(left), right: Box::new(right) };
        }
        Ok(left)
    }

    fn parse_instanceof(&mut self) -> Result<Expression, String> {
        let left = self.parse_power()?;
        if matches!(self.peek(), TokenKind::InstanceOf) {
            self.advance();
            let right = self.parse_power()?;
            return Ok(Expression::Binary { op: BinaryOp::InstanceOf, left: Box::new(left), right: Box::new(right) });
        }
        Ok(left)
    }

    fn parse_power(&mut self) -> Result<Expression, String> {
        let base = self.parse_prefix()?;
        if self.eat(&TokenKind::StarStar) {
            let exp = self.parse_power()?; // right-associative
            return Ok(Expression::Binary { op: BinaryOp::Pow, left: Box::new(base), right: Box::new(exp) });
        }
        Ok(base)
    }

    fn parse_prefix(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            TokenKind::Bang => { self.advance(); let e = self.parse_prefix()?; Ok(Expression::Unary { op: UnaryOp::Not, expr: Box::new(e) }) }
            TokenKind::Tilde => { self.advance(); let e = self.parse_prefix()?; Ok(Expression::Unary { op: UnaryOp::BitNot, expr: Box::new(e) }) }
            TokenKind::Minus => { self.advance(); let e = self.parse_prefix()?; Ok(Expression::Unary { op: UnaryOp::Neg, expr: Box::new(e) }) }
            TokenKind::Plus => { self.advance(); let e = self.parse_prefix()?; Ok(Expression::Unary { op: UnaryOp::Pos, expr: Box::new(e) }) }
            TokenKind::PlusPlus => { self.advance(); let e = self.parse_postfix()?; Ok(Expression::PreUpdate { op: UpdateOp::Inc, expr: Box::new(e) }) }
            TokenKind::MinusMinus => { self.advance(); let e = self.parse_postfix()?; Ok(Expression::PreUpdate { op: UpdateOp::Dec, expr: Box::new(e) }) }
            TokenKind::At => { self.advance(); self.parse_prefix() } // @expr — suppress errors, just evaluate
            TokenKind::Clone => {
                // clone $obj — shallow copy
                self.advance();
                let expr = self.parse_prefix()?;
                // Compile as: __vybe_assign(new_empty_obj, $obj)
                Ok(Expression::Call {
                    callee: Box::new(Expression::Identifier("__clone".to_string())),
                    args: vec![Argument { value: expr, by_ref: false, spread: false, name: None }],
                })
            }
            // Casts: (int) (float) (string) (bool) (array) (object)
            TokenKind::LParen => {
                // peek ahead to see if it's a cast
                if let Some(cast) = self.try_parse_cast() {
                    let e = self.parse_prefix()?;
                    return Ok(Expression::Cast { cast, expr: Box::new(e) });
                }
                self.parse_postfix()
            }
            _ => self.parse_postfix(),
        }
    }

    fn try_parse_cast(&mut self) -> Option<CastKind> {
        // Look for (int) (float) (string) (bool) (array) (object)
        if self.pos + 2 >= self.tokens.len() { return None; }
        if self.tokens[self.pos].kind != TokenKind::LParen { return None; }
        let cast = match &self.tokens[self.pos + 1].kind {
            TokenKind::Identifier(s) => match s.to_lowercase().as_str() {
                "int" | "integer" => CastKind::Int,
                "float" | "double" | "real" => CastKind::Float,
                "string" => CastKind::String,
                "bool" | "boolean" => CastKind::Bool,
                "array" => CastKind::Array,
                "object" => CastKind::Object,
                _ => return None,
            },
            _ => return None,
        };
        if self.tokens[self.pos + 2].kind != TokenKind::RParen { return None; }
        self.pos += 3;
        Some(cast)
    }

    fn parse_postfix(&mut self) -> Result<Expression, String> {
        let mut expr = self.parse_primary()?;
        loop {
            match self.peek().clone() {
                TokenKind::Arrow | TokenKind::NullsafeArrow => {
                    let nullsafe = matches!(self.peek(), TokenKind::NullsafeArrow);
                    self.advance();
                    let prop = self.parse_member_name()?;
                    if matches!(self.peek(), TokenKind::LParen) {
                        // First-class callable: $obj->method(...)
                        if self.is_first_class_callable() {
                            self.advance(); self.advance(); self.advance(); // ( ... )
                            // Return property access — the method ref itself
                            expr = Expression::Property { object: Box::new(expr), name: Box::new(prop), nullsafe };
                        } else {
                            let args = self.parse_args()?;
                            expr = Expression::MethodCall { object: Box::new(expr), method: Box::new(prop), args, nullsafe };
                        }
                    } else {
                        expr = Expression::Property { object: Box::new(expr), name: Box::new(prop), nullsafe };
                    }
                }
                TokenKind::ColonColon => {
                    self.advance();
                    // Foo::class → string of class name
                    if matches!(self.peek(), TokenKind::Class) {
                        self.advance();
                        // expr is the class identifier — convert to string
                        if let Expression::Identifier(name) = &expr {
                            expr = Expression::Str(name.clone());
                        } else {
                            expr = Expression::Str("unknown".into());
                        }
                        continue;
                    }
                    let member = self.parse_member_name()?;
                    if matches!(self.peek(), TokenKind::LParen) {
                        if self.is_first_class_callable() {
                            self.advance(); self.advance(); self.advance();
                            expr = Expression::StaticAccess { class: Box::new(expr), member: Box::new(member) };
                        } else {
                            let args = self.parse_args()?;
                            expr = Expression::StaticCall { class: Box::new(expr), method: Box::new(member), args };
                        }
                    } else {
                        expr = Expression::StaticAccess { class: Box::new(expr), member: Box::new(member) };
                    }
                }
                TokenKind::LBracket => {
                    self.advance();
                    let idx = self.parse_expression()?;
                    self.expect(&TokenKind::RBracket)?;
                    expr = Expression::ArrayAccess { array: Box::new(expr), index: Box::new(idx) };
                }
                TokenKind::LBrace => {
                    // $str{0} — old PHP string index, treat as array access
                    self.advance();
                    let idx = self.parse_expression()?;
                    self.expect(&TokenKind::RBrace)?;
                    expr = Expression::ArrayAccess { array: Box::new(expr), index: Box::new(idx) };
                }
                TokenKind::LParen => {
                    // First-class callable: strlen(...) → just the function reference
                    if self.is_first_class_callable() {
                        self.advance(); self.advance(); self.advance(); // ( ... )
                    } else {
                        let args = self.parse_args()?;
                        expr = Expression::Call { callee: Box::new(expr), args };
                    }
                }
                TokenKind::PlusPlus => { self.advance(); expr = Expression::PostUpdate { op: UpdateOp::Inc, expr: Box::new(expr) }; }
                TokenKind::MinusMinus => { self.advance(); expr = Expression::PostUpdate { op: UpdateOp::Dec, expr: Box::new(expr) }; }
                _ => break,
            }
        }
        Ok(expr)
    }

    fn parse_member_name(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            TokenKind::Variable(s) => { self.advance(); Ok(Expression::Variable(s)) }
            TokenKind::LBrace => {
                self.advance();
                let e = self.parse_expression()?;
                self.expect(&TokenKind::RBrace)?;
                Ok(e)
            }
            _ => {
                let name = self.expect_ident()?;
                Ok(Expression::Identifier(name))
            }
        }
    }

    fn parse_args(&mut self) -> Result<Vec<Argument>, String> {
        self.expect(&TokenKind::LParen)?;
        let mut args = Vec::new();
        while !matches!(self.peek(), TokenKind::RParen | TokenKind::Eof) {
            let spread = self.eat(&TokenKind::Ellipsis);
            let by_ref = self.eat(&TokenKind::Amp);
            // Check for named argument: ident: value
            let (name, value) = if matches!(self.peek(), TokenKind::Identifier(_))
                && matches!(self.peek2(), TokenKind::Colon) {
                let n = self.expect_ident()?;
                self.advance(); // :
                let v = self.parse_expression()?;
                (Some(n), v)
            } else {
                (None, self.parse_expression()?)
            };
            args.push(Argument { value, by_ref, spread, name });
            if !self.eat(&TokenKind::Comma) { break; }
        }
        self.expect(&TokenKind::RParen)?;
        Ok(args)
    }

    fn parse_primary(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            TokenKind::Number(n) => { self.advance(); Ok(Expression::Number(n)) }
            TokenKind::Str(s) => { self.advance(); Ok(Expression::Str(s)) }
            TokenKind::True => { self.advance(); Ok(Expression::Bool(true)) }
            TokenKind::False => { self.advance(); Ok(Expression::Bool(false)) }
            TokenKind::Null => { self.advance(); Ok(Expression::Null) }
            TokenKind::Variable(name) => {
                let name = name.clone();
                if name == "this" {
                    self.advance();
                    return Ok(Expression::This);
                }
                self.advance();
                Ok(Expression::Variable(name))
            }
            // throw as expression (PHP 8.0): $x = $val ?? throw new Exception(...)
            TokenKind::Throw => {
                self.advance();
                let expr = self.parse_expression()?;
                // Wrap as a call to a throw-expression helper — at runtime this throws
                // Compile as: throw expr (same as statement, but returns never)
                Ok(Expression::Call {
                    callee: Box::new(Expression::Identifier("__throw".to_string())),
                    args: vec![Argument { value: expr, by_ref: false, spread: false, name: None }],
                })
            }
            TokenKind::New => {
                self.advance();
                // Anonymous class: new class { ... } or new class(...) { ... }
                if matches!(self.peek(), TokenKind::Class) {
                    self.advance();
                    let args = if matches!(self.peek(), TokenKind::LParen) {
                        self.parse_args()?
                    } else { Vec::new() };
                    // Parse the class body as an inline class declaration
                    let anon_name = format!("__anon_{}", self.pos);
                    let decl = self.parse_class_decl_body(anon_name)?;
                    // Compile as: declare class + instantiate
                    // For now, wrap as New with the anon class name
                    return Ok(Expression::New {
                        class: Box::new(Expression::Identifier(decl.name.clone())),
                        args,
                    });
                }
                let class = self.parse_class_name_expr()?;
                let args = if matches!(self.peek(), TokenKind::LParen) {
                    self.parse_args()?
                } else { Vec::new() };
                Ok(Expression::New { class: Box::new(class), args })
            }
            TokenKind::LParen => {
                self.advance();
                let e = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                Ok(e)
            }
            TokenKind::LBracket => {
                self.advance();
                let elements = self.parse_array_elements(&TokenKind::RBracket)?;
                self.expect(&TokenKind::RBracket)?;
                Ok(Expression::Array(elements))
            }
            TokenKind::Identifier(name) => {
                let name = name.clone();
                // array(...) constructor
                if name.to_lowercase() == "array" && matches!(self.peek2(), TokenKind::LParen) {
                    self.advance();
                    self.advance(); // (
                    let elements = self.parse_array_elements(&TokenKind::RParen)?;
                    self.expect(&TokenKind::RParen)?;
                    return Ok(Expression::Array(elements));
                }
                // isset / empty / unset
                if name.to_lowercase() == "isset" {
                    self.advance();
                    self.expect(&TokenKind::LParen)?;
                    let mut vars = vec![self.parse_expression()?];
                    while self.eat(&TokenKind::Comma) { vars.push(self.parse_expression()?); }
                    self.expect(&TokenKind::RParen)?;
                    return Ok(Expression::Isset(vars));
                }
                if name.to_lowercase() == "empty" {
                    self.advance();
                    self.expect(&TokenKind::LParen)?;
                    let e = self.parse_expression()?;
                    self.expect(&TokenKind::RParen)?;
                    return Ok(Expression::Empty(Box::new(e)));
                }
                if name.to_lowercase() == "unset" {
                    self.advance();
                    self.expect(&TokenKind::LParen)?;
                    let mut vars = vec![self.parse_expression()?];
                    while self.eat(&TokenKind::Comma) { vars.push(self.parse_expression()?); }
                    self.expect(&TokenKind::RParen)?;
                    return Ok(Expression::Unset(vars));
                }
                // self / static / parent as class keyword
                if matches!(name.as_str(), "self" | "static" | "parent") {
                    self.advance();
                    return Ok(Expression::ClassKeyword(name));
                }
                self.advance();
                Ok(Expression::Identifier(name))
            }
            TokenKind::Static => { self.advance(); Ok(Expression::ClassKeyword("static".into())) }
            TokenKind::Function => {
                self.advance();
                let params = self.parse_params()?;
                let uses = self.parse_use_clause()?;
                if self.eat(&TokenKind::Colon) { self.skip_type_hint(); }
                self.expect(&TokenKind::LBrace)?;
                let body = self.parse_block_body()?;
                Ok(Expression::Closure {
                    params,
                    uses,
                    body: Box::new(ClosureBody::Block(body)),
                    is_arrow: false,
                })
            }
            TokenKind::Fn => {
                self.advance();
                let params = self.parse_params()?;
                if self.eat(&TokenKind::Colon) { self.skip_type_hint(); }
                self.expect(&TokenKind::FatArrow)?;
                let body = self.parse_expression()?;
                Ok(Expression::Closure {
                    params,
                    uses: Vec::new(),
                    body: Box::new(ClosureBody::Expr(body)),
                    is_arrow: true,
                })
            }
            TokenKind::Match => {
                self.advance();
                self.expect(&TokenKind::LParen)?;
                let subject = self.parse_expression()?;
                self.expect(&TokenKind::RParen)?;
                self.expect(&TokenKind::LBrace)?;
                let mut arms = Vec::new();
                while !matches!(self.peek(), TokenKind::RBrace | TokenKind::Eof) {
                    let conditions = if self.eat(&TokenKind::Default) {
                        None
                    } else {
                        let mut conds = vec![self.parse_expression()?];
                        while self.eat(&TokenKind::Comma) {
                            if matches!(self.peek(), TokenKind::FatArrow) { break; }
                            conds.push(self.parse_expression()?);
                        }
                        Some(conds)
                    };
                    self.expect(&TokenKind::FatArrow)?;
                    let body = self.parse_expression()?;
                    self.eat(&TokenKind::Comma);
                    arms.push(MatchArm { conditions, body });
                }
                self.expect(&TokenKind::RBrace)?;
                Ok(Expression::Match { subject: Box::new(subject), arms })
            }
            TokenKind::List => {
                self.advance();
                self.expect(&TokenKind::LParen)?;
                let mut items = Vec::new();
                while !matches!(self.peek(), TokenKind::RParen | TokenKind::Eof) {
                    if self.eat(&TokenKind::Comma) {
                        items.push(None);
                    } else {
                        items.push(Some(self.parse_expression()?));
                        self.eat(&TokenKind::Comma);
                    }
                }
                self.expect(&TokenKind::RParen)?;
                Ok(Expression::List(items))
            }
            TokenKind::Ellipsis => {
                self.advance();
                let e = self.parse_expression()?;
                Ok(Expression::Spread(Box::new(e)))
            }
            k => Err(format!("line {}: unexpected token {:?}", self.line(), k)),
        }
    }

    fn parse_class_name_expr(&mut self) -> Result<Expression, String> {
        match self.peek().clone() {
            TokenKind::Identifier(s) => {
                if matches!(s.as_str(), "self" | "static" | "parent") {
                    self.advance();
                    return Ok(Expression::ClassKeyword(s));
                }
                self.advance();
                Ok(Expression::Identifier(s))
            }
            TokenKind::Static => { self.advance(); Ok(Expression::ClassKeyword("static".into())) }
            TokenKind::Variable(s) => { self.advance(); Ok(Expression::Variable(s)) }
            k => Err(format!("line {}: expected class name, got {:?}", self.line(), k)),
        }
    }

    fn parse_array_elements(&mut self, end: &TokenKind) -> Result<Vec<ArrayElement>, String> {
        let mut elements = Vec::new();
        while self.peek() != end && !matches!(self.peek(), TokenKind::Eof) {
            let spread = self.eat(&TokenKind::Ellipsis);
            let by_ref = self.eat(&TokenKind::Amp);
            let expr = self.parse_expression()?;
            // key => value?
            if self.eat(&TokenKind::FatArrow) {
                let value = self.parse_expression()?;
                elements.push(ArrayElement { key: Some(expr), value, by_ref, spread });
            } else {
                elements.push(ArrayElement { key: None, value: expr, by_ref, spread });
            }
            if !self.eat(&TokenKind::Comma) { break; }
        }
        Ok(elements)
    }

    fn parse_use_clause(&mut self) -> Result<Vec<String>, String> {
        let mut uses = Vec::new();
        if self.eat(&TokenKind::Use) {
            self.expect(&TokenKind::LParen)?;
            while !matches!(self.peek(), TokenKind::RParen | TokenKind::Eof) {
                self.eat(&TokenKind::Amp); // by-ref, ignored for now
                uses.push(self.expect_var()?);
                if !self.eat(&TokenKind::Comma) { break; }
            }
            self.expect(&TokenKind::RParen)?;
        }
        Ok(uses)
    }
}
