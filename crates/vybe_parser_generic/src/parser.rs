//! Generic parser engine — produces common AST from tokens + grammar definition.
//!
//! Architecture:
//! - Pratt parser for expressions (driven by OperatorTable)
//! - Pattern-matching parser for statements/declarations (driven by PatternRules)
//! - Both consume the same token stream and produce vybe_parser_generic AST nodes.

use crate::grammar::*;
use crate::lexer::{Token, TokenKind};
use crate::*;

/// Parse tokens into a Module using the given grammar.
pub fn parse(tokens: &[Token], grammar: &GrammarDef) -> Result<Module, ParseError> {
    let mut p = Parser::new(tokens, grammar);
    p.parse_module()
}

#[derive(Debug, Clone)]
pub struct ParseError {
    pub message: String,
    pub line: u32,
    pub col: u32,
}

impl std::fmt::Display for ParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "line {}:{}: {}", self.line + 1, self.col + 1, self.message)
    }
}

struct Parser<'a> {
    tokens: &'a [Token],
    pos: usize,
    grammar: &'a GrammarDef,
}

impl<'a> Parser<'a> {
    fn new(tokens: &'a [Token], grammar: &'a GrammarDef) -> Self {
        Self { tokens, pos: 0, grammar }
    }

    // ── Token access ─────────────────────────────────────────────────────

    fn peek(&self) -> &Token {
        self.tokens.get(self.pos).unwrap_or(&EOF_TOKEN)
    }

    fn peek2(&self) -> &Token {
        self.tokens.get(self.pos + 1).unwrap_or(&EOF_TOKEN)
    }

    fn at_end(&self) -> bool {
        self.peek().kind == TokenKind::Eof
    }

    fn advance(&mut self) -> &Token {
        let tok = self.peek();
        if self.pos < self.tokens.len() { self.pos += 1; }
        // SAFETY: we return from the slice which outlives the borrow
        &self.tokens[self.pos - 1]
    }

    fn eat_keyword(&mut self, kw: &str) -> bool {
        if self.is_keyword(kw) { self.advance(); true } else { false }
    }

    fn expect_keyword(&mut self, kw: &str) -> Result<(), ParseError> {
        if self.eat_keyword(kw) { Ok(()) }
        else { Err(self.error(format!("expected '{}'", kw))) }
    }

    fn eat_op(&mut self, op: &str) -> bool {
        if self.is_op(op) { self.advance(); true } else { false }
    }

    fn expect_op(&mut self, op: &str) -> Result<(), ParseError> {
        if self.eat_op(op) { Ok(()) }
        else { Err(self.error(format!("expected '{}'", op))) }
    }

    fn expect_ident(&mut self) -> Result<String, ParseError> {
        let tok = self.peek().clone();
        match tok.kind {
            TokenKind::Ident => { self.advance(); Ok(tok.text.clone()) }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("result") => { self.advance(); Ok("Result".into()) }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("self") => { self.advance(); Ok("Self".into()) }
            _ => Err(self.error(format!("expected identifier, got '{}'", tok.text)))
        }
    }

    /// Like expect_ident but also accepts keywords — used after `.` for member names
    /// and for declaration names where keywords can be used as names.
    fn expect_name(&mut self) -> Result<String, ParseError> {
        let tok = self.peek().clone();
        match tok.kind {
            TokenKind::Ident | TokenKind::Keyword => { self.advance(); Ok(tok.text.clone()) }
            _ => Err(self.error(format!("expected name, got '{}'", tok.text)))
        }
    }

    fn is_keyword(&self, kw: &str) -> bool {
        let tok = self.peek();
        tok.kind == TokenKind::Keyword && if self.grammar.language.case_sensitive {
            tok.text == kw
        } else {
            tok.text.eq_ignore_ascii_case(kw)
        }
    }

    fn is_op(&self, op: &str) -> bool {
        let tok = self.peek();
        tok.kind == TokenKind::Operator && tok.text == op
    }

    fn error(&self, msg: String) -> ParseError {
        let tok = self.peek();
        ParseError { message: msg, line: tok.line, col: tok.col }
    }

    fn span_from(&self, start_line: u32, start_col: u32) -> Span {
        let tok = self.peek();
        Span { start_line, start_col, end_line: tok.line, end_col: tok.col }
    }

    fn skip_newlines(&mut self) {
        while self.peek().kind == TokenKind::Newline { self.advance(); }
    }

    fn skip_terminators(&mut self) {
        match &self.grammar.language.statement_terminator {
            Terminator::Char(c) => {
                let s = c.to_string();
                while self.is_op(&s) { self.advance(); }
            }
            Terminator::Newline => { self.skip_newlines(); }
            _ => {}
        }
    }

    fn eat_terminator(&mut self) -> bool {
        match &self.grammar.language.statement_terminator {
            Terminator::Char(c) => { let s = c.to_string(); self.eat_op(&s) }
            Terminator::Newline => {
                if self.peek().kind == TokenKind::Newline { self.advance(); true } else { false }
            }
            _ => true,
        }
    }

    // ── Module parsing ───────────────────────────────────────────────────

    fn parse_module(&mut self) -> Result<Module, ParseError> {
        self.skip_newlines();
        let mut name = "main".to_string();
        let mut imports = Vec::new();

        // Optional program/unit header (Pascal)
        if self.is_keyword("program") || self.is_keyword("unit") {
            self.advance();
            name = self.expect_ident()?;
            self.eat_terminator();
        }

        // Uses/imports (multiple)
        while self.is_keyword("uses") || self.is_keyword("import") || self.is_keyword("imports") || self.is_keyword("from") {
            let mut new_imports = self.parse_imports()?;
            imports.append(&mut new_imports);
        }

        // Skip interface/implementation (Pascal units)
        self.eat_keyword("interface");
        self.eat_keyword("implementation");

        let mut body = Vec::new();
        while !self.at_end() {
            self.skip_newlines();
            if self.at_end() { break; }
            // Stop at main block delimiter (Pascal: begin)
            if self.is_keyword(&self.grammar.blocks.open) && !self.grammar.language.indentation_based { break; }
            if let Some(stmt) = self.try_parse_declaration()? {
                body.push(stmt);
            } else if self.is_keyword(&self.grammar.blocks.open) && !self.grammar.language.indentation_based {
                break;
            } else if self.at_end() {
                break;
            } else {
                // Not a declaration — try as a statement (Python/JS/Ruby top-level code)
                body.push(self.parse_statement()?);
            }
            self.eat_terminator();
        }

        // Main body block (Pascal: begin..end)
        if self.is_keyword(&self.grammar.blocks.open) && !self.grammar.language.indentation_based {
            let block = self.parse_block()?;
            body.extend(block);
        }

        // Trailing dot (Pascal)
        self.eat_op(".");

        // Normalize: merge separated method implementations into their class declarations.
        // e.g. `constructor TFoo.Create; begin...end;` merges into the TFoo ClassDecl.
        body = Self::normalize_separated_methods(body);

        Ok(Module {
            name,
            language: self.detect_lang(),
            body,
            imports,
        })
    }

    /// Merge FunctionDecl nodes with dotted names (ClassName.MethodName) into their
    /// parent ClassDecl's members list. This normalizes Pascal-style separated
    /// declarations so the compiler sees one consistent class shape.
    fn normalize_separated_methods(mut stmts: Vec<Statement>) -> Vec<Statement> {
        // First pass: collect class names and their indices
        let mut class_indices: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
        for (i, stmt) in stmts.iter().enumerate() {
            if let StmtKind::ClassDecl { name, .. } = &stmt.kind {
                class_indices.insert(name.clone(), i);
            }
        }

        if class_indices.is_empty() { return stmts; }

        // Second pass: find FunctionDecls with dotted names and extract them
        let mut methods_to_merge: Vec<(String, Statement)> = Vec::new(); // (class_name, stmt)
        let mut to_remove: Vec<usize> = Vec::new();

        for (i, stmt) in stmts.iter().enumerate() {
            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                if let Some(dot_pos) = name.find('.') {
                    let class_name = &name[..dot_pos];
                    if class_indices.contains_key(class_name) {
                        // Rewrite the name to just the method name
                        let method_name = name[dot_pos + 1..].to_string();
                        let mut method_stmt = stmt.clone();
                        if let StmtKind::FunctionDecl { name: ref mut n, .. } = method_stmt.kind {
                            *n = method_name;
                        }
                        methods_to_merge.push((class_name.to_string(), method_stmt));
                        to_remove.push(i);
                    }
                }
            }
        }

        if methods_to_merge.is_empty() { return stmts; }

        // Remove the extracted method stmts (reverse order to preserve indices)
        to_remove.sort();
        to_remove.reverse();
        for i in &to_remove { stmts.remove(*i); }

        // Recalculate class indices after removal
        class_indices.clear();
        for (i, stmt) in stmts.iter().enumerate() {
            if let StmtKind::ClassDecl { name, .. } = &stmt.kind {
                class_indices.insert(name.clone(), i);
            }
        }

        // Third pass: merge methods into class declarations
        for (class_name, method_stmt) in methods_to_merge {
            if let Some(&idx) = class_indices.get(&class_name) {
                if let StmtKind::ClassDecl { members, .. } = &mut stmts[idx].kind {
                    // Find if there's a matching empty-body method signature to replace
                    let method_name = if let StmtKind::FunctionDecl { name, .. } = &method_stmt.kind { name.clone() } else { String::new() };
                    let mut replaced = false;
                    for m in members.iter_mut() {
                        if let StmtKind::FunctionDecl { name, body, .. } = &mut m.kind {
                            if name.eq_ignore_ascii_case(&method_name) && body.is_empty() {
                                // Replace empty signature with full implementation
                                *m = method_stmt.clone();
                                replaced = true;
                                break;
                            }
                        }
                    }
                    if !replaced {
                        // No matching signature — just add it
                        members.push(method_stmt);
                    }
                }
            }
        }

        stmts
    }

    fn detect_lang(&self) -> Lang {
        match self.grammar.language.name.as_str() {
            "pascal" => Lang::Pascal,
            "javascript" => Lang::JavaScript,
            "python" => Lang::Python,
            "ruby" => Lang::Ruby,
            "php" => Lang::PHP,
            "csharp" => Lang::CSharp,
            "vb" => Lang::VB,
            "dart" => Lang::Dart,
            "cobol" => Lang::Cobol,
            _ => Lang::Unknown,
        }
    }

    fn parse_imports(&mut self) -> Result<Vec<Import>, ParseError> {
        let mut imports = Vec::new();
        if self.eat_keyword("uses") {
            let line = self.peek().line;
            loop {
                let name = self.expect_ident()?;
                imports.push(Import { path: name, alias: None, names: Vec::new(), span: Span { start_line: line, start_col: 0, end_line: line, end_col: 0 } });
                if !self.eat_op(",") { break; }
            }
            self.eat_terminator();
        }
        Ok(imports)
    }

    // ── Block parsing ────────────────────────────────────────────────────

    fn parse_block(&mut self) -> Result<Vec<Statement>, ParseError> {
        self.parse_block_kind(None)
    }

    /// Parse a block, optionally expecting `end <kind>` as the closer (VB: End Sub, End If, etc.)
    fn parse_block_kind(&mut self, end_kind: Option<&str>) -> Result<Vec<Statement>, ParseError> {
        let open = &self.grammar.blocks.open;
        let close = self.grammar.blocks.close.clone();

        // VB-style: no explicit open token, body ends at "End <kind>"
        if self.grammar.blocks.close_with_kind && end_kind.is_some() {
            self.skip_newlines();
            let mut stmts = Vec::new();
            while !self.at_block_end(&close, end_kind) && !self.at_end() {
                self.skip_newlines();
                if self.at_block_end(&close, end_kind) { break; }
                stmts.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.consume_block_end(&close, end_kind);
            return Ok(stmts);
        }

        if open == "INDENT" {
            // Indentation-based block
            if let Some(ref prefix) = self.grammar.blocks.prefix {
                self.expect_op(prefix)?;
            }
            self.skip_newlines();
            if self.peek().kind != TokenKind::Indent {
                let stmt = self.parse_statement()?;
                return Ok(vec![stmt]);
            }
            self.advance();
            let mut stmts = Vec::new();
            while self.peek().kind != TokenKind::Dedent && !self.at_end() {
                self.skip_newlines();
                if self.peek().kind == TokenKind::Dedent { break; }
                stmts.push(self.parse_statement()?);
                self.eat_terminator();
            }
            if self.peek().kind == TokenKind::Dedent { self.advance(); }
            Ok(stmts)
        } else if self.is_keyword(open) || self.is_op(open) {
            self.advance();
            self.skip_newlines();
            let mut stmts = Vec::new();
            while !self.at_block_end(&close, end_kind) && !self.at_end() {
                self.skip_newlines();
                if self.at_block_end(&close, end_kind) { break; }
                stmts.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.consume_block_end(&close, end_kind);
            Ok(stmts)
        } else {
            let stmt = self.parse_statement()?;
            Ok(vec![stmt])
        }
    }

    /// Check if we're at a block end — handles both simple (`end`/`}`) and VB-style (`End Sub`)
    fn at_block_end(&self, close: &str, end_kind: Option<&str>) -> bool {
        if self.grammar.blocks.close_with_kind {
            // VB-style: "end" followed by kind keyword
            if self.is_keyword(&close) {
                if let Some(kind) = end_kind {
                    return self.peek2().text.eq_ignore_ascii_case(kind);
                }
                // No specific kind — any "end X" closes
                return self.peek2().kind == TokenKind::Keyword || self.peek2().kind == TokenKind::Ident;
            }
            false
        } else {
            self.is_keyword(close) || self.is_op(close)
        }
    }

    /// Consume the block end tokens
    fn consume_block_end(&mut self, close: &str, end_kind: Option<&str>) {
        if self.grammar.blocks.close_with_kind {
            if self.is_keyword(close) {
                self.advance(); // consume "end"
                if end_kind.is_some() || self.peek().kind == TokenKind::Keyword || self.peek().kind == TokenKind::Ident {
                    self.advance(); // consume the kind keyword (Sub, If, Module, etc.)
                }
            }
        } else {
            if self.is_keyword(close) || self.is_op(close) { self.advance(); }
        }
    }

    /// Parse a block or single statement (for `then`, `do`, `else` etc.)
    fn parse_block_or_stmt(&mut self) -> Result<Vec<Statement>, ParseError> {
        if self.is_keyword(&self.grammar.blocks.open) || self.is_op(&self.grammar.blocks.open) {
            self.parse_block()
        } else if self.grammar.language.indentation_based {
            self.parse_block()
        } else {
            let stmt = self.parse_statement()?;
            Ok(vec![stmt])
        }
    }

    // ── Statement parsing ────────────────────────────────────────────────

    fn parse_statement(&mut self) -> Result<Statement, ParseError> {
        self.skip_newlines();
        let line = self.peek().line;
        let col = self.peek().col;

        // Try declarations first
        if let Some(decl) = self.try_parse_declaration()? {
            return Ok(decl);
        }

        // Try statement patterns
        if let Some(stmt) = self.try_parse_statement_pattern()? {
            return Ok(stmt);
        }

        // Fall back to expression statement or assignment
        self.parse_expr_or_assign()
    }

    fn try_parse_statement_pattern(&mut self) -> Result<Option<Statement>, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;

        // if
        if self.is_keyword("if") {
            return Ok(Some(self.parse_if()?));
        }
        // while
        if self.is_keyword("while") {
            return Ok(Some(self.parse_while()?));
        }
        // for
        if self.is_keyword("for") {
            return Ok(Some(self.parse_for()?));
        }
        // do / repeat
        if self.is_keyword("do") || self.is_keyword("repeat") {
            return Ok(Some(self.parse_do_while()?));
        }
        // case / switch
        if self.is_keyword("case") || self.is_keyword("switch") || self.is_keyword("select") {
            return Ok(Some(self.parse_switch()?));
        }
        // try
        if self.is_keyword("try") {
            return Ok(Some(self.parse_try()?));
        }
        // return
        if self.is_keyword("return") {
            self.advance();
            let val = if !self.at_stmt_end() { Some(self.parse_expression()?) } else { None };
            return Ok(Some(Statement::with_span(StmtKind::Return(val), self.span_from(line, col))));
        }
        // break
        if self.is_keyword("break") {
            self.advance();
            let val = if !self.at_stmt_end() && !self.is_keyword("end") {
                // Ruby: break value
                None // TODO: parse break value for Ruby
            } else { None };
            return Ok(Some(Statement::with_span(StmtKind::Break(val), self.span_from(line, col))));
        }
        // continue / next
        if self.is_keyword("continue") || self.is_keyword("next") {
            self.advance();
            return Ok(Some(Statement::with_span(StmtKind::Continue, self.span_from(line, col))));
        }
        // throw / raise
        if self.is_keyword("throw") || self.is_keyword("raise") {
            self.advance();
            let val = if !self.at_stmt_end() { Some(self.parse_expression()?) } else { None };
            return Ok(Some(Statement::with_span(StmtKind::Throw(val), self.span_from(line, col))));
        }
        // exit (Pascal)
        if self.is_keyword("exit") {
            self.advance();
            let val = if self.eat_op("(") {
                let v = self.parse_expression()?;
                self.expect_op(")")?;
                Some(v)
            } else { None };
            return Ok(Some(Statement::with_span(StmtKind::Exit(val), self.span_from(line, col))));
        }
        // pass (Python)
        if self.is_keyword("pass") {
            self.advance();
            return Ok(Some(Statement::with_span(StmtKind::Empty, self.span_from(line, col))));
        }
        // assert (Python)
        if self.is_keyword("assert") {
            self.advance();
            let cond = self.parse_expression()?;
            let msg = if self.eat_op(",") { Some(self.parse_expression()?) } else { None };
            let mut exprs = vec![cond];
            if let Some(m) = msg { exprs.push(m); }
            return Ok(Some(Statement::with_span(StmtKind::Extra { tag: "assert".into(), exprs, stmts: Vec::new() }, self.span_from(line, col))));
        }
        // del (Python)
        if self.is_keyword("del") {
            self.advance();
            let expr = self.parse_expression()?;
            return Ok(Some(Statement::with_span(StmtKind::Extra { tag: "del".into(), exprs: vec![expr], stmts: Vec::new() }, self.span_from(line, col))));
        }
        // global / nonlocal (Python)
        if self.is_keyword("global") || self.is_keyword("nonlocal") {
            let kw = self.advance().text.clone();
            let mut names = vec![self.expect_ident()?];
            while self.eat_op(",") { names.push(self.expect_ident()?); }
            let exprs = names.into_iter().map(|n| Expression::ident(&n)).collect();
            return Ok(Some(Statement::with_span(StmtKind::Extra { tag: kw, exprs, stmts: Vec::new() }, self.span_from(line, col))));
        }
        // with (Pascal/Python)
        if self.is_keyword("with") {
            self.advance();
            let expr = self.parse_expression()?;
            // Python: with expr as var:
            if self.eat_keyword("as") {
                let _var = self.expect_ident()?;
                // TODO: bind var to expr in body scope
            }
            self.eat_keyword("do"); // Pascal
            let body = self.parse_block_or_stmt()?;
            return Ok(Some(Statement::with_span(StmtKind::With { expr, body }, self.span_from(line, col))));
        }
        // inherited (Pascal)
        if self.is_keyword("inherited") {
            self.advance();
            let method = if self.peek().kind == TokenKind::Ident {
                Some(self.expect_ident()?)
            } else { None };
            let args = if self.eat_op("(") {
                let a = self.parse_expr_list(")")?;
                self.expect_op(")")?;
                a
            } else { Vec::new() };
            let expr = Expression::new(ExprKind::Inherited { method, args });
            return Ok(Some(Statement::with_span(StmtKind::Expr(expr), self.span_from(line, col))));
        }
        // block (begin/end or {})
        if self.is_keyword(&self.grammar.blocks.open) || self.is_op(&self.grammar.blocks.open) {
            let stmts = self.parse_block()?;
            return Ok(Some(Statement::with_span(StmtKind::Block(stmts), self.span_from(line, col))));
        }

        Ok(None)
    }

    /// Check if current position looks like the start of a case value (for switch/case parsing).
    /// A case value is a literal or identifier followed by ':' or ','.
    fn is_case_value_start(&self) -> bool {
        let tok = self.peek();
        let is_value = matches!(tok.kind, TokenKind::IntLit | TokenKind::FloatLit | TokenKind::StringLit | TokenKind::CharLit)
            || (tok.kind == TokenKind::Ident);
        if !is_value { return false; }
        // Check if followed by ':' (case separator) or ',' (multiple values) or '..' (range)
        let next = self.peek2();
        next.kind == TokenKind::Operator && (next.text == ":" || next.text == "," || next.text == "..")
    }

    fn at_stmt_end(&self) -> bool {
        let tok = self.peek();
        tok.kind == TokenKind::Eof
            || tok.kind == TokenKind::Newline
            || tok.kind == TokenKind::Dedent
            || (tok.kind == TokenKind::Operator && matches!(self.grammar.language.statement_terminator, Terminator::Char(c) if tok.text == c.to_string()))
            || self.is_keyword("end")
            || self.is_keyword("else")
            || self.is_keyword("elif")
            || self.is_keyword("elsif")
            || self.is_keyword("elseif")
            || self.is_keyword("until")
            || self.is_keyword("except")
            || self.is_keyword("finally")
            || self.is_keyword("then")
    }

    // ── If ───────────────────────────────────────────────────────────────

    fn parse_if(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'if'

        let need_parens = matches!(self.grammar.language.name.as_str(), "javascript" | "csharp" | "dart" | "php");
        if need_parens { self.expect_op("(")?; }
        let cond = self.parse_expression()?;
        if need_parens { self.expect_op(")")?; }

        self.eat_keyword("then"); // Pascal/VB

        // VB: multi-line if body until ElseIf/Else/End If
        let vb_style = self.grammar.blocks.close_with_kind;

        // VB single-line If: If cond Then stmt [Else stmt]
        // Detected when 'then' is NOT followed by a newline
        let single_line = vb_style && !self.at_stmt_end();

        let then = if single_line {
            vec![self.parse_statement()?]
        } else if vb_style {
            self.parse_stmts_until(&["elseif", "else", "end"])?
        } else {
            self.parse_block_or_stmt()?
        };

        let mut elifs = Vec::new();
        let mut else_ = None;

        if !single_line {
            loop {
                if self.eat_keyword("elif") || self.eat_keyword("elsif") || self.eat_keyword("elseif") {
                    if need_parens { self.expect_op("(")?; }
                    let c = self.parse_expression()?;
                    if need_parens { self.expect_op(")")?; }
                    self.eat_keyword("then");
                    let b = if vb_style {
                        self.parse_stmts_until(&["elseif", "else", "end"])?
                    } else {
                        self.parse_block_or_stmt()?
                    };
                    elifs.push((c, b));
                } else if self.is_keyword("else") && self.peek2().kind == TokenKind::Keyword && self.peek2().text.eq_ignore_ascii_case("if") {
                    self.advance(); // else
                    let nested_if = self.parse_if()?;
                    else_ = Some(vec![nested_if]);
                    break;
                } else if self.eat_keyword("else") {
                    let b = if vb_style {
                        self.parse_stmts_until(&["end"])?
                    } else {
                        self.parse_block_or_stmt()?
                    };
                    else_ = Some(b);
                    break;
                } else {
                    break;
                }
            }
        } else if self.eat_keyword("else") {
            // Single-line Else
            else_ = Some(vec![self.parse_statement()?]);
        }

        // VB: consume End If (only for multi-line)
        if vb_style && !single_line {
            self.eat_keyword("end");
            self.eat_keyword("if");
        }

        Ok(Statement::with_span(StmtKind::If { cond, then, elifs, else_ }, self.span_from(line, col)))
    }

    /// Parse statements until one of the stop keywords is seen.
    /// Used for VB-style blocks where the body is terminated by context keywords.
    fn parse_stmts_until(&mut self, stop_kws: &[&str]) -> Result<Vec<Statement>, ParseError> {
        let mut stmts = Vec::new();
        loop {
            self.skip_newlines();
            if self.at_end() { break; }
            // Check if current token is a stop keyword
            let should_stop = stop_kws.iter().any(|kw| self.is_keyword(kw));
            if should_stop { break; }
            stmts.push(self.parse_statement()?);
            self.eat_terminator();
        }
        Ok(stmts)
    }

    // ── While ────────────────────────────────────────────────────────────

    fn parse_while(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'while'

        let need_parens = matches!(self.grammar.language.name.as_str(), "javascript" | "csharp" | "dart" | "php");
        if need_parens { self.expect_op("(")?; }
        let cond = self.parse_expression()?;
        if need_parens { self.expect_op(")")?; }

        self.eat_keyword("do"); // Pascal
        let body = if self.grammar.blocks.close_with_kind {
            // VB: While...End While or While...Wend
            let stmts = self.parse_stmts_until(&["end", "wend"])?;
            if self.eat_keyword("wend") {
                // old VB style
            } else {
                self.eat_keyword("end");
                self.eat_keyword("while");
            }
            stmts
        } else {
            self.parse_block_or_stmt()?
        };

        Ok(Statement::with_span(StmtKind::While { cond, body }, self.span_from(line, col)))
    }

    // ── For ──────────────────────────────────────────────────────────────

    fn parse_for(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'for'

        // Check for for-in: `for x in expr` / `for (var x of expr)`
        let saved = self.pos;

        // C-family: for (
        if self.eat_op("(") {
            // Could be for(init;cond;update) or for(x in/of expr)
            // Try for-in/of first
            let has_var_kw = self.eat_keyword("var") || self.eat_keyword("let") || self.eat_keyword("const");
            if self.peek().kind == TokenKind::Ident {
                let var = self.peek().text.clone();
                let after_ident = self.pos + 1;
                if self.tokens.get(after_ident).map_or(false, |t| t.text == "in" || t.text == "of") {
                    self.advance(); // consume ident
                    self.advance(); // consume in/of
                    let iter = self.parse_expression()?;
                    self.expect_op(")")?;
                    let body = self.parse_block_or_stmt()?;
                    return Ok(Statement::with_span(StmtKind::ForIn { var, iter, body }, self.span_from(line, col)));
                }
            }
            // C-style for(init; cond; update)
            self.pos = saved;
            self.eat_op("(");
            let init = if !self.is_op(";") {
                Some(Box::new(self.parse_statement()?))
            } else { None };
            self.eat_op(";");
            let cond = if !self.is_op(";") { Some(self.parse_expression()?) } else { None };
            self.eat_op(";");
            let update = if !self.is_op(")") { Some(self.parse_expression()?) } else { None };
            self.expect_op(")")?;
            let body = self.parse_block_or_stmt()?;
            return Ok(Statement::with_span(StmtKind::For { init, cond, update, body }, self.span_from(line, col)));
        }

        // VB: For Each x In collection  or  For Each item As String In list
        // Must check BEFORE expect_ident since 'each' is a keyword
        if self.eat_keyword("each") {
            // already consumed "for", now "each" consumed, read variable
            let each_var = self.expect_name()?;
            // Optional type annotation: As Type
            if self.is_keyword("as") { self.advance(); self.expect_name()?; }
            self.expect_keyword("in")?;
            let iter = self.parse_expression()?;
            let body = if self.grammar.blocks.close_with_kind {
                let stmts = self.parse_stmts_until(&["next"])?;
                self.eat_keyword("next");
                if self.peek().kind == TokenKind::Ident { self.advance(); }
                stmts
            } else {
                self.parse_block_or_stmt()?
            };
            return Ok(Statement::with_span(StmtKind::ForIn { var: each_var, iter, body }, self.span_from(line, col)));
        }

        // Pascal-style: for i := expr to/downto expr do
        // Python-style: for x in expr: / for i, v in enumerate(x):
        let var = self.expect_ident()?;

        // Check for tuple unpacking in for: for a, b in ...
        let mut extra_vars = Vec::new();
        while self.eat_op(",") {
            if self.is_keyword("in") { break; }
            extra_vars.push(self.expect_ident()?);
        }

        if self.eat_keyword("in") {
            // for-in
            let iter = self.parse_expression()?;
            self.eat_keyword("do"); // Pascal
            let body = self.parse_block_or_stmt()?;
            // If tuple unpacking, use combined var name
            let var_name = if extra_vars.is_empty() { var } else {
                let mut all = vec![var];
                all.extend(extra_vars);
                all.join(",")
            };
            return Ok(Statement::with_span(StmtKind::ForIn { var: var_name, iter, body }, self.span_from(line, col)));
        }

        // Pascal for := from to/downto limit do
        // VB: For i As Integer = 1 To 5 — skip optional type annotation
        if self.is_keyword("as") {
            self.advance(); // consume 'as'
            self.expect_name()?; // consume type name
        }
        if let Some(ref assign_op) = self.grammar.assignment.operator {
            self.expect_op(assign_op)?;
        }
        let from = self.parse_expression()?;

        let downto = if self.eat_keyword("downto") { true } else { self.eat_keyword("to"); false };
        let to = self.parse_expression()?;
        // Optional step: Step expr (VB)
        let step_expr = if self.eat_keyword("step") { Some(self.parse_expression()?) } else { None };
        self.eat_keyword("do"); // Pascal
        let body_stmts = if self.grammar.blocks.close_with_kind {
            // VB: body until Next
            let stmts = self.parse_stmts_until(&["next"])?;
            self.eat_keyword("next");
            // Next may have optional variable name
            if self.peek().kind == TokenKind::Ident { self.advance(); }
            stmts
        } else {
            self.parse_block_or_stmt()?
        };

        // Convert Pascal for to common For with init/cond/update
        let init = Some(Box::new(Statement::new(StmtKind::Assign {
            target: Expression::ident(&var),
            value: from,
        })));
        let cmp_op = if downto { BinOp::Ge } else { BinOp::Le };
        let cond = Some(Expression::new(ExprKind::Binary {
            op: cmp_op,
            left: Box::new(Expression::ident(&var)),
            right: Box::new(to),
        }));
        let step_op = if downto { BinOp::Sub } else { BinOp::Add };
        let update = Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&var)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: step_op,
                left: Box::new(Expression::ident(&var)),
                right: Box::new(step_expr.unwrap_or_else(|| Expression::int(1))),
            })),
        }));

        Ok(Statement::with_span(StmtKind::For { init, cond, update, body: body_stmts }, self.span_from(line, col)))
    }

    // ── Do/While / Repeat/Until ──────────────────────────────────────────

    fn parse_do_while(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;

        if self.eat_keyword("repeat") {
            // Pascal repeat..until
            let mut body = Vec::new();
            while !self.is_keyword("until") && !self.at_end() {
                body.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.expect_keyword("until")?;
            let cond = self.parse_expression()?;
            return Ok(Statement::with_span(StmtKind::DoWhile { body, cond, until: true }, self.span_from(line, col)));
        }

        // VB: Do While cond ... Loop  |  Do Until cond ... Loop  |  Do ... Loop While cond  |  Do ... Loop Until cond
        // C-family: do { body } while (cond)
        self.advance(); // consume 'do'

        // VB: Do While cond ... Loop
        if self.is_keyword("while") {
            self.advance(); // consume 'while'
            let cond = self.parse_expression()?;
            self.eat_terminator();
            let mut body = Vec::new();
            while !self.is_keyword("loop") && !self.at_end() {
                body.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.eat_keyword("loop");
            return Ok(Statement::with_span(StmtKind::While { cond, body }, self.span_from(line, col)));
        }

        // VB: Do Until cond ... Loop
        if self.is_keyword("until") {
            self.advance(); // consume 'until'
            let cond = self.parse_expression()?;
            self.eat_terminator();
            let mut body = Vec::new();
            while !self.is_keyword("loop") && !self.at_end() {
                body.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.eat_keyword("loop");
            // Until = while NOT cond
            let neg_cond = Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(cond) });
            return Ok(Statement::with_span(StmtKind::While { cond: neg_cond, body }, self.span_from(line, col)));
        }

        // VB: Do ... Loop While/Until  OR  C-family: do { body } while (cond)
        let body = if self.grammar.blocks.close_with_kind {
            // VB-style: parse body until Loop
            self.eat_terminator();
            let mut stmts = Vec::new();
            while !self.is_keyword("loop") && !self.at_end() {
                stmts.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.eat_keyword("loop");
            stmts
        } else {
            self.parse_block_or_stmt()?
        };

        let until = if self.eat_keyword("until") { true }
                    else { self.expect_keyword("while")?; false };
        let need_parens = matches!(self.grammar.language.name.as_str(), "javascript" | "csharp" | "dart" | "php");
        if need_parens { self.expect_op("(")?; }
        let cond = self.parse_expression()?;
        if need_parens { self.expect_op(")")?; }

        Ok(Statement::with_span(StmtKind::DoWhile { body, cond, until }, self.span_from(line, col)))
    }

    // ── Switch/Case ──────────────────────────────────────────────────────

    fn parse_switch(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        let kw = self.advance().text.clone(); // consume case/switch/select

        // VB: Select Case expr
        if kw.eq_ignore_ascii_case("select") { self.eat_keyword("case"); }

        let need_parens = matches!(self.grammar.language.name.as_str(), "javascript" | "csharp" | "dart" | "php");
        if need_parens { self.expect_op("(")?; }
        let expr = self.parse_expression()?;
        if need_parens { self.expect_op(")")?; }

        self.eat_keyword("of"); // Pascal
        self.eat_op("{");       // C-family

        let mut cases = Vec::new();
        let mut default = None;

        loop {
            self.skip_newlines();
            if self.is_keyword("end") || self.is_op("}") || self.at_end() { break; }

            if self.eat_keyword("else") || self.eat_keyword("otherwise") || self.eat_keyword("default") {
                self.eat_op(":");
                let mut stmts = Vec::new();
                while !self.is_keyword("end") && !self.is_op("}") && !self.at_end() {
                    stmts.push(self.parse_statement()?);
                    self.eat_terminator();
                }
                default = Some(stmts);
                break;
            }

            if self.is_keyword("case") { self.advance(); } // C-family `case val:`

            let mut values = Vec::new();
            loop {
                let v = self.parse_expression()?;
                values.push(v);
                if !self.eat_op(",") { break; }
            }
            self.eat_op(":");

            let mut body = Vec::new();
            if self.is_keyword(&self.grammar.blocks.open) || self.is_op(&self.grammar.blocks.open) {
                body = self.parse_block()?;
            } else {
                loop {
                    if self.is_keyword("end") || self.is_op("}") || self.is_keyword("case")
                        || self.is_keyword("else") || self.is_keyword("otherwise")
                        || self.is_keyword("default") || self.at_end() { break; }
                    // Check if next token starts a new case value (literal/ident followed by ':')
                    if self.is_case_value_start() { break; }
                    body.push(self.parse_statement()?);
                    self.eat_terminator();
                    // C-family: break after case
                    if self.is_keyword("break") { self.advance(); self.eat_terminator(); break; }
                }
            }
            self.eat_terminator();
            cases.push(SwitchCase { values, body });
        }

        if self.grammar.blocks.close_with_kind {
            // VB: End Select
            self.eat_keyword("end");
            self.eat_keyword("select");
        } else {
            self.eat_keyword("end");
            self.eat_op("}");
        }

        Ok(Statement::with_span(StmtKind::Switch { expr, cases, default }, self.span_from(line, col)))
    }

    // ── Try/Catch/Finally ────────────────────────────────────────────────

    fn parse_try(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'try'

        // Try body
        let body = if self.grammar.language.indentation_based {
            self.parse_block()?
        } else {
            self.eat_op("{");
            let mut stmts = Vec::new();
            while !self.is_keyword("except") && !self.is_keyword("catch") && !self.is_keyword("finally")
                  && !self.is_op("}") && !self.at_end() {
                stmts.push(self.parse_statement()?);
                self.eat_terminator();
            }
            self.eat_op("}");
            stmts
        };

        let mut catches = Vec::new();
        let mut finally = None;

        // except (Python/Pascal) / catch (C-family) / rescue (Ruby)
        while self.is_keyword("except") || self.is_keyword("catch") || self.is_keyword("rescue") {
            self.advance();
            let has_paren = self.eat_op("(");
            let mut type_name = None;
            let mut var_name = None;
            // Parse exception type and optional variable binding
            // Python: except ValueError as e:
            // C#/Java: catch (Exception e)
            // Pascal: on E: Exception do
            if self.peek().kind == TokenKind::Ident
                && !self.is_keyword("as")
                && !(self.grammar.language.indentation_based && self.is_op(":")) {
                let first = self.expect_ident()?;
                if self.eat_keyword("as") {
                    // except Type as var
                    type_name = Some(first);
                    var_name = Some(self.expect_ident()?);
                } else if !self.grammar.language.indentation_based && self.eat_op(":") {
                    // catch (var: Type) or on var: Type — NOT for Python
                    type_name = Some(self.expect_ident()?);
                    var_name = Some(first);
                } else {
                    // Just the type name (except ValueError:)
                    type_name = Some(first);
                }
            }
            if has_paren { self.eat_op(")"); }
            self.eat_keyword("do"); // Pascal

            let catch_body = if self.grammar.language.indentation_based {
                self.parse_block()?
            } else {
                self.eat_op("{");
                let mut stmts = Vec::new();
                while !self.is_keyword("end") && !self.is_keyword("except") && !self.is_keyword("catch")
                      && !self.is_keyword("finally") && !self.is_keyword("ensure")
                      && !self.is_op("}") && !self.at_end() {
                    stmts.push(self.parse_statement()?);
                    self.eat_terminator();
                }
                self.eat_op("}");
                stmts
            };
            catches.push(CatchClause { type_name, var_name, body: catch_body });
        }

        // Python: else clause on try (runs if no exception)
        let mut else_ = None;
        if self.is_keyword("else") && !catches.is_empty() {
            self.advance();
            let else_body = if self.grammar.language.indentation_based {
                self.parse_block()?
            } else {
                let s = self.parse_statement()?;
                vec![s]
            };
            else_ = Some(else_body);
        }

        // finally / ensure
        if self.is_keyword("finally") || self.is_keyword("ensure") {
            self.advance();
            let fin = if self.grammar.language.indentation_based {
                self.parse_block()?
            } else {
                self.eat_op("{");
                let mut stmts = Vec::new();
                while !self.is_keyword("end") && !self.is_op("}") && !self.at_end() {
                    stmts.push(self.parse_statement()?);
                    self.eat_terminator();
                }
                self.eat_op("}");
                stmts
            };
            finally = Some(fin);
        }

        if self.grammar.blocks.close_with_kind {
            self.eat_keyword("end");
            self.eat_keyword("try");
        } else {
            self.eat_keyword("end");
        }

        Ok(Statement::with_span(StmtKind::Try { body, catches, else_, finally }, self.span_from(line, col)))
    }

    // ── Declarations ─────────────────────────────────────────────────────

    fn try_parse_declaration(&mut self) -> Result<Option<Statement>, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;

        // Decorators: @expr before def/class (Python, Java, etc.)
        if self.is_op("@") {
            let mut decorators = Vec::new();
            while self.eat_op("@") {
                let name = self.expect_ident()?;
                // Optional args: @decorator(args)
                if self.is_op("(") { let _ = self.parse_call_args()?; }
                decorators.push(name);
                self.skip_newlines();
            }
            // Parse the decorated declaration
            if let Some(mut stmt) = self.try_parse_declaration()? {
                if let StmtKind::FunctionDecl { ref mut modifiers, .. } = stmt.kind {
                    modifiers.decorators = decorators;
                } else if let StmtKind::ClassDecl { ref mut modifiers, .. } = stmt.kind {
                    modifiers.decorators = decorators;
                }
                return Ok(Some(stmt));
            }
        }

        // var section
        if self.is_keyword("var") && !self.grammar.language.case_sensitive {
            // Pascal var section (may have multiple declarations)
            return Ok(Some(self.parse_var_section()?));
        }
        if self.is_keyword("var") || self.is_keyword("let") || self.is_keyword("const") || self.is_keyword("dim") {
            return Ok(Some(self.parse_var_decl()?));
        }

        // const section (Pascal)
        if self.is_keyword("const") {
            return Ok(Some(self.parse_const_section()?));
        }

        // function / procedure / def / sub
        if self.is_keyword("function") || self.is_keyword("procedure") || self.is_keyword("def") || self.is_keyword("sub") {
            return Ok(Some(self.parse_function_decl()?));
        }

        // async function / async def
        if self.is_keyword("async") && (self.peek2().text.eq_ignore_ascii_case("function") || self.peek2().text == "def") {
            return Ok(Some(self.parse_function_decl()?));
        }

        // constructor / destructor (Pascal)
        if self.is_keyword("constructor") || self.is_keyword("destructor") {
            return Ok(Some(self.parse_function_decl()?));
        }

        // class
        if self.is_keyword("class") {
            return Ok(Some(self.parse_class_decl()?));
        }

        // module (VB: Module Name...End Module)
        if self.is_keyword("module") {
            self.advance();
            let mod_name = self.expect_name()?;
            self.eat_terminator();
            let mut members = Vec::new();
            while !self.at_block_end(&self.grammar.blocks.close.clone(), Some("module")) && !self.at_end() {
                self.skip_newlines();
                if self.at_block_end(&self.grammar.blocks.close.clone(), Some("module")) { break; }
                if let Some(decl) = self.try_parse_declaration()? {
                    members.push(decl);
                } else if self.at_end() {
                    break;
                } else {
                    members.push(self.parse_statement()?);
                }
                self.eat_terminator();
            }
            self.consume_block_end(&self.grammar.blocks.close.clone(), Some("module"));
            return Ok(Some(Statement::with_span(StmtKind::ModuleDecl { name: mod_name, body: members }, self.span_from(line, col))));
        }

        // type section (Pascal)
        if self.is_keyword("type") {
            return Ok(Some(self.parse_type_section()?));
        }

        Ok(None)
    }

    fn parse_var_section(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'var'

        // Pascal: var x: Integer; y: String;
        // We emit multiple VarDecl statements wrapped in a Block
        let mut stmts = Vec::new();
        while self.peek().kind == TokenKind::Ident {
            let mut names = vec![self.expect_ident()?];
            while self.eat_op(",") { names.push(self.expect_ident()?); }
            self.expect_op(":")?;
            let type_hint = Some(self.parse_type_name()?);
            let init = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
            self.eat_terminator();
            for name in names {
                stmts.push(Statement::with_span(StmtKind::VarDecl {
                    name, type_hint: type_hint.clone(), init: init.clone(), is_const: false, mutable: true,
                }, self.span_from(line, col)));
            }
        }
        if stmts.len() == 1 { Ok(stmts.pop().unwrap()) }
        else { Ok(Statement::with_span(StmtKind::Block(stmts), self.span_from(line, col))) }
    }

    fn parse_var_decl(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        let is_const = self.is_keyword("const");
        self.advance(); // consume var/let/const/dim

        let name = self.expect_name()?;

        // VB: Dim arr(5) As Integer — array size declaration
        let array_size = if self.is_op("(") && !self.is_keyword("as") {
            // Peek ahead: if it's `(number)` followed by As, it's array decl
            let saved = self.pos;
            self.advance(); // consume (
            if self.peek().kind == TokenKind::IntLit || self.peek().kind == TokenKind::Ident {
                let size_expr = self.parse_expression().ok();
                if self.eat_op(")") && (self.is_keyword("as") || self.at_stmt_end()) {
                    size_expr
                } else {
                    self.pos = saved;
                    None
                }
            } else if self.eat_op(")") {
                // Empty parens: Dim arr() As Integer
                Some(Expression::int(0))
            } else {
                self.pos = saved;
                None
            }
        } else { None };

        let (type_hint, init) = if self.eat_op(":") || self.eat_keyword("as") {
            if self.is_keyword("new") {
                // VB: Dim x As New ClassName(args) or Dim x As New List(Of String)
                self.advance(); // consume 'new'
                let mut class_name = self.expect_name()?;
                // Dotted type: System.Drawing.Point
                while self.eat_op(".") {
                    class_name.push('.');
                    class_name.push_str(&self.expect_name()?);
                }
                // Generic: List(Of String) or Dictionary(Of String, String)
                if self.is_op("(") && self.peek2().text.eq_ignore_ascii_case("of") {
                    self.advance(); // consume (
                    self.advance(); // consume Of
                    let mut generic_types = vec![self.expect_name()?];
                    while self.eat_op(",") {
                        generic_types.push(self.expect_name()?);
                    }
                    self.eat_op(")");
                    class_name = format!("{}(Of {})", class_name, generic_types.join(", "));
                }
                let args = if self.is_op("(") { self.parse_call_args()? } else { Vec::new() };
                (Some(class_name.clone()), Some(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args,
                })))
            } else {
                let t = self.parse_type_name()?;
                let init = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
                (Some(t), init)
            }
        } else {
            let init = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
            (None, init)
        };

        // If array_size was specified, wrap init in array creation
        let init = if let Some(size) = array_size {
            Some(Expression::new(ExprKind::Extra {
                tag: "array_new".to_string(),
                exprs: vec![size],
            }))
        } else {
            init
        };

        Ok(Statement::with_span(StmtKind::VarDecl {
            name, type_hint, init, is_const, mutable: !is_const,
        }, self.span_from(line, col)))
    }

    fn parse_const_section(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'const'

        let mut stmts = Vec::new();
        while self.peek().kind == TokenKind::Ident {
            let name = self.expect_ident()?;
            let type_hint = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
            self.expect_op("=")?;
            let init = self.parse_expression()?;
            self.eat_terminator();
            stmts.push(Statement::with_span(StmtKind::VarDecl {
                name, type_hint, init: Some(init), is_const: true, mutable: false,
            }, self.span_from(line, col)));
        }
        if stmts.len() == 1 { Ok(stmts.pop().unwrap()) }
        else { Ok(Statement::with_span(StmtKind::Block(stmts), self.span_from(line, col))) }
    }

    fn parse_function_decl(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        let mut modifiers = Modifiers::default();

        if self.eat_keyword("async") { modifiers.is_async = true; }

        let kw = self.advance().text.clone(); // function/procedure/def/sub/constructor/destructor
        let is_method_impl = self.peek().kind == TokenKind::Ident && self.peek2().kind == TokenKind::Operator && self.peek2().text == ".";

        let name = if is_method_impl {
            let class = self.expect_name()?;
            self.expect_op(".")?;
            let method = self.expect_name()?;
            format!("{}.{}", class, method)
        } else {
            self.expect_name()?
        };

        let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };

        // Return type: use grammar's return_separator (":" for Pascal, "->" for Python)
        // Don't consume ":" if it's the block prefix (Python: def foo(): means block start, not return type)
        let return_type = if let Some(ref ret_sep) = self.grammar.types.return_separator {
            if self.eat_op(ret_sep) || self.eat_keyword(ret_sep) { Some(self.parse_type_name()?) } else { None }
        } else if self.grammar.blocks.prefix.as_deref() != Some(":") && self.eat_op(":") {
            // Only eat ":" for return type if ":" is NOT the block prefix
            Some(self.parse_type_name()?)
        } else if self.eat_keyword("as") {
            Some(self.parse_type_name()?)
        } else { None };

        self.eat_terminator();

        // Skip forward declarations
        if self.eat_keyword("forward") { self.eat_terminator(); return Ok(Statement::with_span(StmtKind::Empty, self.span_from(line, col))); }

        // Nested declarations (Pascal)
        let mut body = Vec::new();
        while let Some(decl) = self.try_parse_declaration()? {
            body.push(decl);
            self.eat_terminator();
        }

        // Use block-with-kind for VB (End Sub / End Function), plain block for others
        let end_kind = if self.grammar.blocks.close_with_kind { Some(kw.as_str()) } else { None };
        let block = self.parse_block_kind(end_kind)?;
        body.extend(block);
        self.eat_terminator();

        if kw.eq_ignore_ascii_case("constructor") || kw.eq_ignore_ascii_case("destructor") {
            modifiers.extra.push(kw.to_lowercase());
        }

        Ok(Statement::with_span(StmtKind::FunctionDecl {
            name, params, return_type, body, modifiers,
        }, self.span_from(line, col)))
    }

    fn parse_class_decl(&mut self) -> Result<Statement, ParseError> {
        self.advance(); // consume 'class'
        let name = self.expect_name()?;
        self.parse_class_body(name)
    }

    fn parse_class_body(&mut self, name: String) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;

        // Inheritance
        let mut parent = None;
        let mut interfaces = Vec::new();
        // Don't eat ":" as inheritance if it's the block prefix (Python)
        if self.eat_op("(") {
            if !self.is_op(")") {
                let first = self.expect_ident()?;
                parent = Some(first);
                while self.eat_op(",") { interfaces.push(self.expect_ident()?); }
            }
            self.eat_op(")");
        } else if self.eat_keyword("extends") || (self.is_op(":") && self.grammar.blocks.prefix.as_deref() != Some(":")) {
            if self.is_op(":") { self.advance(); }
            let first = self.expect_ident()?;
            parent = Some(first);
            while self.eat_op(",") { interfaces.push(self.expect_ident()?); }
        }
        self.eat_terminator();
        self.skip_newlines();
        if self.eat_keyword("inherits") {
            parent = Some(self.expect_name()?);
        }
        self.eat_terminator();
        self.skip_newlines();
        if self.eat_keyword("implements") {
            loop {
                interfaces.push(self.expect_ident()?);
                if !self.eat_op(",") { break; }
            }
        }

        self.eat_terminator();

        // For indentation-based languages, class body is a regular block
        if self.grammar.language.indentation_based {
            let members = self.parse_block()?;
            self.eat_keyword("end");
            return Ok(Statement::with_span(StmtKind::ClassDecl {
                name, parent, interfaces, members, modifiers: Modifiers::default(),
            }, self.span_from(line, col)));
        }

        self.eat_op("{");

        // Parse class members — signatures only (no bodies, Pascal-style)
        let mut members = Vec::new();
        loop {
            self.skip_newlines();
            // Skip visibility keywords
            if self.is_keyword("public") || self.is_keyword("private") || self.is_keyword("protected") || self.is_keyword("published") {
                self.advance();
                continue;
            }
            if self.grammar.blocks.close_with_kind {
                if self.at_block_end(&self.grammar.blocks.close.clone(), Some("class")) { break; }
            } else {
                if self.is_keyword("end") || self.is_op("}") || self.at_end() { break; }
            }

            // VB: Sub/Function with body inside class (non-separated)
            if self.grammar.blocks.close_with_kind
                && (self.is_keyword("sub") || self.is_keyword("function")) {
                let decl = self.parse_function_decl()?;
                members.push(decl);
                self.eat_terminator();
                continue;
            }

            // Method signatures: function/procedure/constructor/destructor Name(params): Type; directives
            if self.is_keyword("function") || self.is_keyword("procedure")
                || self.is_keyword("constructor") || self.is_keyword("destructor") {
                let kw = self.advance().text.clone();
                let mname = self.expect_ident()?;
                let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                self.eat_terminator();
                // Parse and store directives: virtual; override; abstract; reintroduce; overload; inline; etc.
                let mut modifiers = Modifiers::default();
                loop {
                    if self.is_keyword("virtual") { self.advance(); modifiers.is_virtual = true; self.eat_terminator(); }
                    else if self.is_keyword("override") { self.advance(); modifiers.is_override = true; self.eat_terminator(); }
                    else if self.is_keyword("abstract") { self.advance(); modifiers.is_abstract = true; self.eat_terminator(); }
                    else if self.peek().kind == TokenKind::Ident && matches!(self.peek().text.to_lowercase().as_str(), "reintroduce"|"overload"|"inline"|"cdecl"|"stdcall"|"register"|"dynamic") {
                        modifiers.extra.push(self.advance().text.to_lowercase());
                        self.eat_terminator();
                    }
                    else { break; }
                }
                if kw.eq_ignore_ascii_case("constructor") || kw.eq_ignore_ascii_case("destructor") {
                    modifiers.extra.push(kw.to_lowercase());
                }
                members.push(Statement::with_span(StmtKind::FunctionDecl {
                    name: mname, params, return_type, body: Vec::new(), modifiers,
                }, self.span_from(line, col)));
                continue;
            }
            // operator overloading: `class operator Add(a, b: TFoo): TFoo;` or just `operator`
            if self.peek().kind == TokenKind::Ident && self.peek().text.eq_ignore_ascii_case("operator") {
                self.advance(); // consume 'operator'
                let op_name = self.expect_ident()?;
                let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                self.eat_terminator();
                let mut modifiers = Modifiers::default();
                modifiers.extra.push("operator".into());
                members.push(Statement::with_span(StmtKind::FunctionDecl {
                    name: op_name, params, return_type, body: Vec::new(), modifiers,
                }, self.span_from(line, col)));
                continue;
            }
            // class function/procedure/operator (static)
            if self.is_keyword("class") {
                self.advance();
                // class operator Add(...)
                if self.peek().kind == TokenKind::Ident && self.peek().text.eq_ignore_ascii_case("operator") {
                    self.advance(); // consume 'operator'
                    let op_name = self.expect_ident()?;
                    let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                    let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                    self.eat_terminator();
                    let mut modifiers = Modifiers::default();
                    modifiers.is_static = true;
                    modifiers.extra.push("operator".into());
                    members.push(Statement::with_span(StmtKind::FunctionDecl {
                        name: op_name, params, return_type, body: Vec::new(), modifiers,
                    }, self.span_from(line, col)));
                    continue;
                }
                if self.is_keyword("function") || self.is_keyword("procedure") {
                    let kw = self.advance().text.clone();
                    let mname = self.expect_ident()?;
                    let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                    let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                    self.eat_terminator();
                    let mut modifiers = Modifiers::default();
                    modifiers.is_static = true;
                    members.push(Statement::with_span(StmtKind::FunctionDecl {
                        name: mname, params, return_type, body: Vec::new(), modifiers,
                    }, self.span_from(line, col)));
                    continue;
                } else if self.eat_keyword("var") {
                    // class var
                    if self.peek().kind == TokenKind::Ident {
                        let fname = self.expect_ident()?;
                        if self.eat_op(":") {
                            let t = self.parse_type_name()?;
                            self.eat_terminator();
                            let mut mods = Modifiers::default();
                            mods.is_static = true;
                            members.push(Statement::with_span(StmtKind::VarDecl {
                                name: fname, type_hint: Some(t), init: None, is_const: false, mutable: true,
                            }, self.span_from(line, col)));
                            continue;
                        }
                    }
                }
            }
            // Property declarations
            if let TokenKind::Ident = self.peek().kind {
                if self.peek().text.eq_ignore_ascii_case("property") {
                    self.advance();
                    let pname = self.expect_ident()?;
                    self.expect_op(":")?;
                    let ptype = self.parse_type_name()?;
                    let mut getter = None;
                    let mut setter = None;
                    while let TokenKind::Ident = self.peek().kind {
                        match self.peek().text.to_lowercase().as_str() {
                            "read" => { self.advance(); getter = Some(self.expect_ident()?); }
                            "write" => { self.advance(); setter = Some(self.expect_ident()?); }
                            _ => break,
                        }
                    }
                    self.eat_terminator();
                    members.push(Statement::with_span(StmtKind::PropertyDecl {
                        name: pname, type_hint: Some(ptype), getter, setter,
                    }, self.span_from(line, col)));
                    continue;
                }
            }
            // VB: Property Get/Set
            if self.is_keyword("property") {
                self.advance();
                let pname = self.expect_name()?;
                // VB property may have return type: Property Name() As Type
                let _params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                let _ret = if self.eat_keyword("as") { Some(self.parse_type_name()?) } else { None };
                self.eat_terminator();
                // Parse Get/Set blocks
                let mut getter_body = Vec::new();
                let mut setter_body = Vec::new();
                while !self.at_block_end(&self.grammar.blocks.close.clone(), Some("property")) && !self.at_end() {
                    self.skip_newlines();
                    if self.at_block_end(&self.grammar.blocks.close.clone(), Some("property")) { break; }
                    if self.eat_keyword("get") {
                        self.eat_terminator();
                        while !self.at_block_end(&self.grammar.blocks.close.clone(), Some("get")) && !self.at_end() {
                            getter_body.push(self.parse_statement()?);
                            self.eat_terminator();
                        }
                        self.consume_block_end(&self.grammar.blocks.close.clone(), Some("get"));
                    } else if self.eat_keyword("set") {
                        // Set may have parameter: Set(value As Type)
                        if self.is_op("(") { let _ = self.parse_params()?; }
                        self.eat_terminator();
                        while !self.at_block_end(&self.grammar.blocks.close.clone(), Some("set")) && !self.at_end() {
                            setter_body.push(self.parse_statement()?);
                            self.eat_terminator();
                        }
                        self.consume_block_end(&self.grammar.blocks.close.clone(), Some("set"));
                    } else {
                        break;
                    }
                    self.eat_terminator();
                }
                self.consume_block_end(&self.grammar.blocks.close.clone(), Some("property"));
                // Emit as getter/setter functions
                if !getter_body.is_empty() {
                    let gname = format!("__get_{}", pname);
                    members.push(Statement::new(StmtKind::FunctionDecl {
                        name: gname, params: Vec::new(), return_type: None, body: getter_body, modifiers: Modifiers::default(),
                    }));
                }
                if !setter_body.is_empty() {
                    let sname = format!("__set_{}", pname);
                    let param = Param { name: "value".into(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false };
                    members.push(Statement::new(StmtKind::FunctionDecl {
                        name: sname, params: vec![param], return_type: None, body: setter_body, modifiers: Modifiers::default(),
                    }));
                }
                members.push(Statement::new(StmtKind::PropertyDecl {
                    name: pname, type_hint: _ret, getter: None, setter: None,
                }));
                continue;
            }
            // VB: Shared members
            if self.is_keyword("shared") {
                self.advance();
                continue; // skip modifier, next iteration handles the actual member
            }
            // VB: Overridable/Overrides/MustOverride/NotOverridable
            if self.is_keyword("overridable") || self.is_keyword("overrides")
                || self.is_keyword("mustoverride") || self.is_keyword("notoverridable") {
                self.advance();
                continue;
            }
            // Field declaration: FName: Type; or Name As Type (VB)
            if self.peek().kind == TokenKind::Ident || self.is_keyword("dim") {
                self.eat_keyword("dim"); // optional Dim keyword for fields
                let mut names = vec![self.expect_name()?];
                while self.eat_op(",") {
                    if self.peek().kind == TokenKind::Ident || self.peek().kind == TokenKind::Keyword { names.push(self.expect_name()?); }
                    else { break; }
                }
                if self.eat_op(":") || self.eat_keyword("as") {
                    let type_hint = Some(self.parse_type_name()?);
                    let init = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
                    self.eat_terminator();
                    for fname in names {
                        members.push(Statement::with_span(StmtKind::VarDecl {
                            name: fname, type_hint: type_hint.clone(), init: init.clone(), is_const: false, mutable: true,
                        }, self.span_from(line, col)));
                    }
                } else {
                    break;
                }
            } else {
                break;
            }
        }

        // VB: End Class  |  Pascal/C-family: end / }
        if self.grammar.blocks.close_with_kind {
            self.consume_block_end(&self.grammar.blocks.close.clone(), Some("class"));
        } else {
            self.eat_keyword("end");
            self.eat_op("}");
        }

        Ok(Statement::with_span(StmtKind::ClassDecl {
            name, parent, interfaces, members, modifiers: Modifiers::default(),
        }, self.span_from(line, col)))
    }

    fn parse_type_section(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        self.advance(); // consume 'type'

        let mut stmts = Vec::new();
        while self.peek().kind == TokenKind::Ident {
            let name = self.expect_ident()?;
            self.expect_op("=")?;

            if self.is_keyword("class") {
                self.advance(); // consume 'class'
                let class_stmt = self.parse_class_body(name)?;
                stmts.push(class_stmt);
            } else if self.is_keyword("interface") {
                self.advance();
                // Parse interface
                let parent = if self.eat_op("(") { let p = self.expect_ident()?; self.expect_op(")")?; Some(p) } else { None };
                let mut members = Vec::new();
                while !self.is_keyword("end") && !self.at_end() {
                    if self.is_keyword("function") || self.is_keyword("procedure") {
                        members.push(self.parse_function_decl()?);
                    } else { break; }
                    self.eat_terminator();
                }
                self.eat_keyword("end");
                stmts.push(Statement::with_span(StmtKind::InterfaceDecl { name: name.clone(), parent, members }, self.span_from(line, col)));
            } else if self.eat_op("(") {
                // Enum: (Red, Green, Blue)
                let mut members = Vec::new();
                while !self.is_op(")") && !self.at_end() {
                    let mname = self.expect_ident()?;
                    let value = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
                    members.push(EnumMember { name: mname, value });
                    if !self.eat_op(",") { break; }
                }
                self.expect_op(")")?;
                stmts.push(Statement::with_span(StmtKind::EnumDecl { name, members }, self.span_from(line, col)));
            } else if self.is_keyword("record") {
                self.advance();
                let mut members = Vec::new();
                while !self.is_keyword("end") && !self.at_end() {
                    // Skip visibility
                    if self.is_keyword("public") || self.is_keyword("private") || self.is_keyword("protected") {
                        self.advance(); continue;
                    }
                    // Methods in records
                    if self.is_keyword("function") || self.is_keyword("procedure")
                        || self.is_keyword("constructor") || self.is_keyword("destructor") {
                        let kw = self.advance().text.clone();
                        let mname = self.expect_ident()?;
                        let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                        let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                        self.eat_terminator();
                        let mut modifiers = Modifiers::default();
                        if kw.eq_ignore_ascii_case("constructor") { modifiers.extra.push("constructor".into()); }
                        members.push(Statement::new(StmtKind::FunctionDecl {
                            name: mname, params, return_type, body: Vec::new(), modifiers,
                        }));
                        continue;
                    }
                    // class function/procedure (static) in records
                    if self.is_keyword("class") {
                        self.advance();
                        if self.is_keyword("function") || self.is_keyword("procedure") {
                            let _kw = self.advance().text.clone();
                            let mname = self.expect_ident()?;
                            let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                            let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                            self.eat_terminator();
                            let mut modifiers = Modifiers::default();
                            modifiers.is_static = true;
                            members.push(Statement::new(StmtKind::FunctionDecl {
                                name: mname, params, return_type, body: Vec::new(), modifiers,
                            }));
                            continue;
                        }
                        // class operator
                        if self.peek().kind == TokenKind::Ident && self.peek().text.eq_ignore_ascii_case("operator") {
                            self.advance();
                            let op_name = self.expect_ident()?;
                            let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                            let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                            self.eat_terminator();
                            let mut modifiers = Modifiers::default();
                            modifiers.is_static = true;
                            modifiers.extra.push("operator".into());
                            members.push(Statement::new(StmtKind::FunctionDecl {
                                name: op_name, params, return_type, body: Vec::new(), modifiers,
                            }));
                            continue;
                        }
                    }
                    // operator (non-class) in records
                    if self.peek().kind == TokenKind::Ident && self.peek().text.eq_ignore_ascii_case("operator") {
                        self.advance();
                        let op_name = self.expect_ident()?;
                        let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                        let return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                        self.eat_terminator();
                        let mut modifiers = Modifiers::default();
                        modifiers.extra.push("operator".into());
                        members.push(Statement::new(StmtKind::FunctionDecl {
                            name: op_name, params, return_type, body: Vec::new(), modifiers,
                        }));
                        continue;
                    }
                    // Fields: X, Y: Integer;
                    if self.peek().kind == TokenKind::Ident {
                        let mut names = vec![self.expect_ident()?];
                        while self.eat_op(",") {
                            if self.peek().kind == TokenKind::Ident { names.push(self.expect_ident()?); }
                            else { break; }
                        }
                        if self.eat_op(":") {
                            let t = self.parse_type_name()?;
                            self.eat_terminator();
                            for fname in names {
                                members.push(Statement::new(StmtKind::VarDecl { name: fname, type_hint: Some(t.clone()), init: None, is_const: false, mutable: true }));
                            }
                        } else { break; }
                    } else { break; }
                }
                self.eat_keyword("end");
                stmts.push(Statement::with_span(StmtKind::StructDecl { name, members }, self.span_from(line, col)));
            } else {
                // Type alias
                let target = self.parse_type_name()?;
                stmts.push(Statement::with_span(StmtKind::TypeAlias { name, target }, self.span_from(line, col)));
            }
            self.eat_terminator();
        }

        if stmts.len() == 1 { Ok(stmts.pop().unwrap()) }
        else { Ok(Statement::with_span(StmtKind::Block(stmts), self.span_from(line, col))) }
    }

    fn parse_type_name(&mut self) -> Result<String, ParseError> {
        let mut name = String::new();

        // Pointer type: ^TypeName
        if self.eat_op("^") {
            name.push('^');
            name.push_str(&self.parse_type_name()?);
            return Ok(name);
        }

        // Array type: array of X, array[N] of X
        if self.is_keyword("array") {
            name.push_str(&self.advance().text);
            if self.eat_op("[") {
                name.push('[');
                while !self.is_op("]") && !self.at_end() { name.push_str(&self.advance().text); }
                self.eat_op("]");
                name.push(']');
            }
            if self.eat_keyword("of") {
                name.push_str(" of ");
                name.push_str(&self.parse_type_name()?);
            }
            return Ok(name);
        }

        // Procedural type: procedure or procedure(params)
        if self.is_keyword("procedure") {
            self.advance();
            name = "procedure".to_string();
            if self.eat_op("(") {
                name.push('(');
                let mut depth = 1;
                while depth > 0 && !self.at_end() {
                    let t = self.advance();
                    if t.text == "(" { depth += 1; }
                    if t.text == ")" { depth -= 1; if depth == 0 { break; } }
                    name.push_str(&t.text);
                }
                name.push(')');
            }
            return Ok(name);
        }

        // Function type: function(params): ReturnType
        if self.is_keyword("function") {
            self.advance();
            name = "function".to_string();
            if self.eat_op("(") {
                name.push('(');
                let mut depth = 1;
                while depth > 0 && !self.at_end() {
                    let t = self.advance();
                    if t.text == "(" { depth += 1; }
                    if t.text == ")" { depth -= 1; if depth == 0 { break; } }
                    name.push_str(&t.text);
                }
                name.push(')');
            }
            if self.eat_op(":") {
                name.push_str(": ");
                name.push_str(&self.parse_type_name()?);
            }
            return Ok(name);
        }

        // Simple identifier or keyword type (Integer, String, etc.)
        let tok = self.advance();
        name = tok.text.clone();
        // Generic: TList<Integer>
        if self.eat_op("<") {
            name.push('<');
            name.push_str(&self.parse_type_name()?);
            while self.eat_op(",") {
                name.push(',');
                name.push_str(&self.parse_type_name()?);
            }
            self.expect_op(">")?;
            name.push('>');
        }
        Ok(name)
    }

    fn parse_params(&mut self) -> Result<Vec<Param>, ParseError> {
        let open = &self.grammar.params.open;
        let close = &self.grammar.params.close;
        self.expect_op(open)?;
        let mut params = Vec::new();
        while !self.is_op(close) && !self.at_end() {
            // Check for pass-by modifiers
            let mut pass_by = PassBy::Value;
            for (kw, mode) in &self.grammar.params.pass_by {
                if self.is_keyword(kw) {
                    self.advance();
                    pass_by = match mode.as_str() {
                        "ref" => PassBy::Ref,
                        "const" => PassBy::Const,
                        "out" => PassBy::Out,
                        _ => PassBy::Value,
                    };
                    break;
                }
            }

            // Rest/spread prefix
            let is_rest = if let Some(ref prefix) = self.grammar.params.rest_prefix {
                self.eat_op(prefix)
            } else { false };

            let name = self.expect_ident()?;

            // Multi-name params: (a, b: Integer)
            let mut extra_names = Vec::new();
            if self.grammar.params.multi_name {
                while self.eat_op(self.grammar.params.multi_name_sep.as_deref().unwrap_or(",")) {
                    if self.peek().kind == TokenKind::Ident && !self.is_op(close) {
                        extra_names.push(self.expect_ident()?);
                    } else { break; }
                }
            }

            let type_hint = if let Some(ref sep) = self.grammar.params.name_type_sep {
                if self.eat_op(sep) || self.eat_keyword(sep) {
                    Some(self.parse_type_name()?)
                } else { None }
            } else { None };

            let default = if let Some(ref eq) = self.grammar.params.default_value {
                if self.eat_op(eq) { Some(self.parse_expression()?) } else { None }
            } else { None };

            params.push(Param { name: name.clone(), type_hint: type_hint.clone(), default: default.clone(), pass_by, is_rest });
            for en in extra_names {
                params.push(Param { name: en, type_hint: type_hint.clone(), default: default.clone(), pass_by, is_rest: false });
            }

            let sep = &self.grammar.params.separator;
            if !self.eat_op(sep) { break; }
        }
        self.expect_op(close)?;
        Ok(params)
    }

    // ── Expression / Assignment ──────────────────────────────────────────

    fn parse_expr_or_assign(&mut self) -> Result<Statement, ParseError> {
        let line = self.peek().line;
        let col = self.peek().col;
        let expr = self.parse_expression()?;

        // Python type annotation: x: int = 5 or x: int
        if self.is_op(":") && self.grammar.language.indentation_based {
            if let ExprKind::Ident(ref name) = expr.kind {
                self.advance(); // consume ':'
                let type_name = self.parse_expression()?; // parse type (might be complex)
                let type_str = if let ExprKind::Ident(t) = &type_name.kind { Some(t.clone()) } else { None };
                let init = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
                return Ok(Statement::with_span(StmtKind::VarDecl {
                    name: name.clone(), type_hint: type_str, init, is_const: false, mutable: true,
                }, self.span_from(line, col)));
            }
        }

        // Tuple unpacking: a, b = 1, 2
        if self.is_op(",") {
            let mut targets = vec![expr.clone()];
            while self.eat_op(",") {
                if self.at_stmt_end() { break; }
                targets.push(self.parse_expression()?);
            }
            if let Some(ref op) = self.grammar.assignment.operator {
                if self.is_op(op) {
                    self.advance();
                    let first_val = self.parse_expression()?;
                    let mut values = vec![first_val];
                    while self.eat_op(",") {
                        values.push(self.parse_expression()?);
                    }
                    // Emit as Extra with tag "unpack"
                    let target_arr = Expression::new(ExprKind::Array(targets));
                    let value_arr = Expression::new(ExprKind::Array(values));
                    return Ok(Statement::with_span(StmtKind::Assign { target: target_arr, value: value_arr }, self.span_from(line, col)));
                }
            }
            // Not an assignment — it's a tuple expression statement
            let tuple = Expression::new(ExprKind::Array(targets));
            return Ok(Statement::with_span(StmtKind::Expr(tuple), self.span_from(line, col)));
        }

        // VB-style ambiguity: `=` is both assignment and comparison.
        // When `=` is in the infix table AND is the assignment operator,
        // the Pratt parser already consumed it as Binary(Eq, target, value).
        // At statement level, convert back to assignment.
        if self.grammar.assignment.operator.as_deref() == Some("=")
            && self.grammar.operators.infix.iter().any(|lvl| lvl.ops.iter().any(|o| o == "="))
        {
            if let ExprKind::Binary { op: BinOp::Eq, ref left, ref right } = expr.kind {
                match left.kind {
                    ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. } | ExprKind::Call { .. } => {
                        return Ok(Statement::with_span(
                            StmtKind::Assign { target: (**left).clone(), value: (**right).clone() },
                            self.span_from(line, col),
                        ));
                    }
                    _ => {}
                }
            }
        }

        // Check for assignment
        if let Some(ref op) = self.grammar.assignment.operator {
            if self.is_op(op) {
                self.advance();
                let value = self.parse_expression()?;
                return Ok(Statement::with_span(StmtKind::Assign { target: expr, value }, self.span_from(line, col)));
            }
        }

        // Compound assignment
        for (op_str, bin_op_name) in &self.grammar.assignment.compound {
            if self.is_op(op_str) {
                self.advance();
                let value = self.parse_expression()?;
                let bin_op = match bin_op_name.as_str() {
                    "Add" => BinOp::Add, "Sub" => BinOp::Sub,
                    "Mul" => BinOp::Mul, "Div" => BinOp::Div,
                    "Mod" => BinOp::Mod, "Pow" => BinOp::Pow,
                    _ => BinOp::Add,
                };
                return Ok(Statement::with_span(StmtKind::CompoundAssign { target: expr, op: bin_op, value }, self.span_from(line, col)));
            }
        }

        // C-family `=` assignment (when assignment.operator is "=")
        if self.is_op("=") && self.grammar.assignment.operator.as_deref() == Some("=") {
            self.advance();
            let value = self.parse_expression()?;
            return Ok(Statement::with_span(StmtKind::Assign { target: expr, value }, self.span_from(line, col)));
        }

        Ok(Statement::with_span(StmtKind::Expr(expr), self.span_from(line, col)))
    }

    // ── Expression parsing (Pratt parser) ────────────────────────────────

    pub fn parse_expression(&mut self) -> Result<Expression, ParseError> {
        self.pratt_parse(0)
    }

    fn pratt_parse(&mut self, min_prec: u8) -> Result<Expression, ParseError> {
        let mut left = self.parse_prefix()?;

        loop {
            // Check for postfix operators
            left = self.parse_postfix(left)?;

            // Check for infix operators
            let tok = self.peek();
            if tok.kind == TokenKind::Eof { break; }

            let op_text = tok.text.clone();
            // Handle two-word operators: "not in", "is not"
            if self.is_keyword("not") && self.peek2().text == "in" {
                if let Some((prec, _, _)) = self.lookup_infix("in") {
                    if prec >= min_prec {
                        self.advance(); self.advance(); // consume "not" "in"
                        let next_prec = prec + 1;
                        let right = self.pratt_parse(next_prec)?;
                        left = Expression::new(ExprKind::Binary { op: BinOp::NotIn, left: Box::new(left), right: Box::new(right) });
                        continue;
                    }
                }
            }
            if self.is_keyword("is") && self.peek2().text == "not" {
                if let Some((prec, _, _)) = self.lookup_infix("is") {
                    if prec >= min_prec {
                        self.advance(); self.advance(); // consume "is" "not"
                        let next_prec = prec + 1;
                        let right = self.pratt_parse(next_prec)?;
                        left = Expression::new(ExprKind::Binary { op: BinOp::NotEq, left: Box::new(left), right: Box::new(right) });
                        continue;
                    }
                }
            }
            if let Some((prec, assoc, bin_op)) = self.lookup_infix(&op_text) {
                if prec < min_prec { break; }
                self.advance();
                let next_prec = if assoc == Assoc::Right { prec } else { prec + 1 };
                let right = self.pratt_parse(next_prec)?;
                left = Expression::new(ExprKind::Binary {
                    op: bin_op,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            } else if self.is_op("?") && !self.is_op("?.") && !self.is_op("??") {
                // Ternary: expr ? then : else
                self.advance();
                let then = self.parse_expression()?;
                self.expect_op(":")?;
                let else_ = self.pratt_parse(0)?;
                left = Expression::new(ExprKind::Ternary {
                    cond: Box::new(left), then: Box::new(then), else_: Box::new(else_),
                });
            } else if self.is_keyword("if") && self.grammar.language.indentation_based {
                // Python ternary: value if cond else alternative
                self.advance(); // consume 'if'
                let cond = self.pratt_parse(0)?;
                self.expect_keyword("else")?;
                let else_ = self.pratt_parse(0)?;
                left = Expression::new(ExprKind::Ternary {
                    cond: Box::new(cond), then: Box::new(left), else_: Box::new(else_),
                });
            } else {
                break;
            }
        }
        Ok(left)
    }

    fn parse_prefix(&mut self) -> Result<Expression, ParseError> {
        let tok = self.peek().clone();
        let line = tok.line;
        let col = tok.col;

        // Prefix operators (only for operator/keyword tokens, not string literals)
        if (tok.kind == TokenKind::Operator || tok.kind == TokenKind::Keyword) && self.lookup_prefix(&tok.text).is_some() {
            let unary_op = self.lookup_prefix(&tok.text).unwrap();
            self.advance();
            let expr = self.pratt_parse(100)?; // high precedence for prefix
            return Ok(Expression::with_span(ExprKind::Unary { op: unary_op, expr: Box::new(expr) }, self.span_from(line, col)));
        }

        // Primary expressions
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Result<Expression, ParseError> {
        let tok = self.peek().clone();
        let line = tok.line;
        let col = tok.col;

        match tok.kind {
            TokenKind::IntLit => {
                self.advance();
                let n: i64 = tok.text.parse().unwrap_or(0);
                Ok(Expression::with_span(ExprKind::Lit(Literal::Int(n)), self.span_from(line, col)))
            }
            TokenKind::FloatLit => {
                self.advance();
                let n: f64 = tok.text.parse().unwrap_or(0.0);
                Ok(Expression::with_span(ExprKind::Lit(Literal::Float(n)), self.span_from(line, col)))
            }
            TokenKind::StringLit => {
                self.advance();
                Ok(Expression::with_span(ExprKind::Lit(Literal::Str(tok.text.clone())), self.span_from(line, col)))
            }
            TokenKind::CharLit => {
                self.advance();
                let c = tok.text.chars().next().unwrap_or('\0');
                Ok(Expression::with_span(ExprKind::Lit(Literal::Char(c)), self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("true") => {
                self.advance(); Ok(Expression::with_span(ExprKind::Lit(Literal::Bool(true)), self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("false") => {
                self.advance(); Ok(Expression::with_span(ExprKind::Lit(Literal::Bool(false)), self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("nil") || tok.text == "null" || tok.text == "None" || tok.text == "nothing" || tok.text == "undefined" => {
                self.advance(); Ok(Expression::with_span(ExprKind::Lit(Literal::Null), self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("self") || tok.text == "this" || tok.text.eq_ignore_ascii_case("me") => {
                self.advance(); Ok(Expression::with_span(ExprKind::This, self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("super") || tok.text == "base" || tok.text.eq_ignore_ascii_case("mybase") || tok.text == "parent" => {
                self.advance(); Ok(Expression::with_span(ExprKind::Super, self.span_from(line, col)))
            }
            TokenKind::Keyword if tok.text == "new" => {
                self.advance();
                let class = self.parse_primary()?;
                let args = if self.is_op("(") { self.parse_call_args()? } else { Vec::new() };
                Ok(Expression::with_span(ExprKind::New { class: Box::new(class), args }, self.span_from(line, col)))
            }
            // Lambda: lambda params: expr
            TokenKind::Keyword if tok.text == "lambda" => {
                self.advance();
                // Parse comma-separated parameter names (no parens, no types)
                let mut params = Vec::new();
                while self.peek().kind == TokenKind::Ident {
                    let name = self.expect_ident()?;
                    let default = if self.eat_op("=") { Some(self.parse_expression()?) } else { None };
                    params.push(Param { name, type_hint: None, default, pass_by: PassBy::Value, is_rest: false });
                    if !self.eat_op(",") { break; }
                }
                self.expect_op(":")?;
                let body_expr = self.parse_expression()?;
                let body = vec![Statement::new(StmtKind::Return(Some(body_expr)))];
                Ok(Expression::with_span(ExprKind::Lambda { params, body, is_async: false }, self.span_from(line, col)))
            }
            // Anonymous procedure: procedure(params) begin ... end
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("procedure") && !self.grammar.language.expression_language => {
                self.advance();
                let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                let body = self.parse_block()?;
                Ok(Expression::with_span(ExprKind::Lambda { params, body, is_async: false }, self.span_from(line, col)))
            }
            // Anonymous function: function(params): Type begin ... end
            TokenKind::Keyword if tok.text.eq_ignore_ascii_case("function") && self.peek2().kind == TokenKind::Operator && self.peek2().text == "(" => {
                self.advance();
                let params = if self.is_op("(") { self.parse_params()? } else { Vec::new() };
                let _return_type = if self.eat_op(":") { Some(self.parse_type_name()?) } else { None };
                let body = self.parse_block()?;
                Ok(Expression::with_span(ExprKind::Lambda { params, body, is_async: false }, self.span_from(line, col)))
            }
            TokenKind::Ident | TokenKind::Keyword => {
                // Identifiers and keyword-identifiers (Result, etc.)
                self.advance();
                Ok(Expression::with_span(ExprKind::Ident(tok.text.clone()), self.span_from(line, col)))
            }
            TokenKind::Operator if tok.text == "(" => {
                // Parenthesized expression or tuple literal
                self.advance();
                if self.is_op(")") {
                    // Empty tuple ()
                    self.advance();
                    return Ok(Expression::with_span(ExprKind::Array(Vec::new()), self.span_from(line, col)));
                }
                let first = self.parse_expression()?;
                if self.eat_op(",") {
                    // Tuple: (a, b, ...)
                    let mut items = vec![first];
                    while !self.is_op(")") && !self.at_end() {
                        items.push(self.parse_expression()?);
                        if !self.eat_op(",") { break; }
                    }
                    self.expect_op(")")?;
                    Ok(Expression::with_span(ExprKind::Array(items), self.span_from(line, col)))
                } else {
                    // Just parenthesized expression
                    self.expect_op(")")?;
                    Ok(first)
                }
            }
            TokenKind::Operator if tok.text == "[" => {
                // Array/list literal or list comprehension
                self.advance();
                if self.is_op("]") {
                    self.advance();
                    return Ok(Expression::with_span(ExprKind::Array(Vec::new()), self.span_from(line, col)));
                }
                let first = self.parse_expression()?;
                // Check for comprehension: [expr for var in iter]
                if self.is_keyword("for") {
                    self.advance(); // consume 'for'
                    let var = self.expect_ident()?;
                    self.expect_keyword("in")?;
                    let iter = self.parse_expression()?;
                    // Optional filter: if cond
                    let filter = if self.is_keyword("if") {
                        self.advance();
                        Some(self.parse_expression()?)
                    } else { None };
                    self.expect_op("]")?;
                    // Represent as Extra
                    let mut exprs = vec![first, Expression::ident(&var), iter];
                    if let Some(f) = filter { exprs.push(f); }
                    return Ok(Expression::with_span(ExprKind::Extra { tag: "listcomp".into(), exprs }, self.span_from(line, col)));
                }
                // Regular list
                let mut items = vec![first];
                while self.eat_op(",") {
                    if self.is_op("]") { break; }
                    items.push(self.parse_expression()?);
                }
                self.expect_op("]")?;
                Ok(Expression::with_span(ExprKind::Array(items), self.span_from(line, col)))
            }
            TokenKind::Operator if tok.text == "{" => {
                // Dict/object literal: {key: value, ...} or set {1, 2, 3}
                self.advance();
                let mut pairs = Vec::new();
                while !self.is_op("}") && !self.at_end() {
                    let key = self.parse_expression()?;
                    if self.eat_op(":") {
                        let value = self.parse_expression()?;
                        pairs.push((key, value));
                    } else {
                        // Set literal or single-element: treat as array
                        let mut items = vec![key];
                        while self.eat_op(",") {
                            if self.is_op("}") { break; }
                            items.push(self.parse_expression()?);
                        }
                        self.expect_op("}")?;
                        return Ok(Expression::with_span(ExprKind::Array(items), self.span_from(line, col)));
                    }
                    if !self.eat_op(",") { break; }
                }
                self.expect_op("}")?;
                Ok(Expression::with_span(ExprKind::Object(pairs), self.span_from(line, col)))
            }
            _ => {
                Err(self.error(format!("unexpected token: '{}'", tok.text)))
            }
        }
    }

    fn parse_postfix(&mut self, mut expr: Expression) -> Result<Expression, ParseError> {
        loop {
            let tok = self.peek();
            // Member access: .field
            if let Some(ref ma) = self.grammar.expressions.member_access {
                if self.is_op(ma) {
                    self.advance();
                    let field = self.expect_name()?;
                    // Check for method call: obj.method(args)
                    if self.is_op(self.grammar.expressions.call_open.as_deref().unwrap_or("(")) {
                        let args = self.parse_call_args()?;
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(expr), field, null_safe: false,
                            })),
                            args,
                        });
                    } else {
                        expr = Expression::new(ExprKind::Member { object: Box::new(expr), field, null_safe: false });
                    }
                    continue;
                }
            }
            // Optional chain: ?.field
            if let Some(ref oc) = self.grammar.expressions.optional_chain {
                if self.is_op(oc) {
                    self.advance();
                    let field = self.expect_name()?;
                    expr = Expression::new(ExprKind::Member { object: Box::new(expr), field, null_safe: true });
                    continue;
                }
            }
            // Index: [expr] — skip if index_open == call_open (VB uses () for both)
            if let Some(ref io) = self.grammar.expressions.index_open {
                let call_open = self.grammar.expressions.call_open.as_deref().unwrap_or("(");
                if self.is_op(io) && io != call_open {
                    self.advance();
                    let close = self.grammar.expressions.index_close.as_deref().unwrap_or("]");
                    // Check for slice: x[start:end] or x[start:end:step]
                    if self.is_op(":") || {
                        // Peek ahead: parse expr, check if followed by ':'
                        let saved = self.pos;
                        let has_colon = if !self.is_op(close) {
                            let _ = self.parse_expression();
                            let r = self.is_op(":");
                            self.pos = saved;
                            r
                        } else { false };
                        has_colon
                    } {
                        // Slice expression — parse as call to __slice(start, end, step)
                        let start = if self.is_op(":") {
                            Expression::new(ExprKind::Lit(Literal::Null))
                        } else {
                            self.parse_expression()?
                        };
                        self.eat_op(":");
                        let end = if self.is_op(close) || self.is_op(":") {
                            Expression::new(ExprKind::Lit(Literal::Null))
                        } else {
                            self.parse_expression()?
                        };
                        let step = if self.eat_op(":") {
                            if self.is_op(close) {
                                Expression::new(ExprKind::Lit(Literal::Null))
                            } else {
                                self.parse_expression()?
                            }
                        } else {
                            Expression::new(ExprKind::Lit(Literal::Null))
                        };
                        self.expect_op(close)?;
                        // Represent slice as Extra with tag "slice"
                        expr = Expression::new(ExprKind::Extra {
                            tag: "slice".to_string(),
                            exprs: vec![expr, start, end, step],
                        });
                        continue;
                    }
                    // Regular index
                    let index = self.parse_expression()?;
                    self.expect_op(close)?;
                    expr = Expression::new(ExprKind::Index { object: Box::new(expr), index: Box::new(index) });
                    continue;
                }
            }
            // Function call: (args)
            if let Some(ref co) = self.grammar.expressions.call_open {
                if self.is_op(co) {
                    let args = self.parse_call_args()?;
                    expr = Expression::new(ExprKind::Call { callee: Box::new(expr), args });
                    continue;
                }
            }
            // Deref: ^
            if let Some(ref d) = self.grammar.expressions.deref {
                if self.is_op(d) {
                    self.advance();
                    expr = Expression::new(ExprKind::Unary { op: UnaryOp::Deref, expr: Box::new(expr) });
                    continue;
                }
            }
            // Postfix operators (++, --)
            for pop in &self.grammar.operators.postfix {
                if self.is_op(pop) {
                    self.advance();
                    let op = match pop.as_str() {
                        "++" => UnaryOp::PostInc,
                        "--" => UnaryOp::PostDec,
                        _ => continue,
                    };
                    expr = Expression::new(ExprKind::Unary { op, expr: Box::new(expr) });
                    continue;
                }
            }
            break;
        }
        Ok(expr)
    }

    fn parse_call_args(&mut self) -> Result<Vec<Expression>, ParseError> {
        let open = self.grammar.expressions.call_open.as_deref().unwrap_or("(");
        let close = self.grammar.expressions.call_close.as_deref().unwrap_or(")");
        self.expect_op(open)?;
        let args = self.parse_expr_list(close)?;
        self.expect_op(close)?;
        Ok(args)
    }

    fn parse_expr_list(&mut self, close: &str) -> Result<Vec<Expression>, ParseError> {
        let mut items = Vec::new();
        while !self.is_op(close) && !self.at_end() {
            items.push(self.parse_expression()?);
            if !self.eat_op(",") { break; }
        }
        Ok(items)
    }

    // ── Operator lookup ──────────────────────────────────────────────────

    fn lookup_infix(&self, op: &str) -> Option<(u8, Assoc, BinOp)> {
        for level in &self.grammar.operators.infix {
            for lop in &level.ops {
                let matches = if self.grammar.language.case_sensitive {
                    lop == op
                } else {
                    lop.eq_ignore_ascii_case(op)
                };
                if matches {
                    let bin_op = str_to_binop(op, &self.grammar.language.name);
                    return Some((level.precedence, level.assoc, bin_op));
                }
            }
        }
        None
    }

    fn lookup_prefix(&self, op: &str) -> Option<UnaryOp> {
        for pop in &self.grammar.operators.prefix {
            let matches = if self.grammar.language.case_sensitive {
                pop == op
            } else {
                pop.eq_ignore_ascii_case(op)
            };
            if matches {
                return Some(match op.to_lowercase().as_str() {
                    "not" | "!" => UnaryOp::Not,
                    "-" => UnaryOp::Neg,
                    "+" => UnaryOp::Pos,
                    "~" => UnaryOp::BitNot,
                    "++" => UnaryOp::PreInc,
                    "--" => UnaryOp::PreDec,
                    "typeof" => UnaryOp::Typeof,
                    "@" => UnaryOp::AddrOf,
                    _ => return None,
                });
            }
        }
        None
    }
}

fn str_to_binop(s: &str, lang: &str) -> BinOp {
    let lower = s.to_lowercase();
    match lower.as_str() {
        "+" => BinOp::Add, "-" => BinOp::Sub,
        "*" => BinOp::Mul, "/" => BinOp::Div,
        "div" => BinOp::IDiv, "%" | "mod" => BinOp::Mod,
        "**" => BinOp::Pow,
        "=" | "==" | "===" => BinOp::Eq,
        "<>" | "!=" | "!==" => BinOp::NotEq,
        "<" => BinOp::Lt, ">" => BinOp::Gt,
        "<=" => BinOp::Le, ">=" => BinOp::Ge,
        "and" | "&&" | "andalso" => BinOp::And, "or" | "||" | "orelse" => BinOp::Or,
        "xor" => BinOp::Xor,
        "&" if lang == "vb" => BinOp::Concat,
        "&" => BinOp::BitAnd, "|" => BinOp::BitOr,
        "^" if lang == "vb" => BinOp::Pow, // VB: ^ is exponent
        "^" => BinOp::BitXor,
        "\\" => BinOp::IDiv, // VB integer division
        "shl" | "<<" => BinOp::Shl, "shr" | ">>" => BinOp::Shr,
        "in" => BinOp::In, "not in" => BinOp::NotIn,
        "??" => BinOp::NullCoalesce,
        "." => BinOp::Concat,
        "is" => BinOp::Eq, // simplified
        "isnot" => BinOp::NotEq,
        "as" => BinOp::Eq, // simplified
        "instanceof" => BinOp::In, // simplified
        _ => BinOp::Add, // fallback
    }
}

static EOF_TOKEN: Token = Token { kind: TokenKind::Eof, text: String::new(), line: 0, col: 0 };

#[cfg(test)]
mod tests {
    use super::*;
    use crate::lexer::tokenize;
    use crate::grammar::*;

    fn pascal_grammar() -> GrammarDef {
        GrammarDef {
            language: LanguageSpec {
                name: "pascal".into(),
                case_sensitive: false,
                statement_terminator: Terminator::Char(';'),
                indentation_based: false,
                expression_language: false,
            },
            lexer: LexerSpec {
                comment_line: vec!["//".into()],
                comment_block: vec![("{".into(), "}".into())],
                string_delimiters: vec!["'".into()],
                string_escape: Some("''".into()),
                triple_string: Vec::new(),
                string_prefixes: Vec::new(),
                interpolation: None,
                template_string: None,
                char_prefix: Some("#".into()),
                hex_prefix: Some("$".into()),
                keywords: vec![
                    "program".into(),"begin".into(),"end".into(),"var".into(),"const".into(),"type".into(),
                    "procedure".into(),"function".into(),"constructor".into(),"destructor".into(),
                    "if".into(),"then".into(),"else".into(),
                    "for".into(),"to".into(),"downto".into(),"do".into(),"in".into(),
                    "while".into(),"repeat".into(),"until".into(),
                    "case".into(),"of".into(),"otherwise".into(),
                    "class".into(),"record".into(),"interface".into(),"inherited".into(),
                    "override".into(),"virtual".into(),"abstract".into(),
                    "try".into(),"except".into(),"finally".into(),"raise".into(),
                    "and".into(),"or".into(),"not".into(),"xor".into(),"div".into(),"mod".into(),"shl".into(),"shr".into(),
                    "nil".into(),"true".into(),"false".into(),
                    "exit".into(),"break".into(),"continue".into(),
                    "with".into(),"is".into(),"as".into(),
                    "forward".into(),"result".into(),"self".into(),
                    "public".into(),"private".into(),"protected".into(),"published".into(),
                    "array".into(),"string".into(),"integer".into(),"real".into(),"boolean".into(),"char".into(),
                    "writeln".into(),
                ],
                operators: vec![
                    ":=".into(),"+=".into(),"-=".into(),"*=".into(),"/=".into(),
                    "<>".into(),"<=".into(),">=".into(),"..".into(),
                    "+".into(),"-".into(),"*".into(),"/".into(),
                    "=".into(),"<".into(),">".into(),
                    "(".into(),")".into(),"[".into(),"]".into(),
                    ".".into(),",".into(),";".into(),":".into(),"^".into(),"@".into(),
                ],
            },
            operators: OperatorTable {
                prefix: vec!["not".into(), "-".into(), "@".into()],
                postfix: Vec::new(),
                infix: vec![
                    InfixLevel { precedence: 1, ops: vec!["or".into(), "xor".into()], assoc: Assoc::Left },
                    InfixLevel { precedence: 2, ops: vec!["and".into()], assoc: Assoc::Left },
                    InfixLevel { precedence: 3, ops: vec!["=".into(),"<>".into(),"<".into(),">".into(),"<=".into(),">=".into(),"in".into(),"is".into(),"as".into()], assoc: Assoc::Left },
                    InfixLevel { precedence: 4, ops: vec!["+".into(), "-".into()], assoc: Assoc::Left },
                    InfixLevel { precedence: 5, ops: vec!["*".into(), "/".into(), "div".into(), "mod".into(), "shl".into(), "shr".into()], assoc: Assoc::Left },
                ],
            },
            blocks: BlockSpec { open: "begin".into(), close: "end".into(), prefix: None, close_with_kind: false },
            types: TypeSpec { position: TypePosition::After, separator: Some(":".into()), return_separator: None },
            statements: Vec::new(),
            declarations: Vec::new(),
            expressions: ExpressionSpec {
                member_access: Some(".".into()),
                optional_chain: None,
                index_open: Some("[".into()),
                index_close: Some("]".into()),
                call_open: Some("(".into()),
                call_close: Some(")".into()),
                deref: Some("^".into()),
                primary_forms: Vec::new(),
            },
            params: ParamSpec {
                open: "(".into(), close: ")".into(), separator: ";".into(),
                name_type_sep: Some(":".into()), type_position: TypePosition::After,
                default_value: Some("=".into()),
                rest_prefix: None, kwargs_prefix: None,
                multi_name: true, multi_name_sep: Some(",".into()),
                pass_by: [("var".into(), "ref".into()), ("const".into(), "const".into())].into_iter().collect(),
            },
            assignment: AssignmentSpec {
                operator: Some(":=".into()),
                compound: [("+=".into(),"Add".into()),("-=".into(),"Sub".into()),("*=".into(),"Mul".into()),("/=".into(),"Div".into())].into_iter().collect(),
                walrus: None,
            },
            program: ProgramSpec { header: None, uses: None, body: None },
        }
    }

    #[test]
    fn parse_pascal_hello() {
        let g = pascal_grammar();
        let tokens = tokenize("program Test; begin WriteLn('hello'); end.", &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        assert_eq!(module.name, "Test");
        assert!(!module.body.is_empty());
        // The body should contain an expression statement: WriteLn('hello')
        if let StmtKind::Expr(ref expr) = module.body[0].kind {
            if let ExprKind::Call { ref callee, ref args } = expr.kind {
                if let ExprKind::Ident(ref name) = callee.kind {
                    assert!(name.eq_ignore_ascii_case("WriteLn"));
                }
                assert_eq!(args.len(), 1);
            } else { panic!("expected Call, got {:?}", expr.kind); }
        } else { panic!("expected Expr, got {:?}", module.body[0].kind); }
    }

    #[test]
    fn parse_pascal_var_and_assign() {
        let g = pascal_grammar();
        let tokens = tokenize("program T; var x: Integer; begin x := 42; end.", &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        // Should have var decl + assignment in body
        assert!(module.body.len() >= 2);
    }

    #[test]
    fn parse_pascal_function() {
        let g = pascal_grammar();
        let src = "program T; function Add(a, b: Integer): Integer; begin Result := a + b; end; begin WriteLn(Add(3, 4)); end.";
        let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        // Should have function decl + call
        let has_func = module.body.iter().any(|s| matches!(s.kind, StmtKind::FunctionDecl { .. }));
        assert!(has_func, "expected function declaration");
    }

    #[test]
    fn parse_pascal_if_else() {
        let g = pascal_grammar();
        let src = "program T; var x: Integer; begin x := 5; if x > 3 then WriteLn('yes') else WriteLn('no'); end.";
        let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        let has_if = module.body.iter().any(|s| matches!(s.kind, StmtKind::If { .. }));
        assert!(has_if, "expected if statement");
    }

    #[test]
    fn parse_pascal_for_loop() {
        let g = pascal_grammar();
        let src = "program T; var i: Integer; begin for i := 1 to 5 do WriteLn(i); end.";
        let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        let has_for = module.body.iter().any(|s| matches!(s.kind, StmtKind::For { .. }));
        assert!(has_for, "expected for loop");
    }

    #[test]
    fn parse_pascal_class() {
        let g = pascal_grammar();
        let src = r#"program T;
type TFoo = class
  public
    FVal: Integer;
    constructor Create;
    function GetVal: Integer;
  end;
constructor TFoo.Create;
begin
end;
function TFoo.GetVal: Integer;
begin
  Result := FVal;
end;
begin
end."#;
        let tokens = tokenize(src, &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        let has_class = module.body.iter().any(|s| matches!(s.kind, StmtKind::ClassDecl { .. }));
        assert!(has_class, "expected class declaration, got: {:?}", module.body.iter().map(|s| std::mem::discriminant(&s.kind)).collect::<Vec<_>>());
    }

    #[test]
    fn parse_pascal_expression_precedence() {
        let g = pascal_grammar();
        let tokens = tokenize("program T; begin WriteLn(2 + 3 * 4); end.", &g.lexer, &g.language.statement_terminator, false, false);
        let module = parse(&tokens, &g).unwrap();
        // 2 + (3 * 4) = 14, not (2 + 3) * 4 = 20
        // The AST should have Binary(+, 2, Binary(*, 3, 4))
        if let StmtKind::Expr(ref expr) = module.body[0].kind {
            if let ExprKind::Call { ref args, .. } = expr.kind {
                if let ExprKind::Binary { ref op, ref right, .. } = args[0].kind {
                    assert_eq!(*op, BinOp::Add);
                    if let ExprKind::Binary { ref op, .. } = right.kind {
                        assert_eq!(*op, BinOp::Mul);
                    } else { panic!("expected nested Mul"); }
                } else { panic!("expected Binary Add"); }
            }
        }
    }
}
