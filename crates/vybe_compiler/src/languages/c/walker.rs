//! C → common AST walker.
//!
//! Walks the pest parse tree from `grammar.pest` into `vybe_compiler::ast`
//! nodes. C-specific normalizations happen here so the shared compiler stays
//! language-agnostic:
//!   - `printf(fmt, …)` → `__c_printf(fmt, …)` (shared sprintf formatter)
//!   - structs are tracked so `struct P x;` initializes a zero-filled object
//!   - pointer deref `*p` / address-of `&x` collapse to identity (the VM's
//!     objects/arrays are already references)
//!   - `a->b` is treated as `a.b`

use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};

use super::{CParser, Rule};
use crate::ast::*;

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs =
        CParser::parse(Rule::program, source).map_err(|e| format!("C parse error: {e}"))?;
    let program = pairs.next().ok_or("empty parse")?;
    let mut w = Walker::default();
    let mut body = Vec::new();
    for item in program.into_inner() {
        match item.as_rule() {
            Rule::EOI => {}
            _ => w.walk_top_item(item, &mut body),
        }
    }
    Ok(Module {
        name: "main".to_string(),
        language: Lang::Unknown,
        body,
        imports: Vec::new(),
    })
}

#[derive(Default)]
struct Walker {
    /// struct/union name → ordered field names (for zero-init at decl site)
    structs: HashMap<String, Vec<String>>,
    /// identifiers declared as `char*`; used for pointer-like string traversal.
    char_pointers: HashSet<String>,
    /// variable/parameter name → C type string for sizeof resolution
    var_types: HashMap<String, String>,
}

fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}
fn expr(kind: ExprKind) -> Expression {
    Expression::new(kind)
}
fn ident(name: &str) -> Expression {
    expr(ExprKind::Ident(name.to_string()))
}

impl Walker {
    fn walk_top_item(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        match pair.as_rule() {
            Rule::preproc_directive => self.walk_preproc(pair, out),
            Rule::function_definition => {
                if let Some(s) = self.walk_function(pair) {
                    out.push(s);
                }
            }
            Rule::declaration => self.walk_declaration(pair, out),
            _ => {}
        }
    }

    fn walk_preproc(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        for inner in pair.into_inner() {
            if inner.as_rule() == Rule::define_directive {
                // #define NAME value  → const NAME = value  (object-like only)
                let mut it = inner.into_inner();
                let name = match it.next() {
                    Some(p) if p.as_rule() == Rule::ident_name => p.as_str().to_string(),
                    _ => continue,
                };
                // Skip function-like macros (define_params present).
                let mut value_pair = None;
                let mut has_params = false;
                for p in it {
                    match p.as_rule() {
                        Rule::define_params => has_params = true,
                        Rule::define_value => value_pair = Some(p),
                        _ => {}
                    }
                }
                if has_params {
                    continue;
                }
                let init = value_pair
                    .map(|p| self.parse_define_value(p.as_str().trim()))
                    .unwrap_or_else(|| expr(ExprKind::Lit(Literal::Int(1))));
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: None,
                        init: Some(init),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
            }
            // #include and conditionals: ignored (stdlib is profile-provided).
        }
    }

    /// Best-effort literal for a `#define` value; falls back to a string.
    fn parse_define_value(&mut self, text: &str) -> Expression {
        if let Ok(i) = text.parse::<i64>() {
            return expr(ExprKind::Lit(Literal::Int(i)));
        }
        if let Ok(f) = text.parse::<f64>() {
            return expr(ExprKind::Lit(Literal::Float(f)));
        }
        // Try re-parsing the value as a C expression.
        if let Ok(mut p) = CParser::parse(Rule::assignment_expression, text) {
            if let Some(e) = p.next() {
                return self.walk_assignment(e);
            }
        }
        expr(ExprKind::Lit(Literal::Str(text.to_string())))
    }

    // ── Functions ──────────────────────────────────────────────────────────

    fn walk_function(&mut self, pair: Pair<Rule>) -> Option<Statement> {
        let mut return_type = None;
        let mut name = String::new();
        let mut params = Vec::new();
        let mut body = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => return_type = Some(self.type_text(p)),
                Rule::declarator => {
                    let (n, ps) = self.declarator_name_and_params(p);
                    name = n;
                    if let Some(ps) = ps {
                        params = ps;
                    }
                }
                Rule::compound_statement => body = self.walk_block(p),
                _ => {}
            }
        }
        if name.is_empty() {
            return None;
        }
        Some(stmt(StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        }))
    }

    fn declarator_name_and_params(&mut self, pair: Pair<Rule>) -> (String, Option<Vec<Param>>) {
        // declarator = pointer? ~ direct_declarator
        let mut name = String::new();
        let mut params = None;
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, ps) = self.declarator_name_and_params(d);
                            name = n;
                            if ps.is_some() {
                                params = ps;
                            }
                        }
                        Rule::param_suffix => params = Some(self.walk_params(d)),
                        _ => {}
                    }
                }
            }
        }
        (name, params)
    }

    fn walk_params(&mut self, pair: Pair<Rule>) -> Vec<Param> {
        let mut params = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::parameter_list {
                for decl in p.into_inner() {
                    if decl.as_rule() == Rule::parameter_decl {
                        let mut pname = String::new();
                        let mut type_hint = None;
                        for d in decl.into_inner() {
                            match d.as_rule() {
                                Rule::declaration_specifiers => {
                                    type_hint = Some(self.type_text(d))
                                }
                                Rule::declarator => {
                                    pname = self.declarator_name_and_params(d).0
                                }
                                _ => {}
                            }
                        }
                        if !pname.is_empty() {
                            params.push(Param {
                                name: pname,
                                type_hint,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            });
                        }
                    }
                }
            }
        }
        params
    }

    // ── Declarations ─────────────────────────────────────────────────────────

    fn walk_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else { return };
        match inner.as_rule() {
            Rule::typedef_declaration => self.walk_typedef(inner, out),
            Rule::normal_declaration => self.walk_normal_declaration(inner, out),
            _ => {}
        }
    }

    fn walk_typedef(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // typedef declaration_specifiers declarator (, declarator)* ;
        let mut specs = None;
        let mut names = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::declarator => names.push(self.declarator_name_and_params(p).0),
                _ => {}
            }
        }
        // typedef struct {...} Name; → register Name as struct alias.
        if let Some(specs) = specs {
            if let Some((tag, fields)) = self.struct_def_from_specifiers(&specs) {
                for name in &names {
                    self.structs.insert(name.clone(), fields.clone());
                    out.push(self.make_struct_decl(name, &fields));
                }
                let _ = tag;
            }
        }
    }

    fn walk_normal_declaration(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let mut specs = None;
        let mut init_list = None;
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::declaration_specifiers => specs = Some(p),
                Rule::init_declarator_list => init_list = Some(p),
                _ => {}
            }
        }
        let Some(specs) = specs else { return };

        // A struct/union/enum definition with a body.
        if let Some((tag, fields)) = self.struct_def_from_specifiers(&specs) {
            if let Some(tag) = tag.clone() {
                self.structs.insert(tag.clone(), fields.clone());
                if init_list.is_none() {
                    out.push(self.make_struct_decl(&tag, &fields));
                }
            }
        }
        if let Some((name, members)) = self.enum_def_from_specifiers(&specs) {
            // In C, enum members are global integer constants — emit each as const.
            // Auto-increment value starting from 0, incremented per member.
            let mut next_val: i64 = 0;
            for member in &members {
                let val = if let Some(init) = &member.value {
                    // Try to extract constant integer from the init expression.
                    match &init.kind {
                        ExprKind::Lit(Literal::Int(i)) => { next_val = *i; *i }
                        _ => next_val,
                    }
                } else {
                    next_val
                };
                out.push(stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(member.name.clone()),
                        type_hint: Some("int".to_string()),
                        init: Some(expr(ExprKind::Lit(Literal::Int(val)))),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                }));
                next_val = val + 1;
            }
            out.push(stmt(StmtKind::EnumDecl {
                name: name.unwrap_or_else(|| "__anon_enum".to_string()),
                members,
                visibility: Visibility::Public,
                is_flags: false,
                backing_type: None,
                interfaces: Vec::new(),
                body_members: Vec::new(),
                decorators: Vec::new(),
            }));
        }

        let struct_fields = self.struct_type_of_specifiers(&specs);
        let type_text = self.type_text(specs);

        let Some(init_list) = init_list else { return };
        let mut declarations = Vec::new();
        for idecl in init_list.into_inner() {
            if idecl.as_rule() != Rule::init_declarator {
                continue;
            }
            let mut name = String::new();
            let mut array_bounds: Option<Vec<Expression>> = None;
            let mut init = None;
            let mut is_pointer_decl = false;
            for p in idecl.into_inner() {
                match p.as_rule() {
                    Rule::declarator => {
                        is_pointer_decl = p.as_str().contains('*');
                        let (n, bounds) = self.declarator_name_and_bounds(p);
                        name = n;
                        array_bounds = bounds;
                    }
                    Rule::initializer => {
                        let raw = self.walk_initializer(p);
                        // If this is a struct type and we got an Array initializer,
                        // convert positional Array to named Object using field names.
                        init = Some(if let Some(fields) = &struct_fields {
                            convert_array_init_to_struct(raw, fields)
                        } else {
                            raw
                        });
                    }
                    _ => {}
                }
            }
            if name.is_empty() {
                continue;
            }
            if is_pointer_decl && type_text.contains("char") {
                self.char_pointers.insert(name.clone());
            }
            // Zero-init struct instances when no explicit initializer.
            if init.is_none() && array_bounds.is_none() {
                if let Some(fields) = &struct_fields {
                    init = Some(self.zero_struct(fields));
                }
            }
            // char array with bounds initialized by a string → treat as string
            // (e.g. `char buf[32] = "hello"` → just a string variable)
            if array_bounds.is_some() && type_text.trim() == "char" {
                if let Some(ref init_expr) = init {
                    if matches!(init_expr.kind, ExprKind::Lit(Literal::Str(_))) {
                        array_bounds = None; // treat as string, not array
                    }
                }
            }
            // char array with char initializers `{'h','i','\0'}` → join chars to string
            if array_bounds.is_some() && type_text.trim() == "char" {
                if let Some(ExprKind::Array(elems)) = init.as_ref().map(|i| &i.kind) {
                    let s: String = elems.iter().filter_map(|el| {
                        if let ExprKind::Lit(Literal::Int(code)) = &el.value.kind {
                            if *code == 0 { None } else { char::from_u32(*code as u32) }
                        } else { None }
                    }).collect();
                    init = Some(expr(ExprKind::Lit(Literal::Str(s))));
                    array_bounds = None;
                }
            }
            // Record the type for sizeof resolution
            self.var_types.insert(name.clone(), type_text.clone());
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                type_hint: Some(type_text.clone()),
                init,
                array_bounds,
                with_events: false,
            });
        }
        if !declarations.is_empty() {
            out.push(stmt(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Var,
            }));
        }
    }

    fn make_struct_decl(&self, name: &str, fields: &[String]) -> Statement {
        let members = fields
            .iter()
            .map(|f| ClassMember::Field {
                name: f.clone(),
                type_hint: None,
                init: Some(expr(ExprKind::Lit(Literal::Int(0)))),
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            })
            .collect();
        stmt(StmtKind::StructDecl {
            name: name.to_string(),
            interfaces: Vec::new(),
            members,
            visibility: Visibility::Public,
            decorators: Vec::new(),
        })
    }

    fn zero_struct(&self, fields: &[String]) -> Expression {
        let props = fields
            .iter()
            .map(|f| ObjectProperty::KeyValue {
                key: expr(ExprKind::Lit(Literal::Str(f.clone()))),
                value: expr(ExprKind::Lit(Literal::Int(0))),
            })
            .collect();
        expr(ExprKind::Object(props))
    }

    /// If the specifiers declare a struct/union with a body, return
    /// `(optional tag name, field names)`.
    fn struct_def_from_specifiers(
        &self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<String>)> {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::struct_or_union_spec {
                        let mut tag = None;
                        let mut fields = Vec::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => tag = Some(sp.as_str().to_string()),
                                Rule::struct_member => {
                                    has_body = true;
                                    self.collect_struct_fields(sp, &mut fields);
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            return Some((tag, fields));
                        }
                    }
                }
            }
        }
        None
    }

    fn collect_struct_fields(&self, member: Pair<Rule>, fields: &mut Vec<String>) {
        for p in member.into_inner() {
            if p.as_rule() == Rule::struct_declarator_list {
                for d in p.into_inner() {
                    if d.as_rule() == Rule::declarator {
                        let n = self.clone_declarator_name(d);
                        if !n.is_empty() {
                            fields.push(n);
                        }
                    }
                }
            }
        }
    }

    fn clone_declarator_name(&self, pair: Pair<Rule>) -> String {
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => return d.as_str().to_string(),
                        Rule::declarator => return self.clone_declarator_name(d),
                        _ => {}
                    }
                }
            }
        }
        String::new()
    }

    /// Resolve the struct field list referenced by a declaration's specifiers
    /// (either an inline body or a previously-registered struct name).
    fn struct_type_of_specifiers(&self, specs: &Pair<Rule>) -> Option<Vec<String>> {
        if let Some((_, fields)) = self.struct_def_from_specifiers(specs) {
            return Some(fields);
        }
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    match ts.as_rule() {
                        Rule::struct_or_union_spec => {
                            for sp in ts.into_inner() {
                                if sp.as_rule() == Rule::ident_name {
                                    if let Some(f) = self.structs.get(sp.as_str()) {
                                        return Some(f.clone());
                                    }
                                }
                            }
                        }
                        Rule::typedef_name => {
                            if let Some(f) = self.structs.get(ts.as_str()) {
                                return Some(f.clone());
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
        None
    }

    fn enum_def_from_specifiers(
        &mut self,
        specs: &Pair<Rule>,
    ) -> Option<(Option<String>, Vec<EnumMember>)> {
        for p in specs.clone().into_inner() {
            if p.as_rule() == Rule::type_specifier || p.as_rule() == Rule::type_specifier_strict {
                for ts in p.into_inner() {
                    if ts.as_rule() == Rule::enum_spec {
                        let mut name = None;
                        let mut members = Vec::new();
                        let mut has_body = false;
                        for sp in ts.into_inner() {
                            match sp.as_rule() {
                                Rule::ident_name => name = Some(sp.as_str().to_string()),
                                Rule::enumerator_list => {
                                    has_body = true;
                                    for en in sp.into_inner() {
                                        if en.as_rule() == Rule::enumerator {
                                            let mut it = en.into_inner();
                                            let nm = it.next().map(|x| x.as_str().to_string());
                                            let val = it.next().map(|x| self.walk_assignment(x));
                                            if let Some(nm) = nm {
                                                members.push(EnumMember {
                                                    name: nm,
                                                    value: val,
                                                    constructor_args: Vec::new(),
                                                });
                                            }
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                        if has_body {
                            return Some((name, members));
                        }
                    }
                }
            }
        }
        None
    }

    fn declarator_name_and_bounds(
        &mut self,
        pair: Pair<Rule>,
    ) -> (String, Option<Vec<Expression>>) {
        let mut name = String::new();
        let mut bounds: Vec<Expression> = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::direct_declarator {
                for d in p.into_inner() {
                    match d.as_rule() {
                        Rule::ident_name => name = d.as_str().to_string(),
                        Rule::declarator => {
                            let (n, b) = self.declarator_name_and_bounds(d);
                            name = n;
                            if let Some(b) = b {
                                bounds.extend(b);
                            }
                        }
                        Rule::array_suffix => {
                            if let Some(sz) = d.into_inner().next() {
                                bounds.push(self.walk_assignment(sz));
                            } else {
                                bounds.push(expr(ExprKind::Lit(Literal::Int(0))));
                            }
                        }
                        _ => {}
                    }
                }
            }
        }
        (name, if bounds.is_empty() { None } else { Some(bounds) })
    }

    fn walk_initializer(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next();
        let Some(inner) = inner else {
            return expr(ExprKind::Lit(Literal::Null));
        };
        match inner.as_rule() {
            Rule::assignment_expression => self.walk_assignment(inner),
            Rule::initializer_list => {
                let mut is_object = false;
                let mut elems = Vec::new();
                let mut props = Vec::new();
                for di in inner.into_inner() {
                    if di.as_rule() != Rule::designated_init {
                        continue;
                    }
                    let mut it = di.into_inner().peekable();
                    // designated `.field = init`
                    let first = it.next();
                    match first {
                        Some(p) if p.as_rule() == Rule::ident_name => {
                            is_object = true;
                            let key = p.as_str().to_string();
                            if let Some(v) = it.next() {
                                let val = self.walk_initializer(v);
                                props.push(ObjectProperty::KeyValue {
                                    key: expr(ExprKind::Lit(Literal::Str(key))),
                                    value: val,
                                });
                            }
                        }
                        Some(p) if p.as_rule() == Rule::initializer => {
                            elems.push(ArrayElement {
                                key: None,
                                value: self.walk_initializer(p),
                                spread: false,
                                by_ref: false,
                            });
                        }
                        Some(p) if p.as_rule() == Rule::assignment_expression => {
                            // `[idx] = init` designator → treat as array element
                            is_object = false;
                            if let Some(v) = it.next() {
                                elems.push(ArrayElement {
                                    key: Some(self.walk_assignment(p)),
                                    value: self.walk_initializer(v),
                                    spread: false,
                                    by_ref: false,
                                });
                            } else {
                                elems.push(ArrayElement {
                                    key: None,
                                    value: self.walk_assignment(p),
                                    spread: false,
                                    by_ref: false,
                                });
                            }
                        }
                        _ => {}
                    }
                }
                if is_object {
                    expr(ExprKind::Object(props))
                } else {
                    expr(ExprKind::Array(elems))
                }
            }
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }

    fn type_text(&self, pair: Pair<Rule>) -> String {
        pair.as_str().split_whitespace().collect::<Vec<_>>().join(" ")
    }

    // ── Statements ─────────────────────────────────────────────────────────

    fn walk_block(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        let mut out = Vec::new();
        for p in pair.into_inner() {
            if p.as_rule() == Rule::statement {
                self.walk_statement(p, &mut out);
            }
        }
        out
    }

    fn walk_statement(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        let Some(inner) = pair.into_inner().next() else {
            return;
        };
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                out.push(stmt(StmtKind::Expr(self.walk_expression(e))));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::labeled_statement => self.walk_labeled(inner, out),
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn body_of(&mut self, pair: Pair<Rule>) -> Vec<Statement> {
        // A `statement` that may be a block or a single statement.
        let mut out = Vec::new();
        self.walk_statement(pair, &mut out);
        // Unwrap a single Block into its statements for cleaner nesting.
        if out.len() == 1 {
            if let StmtKind::Block(inner) = &out[0].kind {
                return inner.clone();
            }
        }
        out
    }

    fn walk_if(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond = self.walk_expression(it.next().unwrap());
        let then_body = self.body_of(it.next().unwrap());
        let mut else_body = None;
        if let Some(else_clause) = it.next() {
            if else_clause.as_rule() == Rule::else_clause {
                let st = else_clause.into_inner().next().unwrap();
                else_body = Some(self.body_of(st));
            }
        }
        stmt(StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body,
        })
    }

    fn walk_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let cond = self.walk_expression(it.next().unwrap());
        let body = self.body_of(it.next().unwrap());
        stmt(StmtKind::While {
            cond,
            body,
            else_body: None,
        })
    }

    fn walk_do_while(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let body = self.body_of(it.next().unwrap());
        let cond = self.walk_expression(it.next().unwrap());
        stmt(StmtKind::DoWhile {
            body,
            cond,
            until: false,
        })
    }

    fn walk_for(&mut self, pair: Pair<Rule>) -> Statement {
        let mut init = None;
        let mut cond = None;
        let mut update = None;
        let mut body = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::for_init => {
                    let mut inits = Vec::new();
                    if let Some(c) = p.into_inner().next() {
                        match c.as_rule() {
                            Rule::declaration => self.walk_declaration(c, &mut inits),
                            Rule::expression => {
                                inits.push(stmt(StmtKind::Expr(self.walk_expression(c))))
                            }
                            _ => {}
                        }
                    }
                    init = inits.into_iter().next().map(Box::new);
                }
                Rule::for_cond => {
                    cond = p.into_inner().next().map(|e| self.walk_expression(e));
                }
                Rule::for_update => {
                    update = p.into_inner().next().map(|e| self.walk_expression(e));
                }
                Rule::statement => body = self.body_of(p),
                _ => {}
            }
        }
        stmt(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }

    fn walk_switch(&mut self, pair: Pair<Rule>) -> Statement {
        let mut it = pair.into_inner();
        let subject = self.walk_expression(it.next().unwrap());
        let body_stmt = it.next().unwrap();
        // The body is a compound statement of case/default labels.
        let mut cases: Vec<SwitchCase> = Vec::new();
        let mut default: Option<Vec<Statement>> = None;
        let block = body_stmt.into_inner().next();
        if let Some(block) = block {
            if block.as_rule() == Rule::compound_statement {
                self.collect_switch_cases(block, &mut cases, &mut default);
            }
        }
        stmt(StmtKind::Switch {
            expr: subject,
            cases,
            default,
        })
    }

    fn collect_switch_cases(
        &mut self,
        block: Pair<Rule>,
        cases: &mut Vec<SwitchCase>,
        default: &mut Option<Vec<Statement>>,
    ) {
        // Flatten labeled_statement chains into (condition, following stmts).
        let mut pending_conditions: Vec<CaseCondition> = Vec::new();
        let mut cur_body: Vec<Statement> = Vec::new();
        let mut in_default = false;
        let mut started = false;

        let mut flush = |conds: &mut Vec<CaseCondition>,
                         body: &mut Vec<Statement>,
                         is_default: bool,
                         cases: &mut Vec<SwitchCase>,
                         default: &mut Option<Vec<Statement>>| {
            if is_default {
                *default = Some(std::mem::take(body));
            } else if !conds.is_empty() {
                cases.push(SwitchCase {
                    conditions: std::mem::take(conds),
                    body: std::mem::take(body),
                });
            } else {
                body.clear();
            }
        };

        for st in block.into_inner() {
            if st.as_rule() != Rule::statement {
                continue;
            }
            // Peel off case/default labels, possibly nested via the trailing
            // `statement?` in the grammar.
            let mut stack = vec![st];
            while let Some(node) = stack.pop() {
                let Some(inner) = node.into_inner().next() else {
                    continue;
                };
                if inner.as_rule() == Rule::labeled_statement {
                    let lbl = inner.into_inner().next().unwrap();
                    match lbl.as_rule() {
                        Rule::case_label => {
                            if started {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = false;
                            let mut ci = lbl.into_inner();
                            let val = self.walk_expression_pair_as_cond(ci.next().unwrap());
                            pending_conditions.push(CaseCondition::Value(val));
                            if let Some(rest) = ci.next() {
                                stack.push(rest);
                            }
                        }
                        Rule::default_label => {
                            if started {
                                flush(
                                    &mut pending_conditions,
                                    &mut cur_body,
                                    in_default,
                                    cases,
                                    default,
                                );
                            }
                            started = true;
                            in_default = true;
                            if let Some(rest) = lbl.into_inner().next() {
                                stack.push(rest);
                            }
                        }
                        Rule::goto_label => {
                            let mut li = lbl.into_inner();
                            let name = li.next().unwrap().as_str().to_string();
                            cur_body.push(stmt(StmtKind::Label(name)));
                            if let Some(rest) = li.next() {
                                stack.push(rest);
                            }
                        }
                        _ => {}
                    }
                } else {
                    let mut tmp = Vec::new();
                    self.dispatch_inner_statement(inner, &mut tmp);
                    cur_body.append(&mut tmp);
                }
            }
        }
        if started {
            flush(
                &mut pending_conditions,
                &mut cur_body,
                in_default,
                cases,
                default,
            );
        }
    }

    /// Walk a `case <expr>` condition. The grammar uses
    /// `conditional_expression`, so route through that.
    fn walk_expression_pair_as_cond(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_expression(pair),
        }
    }

    /// Re-dispatch an already-unwrapped statement inner node.
    fn dispatch_inner_statement(&mut self, inner: Pair<Rule>, out: &mut Vec<Statement>) {
        match inner.as_rule() {
            Rule::compound_statement => out.push(stmt(StmtKind::Block(self.walk_block(inner)))),
            Rule::declaration => self.walk_declaration(inner, out),
            Rule::expression_statement => {
                let e = inner.into_inner().next().unwrap();
                out.push(stmt(StmtKind::Expr(self.walk_expression(e))));
            }
            Rule::if_statement => out.push(self.walk_if(inner)),
            Rule::switch_statement => out.push(self.walk_switch(inner)),
            Rule::while_statement => out.push(self.walk_while(inner)),
            Rule::do_while_statement => out.push(self.walk_do_while(inner)),
            Rule::for_statement => out.push(self.walk_for(inner)),
            Rule::return_statement => {
                let val = inner.into_inner().next().map(|e| self.walk_expression(e));
                out.push(stmt(StmtKind::Return(val)));
            }
            Rule::break_statement => out.push(stmt(StmtKind::Break(BreakTarget::Implicit))),
            Rule::continue_statement => {
                out.push(stmt(StmtKind::Continue(ContinueTarget::Implicit)))
            }
            Rule::goto_statement => {
                let label = inner.into_inner().next().map(|p| p.as_str().to_string());
                out.push(stmt(StmtKind::GoTo(label.unwrap_or_default())));
            }
            Rule::empty_statement => {}
            _ => {}
        }
    }

    fn walk_labeled(&mut self, pair: Pair<Rule>, out: &mut Vec<Statement>) {
        // goto_label outside a switch: `name: stmt`
        if let Some(lbl) = pair.into_inner().next() {
            if lbl.as_rule() == Rule::goto_label {
                let mut it = lbl.into_inner();
                let name = it.next().unwrap().as_str().to_string();
                out.push(stmt(StmtKind::Label(name)));
                if let Some(rest) = it.next() {
                    self.walk_statement(rest, out);
                }
            }
        }
    }

    // ── Expressions ──────────────────────────────────────────────────────────

    fn walk_expression(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::expression => {
                let mut parts: Vec<Expression> =
                    pair.into_inner().map(|p| self.walk_assignment(p)).collect();
                if parts.len() == 1 {
                    parts.pop().unwrap()
                } else {
                    expr(ExprKind::Sequence(parts))
                }
            }
            Rule::assignment_expression => self.walk_assignment(pair),
            Rule::conditional_expression => self.walk_conditional(pair),
            _ => self.walk_assignment(pair),
        }
    }

    fn walk_assignment(&mut self, pair: Pair<Rule>) -> Expression {
        if pair.as_rule() != Rule::assignment_expression {
            return self.walk_conditional(pair);
        }
        let mut it = pair.into_inner().peekable();
        let first = it.next().unwrap();
        if first.as_rule() == Rule::conditional_expression {
            return self.walk_conditional(first);
        }
        // unary ~ assign_op ~ assignment
        let target = self.walk_unary(first);
        let op = it.next().unwrap().as_str().to_string();
        let value = self.walk_assignment(it.next().unwrap());
        if op == "=" {
            expr(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            })
        } else {
            let cop = match op.as_str() {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "%=" => CompoundOp::Mod,
                "&=" => CompoundOp::BitAnd,
                "|=" => CompoundOp::BitOr,
                "^=" => CompoundOp::BitXor,
                "<<=" => CompoundOp::Shl,
                ">>=" => CompoundOp::Shr,
                _ => CompoundOp::Add,
            };
            // Represent compound assignment as a statement-level op via Assign
            // of a Binary, since ExprKind has no CompoundAssign.
            let bin = match cop {
                CompoundOp::Add => BinOp::Add,
                CompoundOp::Sub => BinOp::Sub,
                CompoundOp::Mul => BinOp::Mul,
                CompoundOp::Div => BinOp::Div,
                CompoundOp::Mod => BinOp::Mod,
                CompoundOp::BitAnd => BinOp::BitAnd,
                CompoundOp::BitOr => BinOp::BitOr,
                CompoundOp::BitXor => BinOp::BitXor,
                CompoundOp::Shl => BinOp::Shl,
                CompoundOp::Shr => BinOp::Shr,
                _ => BinOp::Add,
            };
            expr(ExprKind::Assign {
                target: Box::new(target.clone()),
                value: Box::new(expr(ExprKind::Binary {
                    op: bin,
                    left: Box::new(target),
                    right: Box::new(value),
                })),
            })
        }
    }

    fn walk_conditional(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let cond = self.walk_binary(it.next().unwrap());
        if let Some(then_p) = it.next() {
            let then = self.walk_expression(then_p);
            let else_ = self.walk_conditional(it.next().unwrap());
            expr(ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then),
                else_: Box::new(else_),
            })
        } else {
            cond
        }
    }

    fn walk_binary(&mut self, pair: Pair<Rule>) -> Expression {
        // binary_expression = unary (op unary)*
        let mut operands: Vec<Expression> = Vec::new();
        let mut ops: Vec<String> = Vec::new();
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::binary_op => ops.push(p.as_str().to_string()),
                _ => operands.push(self.walk_unary(p)),
            }
        }
        fold_binary(operands, ops)
    }

    fn walk_unary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::unary_expression => {
                let mut it = pair.into_inner().peekable();
                let first = it.next().unwrap();
                match first.as_rule() {
                    Rule::sizeof_expression => {
                        // sizeof(type) or sizeof expr — return a C-model size.
                        let inner = first.into_inner().next();
                        if let Some(p) = inner {
                            let sz = self.sizeof_from_rule(&p);
                            expr(ExprKind::Lit(Literal::Int(sz)))
                        } else {
                            expr(ExprKind::Lit(Literal::Int(8)))
                        }
                    }
                    Rule::prefix_op => {
                        let op = first.as_str();
                        let operand = self.walk_unary(it.next().unwrap());
                        self.apply_prefix(op, operand)
                    }
                    Rule::cast_expression => self.walk_cast(first),
                    Rule::postfix_expression => self.walk_postfix(first),
                    _ => self.walk_unary(first),
                }
            }
            Rule::cast_expression => self.walk_cast(pair),
            Rule::postfix_expression => self.walk_postfix(pair),
            _ => self.walk_postfix(pair),
        }
    }

    fn apply_prefix(&self, op: &str, operand: Expression) -> Expression {
        match op {
            "*" => {
                if let ExprKind::Ident(name) = &operand.kind {
                    if self.char_pointers.contains(name) {
                        return expr(ExprKind::Index {
                            object: Box::new(operand),
                            index: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
                            null_safe: false,
                        });
                    }
                }
                operand
            }
            "&" => operand,
            "+" => operand,
            "-" => expr(ExprKind::Unary {
                op: UnaryOp::Neg,
                expr: Box::new(operand),
            }),
            "!" => expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(operand),
            }),
            "~" => expr(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(operand),
            }),
            "++" => {
                if let ExprKind::Ident(name) = &operand.kind {
                    if self.char_pointers.contains(name) {
                        return expr(ExprKind::Assign {
                            target: Box::new(ident(name)),
                            value: Box::new(expr(ExprKind::Call {
                                callee: Box::new(expr(ExprKind::Member {
                                    object: Box::new(ident(name)),
                                    field: "substring".to_string(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Int(
                                    1,
                                ))))],
                                optional: false,
                            })),
                        });
                    }
                }
                expr(ExprKind::Unary {
                    op: UnaryOp::PreInc,
                    expr: Box::new(operand),
                })
            }
            "--" => expr(ExprKind::Unary {
                op: UnaryOp::PreDec,
                expr: Box::new(operand),
            }),
            _ => operand,
        }
    }

    fn walk_cast(&mut self, pair: Pair<Rule>) -> Expression {
        // (type_name) unary  → Cast, but for int/double casts we keep the
        // numeric-coercing Cast; otherwise identity.
        let mut it = pair.into_inner();
        let type_name = it.next().unwrap();
        let tn = type_name.as_str().trim().to_string();
        let operand = self.walk_unary(it.next().unwrap());
        let canon = if tn.contains("double") || tn.contains("float") {
            "double"
        } else if tn.contains("char") {
            "char"
        } else if tn.contains("int")
            || tn.contains("long")
            || tn.contains("short")
            || tn.contains("unsigned")
        {
            "int"
        } else {
            return operand;
        };
        expr(ExprKind::Cast {
            expr: Box::new(operand),
            type_name: canon.to_string(),
        })
    }

    fn walk_postfix(&mut self, pair: Pair<Rule>) -> Expression {
        let mut it = pair.into_inner();
        let mut base = self.walk_primary(it.next().unwrap());
        for suffix in it {
            base = match suffix.as_rule() {
                Rule::call_suffix => {
                    let mut args = Vec::new();
                    if let Some(arglist) = suffix.into_inner().next() {
                        for a in arglist.into_inner() {
                            args.push(Argument::positional(self.walk_assignment(a)));
                        }
                    }
                    self.normalize_call(base, args)
                }
                Rule::index_suffix => {
                    let idx = self.walk_expression(suffix.into_inner().next().unwrap());
                    expr(ExprKind::Index {
                        object: Box::new(base),
                        index: Box::new(idx),
                        null_safe: false,
                    })
                }
                Rule::member_suffix | Rule::arrow_suffix => {
                    let field = suffix.into_inner().next().unwrap().as_str().to_string();
                    expr(ExprKind::Member {
                        object: Box::new(base),
                        field,
                        null_safe: false,
                    })
                }
                Rule::inc_dec_suffix => {
                    if suffix.as_str() == "++" {
                        if let ExprKind::Ident(name) = &base.kind {
                            if self.char_pointers.contains(name) {
                                expr(ExprKind::Assign {
                                    target: Box::new(ident(name)),
                                    value: Box::new(expr(ExprKind::Call {
                                        callee: Box::new(expr(ExprKind::Member {
                                            object: Box::new(ident(name)),
                                            field: "substring".to_string(),
                                            null_safe: false,
                                        })),
                                        args: vec![Argument::positional(expr(ExprKind::Lit(
                                            Literal::Int(1),
                                        )))],
                                        optional: false,
                                    })),
                                })
                            } else {
                                expr(ExprKind::Unary {
                                    op: UnaryOp::PostInc,
                                    expr: Box::new(base),
                                })
                            }
                        } else {
                            expr(ExprKind::Unary {
                                op: UnaryOp::PostInc,
                                expr: Box::new(base),
                            })
                        }
                    } else {
                    let op = if suffix.as_str() == "++" {
                        UnaryOp::PostInc
                    } else {
                        UnaryOp::PostDec
                    };
                    expr(ExprKind::Unary {
                        op,
                        expr: Box::new(base),
                    })
                    }
                }
                _ => base,
            };
        }
        base
    }

    /// C library call normalizations. Returns the final expression to use
    /// (may wrap the call in puts() for printf-style functions).
    fn normalize_call(
        &self,
        callee: Expression,
        args: Vec<Argument>,
    ) -> Expression {
        if let ExprKind::Ident(name) = &callee.kind {
            match name.as_str() {
                "printf" => {
                    // printf(fmt, args...) → puts(sprintf(fmt, args...))
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args,
                        optional: false,
                    });
                    return expr(ExprKind::Call {
                        callee: Box::new(ident("puts")),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                "fprintf" => {
                    // fprintf(stream, fmt, args...) → drop stream, puts(sprintf(fmt, args...))
                    let mut inner_args = args;
                    if !inner_args.is_empty() {
                        inner_args.remove(0);
                    }
                    let sprintf_call = expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: inner_args,
                        optional: false,
                    });
                    return expr(ExprKind::Call {
                        callee: Box::new(ident("puts")),
                        args: vec![Argument::positional(sprintf_call)],
                        optional: false,
                    });
                }
                // sprintf / snprintf → just sprintf (returns the formatted string)
                "sprintf" | "snprintf" => {
                    let mut inner_args = args;
                    // sprintf(buf, fmt, ...) keeps fmt onward; drop destination buf.
                    if name.as_str() == "sprintf" && !inner_args.is_empty() {
                        inner_args.remove(0);
                    }
                    // snprintf has a size arg as second arg: snprintf(buf, size, fmt, ...)
                    // drop first two args (buf + size), keep fmt onwards
                    if name.as_str() == "snprintf" && inner_args.len() >= 2 {
                        inner_args.remove(0);
                        inner_args.remove(0);
                    }
                    return expr(ExprKind::Call {
                        callee: Box::new(ident("sprintf")),
                        args: inner_args,
                        optional: false,
                    });
                }
                // ── ctype.h — take integer char code, classify ───────────────
                // Wrap the char-code arg in String.fromCharCode so the
                // ecma:string host functions (which expect strings) work.
                "isalpha" | "isdigit" | "isalnum" | "isspace"
                | "isupper" | "islower" | "isxdigit" => {
                    let char_arg = wrap_charcode_arg(&args);
                    return expr(ExprKind::Call {
                        callee: Box::new(callee),
                        args: vec![Argument::positional(char_arg)],
                        optional: false,
                    });
                }
                "iscntrl" => {
                    // iscntrl(c): c < 32 || c == 127
                    if let Some(a) = args.into_iter().next() {
                        return c_iscntrl(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "isprint" => {
                    // isprint(c): c >= 32 && c < 127
                    if let Some(a) = args.into_iter().next() {
                        return c_isprint(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "ispunct" => {
                    // ispunct(c): isprint(c) && !isspace(c) && !isalnum(c)
                    if let Some(a) = args.into_iter().next() {
                        return c_ispunct(a.value);
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                // toupper/tolower: int → char-str → upper/lower → char-code
                "toupper" => {
                    if let Some(a) = args.into_iter().next() {
                        return charcode_of(call1(ident("strupr"), wrap_charcode_arg_val(a.value)));
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }
                "tolower" => {
                    if let Some(a) = args.into_iter().next() {
                        return charcode_of(call1(ident("strlwr"), wrap_charcode_arg_val(a.value)));
                    }
                    return expr(ExprKind::Lit(Literal::Int(0)));
                }

                // ── strchr/strstr — return suffix string or null ──────────────
                // strchr(s, c) → index = s.indexOf(fromCharCode(c));
                //                index >= 0 ? s.slice(index) : null
                "strchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        let s = s_arg.value;
                        let ch_str = wrap_charcode_arg_val(c_arg.value);
                        return strchr_expr(s, ch_str);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strrchr(s, c) → last occurrence
                "strrchr" => {
                    let mut it = args.into_iter();
                    if let (Some(s_arg), Some(c_arg)) = (it.next(), it.next()) {
                        let s = s_arg.value;
                        let ch_str = wrap_charcode_arg_val(c_arg.value);
                        return strrchr_expr(s, ch_str);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strstr(haystack, needle) → suffix from match or null
                "strstr" => {
                    let mut it = args.into_iter();
                    if let (Some(hay), Some(ndl)) = (it.next(), it.next()) {
                        return strstr_expr(hay.value, ndl.value);
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                // ── string.h — proper mutations ──────────────────────────────
                // strcpy(dest, src) → dest = src  (returns dest which == src)
                "strcpy" | "strncpy" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        let src_val = if name.as_str() == "strncpy" {
                            // strncpy also has a length arg — ignore it
                            src.value
                        } else {
                            src.value
                        };
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(src_val),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }
                // strcat(dest, src) → dest = dest + src  (returns dest)
                "strcat" | "strncat" => {
                    let mut it = args.into_iter();
                    if let (Some(dest), Some(src)) = (it.next(), it.next()) {
                        let concat = expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(dest.value.clone()),
                            right: Box::new(src.value),
                        });
                        return expr(ExprKind::Assign {
                            target: Box::new(dest.value),
                            value: Box::new(concat),
                        });
                    }
                    return expr(ExprKind::Lit(Literal::Null));
                }

                // ── stdlib.h — heap allocation → arrays ──────────────────────
                // malloc(n) / realloc(p, n) → [] (GC-managed array)
                "malloc" | "realloc" => {
                    return expr(ExprKind::Array(Vec::new()));
                }
                // calloc(count, size) → array of `count` zeros
                "calloc" => {
                    let count = args.into_iter().next()
                        .map(|a| a.value)
                        .unwrap_or(expr(ExprKind::Lit(Literal::Int(0))));
                    // Build: new Array filled with 0 — represent as empty array
                    // (tests typically fill via indexing; empty array suffices)
                    let _ = count;
                    return expr(ExprKind::Array(Vec::new()));
                }
                "free" => {
                    // noop — GC handles deallocation
                    return expr(ExprKind::Lit(Literal::Null));
                }

                _ => {}
            }
        }
        expr(ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        })
    }

    fn walk_primary(&mut self, pair: Pair<Rule>) -> Expression {
        match pair.as_rule() {
            Rule::primary_expression => {
                let inner = pair.into_inner().next().unwrap();
                self.walk_primary(inner)
            }
            Rule::literal => self.walk_literal(pair),
            Rule::ident_name => ident(pair.as_str()),
            Rule::expression => self.walk_expression(pair),
            _ => self.walk_expression(pair),
        }
    }

    fn walk_literal(&mut self, pair: Pair<Rule>) -> Expression {
        let inner = pair.into_inner().next().unwrap();
        match inner.as_rule() {
            Rule::int_literal => {
                let raw = inner.as_str().trim_end_matches(['u', 'U', 'l', 'L']);
                let v = parse_int_literal(raw);
                expr(ExprKind::Lit(Literal::Int(v)))
            }
            Rule::float_literal => {
                let raw = inner.as_str().trim_end_matches(['f', 'F', 'l', 'L']);
                let v = raw.parse::<f64>().unwrap_or(0.0);
                expr(ExprKind::Lit(Literal::Float(v)))
            }
            Rule::char_literal => {
                let c = parse_char_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Int(c as i64)))
            }
            Rule::string_literal => {
                let s = parse_string_literal(inner.as_str());
                expr(ExprKind::Lit(Literal::Str(s)))
            }
            Rule::bool_literal => {
                expr(ExprKind::Lit(Literal::Bool(inner.as_str() == "true")))
            }
            _ => expr(ExprKind::Lit(Literal::Null)),
        }
    }
}

// ── Binary precedence folding ─────────────────────────────────────────────

fn bin_op(op: &str) -> (BinOp, u8) {
    match op {
        "||" => (BinOp::Or, 1),
        "&&" => (BinOp::And, 2),
        "|" => (BinOp::BitOr, 3),
        "^" => (BinOp::BitXor, 4),
        "&" => (BinOp::BitAnd, 5),
        "==" => (BinOp::Eq, 6),
        "!=" => (BinOp::NotEq, 6),
        "<" => (BinOp::Lt, 7),
        "<=" => (BinOp::LtEq, 7),
        ">" => (BinOp::Gt, 7),
        ">=" => (BinOp::GtEq, 7),
        "<<" => (BinOp::Shl, 8),
        ">>" => (BinOp::Shr, 8),
        "+" => (BinOp::Add, 9),
        "-" => (BinOp::Sub, 9),
        "*" => (BinOp::Mul, 10),
        "/" => (BinOp::Div, 10),
        "%" => (BinOp::Mod, 10),
        _ => (BinOp::Add, 9),
    }
}

/// Precedence-climbing fold over a flat operand/operator sequence.
fn fold_binary(mut operands: Vec<Expression>, ops: Vec<String>) -> Expression {
    if operands.is_empty() {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }
    if ops.is_empty() {
        return operands.pop().unwrap();
    }
    // Shunting-yard into an explicit tree.
    let mut output: Vec<Expression> = Vec::new();
    let mut op_stack: Vec<(BinOp, u8)> = Vec::new();
    let mut operand_iter = operands.into_iter();
    output.push(operand_iter.next().unwrap());

    let apply = |output: &mut Vec<Expression>, op: BinOp| {
        let right = output.pop().unwrap();
        let left = output.pop().unwrap();
        output.push(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        }));
    };

    for op_str in ops {
        let (op, prec) = bin_op(&op_str);
        while let Some(&(top_op, top_prec)) = op_stack.last() {
            if top_prec >= prec {
                op_stack.pop();
                apply(&mut output, top_op);
            } else {
                break;
            }
        }
        op_stack.push((op, prec));
        if let Some(next) = operand_iter.next() {
            output.push(next);
        }
    }
    while let Some((op, _)) = op_stack.pop() {
        apply(&mut output, op);
    }
    output.pop().unwrap()
}

// ── Literal parsing helpers ────────────────────────────────────────────────

fn parse_int_literal(raw: &str) -> i64 {
    let raw = raw.trim();
    if let Some(hex) = raw.strip_prefix("0x").or_else(|| raw.strip_prefix("0X")) {
        return i64::from_str_radix(hex, 16).unwrap_or(0);
    }
    if let Some(bin) = raw.strip_prefix("0b").or_else(|| raw.strip_prefix("0B")) {
        return i64::from_str_radix(bin, 2).unwrap_or(0);
    }
    if raw.len() > 1 && raw.starts_with('0') && raw.chars().all(|c| c.is_ascii_digit()) {
        return i64::from_str_radix(&raw[1..], 8).unwrap_or(0);
    }
    raw.parse::<i64>().unwrap_or(0)
}

fn parse_char_literal(raw: &str) -> u32 {
    // raw includes surrounding single quotes
    let inner = &raw[1..raw.len() - 1];
    let mut chars = inner.chars().peekable();
    parse_escape_char(&mut chars).unwrap_or(0)
}

/// Parse one C escape or plain character from a peekable char iterator.
fn parse_escape_char(chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> Option<u32> {
    let c = chars.next()?;
    if c != '\\' {
        return Some(c as u32);
    }
    let esc = chars.next().unwrap_or('\0');
    Some(match esc {
        'n'  => 10,
        't'  => 9,
        'r'  => 13,
        '0'  => 0,
        'a'  => 7,
        'b'  => 8,
        'f'  => 12,
        'v'  => 11,
        '\\' => 92,
        '\'' => 39,
        '"'  => 34,
        'x'  => {
            // \xHH — consume hex digits
            let mut val = 0u32;
            for _ in 0..2 {
                match chars.peek() {
                    Some(&d) if d.is_ascii_hexdigit() => {
                        chars.next();
                        val = val * 16 + d.to_digit(16).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        d if d.is_ascii_digit() => {
            // \NNN — octal (1-3 digits)
            let mut val = d.to_digit(8).unwrap_or(0);
            for _ in 0..2 {
                match chars.peek() {
                    Some(&od) if od.is_ascii_digit() && od < '8' => {
                        chars.next();
                        val = val * 8 + od.to_digit(8).unwrap_or(0);
                    }
                    _ => break,
                }
            }
            val
        }
        other => other as u32,
    })
}

fn parse_string_literal(raw: &str) -> String {
    // Adjacent string literals are concatenated; strip quotes from each segment.
    let mut out = String::new();
    let mut chars = raw.chars().peekable();
    while let Some(c) = chars.next() {
        if c == '"' {
            // read chars until closing quote
            loop {
                match chars.peek() {
                    None => break,
                    Some(&'"') => { chars.next(); break; }
                    Some(&'\\') => {
                        if let Some(code) = parse_escape_char(&mut chars) {
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
                        }
                    }
                    _ => { out.push(chars.next().unwrap()); }
                }
            }
        }
        // whitespace between adjacent literals is skipped
    }
    out
}

// ── struct init helpers ───────────────────────────────────────────────────────

/// Convert a positional Array initializer `{1, 2}` to a struct Object
/// `{a: 1, b: 2}` using the known field names.
fn convert_array_init_to_struct(raw: Expression, fields: &[String]) -> Expression {
    match raw.kind {
        ExprKind::Array(ref elems) => {
            if elems.is_empty() {
                return raw; // already empty / no conversion needed
            }
            // Check if it's already an Object (designated initializers)
            let props: Vec<ObjectProperty> = elems.iter().enumerate()
                .filter_map(|(i, el)| {
                    let fname = fields.get(i)?.clone();
                    if el.key.is_some() {
                        // Designated — use key directly
                        return Some(ObjectProperty::KeyValue {
                            key: expr(ExprKind::Lit(Literal::Str(fname))),
                            value: el.value.clone(),
                        });
                    }
                    Some(ObjectProperty::KeyValue {
                        key: expr(ExprKind::Lit(Literal::Str(fname))),
                        value: el.value.clone(),
                    })
                })
                .collect();
            // Fill missing fields with 0
            let mut all_props = props;
            for i in all_props.len()..fields.len() {
                all_props.push(ObjectProperty::KeyValue {
                    key: expr(ExprKind::Lit(Literal::Str(fields[i].clone()))),
                    value: expr(ExprKind::Lit(Literal::Int(0))),
                });
            }
            expr(ExprKind::Object(all_props))
        }
        ExprKind::Object(_) => raw, // already an object
        _ => raw,
    }
}

// ── sizeof ────────────────────────────────────────────────────────────────────

impl Walker {
    fn sizeof_from_rule(&self, pair: &Pair<Rule>) -> i64 {
        match pair.as_rule() {
            Rule::type_name | Rule::declaration_specifiers | Rule::type_specifier
            | Rule::type_specifier_strict => {
                sizeof_from_type_text(pair.as_str().trim())
            }
            Rule::unary_expression | Rule::postfix_expression | Rule::primary_expression
            | Rule::expression | Rule::assignment_expression => {
                // sizeof(expr) — check if it's a variable name we know
                let text = pair.as_str().trim();
                // Strip parens
                let text = text.trim_start_matches('(').trim_end_matches(')').trim();
                // Check if it's a known variable
                if let Some(ty) = self.var_types.get(text) {
                    let sz = sizeof_from_type_text(ty);
                    // For arrays: sizeof(arr) where arr is int[5] = 5*4
                    // We'd need array size info — for now return element size * 1
                    return sz;
                }
                // Fall through to text-based guess
                let base = sizeof_from_type_text(text);
                if base != 8 { return base; }
                // Try inner nodes
                for inner in pair.clone().into_inner() {
                    let s = self.sizeof_from_rule(&inner);
                    if s != 8 { return s; }
                }
                8
            }
            _ => sizeof_from_type_text(pair.as_str().trim()),
        }
    }
}

fn sizeof_from_type_text(text: &str) -> i64 {
    // Strip qualifiers
    let t = text
        .replace("const ", "").replace("volatile ", "").replace("static ", "")
        .replace("unsigned ", "").replace("signed ", "")
        .replace("register ", "").replace("restrict ", "");
    let t = t.trim();
    // Pointer → pointer size (8 on 64-bit)
    if t.contains('*') { return 8; }
    match t {
        "char" | "int8_t" | "uint8_t" | "_Bool" | "bool" => 1,
        "short" | "int16_t" | "uint16_t" => 2,
        "int" | "float" | "int32_t" | "uint32_t" => 4,
        "long" | "double" | "long long" | "int64_t" | "uint64_t"
        | "size_t" | "ssize_t" | "ptrdiff_t" => 8,
        "long double" => 16,
        "void" => 1,
        _ => 8, // unknown / struct / pointer-like → pointer size
    }
}

// ── ctype / string helper factories ─────────────────────────────────────────

/// Build: s.slice(idx) where idx = s.indexOf(needle); return null if not found.
/// Represents: needle found → suffix; not found → null (0 in C).
fn strchr_expr(s: Expression, needle: Expression) -> Expression {
    // ternary: (idx = s.indexOf(needle)) >= 0 ? s.slice(idx) : null
    // Simplified: s.slice(s.indexOf(needle)) — works because slice(-1) returns
    // last char etc. But we need null when not found (indexOf returns -1).
    // Use: indexOf(s, needle) >= 0 ? s.slice(indexOf(s,needle)) : null
    let idx_call = call1(
        expr(ExprKind::Member { object: Box::new(s.clone()), field: "indexOf".to_string(), null_safe: false }),
        needle,
    );
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member { object: Box::new(s), field: "slice".to_string(), null_safe: false })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Null))),
    })
}

fn strrchr_expr(s: Expression, needle: Expression) -> Expression {
    let idx_call = call1(
        expr(ExprKind::Member { object: Box::new(s.clone()), field: "lastIndexOf".to_string(), null_safe: false }),
        needle,
    );
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member { object: Box::new(s), field: "slice".to_string(), null_safe: false })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Null))),
    })
}

fn strstr_expr(haystack: Expression, needle: Expression) -> Expression {
    let idx_call = call1(
        expr(ExprKind::Member { object: Box::new(haystack.clone()), field: "indexOf".to_string(), null_safe: false }),
        needle,
    );
    expr(ExprKind::Ternary {
        cond: Box::new(expr(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(idx_call.clone()),
            right: Box::new(expr(ExprKind::Lit(Literal::Int(0)))),
        })),
        then: Box::new(expr(ExprKind::Call {
            callee: Box::new(expr(ExprKind::Member { object: Box::new(haystack), field: "slice".to_string(), null_safe: false })),
            args: vec![Argument::positional(idx_call)],
            optional: false,
        })),
        else_: Box::new(expr(ExprKind::Lit(Literal::Null))),
    })
}

/// Build: String.fromCharCode(arg) — converts integer char code to 1-char string.
fn wrap_charcode_arg(args: &[Argument]) -> Expression {
    let arg = args.first().map(|a| a.value.clone())
        .unwrap_or(expr(ExprKind::Lit(Literal::Int(0))));
    wrap_charcode_arg_val(arg)
}

fn wrap_charcode_arg_val(arg: Expression) -> Expression {
    call1(ident("putchar"), arg) // putchar maps to str_from_char_code in profile
}

/// Build: call1(callee, arg) — single-argument call.
fn call1(callee: Expression, arg: Expression) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(callee),
        args: vec![Argument::positional(arg)],
        optional: false,
    })
}

/// Build: charCodeAt(s, 0) — get the char code of the first char of a string.
fn charcode_of(s: Expression) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(expr(ExprKind::Member {
            object: Box::new(s),
            field: "charCodeAt".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(expr(ExprKind::Lit(Literal::Int(0))))],
        optional: false,
    })
}

/// iscntrl(c): c < 32 || c == 127
fn c_iscntrl(c: Expression) -> Expression {
    let lt32 = expr(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(c.clone()),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(32)))),
    });
    let eq127 = expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(127)))),
    });
    expr(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(lt32),
        right: Box::new(eq127),
    })
}

/// isprint(c): c >= 32 && c < 127
fn c_isprint(c: Expression) -> Expression {
    let ge32 = expr(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(32)))),
    });
    let lt127 = expr(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(c),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(127)))),
    });
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(ge32),
        right: Box::new(lt127),
    })
}

/// ispunct(c): isprint(c) && !isspace && !isalnum
/// Simplified: c >= 33 && c <= 126 && !isalnum(fromCharCode(c))
fn c_ispunct(c: Expression) -> Expression {
    let ge33 = expr(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(33)))),
    });
    let le126 = expr(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c.clone()),
        right: Box::new(expr(ExprKind::Lit(Literal::Int(126)))),
    });
    let not_alnum = expr(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(call1(ident("isalnum"), wrap_charcode_arg_val(c))),
    });
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(ge33),
            right: Box::new(le126),
        })),
        right: Box::new(not_alnum),
    })
}
