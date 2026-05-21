//! Fortran walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.

use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use crate::ast::*;
use super::{FortranParser, Rule};

const FORTRAN_TBP_IMPL_HANDLE_PREFIX: &str = "__fortran_tbp_impl:";
const FORTRAN_IO_BUFFER_GLOBAL: &str = "__vybe_fortran_io_buffer";
const FORTRAN_ARRAY_RESULT_PARAM: &str = "__fortran_array_result";
const FORTRAN_ARRAY_INDEXING: ArrayIndexSemantics = ArrayIndexSemantics::ONE_BASED;

/// Convert legacy Hollerith literals like `4HTEST` to standard string literals
/// `"TEST"`. Pest can't match a runtime-determined character count, so we
/// preprocess the source. We respect string literals (don't rewrite inside
/// `'...'` or `"..."`) and comments (anything after `!` to end of line).
fn rewrite_hollerith(src: &str) -> String {
    let mut out = String::with_capacity(src.len());
    let bytes = src.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];
        // Skip line comments.
        if b == b'!' {
            while i < bytes.len() && bytes[i] != b'\n' {
                out.push(bytes[i] as char);
                i += 1;
            }
            continue;
        }
        // Skip string literals.
        if b == b'\'' || b == b'"' {
            let quote = b;
            out.push(b as char);
            i += 1;
            while i < bytes.len() {
                let c = bytes[i];
                out.push(c as char);
                i += 1;
                if c == quote {
                    // Fortran doubled-quote escape.
                    if i < bytes.len() && bytes[i] == quote {
                        out.push(quote as char);
                        i += 1;
                        continue;
                    }
                    break;
                }
            }
            continue;
        }
        // Try Hollerith: must be at start of token (preceded by non-alnum).
        if b.is_ascii_digit() {
            let prev_is_word = i > 0 && {
                let p = bytes[i - 1];
                p.is_ascii_alphanumeric() || p == b'_' || p == b'.'
            };
            if !prev_is_word {
                let mut j = i;
                while j < bytes.len() && bytes[j].is_ascii_digit() { j += 1; }
                if j < bytes.len() && (bytes[j] == b'H' || bytes[j] == b'h') {
                    if let Ok(count) = std::str::from_utf8(&bytes[i..j]).unwrap().parse::<usize>() {
                        let after_h = j + 1;
                        let end = (after_h + count).min(bytes.len());
                        if end - after_h == count {
                            let text = &src[after_h..end];
                            // Emit as quoted string with escaped quotes.
                            out.push('"');
                            for ch in text.chars() {
                                if ch == '"' { out.push('\\'); }
                                out.push(ch);
                            }
                            out.push('"');
                            i = end;
                            continue;
                        }
                    }
                }
            }
        }
        out.push(b as char);
        i += 1;
    }
    out
}

pub fn parse(source: &str) -> Result<Module, String> {
    let preprocessed = rewrite_hollerith(source);
    let source = preprocessed.as_str();
    let mut pairs = FortranParser::parse(Rule::program, source)
        .map_err(|e| format!("Fortran parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut name = String::new();

    for pair in program.into_inner() {
        match pair.as_rule() {
            Rule::EOI | Rule::NEWLINE => {}
            Rule::statement_line => {
                for inner in pair.into_inner().filter(|p| meaningful(p)) {
                    walk_top(inner, &mut name, &mut body, &mut imports)?;
                }
            }
            _ => {}
        }
    }

    bind_top_level_type_bound_procedures(&mut body);
    lower_fortran_body_intrinsics(&[], &mut body);
    lower_fortran_array_semantics(&[], &mut body);
    repair_remaining_fortran_array_calls(&mut body);
    lower_fortran_array_return_calls(&mut body);
    lower_fortran_array_assignments(&mut body);
    lower_fortran_array_call_arguments(&mut body);
    lower_fortran_array_return_calls(&mut body);
    lower_fortran_array_expressions(&mut body);
    lower_fortran_scalar_array_assignments(&mut body);
    body.insert(0, Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(FORTRAN_IO_BUFFER_GLOBAL)],
        value: Expression::string(""),
    }));

    Ok(Module { name, language: Lang::Fortran, body, imports })
}

fn walk_top(pair: Pair<Rule>, name: &mut String, body: &mut Vec<Statement>, imports: &mut Vec<Import>) -> Result<(), String> {
    match pair.as_rule() {
        Rule::program_unit => {
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => { if name.is_empty() { *name = p.as_str().to_string(); } }
                    Rule::statement_line => {
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            walk_top(s, name, body, imports)?;
                        }
                    }
                    _ => { if let Some(st) = walk_stmt(p)? { body.push(st); } }
                }
            }
        }
        Rule::module_unit => {
            let mut mname = String::new();
            let mut members = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                match p.as_rule() {
                    Rule::identifier => { if mname.is_empty() { mname = p.as_str().to_string(); } }
                    Rule::statement_line => {
                        for s in p.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? {
                                members.push(to_class_member(st));
                            }
                        }
                    }
                    _ => { if let Some(st) = walk_stmt(p)? { members.push(to_class_member(st)); } }
                }
            }
            bind_module_type_bound_procedures(&mut members);
            body.push(Statement::new(StmtKind::ModuleDecl { name: mname, members, visibility: Visibility::Public }));
        }
        Rule::use_statement => {
            let mut parts = pair.into_inner().filter(|p| meaningful(p));
            let mname = parts.next().map(|p| p.as_str().to_string()).unwrap_or_default();
            let mut names = Vec::new();
            for p in parts {
                if p.as_rule() == Rule::use_name_list {
                    for np in p.into_inner() {
                        if np.as_rule() == Rule::use_name {
                            let mut ni = np.into_inner();
                            let n = ni.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                            let a = ni.next().map(|p| p.as_str().to_string());
                            names.push(ImportName { name: n, alias: a });
                        }
                    }
                }
            }
            if names.is_empty() {
                imports.push(Import { kind: ImportKind::Simple { path: mname, alias: None }, span: Span::default() });
            } else {
                imports.push(Import { kind: ImportKind::Named { path: mname, names, level: 0 }, span: Span::default() });
            }
        }
        _ => { if let Some(st) = walk_stmt(pair)? { body.push(st); } }
    }
    Ok(())
}

fn walk_stmt(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::implicit_statement | Rule::contains_statement | Rule::NEWLINE => Ok(None),
        Rule::var_declaration => walk_var_decl(pair).map(Some),
        Rule::procedure_decl => walk_procedure_decl(pair).map(Some),
        Rule::assignment_statement => walk_assign(pair).map(Some),
        Rule::if_statement => walk_if(pair).map(Some),
        Rule::do_statement => walk_do(pair).map(Some),
        Rule::do_while_statement => walk_do_while(pair).map(Some),
        Rule::select_case_statement => walk_select(pair).map(Some),
        Rule::print_statement => walk_print(pair).map(Some),
        Rule::write_statement => walk_write(pair).map(Some),
        Rule::read_statement => Ok(Some(Statement::new(StmtKind::Empty))),
        Rule::call_statement => walk_call(pair).map(Some),
        Rule::subroutine_decl => walk_sub(pair).map(Some),
        Rule::function_decl => walk_func(pair).map(Some),
        Rule::type_decl => walk_type(pair).map(Some),
        Rule::interface_decl | Rule::abstract_interface_decl => walk_interface_decl(pair),
        Rule::allocate_statement => walk_allocate_stmt(pair).map(Some),
        Rule::deallocate_statement => walk_deallocate_stmt(pair).map(Some),
        Rule::return_statement => {
            let e = pair.into_inner().filter(|p| meaningful(p)).next().map(walk_expr).transpose()?;
            Ok(Some(Statement::new(StmtKind::Return(e))))
        }
        Rule::yield_statement => {
            let value = pair.into_inner().filter(|p| meaningful(p)).next().map(walk_expr).transpose()?;
            Ok(Some(Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Yield(value.map(Box::new)),
            )))))
        }
        Rule::cycle_statement => Ok(Some(Statement::new(StmtKind::Continue(ContinueTarget::Implicit)))),
        Rule::exit_statement => Ok(Some(Statement::new(StmtKind::Break(BreakTarget::Implicit)))),
        Rule::stop_statement => Ok(Some(Statement::new(StmtKind::Return(None)))),
        Rule::expression_statement => {
            let e = walk_expr(pair.into_inner().next().ok_or("empty expr")?)?;
            if let Some(stmt) = lower_intrinsic_statement(&e) {
                return Ok(Some(stmt));
            }
            Ok(Some(Statement::new(StmtKind::Expr(e))))
        }
        Rule::statement_line => {
            let mut stmts = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if let Some(s) = walk_stmt(p)? { stmts.push(s); }
            }
            match stmts.len() {
                0 => Ok(None),
                1 => Ok(stmts.into_iter().next()),
                _ => Ok(Some(Statement::new(StmtKind::Block(stmts)))),
            }
        }
        Rule::use_statement => Ok(None),
        Rule::program_unit | Rule::module_unit => {
            let mut body = Vec::new();
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if p.as_rule() == Rule::statement_line {
                    for s in p.into_inner().filter(|p| meaningful(p)) {
                        if let Some(st) = walk_stmt(s)? { body.push(st); }
                    }
                } else if p.as_rule() != Rule::identifier {
                    if let Some(st) = walk_stmt(p)? { body.push(st); }
                }
            }
            Ok(Some(Statement::new(StmtKind::Block(body))))
        }
        _ => Ok(None),
    }
}

fn walk_body<'a>(pairs: impl Iterator<Item = Pair<'a, Rule>>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for p in pairs {
        match p.as_rule() {
            Rule::statement_line => {
                for s in p.into_inner().filter(|p| meaningful(p)) {
                    if let Some(st) = walk_stmt(s)? { body.push(st); }
                }
            }
            Rule::identifier => {}
            _ => { if let Some(st) = walk_stmt(p)? { body.push(st); } }
        }
    }
    Ok(body)
}

fn walk_inline_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt in pair.into_inner().filter(|p| meaningful(p)) {
        if let Some(lowered) = walk_stmt(stmt)? {
            body.push(lowered);
        }
    }
    Ok(body)
}

/// Extract the declared length N from a `character(len=N)` or
/// `character*N` type hint, lowercased and whitespace-stripped.
/// Returns `None` for `character` (no length), `character(len=*)`,
/// `character(len=:)`, or non-character types.
fn parse_character_len(type_hint: &str) -> Option<i64> {
    let s: String = type_hint.chars().filter(|c| !c.is_whitespace()).collect();
    let lower = s.to_ascii_lowercase();
    if !lower.starts_with("character") {
        return None;
    }
    // `character*N` form
    if let Some(rest) = lower.strip_prefix("character*") {
        let num: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
        if !num.is_empty() {
            return num.parse().ok();
        }
    }
    // `character(len=N)` or `character(N)` form. Take what's inside the
    // outermost parens and look for a numeric length.
    let lp = lower.find('(')?;
    let rp = lower.rfind(')')?;
    if rp <= lp {
        return None;
    }
    let inside = &lower[lp + 1..rp];
    // Common forms: `len=N`, `n,kind=k`, `N`, `len=*`, `len=:`.
    let first_clause = inside.split(',').next().unwrap_or(inside);
    let val = if let Some(eq) = first_clause.split_once('=') {
        eq.1
    } else {
        first_clause
    };
    if val == "*" || val == ":" {
        return None;
    }
    val.parse().ok()
}

fn parse_derived_type_name(type_hint: &str) -> Option<String> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let prefix_len = if lower.starts_with("type(") {
        5
    } else if lower.starts_with("class(") {
        6
    } else {
        return None;
    };

    let suffix = &trimmed[prefix_len..];
    let end = suffix.find(')')?;
    let name = suffix[..end].trim();
    if name.is_empty() || name == "*" {
        return None;
    }
    Some(name.to_string())
}

fn parse_fortran_dimension_spec_list(pair: Pair<Rule>) -> Result<(Vec<Expression>, Option<Expression>), String> {
    let mut dim_bounds = Vec::new();
    let mut dim_size = None;
    for spec in pair.into_inner().filter(|p| meaningful(p)) {
        if spec.as_rule() != Rule::dimension_spec {
            continue;
        }
        let exprs: Vec<Pair<Rule>> = spec
            .clone()
            .into_inner()
            .filter(|p| meaningful(p))
            .collect();
        let this_size = if exprs.len() == 1 {
            walk_expr(exprs.into_iter().next().unwrap())?
        } else if exprs.len() == 2 {
            let lo = walk_expr(exprs[0].clone())?;
            let hi = walk_expr(exprs[1].clone())?;
            let sub = Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(hi),
                right: Box::new(lo),
            });
            Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(sub),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
            })
        } else {
            continue;
        };
        dim_bounds.push(this_size.clone());
        dim_size = Some(match dim_size.take() {
            Some(prev) => Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(prev),
                right: Box::new(this_size),
            }),
            None => this_size,
        });
    }
    Ok((dim_bounds, dim_size))
}

fn has_deferred_fortran_dimension_spec(pair: &Pair<Rule>) -> bool {
    pair.clone()
        .into_inner()
        .filter(|p| meaningful(p))
        .any(|spec| {
            spec.as_rule() == Rule::dimension_spec
                && spec.into_inner().filter(|p| meaningful(p)).next().is_none()
        })
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut inner = pair.into_inner();
    let type_hint = inner.next().map(|p| p.as_str().trim().to_string());
    let mut is_pointer = false;
    let mut is_allocatable = false;
    let mut has_intent = false;
    let mut attr_dim_bounds: Vec<Expression> = Vec::new();
    let mut has_attr_array_bounds = false;
    let mut has_attr_deferred_array_bounds = false;
    let mut trailing = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::var_attributes => {
                for attr in p.into_inner().filter(|attr| meaningful(attr)) {
                    match attr.as_rule() {
                        Rule::var_attribute => {
                            let attr_text = attr.as_str().trim().to_ascii_lowercase();
                            if attr_text == "pointer" {
                                is_pointer = true;
                            } else if attr_text == "allocatable" {
                                is_allocatable = true;
                            } else if attr_text.starts_with("intent(") {
                                has_intent = true;
                            }
                            for child in attr.into_inner().filter(|child| meaningful(child)) {
                                if child.as_rule() == Rule::dimension_spec_list {
                                    if has_deferred_fortran_dimension_spec(&child) {
                                        has_attr_deferred_array_bounds = true;
                                        has_attr_array_bounds = true;
                                    }
                                    let (bounds, _) = parse_fortran_dimension_spec_list(child)?;
                                    if !bounds.is_empty() {
                                        attr_dim_bounds = bounds;
                                        has_attr_array_bounds = true;
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => trailing.push(p),
        }
    }
    let mut declarations = Vec::new();
    for p in trailing {
        if p.as_rule() == Rule::var_declarator_list {
            for d in p.into_inner() {
                if d.as_rule() == Rule::var_declarator {
                    let mut di = d.into_inner().filter(|p| meaningful(p));
                    let nm = di.next().map(|p| p.as_str().to_string()).unwrap_or_default();
                    let mut init = None;
                    let mut dim_bounds: Vec<Expression> = attr_dim_bounds.clone();
                    let mut has_array_bounds = has_attr_array_bounds || has_attr_deferred_array_bounds;
                    for pp in di {
                        match pp.as_rule() {
                            Rule::dimension_spec_list => {
                                if has_deferred_fortran_dimension_spec(&pp) {
                                    has_array_bounds = true;
                                    dim_bounds.clear();
                                }
                                let (bounds, _) = parse_fortran_dimension_spec_list(pp)?;
                                if !bounds.is_empty() {
                                    has_array_bounds = true;
                                    dim_bounds = bounds;
                                }
                            }
                            Rule::codimension_spec_list => { /* PGAS — model as scalar/array */ }
                            _ => {
                                init = Some(walk_expr(pp)?);
                            }
                        }
                    }
                    // Fixed Fortran arrays should be initialized from their declared
                    // per-dimension bounds in the compiler so multidimensional arrays
                    // allocate nested runtime arrays instead of a single flat buffer.
                    if init.is_none() {
                        if !has_array_bounds && !is_pointer && !is_allocatable && !has_intent {
                            if let Some(class_name) = type_hint.as_deref().and_then(parse_derived_type_name) {
                                init = Some(Expression::new(ExprKind::New {
                                    class: Box::new(Expression::new(ExprKind::Ident(class_name))),
                                    args: Vec::new(),
                                }));
                            }
                        }
                    }
                    // Fortran `character(len=N) :: s = 'literal'` — pad the literal
                    // with trailing blanks so `len(s)` returns the declared length.
                    // Pure JS-shape rewrite: wrap the init in `s.padEnd(N, ' ')`.
                    if !has_array_bounds {
                        if let Some(ref t) = type_hint {
                            if let Some(declared_len) = parse_character_len(t) {
                                if let Some(ref existing) = init {
                                    let padded = Expression::new(ExprKind::Call {
                                        callee: Box::new(Expression::new(ExprKind::Member {
                                            object: Box::new(existing.clone()),
                                            field: "padEnd".into(),
                                            null_safe: false,
                                        })),
                                        args: vec![
                                            Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(declared_len)))),
                                            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(" ".into())))),
                                        ],
                                        optional: false,
                                    });
                                    init = Some(padded);
                                } else {
                                    // No init — synthesize a string of N spaces so len(s) returns N.
                                    let spaces: String = " ".repeat(declared_len as usize);
                                    init = Some(Expression::new(ExprKind::Lit(Literal::Str(spaces))));
                                }
                            }
                        }
                    }
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(nm),
                        type_hint: type_hint.clone(),
                        init,
                        array_bounds: has_array_bounds.then_some(dim_bounds),
                        with_events: false,
                    });
                }
            }
        }
    }
    Ok(Statement::new(StmtKind::VarDecl { declarations, kind: VarDeclKind::Dim }))
}

fn walk_procedure_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut interface_name: Option<String> = None;
    let mut declarations = Vec::new();

    for child in pair.into_inner().filter(|p| meaningful(p)) {
        match child.as_rule() {
            Rule::identifier if interface_name.is_none() => {
                interface_name = Some(child.as_str().to_string());
            }
            Rule::proc_decl_item => {
                let mut item_name: Option<String> = None;
                let mut init: Option<Expression> = None;
                for item_child in child.into_inner().filter(|p| meaningful(p)) {
                    match item_child.as_rule() {
                        Rule::identifier if item_name.is_none() => {
                            item_name = Some(item_child.as_str().to_string());
                        }
                        Rule::identifier => {
                            init = Some(Expression::ident(item_child.as_str()));
                        }
                        _ => {
                            let text = item_child.as_str().trim();
                            if text.eq_ignore_ascii_case("null()") {
                                init = Some(Expression::null());
                            }
                        }
                    }
                }

                if let Some(name) = item_name {
                    let type_hint = interface_name
                        .as_ref()
                        .map(|iface| format!("procedure({iface})"))
                        .or_else(|| Some("procedure".to_string()));
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint,
                        init,
                        array_bounds: None,
                        with_events: false,
                    });
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Dim,
    }))
}

fn parse_fortran_visibility(text: &str) -> Option<Visibility> {
    match text.trim().to_ascii_lowercase().as_str() {
        "public" => Some(Visibility::Public),
        "private" => Some(Visibility::Private),
        "protected" => Some(Visibility::Protected),
        _ => None,
    }
}

fn apply_fortran_type_attribute(
    pair: Pair<Rule>,
    modifiers: &mut ClassModifiers,
    parents: &mut Vec<String>,
) {
    if let Some(visibility) = parse_fortran_visibility(pair.as_str()) {
        modifiers.visibility = visibility;
        return;
    }

    if pair.as_str().trim().eq_ignore_ascii_case("abstract") {
        modifiers.is_abstract = true;
        return;
    }

    for child in pair.into_inner().filter(|child| meaningful(child)) {
        if child.as_rule() == Rule::identifier {
            parents.push(child.as_str().to_string());
        }
    }
}

fn apply_fortran_type_bound_attribute(text: &str, modifiers: &mut Modifiers) {
    if let Some(visibility) = parse_fortran_visibility(text) {
        modifiers.visibility = visibility;
        return;
    }

    let attr = text.trim();
    if attr.eq_ignore_ascii_case("deferred") {
        modifiers.is_abstract = true;
    } else if attr.eq_ignore_ascii_case("non_overridable") {
        modifiers.is_not_overridable = true;
    }
}

fn walk_assign(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let value = walk_expr(parts.pop().ok_or("missing rhs")?)?;
    let mut target = Expression::new(ExprKind::Ident(parts[0].as_str().to_string()));
    for p in &parts[1..] {
        if p.as_rule() == Rule::member_or_index {
            let mut inner = p.clone().into_inner().filter(|m| meaningful(m));
            let first = inner.next();
            match first {
                Some(m) if matches!(m.as_rule(), Rule::identifier | Rule::designator_name) => {
                    target = Expression::new(ExprKind::Member {
                        object: Box::new(target), field: m.as_str().to_string(), null_safe: false,
                    });
                }
                Some(m) if m.as_rule() == Rule::argument_list => {
                    let mut indices = Vec::new();
                    for arg in m.into_inner().filter(|item| item.as_rule() == Rule::argument) {
                        let (_, index) = walk_argument_expr(arg)?;
                        indices.push(index);
                    }
                    for index in indices {
                        target = Expression::new(ExprKind::Index {
                            object: Box::new(target),
                            index: Box::new(index),
                            null_safe: false,
                        });
                    }
                }
                Some(m) => {
                    let idx = walk_expr(m)?;
                    target = Expression::new(ExprKind::Index {
                        object: Box::new(target),
                        index: Box::new(idx),
                        null_safe: false,
                    });
                }
                None => {}
            }
        }
    }
    Ok(Statement::new(StmtKind::Assign { targets: vec![target], value }))
}

fn walk_if(pair: Pair<Rule>) -> Result<Statement, String> {
    let inner: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // Find the condition expression — first non-keyword child
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;
    for p in inner {
        match p.as_rule() {
            Rule::expression | Rule::logical_or | Rule::logical_and | Rule::logical_not
            | Rule::comparison | Rule::addition | Rule::multiplication | Rule::power
            | Rule::concat | Rule::unary | Rule::primary_expr => {
                if cond.is_none() {
                    cond = Some(walk_expr(p)?);
                }
            }
            Rule::statement_line => {
                for s in p.into_inner().filter(|p| meaningful(p)) {
                    if let Some(st) = walk_stmt(s)? { then_body.push(st); }
                }
            }
            Rule::elseif_clause => {
                let ei: Vec<Pair<Rule>> = p.into_inner().filter(|p| meaningful(p)).collect();
                let mut ec = None;
                let mut eb = Vec::new();
                for e in ei {
                    if is_expr_rule(e.as_rule()) && ec.is_none() {
                        ec = Some(walk_expr(e)?);
                    } else if e.as_rule() == Rule::statement_line {
                        for s in e.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { eb.push(st); }
                        }
                    }
                }
                if let Some(c) = ec {
                    elifs.push((c, eb));
                }
            }
            Rule::else_clause => {
                let mut eb = Vec::new();
                for e in p.into_inner().filter(|p| meaningful(p)) {
                    if e.as_rule() == Rule::statement_line {
                        for s in e.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { eb.push(st); }
                        }
                    }
                }
                else_body = Some(eb);
            }
            // Single-line if body (e.g., print_statement, assignment_statement)
            Rule::print_statement | Rule::write_statement | Rule::call_statement
            | Rule::assignment_statement | Rule::return_statement | Rule::cycle_statement
            | Rule::exit_statement | Rule::stop_statement | Rule::expression_statement => {
                if let Some(st) = walk_stmt(p)? { then_body.push(st); }
            }
            _ => {} // skip keywords (kw_if, kw_then, kw_end, etc.)
        }
    }
    Ok(Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body, elifs, else_body
    }))
}

fn walk_do(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    // Collect: identifier, expression, expression [, expression], statement_line*
    let mut var = String::new();
    let mut exprs = Vec::new();
    let mut body_parts = Vec::new();
    let mut body = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::identifier if var.is_empty() => { var = p.as_str().to_string(); }
            Rule::statement_line => { body_parts.push(p); }
            Rule::inline_statement_list => { body.extend(walk_inline_statement_list(p)?); }
            _ if is_expr_rule(p.as_rule()) => { exprs.push(p); }
            Rule::identifier => {} // end do name
            _ => {} // skip kw_do, kw_end etc.
        }
    }
    body.extend(walk_body(body_parts.into_iter())?);
    if var.is_empty() {
        return Ok(Statement::new(StmtKind::While {
            cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
            body,
            else_body: None,
        }));
    }

    let start = if !exprs.is_empty() { walk_expr(exprs.remove(0))? } else { Expression::new(ExprKind::Lit(Literal::Int(0))) };
    let end_e = if !exprs.is_empty() { walk_expr(exprs.remove(0))? } else { Expression::new(ExprKind::Lit(Literal::Int(0))) };
    let step_expr = if !exprs.is_empty() { Some(walk_expr(exprs.remove(0))?) } else { None };
    let init = Some(Box::new(Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Ident(var.clone()))],
        value: start,
    })));
    let cond = Some(Expression::new(ExprKind::Binary {
        left: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
        op: BinOp::LtEq,
        right: Box::new(end_e),
    }));
    let sv = step_expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(1))));
    // i = i + step as an Assign expression
    let update = Some(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::Ident(var.clone()))),
        value: Box::new(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::new(ExprKind::Ident(var))),
            op: BinOp::Add,
            right: Box::new(sv),
        })),
    }));
    Ok(Statement::new(StmtKind::For { init, cond, update, body }))
}

fn walk_do_while(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut cond = None;
    let mut body_parts = Vec::new();
    let mut body = Vec::new();
    for p in parts {
        if is_expr_rule(p.as_rule()) && cond.is_none() {
            cond = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::statement_line {
            body_parts.push(p);
        } else if p.as_rule() == Rule::inline_statement_list {
            body.extend(walk_inline_statement_list(p)?);
        }
        // skip kw_do, kw_while, kw_end
    }
    // If condition not found, emit "false" so the loop immediately exits (never infinite)
    let cond = cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(false))));
    body.extend(walk_body(body_parts.into_iter())?);
    Ok(Statement::new(StmtKind::While { cond, body, else_body: None }))
}

fn walk_select(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut expr = None;
    let mut case_pairs = Vec::new();
    for p in parts {
        if is_expr_rule(p.as_rule()) && expr.is_none() {
            expr = Some(walk_expr(p)?);
        } else if p.as_rule() == Rule::case_block {
            case_pairs.push(p);
        }
    }
    let expr = expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
    let mut cases = Vec::new();
    for p in case_pairs {
        {
            let mut conds = Vec::new();
            let mut cbody = Vec::new();
            let mut is_default = false;
            for c in p.into_inner().filter(|p| meaningful(p)) {
                match c.as_rule() {
                    Rule::case_value_list => {
                        for cv in c.into_inner() {
                            if cv.as_rule() == Rule::case_value {
                                if let Some(first) = cv.into_inner().next() {
                                    conds.push(CaseCondition::Value(walk_expr(first)?));
                                }
                            }
                        }
                    }
                    Rule::statement_line => {
                        for s in c.into_inner().filter(|p| meaningful(p)) {
                            if let Some(st) = walk_stmt(s)? { cbody.push(st); }
                        }
                    }
                    _ => { if c.as_str().to_lowercase().contains("default") { is_default = true; } }
                }
            }
            if is_default { conds.clear(); }
            cases.push(SwitchCase { conditions: conds, body: cbody });
        }
    }
    Ok(Statement::new(StmtKind::Switch { expr, cases, default: None }))
}

fn walk_print(pair: Pair<Rule>) -> Result<Statement, String> {
    let raw = pair.as_str().trim_start().to_ascii_lowercase();
    let mut args = Vec::new();
    let mut explicit_format = false;
    let mut advance = true;
    let mut format_spec = None;

    match pair.as_rule() {
        Rule::print_statement => {
            let list_directed = raw.strip_prefix("print")
                .map(|rest| rest.trim_start().starts_with('*'))
                .unwrap_or(false);
            let mut skipped_format = false;
            for p in pair.into_inner().filter(|p| meaningful(p)) {
                if !list_directed && !skipped_format && p.as_rule() == Rule::string_literal {
                    explicit_format = true;
                    format_spec = Some(parse_fortran_string_literal_text(p.as_str()));
                    skipped_format = true;
                    continue;
                }
                if is_expr_rule(p.as_rule()) || p.as_rule() == Rule::expression {
                    args.push(walk_expr(p)?);
                }
            }
        }
        _ => {}
    }

    let text = build_fortran_text_expr(&args, explicit_format, format_spec.as_deref());
    let callee = if advance { "__fortran_emitln" } else { "__fortran_emit" };
    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(callee)),
        args: vec![Argument::positional(text)],
        optional: false,
    }))))
}

fn walk_write(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut args = Vec::new();
    let mut explicit_format = false;
    let mut advance = true;
    let mut file_number = None;
    let mut format_spec = None;
    let mut positional_spec_index = 0usize;

    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::write_spec => parse_fortran_write_spec(
                &p,
                &mut explicit_format,
                &mut advance,
                &mut file_number,
                &mut format_spec,
                &mut positional_spec_index,
            )?,
            rule if is_expr_rule(rule) || rule == Rule::expression => args.push(walk_expr(p)?),
            _ => {}
        }
    }

    let text = build_fortran_text_expr(&args, explicit_format, format_spec.as_deref());
    if let Some(file_number) = file_number {
        Ok(Statement::new(StmtKind::PrintFile {
            file_number,
            items: vec![text],
        }))
    } else {
        let callee = if advance { "__fortran_emitln" } else { "__fortran_emit" };
        Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(callee)),
            args: vec![Argument::positional(text)],
            optional: false,
        }))))
    }
}

fn parse_fortran_string_literal_text(raw: &str) -> String {
    if raw.len() < 2 {
        return raw.to_string();
    }
    raw[1..raw.len() - 1]
        .replace("''", "'")
        .replace("\"\"", "\"")
}

fn parse_fortran_write_spec(
    spec: &Pair<Rule>,
    explicit_format: &mut bool,
    advance: &mut bool,
    file_number: &mut Option<Expression>,
    format_spec: &mut Option<String>,
    positional_spec_index: &mut usize,
) -> Result<(), String> {
    let raw = spec.as_str().trim();
    if raw.eq("*") {
        *positional_spec_index += 1;
        return Ok(());
    }

    let lowered = raw.to_ascii_lowercase();
    if lowered.starts_with("advance") {
        if lowered.contains("'no'") || lowered.contains("\"no\"") {
            *advance = false;
        }
        return Ok(());
    }

    let mut named_spec = None;
    let mut expr = None;
    for item in spec.clone().into_inner().filter(|p| meaningful(p)) {
        match item.as_rule() {
            Rule::identifier => named_spec = Some(item.as_str().to_ascii_lowercase()),
            rule if is_expr_rule(rule) || rule == Rule::expression => expr = Some(walk_expr(item)?),
            _ => {}
        }
    }

    match named_spec.as_deref() {
        Some("unit") => {
            if let Some(expr) = expr {
                *file_number = Some(expr);
            }
        }
        Some("fmt") => {
            *explicit_format = true;
            if raw.contains('"') || raw.contains('\'') {
                if let Some(eq_idx) = raw.find('=') {
                    *format_spec = Some(parse_fortran_string_literal_text(raw[eq_idx + 1..].trim()));
                }
            }
        }
        Some(_) => {}
        None => {
            if *positional_spec_index == 0 {
                *file_number = expr;
            } else if *positional_spec_index == 1 {
                *explicit_format = true;
                if raw.starts_with('"') || raw.starts_with('\'') {
                    *format_spec = Some(parse_fortran_string_literal_text(raw));
                }
            }
            *positional_spec_index += 1;
        }
    }

    Ok(())
}

#[derive(Debug, Clone)]
enum FortranFormatChunk {
    Data {
        descriptor: String,
        repeat: usize,
        width: Option<usize>,
        precision: Option<usize>,
    },
    Spaces(usize),
    Newline,
    Literal(String),
}

fn build_fortran_text_expr(args: &[Expression], explicit_format: bool, format_spec: Option<&str>) -> Expression {
    if explicit_format {
        if let Some(format_spec) = format_spec {
            if let Some(formatted) = build_fortran_formatted_io_text(args, format_spec) {
                return formatted;
            }
        }
    }
    build_fortran_io_text(args, explicit_format)
}

fn build_fortran_io_text(args: &[Expression], explicit_format: bool) -> Expression {
    let mut parts = args.iter().cloned().map(stringify_fortran_io_expr);
    let Some(mut text) = parts.next() else {
        return Expression::string("");
    };

    for part in parts {
        if explicit_format {
            text = concat_fortran_io_expr(text, part);
        } else {
            text = concat_fortran_io_expr(text, Expression::string(" "));
            text = concat_fortran_io_expr(text, part);
        }
    }
    text
}

fn build_fortran_formatted_io_text(args: &[Expression], format_spec: &str) -> Option<Expression> {
    let chunks = parse_fortran_format_chunks(format_spec)?;
    let mut printf_format = String::new();
    let mut printf_args = Vec::new();
    let mut arg_index = 0usize;

    for chunk in chunks {
        match chunk {
            FortranFormatChunk::Spaces(count) => printf_format.push_str(&" ".repeat(count)),
            FortranFormatChunk::Newline => printf_format.push('\n'),
            FortranFormatChunk::Literal(text) => printf_format.push_str(&escape_printf_literal(&text)),
            FortranFormatChunk::Data {
                descriptor,
                repeat,
                width,
                precision,
            } => {
                let spec = fortran_descriptor_to_printf(&descriptor, width, precision)?;
                let expand_single_array = repeat > 1 && args.len().saturating_sub(arg_index) == 1;
                if !expand_single_array && args.len().saturating_sub(arg_index) < repeat {
                    return None;
                }
                for offset in 0..repeat {
                    printf_format.push_str(&spec);
                    let arg_expr = if expand_single_array {
                        Expression::new(ExprKind::Index {
                            object: Box::new(args[arg_index].clone()),
                            index: Box::new(Expression::int((offset + 1) as i64)),
                            null_safe: false,
                        })
                    } else {
                        args[arg_index + offset].clone()
                    };
                    printf_args.push(Argument::positional(arg_expr));
                }
                arg_index += if expand_single_array { 1 } else { repeat };
            }
        }
    }

    if arg_index != args.len() {
        return None;
    }

    let mut call_args = Vec::with_capacity(printf_args.len() + 1);
    call_args.push(Argument::positional(Expression::string(&printf_format)));
    call_args.extend(printf_args);
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("sprintf")),
        args: call_args,
        optional: false,
    }))
}

fn parse_fortran_format_chunks(format_spec: &str) -> Option<Vec<FortranFormatChunk>> {
    let mut source = format_spec.trim();
    if source.starts_with('(') && source.ends_with(')') && source.len() >= 2 {
        source = &source[1..source.len() - 1];
    }

    let bytes = source.as_bytes();
    let mut index = 0usize;
    let mut chunks = Vec::new();

    while index < bytes.len() {
        let ch = bytes[index] as char;
        if ch.is_ascii_whitespace() || ch == ',' {
            index += 1;
            continue;
        }
        if ch == '/' {
            chunks.push(FortranFormatChunk::Newline);
            index += 1;
            continue;
        }
        if ch == '\'' || ch == '"' {
            let quote = ch;
            index += 1;
            let start = index;
            while index < bytes.len() && bytes[index] as char != quote {
                index += 1;
            }
            if index > bytes.len() {
                return None;
            }
            chunks.push(FortranFormatChunk::Literal(source[start..index].to_string()));
            if index < bytes.len() {
                index += 1;
            }
            continue;
        }

        let repeat_start = index;
        while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
            index += 1;
        }
        let repeat_prefix = if repeat_start < index {
            source[repeat_start..index].parse::<usize>().ok()
        } else {
            None
        };
        if index >= bytes.len() {
            break;
        }

        let mut descriptor = String::new();
        descriptor.push((bytes[index] as char).to_ascii_lowercase());
        index += 1;
        if descriptor == "e" && index < bytes.len() {
            let suffix = (bytes[index] as char).to_ascii_lowercase();
            if suffix == 's' {
                descriptor.push(suffix);
                index += 1;
            }
        }

        if descriptor == "x" {
            chunks.push(FortranFormatChunk::Spaces(repeat_prefix.unwrap_or(1)));
            continue;
        }

        let width_start = index;
        while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
            index += 1;
        }
        let width = if width_start < index {
            source[width_start..index].parse::<usize>().ok()
        } else {
            None
        };

        let precision = if index < bytes.len() && bytes[index] as char == '.' {
            index += 1;
            let precision_start = index;
            while index < bytes.len() && (bytes[index] as char).is_ascii_digit() {
                index += 1;
            }
            if precision_start < index {
                source[precision_start..index].parse::<usize>().ok()
            } else {
                None
            }
        } else {
            None
        };

        chunks.push(FortranFormatChunk::Data {
            descriptor,
            repeat: repeat_prefix.unwrap_or(1),
            width,
            precision,
        });
    }

    Some(chunks)
}

fn fortran_descriptor_to_printf(
    descriptor: &str,
    width: Option<usize>,
    precision: Option<usize>,
) -> Option<String> {
    let mut spec = String::from("%");
    if let Some(width) = width {
        spec.push_str(&width.to_string());
    }
    match descriptor {
        "a" => {
            spec.push('s');
            Some(spec)
        }
        "i" => {
            spec.push('d');
            Some(spec)
        }
        "f" => {
            if let Some(precision) = precision {
                spec.push('.');
                spec.push_str(&precision.to_string());
            }
            spec.push('f');
            Some(spec)
        }
        "e" | "es" | "d" | "g" => {
            if let Some(precision) = precision {
                spec.push('.');
                spec.push_str(&precision.to_string());
            }
            spec.push('e');
            Some(spec)
        }
        "l" => {
            spec.push('s');
            Some(spec)
        }
        _ => None,
    }
}

fn escape_printf_literal(text: &str) -> String {
    text.replace('%', "%%")
}

fn stringify_fortran_io_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__str__")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn concat_fortran_io_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Concat,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn walk_call(pair: Pair<Rule>) -> Result<Statement, String> {
    let inner = pair.into_inner().filter(|p| meaningful(p));
    let mut callee: Option<Expression> = None;
    let mut args = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::identifier | Rule::designator_name => {
                if p.as_str().eq_ignore_ascii_case("call") {
                    continue;
                }
                callee = Some(match callee.take() {
                    None => Expression::new(ExprKind::Ident(p.as_str().to_string())),
                    Some(expr) => Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    }),
                });
            }
            Rule::argument_list => {
                for a in p.into_inner() {
                    if a.as_rule() == Rule::argument {
                        let (name, value) = walk_argument_expr(a)?;
                        args.push(Argument { name, value, by_ref: false, spread: false });
                    }
                }
            }
            _ => {}
        }
    }
    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(callee.ok_or("missing call name")?),
        args, optional: false,
    }))))
}

fn walk_argument_expr(pair: Pair<Rule>) -> Result<(Option<String>, Expression), String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    if inner.is_empty() {
        return Err("empty arg".to_string());
    }

    if inner.len() == 1 {
        return Ok((None, walk_argument_value(inner.pop().unwrap())?));
    }

    let first = inner.remove(0);
    if first.as_rule() == Rule::identifier && inner.len() == 1 {
        return Ok((Some(first.as_str().to_string()), walk_argument_value(inner.pop().unwrap())?));
    }

    Ok((None, walk_argument_value(inner.pop().unwrap())?))
}

fn walk_argument_value(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::slice_arg => walk_slice_arg(pair),
        _ => walk_expr(pair),
    }
}

fn walk_slice_arg(pair: Pair<Rule>) -> Result<Expression, String> {
    let text = pair.as_str().trim();
    let segments: Vec<&str> = text.split(':').collect();
    let mut exprs = Vec::new();
    for part in pair.into_inner().filter(|p| meaningful(p)) {
        exprs.push(walk_expr(part)?);
    }
    let mut exprs = exprs.into_iter();

    let mut lower = None;
    let mut upper = None;
    let mut step = None;

    if segments.get(0).is_some_and(|segment| !segment.trim().is_empty()) {
        lower = exprs.next();
    }
    if segments.get(1).is_some_and(|segment| !segment.trim().is_empty()) {
        upper = exprs.next();
    }
    if segments.get(2).is_some_and(|segment| !segment.trim().is_empty()) {
        step = exprs.next();
    }

    Ok(Expression::new(ExprKind::Slice {
        lower: lower.map(Box::new),
        upper: upper.map(Box::new),
        step: step.map(Box::new),
    }))
}

fn walk_allocate_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    walk_allocator_stmt(pair, "allocate")
}

fn walk_deallocate_stmt(pair: Pair<Rule>) -> Result<Statement, String> {
    walk_allocator_stmt(pair, "deallocate")
}

fn walk_allocator_stmt(pair: Pair<Rule>, intrinsic_name: &str) -> Result<Statement, String> {
    let mut args = Vec::new();
    for item in pair.into_inner().filter(|p| meaningful(p)) {
        match item.as_rule() {
            Rule::alloc_item => args.push(Argument::positional(walk_alloc_item_expr(item)?)),
            Rule::identifier => args.push(Argument::positional(Expression::ident(item.as_str()))),
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(intrinsic_name)),
        args,
        optional: false,
    }))))
}

fn walk_alloc_item_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner().filter(|p| meaningful(p));
    let ident = inner.next().ok_or("missing allocate target")?.as_str().to_string();
    let target = Expression::ident(&ident);

    let mut dims = Vec::new();
    for child in inner {
        if child.as_rule() != Rule::dimension_spec_list {
            continue;
        }
        for dim in child.into_inner().filter(|p| meaningful(p)) {
            if let Some(expr) = walk_dimension_spec_expr(dim)? {
                dims.push(Argument::positional(expr));
            }
        }
    }

    if dims.is_empty() {
        Ok(target)
    } else {
        Ok(Expression::new(ExprKind::Call {
            callee: Box::new(target),
            args: dims,
            optional: false,
        }))
    }
}

fn walk_dimension_spec_expr(pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    match pair.as_rule() {
        Rule::dimension_spec => {
            for child in pair.into_inner().filter(|p| meaningful(p)) {
                if is_expr_rule(child.as_rule()) || child.as_rule() == Rule::expression {
                    return Ok(Some(walk_expr(child)?));
                }
            }
            Ok(None)
        }
        rule if is_expr_rule(rule) || rule == Rule::expression => walk_expr(pair).map(Some),
        _ => Ok(None),
    }
}

fn walk_sub(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        if p.as_rule() == Rule::identifier && nm.is_empty() {
            nm = p.as_str().to_string();
        } else if p.as_rule() == Rule::param_list {
            for pp in p.into_inner() {
                if pp.as_rule() == Rule::identifier {
                    params.push(Param { name: pp.as_str().to_string(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false });
                }
            }
        } else { rest.push(p); }
    }
    apply_fortran_param_declaration_modes(&mut params, &rest);
    let mut body = walk_body(rest.into_iter())?;
    bind_fortran_param_declarations(&mut params, &mut body);
    promote_mutated_fortran_params(&mut params, &body);
    lower_fortran_body_intrinsics(&params, &mut body);
    lower_fortran_array_semantics(&params, &mut body);
    let is_generator = body_has_yield(&body);
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm, params, return_type: None, body, modifiers: Modifiers::default(),
        handles: vec![], is_async: false, is_generator, is_sub: true,
    }))
}

fn walk_func(pair: Pair<Rule>) -> Result<Statement, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().filter(|p| meaningful(p)).collect();
    let mut nm = String::new();
    let mut params = Vec::new();
    let mut rt = None;
    let mut result_name = None;
    let mut rest: Vec<Pair<Rule>> = Vec::new();
    for p in parts {
        match p.as_rule() {
            Rule::type_spec => { rt = Some(p.as_str().trim().to_string()); }
            Rule::identifier => { if nm.is_empty() { nm = p.as_str().to_string(); } }
            Rule::param_list => {
                for pp in p.into_inner() {
                    if pp.as_rule() == Rule::identifier {
                        params.push(Param { name: pp.as_str().to_string(), type_hint: None, default: None, pass_by: PassBy::Value, is_rest: false, is_kwargs: false, is_optional: false, is_nullable: false });
                    }
                }
            }
            Rule::proc_suffix | Rule::result_clause => {
                let suffix = p.as_str().trim().to_ascii_lowercase();
                if suffix.starts_with("result(") {
                    result_name = p.into_inner()
                        .filter(|child| meaningful(child))
                        .find(|child| child.as_rule() == Rule::identifier)
                        .map(|child| child.as_str().to_string());
                }
            }
            _ => { rest.push(p); }
        }
    }
    apply_fortran_param_declaration_modes(&mut params, &rest);
    let mut body = walk_body(rest.into_iter())?;
    bind_fortran_param_declarations(&mut params, &mut body);
    promote_mutated_fortran_params(&mut params, &body);
    lower_fortran_body_intrinsics(&params, &mut body);
    normalize_fortran_function_result(&nm, result_name.as_deref(), &mut rt, &mut body);
    lower_fortran_array_semantics(&params, &mut body);
    let is_generator = body_has_yield(&body);
    Ok(Statement::new(StmtKind::FunctionDecl {
        name: nm, params, return_type: rt, body, modifiers: Modifiers::default(),
        handles: vec![], is_async: false, is_generator, is_sub: false,
    }))
}

fn walk_interface_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let mut name: Option<String> = None;
    let mut members = Vec::new();

    for child in pair.into_inner().filter(|p| meaningful(p)) {
        match child.as_rule() {
            Rule::identifier => {
                if name.is_none() {
                    name = Some(child.as_str().to_string());
                }
            }
            Rule::interface_designator => {
                if name.is_none() {
                    name = Some(child.as_str().trim().to_string());
                }
            }
            Rule::subroutine_decl | Rule::function_decl => {
                members.push(walk_interface_member(child)?);
            }
            Rule::procedure_decl => {
                members.extend(walk_interface_procedure_members(child)?);
            }
            _ => {}
        }
    }

    if members.is_empty() {
        return Ok(Some(Statement::new(StmtKind::Empty)));
    }

    let interface_name = name.unwrap_or_else(|| {
        members
            .first()
            .and_then(|member| match member {
                InterfaceMember::Method { name, .. } => Some(name.clone()),
                _ => None,
            })
            .unwrap_or_else(|| "__fortran_interface".to_string())
    });

    Ok(Some(Statement::new(StmtKind::InterfaceDecl {
        name: interface_name,
        parents: vec![],
        members,
        decorators: vec![],
    })))
}

fn walk_interface_member(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let decl = match pair.as_rule() {
        Rule::subroutine_decl => walk_sub(pair)?,
        Rule::function_decl => walk_func(pair)?,
        _ => return Err("unsupported Fortran interface member".to_string()),
    };

    match decl.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            is_sub,
            ..
        } => Ok(InterfaceMember::Method {
            name,
            params,
            return_type,
            is_sub,
            signature_source: None,
        }),
        _ => Err("expected function declaration in Fortran interface".to_string()),
    }
}

fn walk_interface_procedure_members(pair: Pair<Rule>) -> Result<Vec<InterfaceMember>, String> {
    let decl = walk_procedure_decl(pair)?;
    let StmtKind::VarDecl { declarations, .. } = decl.kind else {
        return Err("expected procedure declaration in Fortran interface".to_string());
    };

    let mut members = Vec::new();
    for declaration in declarations {
        let BindingPattern::Ident(name) = declaration.pattern else {
            continue;
        };

        let signature_source = declaration.type_hint.as_deref().and_then(|type_hint| {
            type_hint
                .strip_prefix("procedure(")
                .and_then(|rest| rest.strip_suffix(')'))
                .map(|source| source.trim().to_string())
        });

        members.push(InterfaceMember::Method {
            name,
            params: vec![],
            return_type: None,
            is_sub: true,
            signature_source,
        });
    }

    Ok(members)
}

fn walk_type(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut nm = String::new();
    let mut members = Vec::new();
    let mut parents = Vec::new();
    let mut modifiers = ClassModifiers::default();
    for p in pair.into_inner().filter(|p| meaningful(p)) {
        match p.as_rule() {
            Rule::identifier => { if nm.is_empty() { nm = p.as_str().to_string(); } }
            Rule::type_attribute => {
                apply_fortran_type_attribute(p, &mut modifiers, &mut parents);
            }
            Rule::type_member => {
                for m in p.into_inner() {
                    if matches!(m.as_rule(), Rule::var_declaration | Rule::procedure_decl) {
                        let decl = match m.as_rule() {
                            Rule::var_declaration => walk_var_decl(m)?,
                            Rule::procedure_decl => walk_procedure_decl(m)?,
                            _ => unreachable!(),
                        };
                        if let StmtKind::VarDecl { declarations, .. } = &decl.kind {
                            for d in declarations {
                                if let BindingPattern::Ident(fname) = &d.pattern {
                                    let field_type_hint = d.type_hint.as_ref().map(|type_hint| {
                                        if d.array_bounds.is_some() && !type_hint.trim_end().ends_with("()") {
                                            format!("{}()", type_hint.trim())
                                        } else {
                                            type_hint.clone()
                                        }
                                    });
                                    members.push(ClassMember::Field {
                                        name: fname.clone(), type_hint: field_type_hint, init: d.init.clone(),
                                        modifiers: Modifiers::default(), with_events: false, array_bounds: d.array_bounds.clone(),
                                    });
                                }
                            }
                        }
                    }
                }
            }
            Rule::type_bound_procedure => {
                let raw_rule = p.as_str().trim().to_ascii_lowercase();
                let is_generic_binding = raw_rule.starts_with("generic");
                let is_final_binding = raw_rule.starts_with("final");
                let mut method_bindings = Vec::new();
                let mut method_modifiers = Modifiers::default();
                let mut generic_or_final_names = Vec::new();
                for child in p.into_inner().filter(|p| meaningful(p)) {
                    match child.as_rule() {
                        Rule::tbp_attribute => {
                            apply_fortran_type_bound_attribute(child.as_str(), &mut method_modifiers);
                        }
                        Rule::tbp_binding => {
                            if let Some((public_name, implementation_name)) = parse_tbp_binding(child) {
                                method_bindings.push((public_name, implementation_name));
                            }
                        }
                        Rule::identifier | Rule::designator_name if is_generic_binding || is_final_binding => {
                            generic_or_final_names.push(child.as_str().to_string());
                        }
                        _ => {}
                    }
                }

                if method_bindings.is_empty() {
                    if is_generic_binding {
                        if let Some((public_name, implementation_names)) = generic_or_final_names.split_first() {
                            for implementation_name in implementation_names {
                                method_bindings.push((public_name.clone(), implementation_name.clone()));
                            }
                        }
                    } else {
                        for name in generic_or_final_names {
                            method_bindings.push((name.clone(), name));
                        }
                    }
                }

                method_bindings.dedup_by(|left, right| {
                    left.0.eq_ignore_ascii_case(&right.0)
                        && left.1.eq_ignore_ascii_case(&right.1)
                });
                for (method_name, implementation_name) in method_bindings {
                    members.push(ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
                        name: method_name,
                        params: vec![], return_type: None, body: vec![],
                        modifiers: method_modifiers.clone(), handles: vec![type_bound_impl_handle(&implementation_name)], is_async: false, is_generator: false, is_sub: true,
                    }))));
                }
            }
            _ => {}
        }
    }
    Ok(Statement::new(StmtKind::ClassDecl {
        name: nm, parents, interfaces: vec![], members, modifiers,
        decorators: vec![],
    }))
}

fn parse_tbp_binding(pair: Pair<Rule>) -> Option<(String, String)> {
    let mut names = pair.into_inner()
        .filter(|binding_part| matches!(binding_part.as_rule(), Rule::identifier | Rule::designator_name))
        .map(|binding_part| binding_part.as_str().to_string());
    let public_name = names.next()?;
    let implementation_name = names.next().unwrap_or_else(|| public_name.clone());
    Some((public_name, implementation_name))
}

fn type_bound_impl_handle(name: &str) -> String {
    format!("{FORTRAN_TBP_IMPL_HANDLE_PREFIX}{name}")
}

fn type_bound_impl_name(handles: &[String]) -> Option<&str> {
    handles.iter().find_map(|handle| handle.strip_prefix(FORTRAN_TBP_IMPL_HANDLE_PREFIX))
}

fn bind_module_type_bound_procedures(members: &mut [ClassMember]) {
    let mut procedures: HashMap<String, Vec<Statement>> = HashMap::new();
    let mut bound_impl_targets: Vec<(String, String)> = Vec::new();
    for member in members.iter() {
        if let ClassMember::Method(stmt) = member {
            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                procedures.entry(name.to_ascii_lowercase()).or_default().push((**stmt).clone());
            }
        }
    }

    for member in members.iter_mut() {
        let ClassMember::NestedType(stmt) = member else {
            continue;
        };
        let StmtKind::ClassDecl { name, members: class_members, .. } = &mut stmt.kind else {
            continue;
        };

        for class_member in class_members.iter_mut() {
            let ClassMember::Method(method_stmt) = class_member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                body,
                modifiers,
                handles,
                ..
            } = &method_stmt.kind else {
                continue;
            };

            if !body.is_empty() {
                continue;
            }

            let implementation_name = type_bound_impl_name(handles).unwrap_or(method_name);
            let Some(candidates) = procedures.get(&implementation_name.to_ascii_lowercase()) else {
                continue;
            };
            let Some(candidate) = candidates.iter().find(|candidate| function_decl_targets_type(candidate, name)) else {
                continue;
            };
            bound_impl_targets.push((implementation_name.to_ascii_lowercase(), name.clone()));
            let StmtKind::FunctionDecl {
                params,
                return_type,
                body,
                is_async,
                is_generator,
                is_sub,
                ..
            } = &candidate.kind else {
                continue;
            };

            *method_stmt = Box::new(Statement {
                kind: StmtKind::FunctionDecl {
                    name: method_name.clone(),
                    params: params.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: Vec::new(),
                    is_async: *is_async,
                    is_generator: *is_generator,
                    is_sub: *is_sub,
                },
                span: method_stmt.span,
            });

            if let StmtKind::FunctionDecl { params, .. } = &mut method_stmt.kind {
                if let Some(first_param) = params.first_mut() {
                    first_param.pass_by = PassBy::Value;
                }
            }
        }
    }

    for member in members.iter_mut() {
        let ClassMember::Method(stmt) = member else {
            continue;
        };

        let should_demote_self_ref = {
            let StmtKind::FunctionDecl { name, params, body, .. } = &stmt.kind else {
                continue;
            };
            bound_impl_targets.iter().any(|(implementation_name, class_name)| {
                name.eq_ignore_ascii_case(implementation_name)
                    && function_decl_targets_type_parts(params, body, class_name)
            })
        };

        if !should_demote_self_ref {
            continue;
        }

        let StmtKind::FunctionDecl { params, .. } = &mut stmt.kind else {
            continue;
        };
        if let Some(first_param) = params.first_mut() {
            first_param.pass_by = PassBy::Value;
        }
    }
}

fn bind_top_level_type_bound_procedures(body: &mut [Statement]) {
    let mut procedures: HashMap<String, Vec<Statement>> = HashMap::new();
    let mut bound_impl_targets: Vec<(String, String)> = Vec::new();
    for stmt in body.iter() {
        if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
            procedures
                .entry(name.to_ascii_lowercase())
                .or_default()
                .push(stmt.clone());
        }
    }

    for stmt in body.iter_mut() {
        let StmtKind::ClassDecl { name, members, .. } = &mut stmt.kind else {
            continue;
        };

        for member in members.iter_mut() {
            let ClassMember::Method(method_stmt) = member else {
                continue;
            };
            let StmtKind::FunctionDecl {
                name: method_name,
                body,
                modifiers,
                handles,
                ..
            } = &method_stmt.kind else {
                continue;
            };

            if !body.is_empty() {
                continue;
            }

            let implementation_name = type_bound_impl_name(handles).unwrap_or(method_name);
            let Some(candidates) = procedures.get(&implementation_name.to_ascii_lowercase()) else {
                continue;
            };
            let Some(candidate) = candidates
                .iter()
                .find(|candidate| function_decl_targets_type(candidate, name))
            else {
                continue;
            };
            bound_impl_targets.push((implementation_name.to_ascii_lowercase(), name.clone()));
            let StmtKind::FunctionDecl {
                params,
                return_type,
                body,
                is_async,
                is_generator,
                is_sub,
                ..
            } = &candidate.kind else {
                continue;
            };

            *method_stmt = Box::new(Statement {
                kind: StmtKind::FunctionDecl {
                    name: method_name.clone(),
                    params: params.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: Vec::new(),
                    is_async: *is_async,
                    is_generator: *is_generator,
                    is_sub: *is_sub,
                },
                span: method_stmt.span,
            });

            if let StmtKind::FunctionDecl { params, .. } = &mut method_stmt.kind {
                if let Some(first_param) = params.first_mut() {
                    first_param.pass_by = PassBy::Value;
                }
            }
        }
    }

    for stmt in body.iter_mut() {
        let should_demote_self_ref = {
            let StmtKind::FunctionDecl { name, params, body, .. } = &stmt.kind else {
                continue;
            };
            bound_impl_targets.iter().any(|(implementation_name, class_name)| {
                name.eq_ignore_ascii_case(implementation_name)
                    && function_decl_targets_type_parts(params, body, class_name)
            })
        };

        if !should_demote_self_ref {
            continue;
        }

        let StmtKind::FunctionDecl { params, .. } = &mut stmt.kind else {
            continue;
        };
        if let Some(first_param) = params.first_mut() {
            first_param.pass_by = PassBy::Value;
        }
    }
}

fn function_decl_targets_type(stmt: &Statement, class_name: &str) -> bool {
    let StmtKind::FunctionDecl { params, body, .. } = &stmt.kind else {
        return false;
    };
    function_decl_targets_type_parts(params, body, class_name)
}

fn function_decl_targets_type_parts(params: &[Param], body: &[Statement], class_name: &str) -> bool {
    let Some(first_param) = params.first() else {
        return false;
    };

    if first_param.type_hint.as_deref()
        .and_then(parse_derived_type_name)
        .is_some_and(|target_type| target_type.eq_ignore_ascii_case(class_name))
    {
        return true;
    }

    body.iter().find_map(|statement| {
        let StmtKind::VarDecl { declarations, .. } = &statement.kind else {
            return None;
        };
        declarations.iter().find_map(|declaration| {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                return None;
            };
            if !name.eq_ignore_ascii_case(&first_param.name) {
                return None;
            }
            declaration.type_hint.as_deref().and_then(parse_derived_type_name)
        })
    }).is_some_and(|target_type| target_type.eq_ignore_ascii_case(class_name))
}

// ── Expressions ────────────────────────────────────────────────────────────

fn walk_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::expression | Rule::logical_or | Rule::logical_and | Rule::logical_not
        | Rule::comparison | Rule::addition | Rule::multiplication | Rule::power | Rule::concat
        | Rule::unary => walk_binop(pair),
        Rule::primary_expr => {
            // primary_atom followed by zero or more postfix_op (member, call, coindex).
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let atom = inner.next().ok_or("empty primary")?;
            let mut expr = walk_expr(atom)?;
            for op in inner {
                if op.as_rule() != Rule::postfix_op { continue; }
                let mut op_inner = op.clone().into_inner().filter(|p| meaningful(p));
                let first = op_inner.next();
                match first {
                    Some(p) if matches!(p.as_rule(), Rule::identifier | Rule::designator_name) => {
                        // %field — member access
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                    Some(p) if p.as_rule() == Rule::argument_list => {
                        // (args) — call or index. Treat as Call by default;
                        // codegen for arrays sees Call(arr, [idx]) and emits index access.
                        let mut args = Vec::new();
                        for a in p.into_inner() {
                            if a.as_rule() == Rule::argument {
                                let (_, value) = walk_argument_expr(a)?;
                                args.push(Argument::positional(value));
                            }
                        }
                        if let Some(lowered) = lower_intrinsic_expr_call(&expr, &args) {
                            expr = lowered;
                        } else {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args,
                                optional: false,
                            });
                        }
                    }
                    Some(_) => { /* ignore */ }
                    None => {
                        // Empty `()` — call with no args
                        if let Some(lowered) = lower_intrinsic_expr_call(&expr, &[]) {
                            expr = lowered;
                        } else {
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args: Vec::new(),
                                optional: false,
                            });
                        }
                    }
                }
            }
            Ok(expr)
        }
        Rule::primary_atom => walk_expr(pair.into_inner().next().ok_or("empty atom")?),
        Rule::complex_literal => {
            // `(re, im)` complex constant — synthesise as a call to a
            // compiler-known `cmplx(re, im)` so the rest of the
            // pipeline treats it like the standard intrinsic. Walker
            // doesn't carry a Complex literal kind — the AST stays
            // language-neutral.
            let mut parts = pair.into_inner().filter(|p| meaningful(p));
            let re = walk_expr(parts.next().ok_or("complex: missing real")?)?;
            let im = walk_expr(parts.next().ok_or("complex: missing imag")?)?;
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("cmplx".to_string()))),
                args: vec![Argument::positional(re), Argument::positional(im)],
                optional: false,
            }));
        }
        Rule::array_constructor => {
            // `[a, b, c]` / `(/ a, b, c /)` / implied-do — walker
            // builds a flat `ExprKind::Array`. Implied-do `(i, i=1,n)`
            // is collapsed to its trip-count expression for now (the
            // proper expansion is a comprehension, pending a
            let mut elems: Vec<crate::ast::ArrayElement> = Vec::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::array_constructor_body) {
                    for v in p.into_inner() {
                        if matches!(v.as_rule(), Rule::array_constructor_value) {
                            // Walk the first expression child as the
                            // element value (implied-do collapses to
                            // its body expression for the MVP).
                            if let Some(inner) = v.into_inner()
                                .filter(|q| meaningful(q))
                                .find(|q| is_expr_rule(q.as_rule())
                                    || matches!(q.as_rule(), Rule::expression))
                            {
                                elems.push(crate::ast::ArrayElement {
                                    key: None,
                                    value: walk_expr(inner)?,
                                    spread: false,
                                    by_ref: false,
                                });
                            }
                        }
                    }
                }
            }
            Ok(Expression::new(ExprKind::Array(elems)))
        }
        Rule::literal => walk_expr(pair.into_inner().next().ok_or("empty literal")?),
        Rule::logical_literal => {
            Ok(Expression::new(ExprKind::Lit(Literal::Bool(pair.as_str().to_lowercase().contains("true")))))
        }
        Rule::number_literal => {
            let s = pair.as_str().trim();
            let clean = s.split('_').next().unwrap_or(s);
            if clean.contains('.') || clean.to_lowercase().contains('e') || clean.to_lowercase().contains('d') {
                let n: f64 = clean.replace('d', "e").replace('D', "E").parse().unwrap_or(0.0);
                Ok(Expression::new(ExprKind::Lit(Literal::Float(n))))
            } else {
                let n: i64 = clean.parse().unwrap_or(0);
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            }
        }
        Rule::string_literal => {
            let s = pair.as_str();
            let inner = &s[1..s.len()-1];
            Ok(Expression::new(ExprKind::Lit(Literal::Str(inner.replace("''", "'").replace("\"\"", "\"")))))
        }
        Rule::boz_literal => {
            // `b'..'` / `o'..'` / `z'..'` — bit / octal / hex literal.
            let s = pair.as_str();
            let prefix = s.chars().next().unwrap_or('z').to_ascii_lowercase();
            let body = &s[1..];
            let trimmed = body.trim_matches(|c: char| c == '\'' || c == '"');
            let radix = match prefix { 'b' => 2, 'o' => 8, _ => 16 };
            let n = i64::from_str_radix(trimmed, radix).unwrap_or(0);
            Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
        }
        Rule::identifier => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::designator_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::function_call_or_subscript => {
            let mut inner = pair.into_inner().filter(|p| meaningful(p));
            let nm = inner.next().ok_or("missing fn")?.as_str().to_string();
            let mut args = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for a in p.into_inner() {
                        if a.as_rule() == Rule::argument {
                            let (_, value) = walk_argument_expr(a)?;
                            args.push(Argument::positional(value));
                        }
                    }
                }
            }
            let callee = Expression::new(ExprKind::Ident(nm));
            if let Some(lowered) = lower_intrinsic_expr_call(&callee, &args) {
                Ok(lowered)
            } else {
                Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(callee), args, optional: false,
                }))
            }
        }
        Rule::argument => Ok(walk_argument_expr(pair)?.1),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn lower_intrinsic_statement(expr: &Expression) -> Option<Statement> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name.eq_ignore_ascii_case("nullify") {
        if args.is_empty() {
            return None;
        }

        let assigns = args.iter().map(|arg| Statement::new(StmtKind::Assign {
            targets: vec![arg.value.clone()],
            value: Expression::null(),
        })).collect::<Vec<_>>();
        return Some(Statement::new(StmtKind::Block(assigns)));
    }

    if name.eq_ignore_ascii_case("open") {
        return lower_open_intrinsic_statement(expr, args);
    }

    if name.eq_ignore_ascii_case("close") {
        return Some(Statement::new(StmtKind::CloseFile(
            args.iter().find(|arg| arg.name.as_deref().is_none_or(|name| name.eq_ignore_ascii_case("unit")))
                .map(|arg| arg.value.clone()),
        )));
    }

    None
}

fn is_fortran_allocator_intrinsic_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("allocate") || name.eq_ignore_ascii_case("deallocate"))
    )
}

fn lower_open_intrinsic_statement(expr: &Expression, args: &[Argument]) -> Option<Statement> {
    let file_number_arg = args.iter().find(|arg| {
        arg.name.as_deref().is_some_and(|name| {
            name.eq_ignore_ascii_case("unit") || name.eq_ignore_ascii_case("newunit")
        })
    }).or_else(|| args.first());
    let file_number = file_number_arg?.value.clone();

    let path = args.iter().find(|arg| {
        arg.name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case("file"))
    }).map(|arg| arg.value.clone()).unwrap_or_else(|| {
        Expression::string(&format!(
            "__fortran_scratch_{}_{}.tmp",
            expr.span.start_line.max(1),
            expr.span.start_col.max(1),
        ))
    });

    let mode = infer_fortran_open_mode(args);
    let open_stmt = Statement::new(StmtKind::OpenFile {
        path,
        mode,
        file_number: file_number.clone(),
    });

    if file_number_arg.and_then(|arg| arg.name.as_deref()).is_some_and(|name| name.eq_ignore_ascii_case("newunit")) {
        let assigned_unit = Expression::int((expr.span.start_line.max(1) as i64) * 1000 + expr.span.start_col.max(1) as i64);
        return Some(Statement::new(StmtKind::Block(vec![
            Statement::new(StmtKind::Assign {
                targets: vec![file_number],
                value: assigned_unit,
            }),
            open_stmt,
        ])));
    }

    Some(open_stmt)
}

fn infer_fortran_open_mode(args: &[Argument]) -> FileMode {
    for arg in args {
        let Some(name) = arg.name.as_deref() else {
            continue;
        };
        let ExprKind::Lit(Literal::Str(value)) = &arg.value.kind else {
            continue;
        };
        let value = value.to_ascii_lowercase();
        if name.eq_ignore_ascii_case("position") && value == "append" {
            return FileMode::Append;
        }
        if name.eq_ignore_ascii_case("action") {
            return match value.as_str() {
                "read" => FileMode::Input,
                "write" | "readwrite" => FileMode::Output,
                _ => FileMode::Output,
            };
        }
        if name.eq_ignore_ascii_case("access") && value == "stream" {
            return FileMode::Binary;
        }
        if name.eq_ignore_ascii_case("form") && value == "unformatted" {
            return FileMode::Binary;
        }
        if name.eq_ignore_ascii_case("status") {
            return match value.as_str() {
                "old" => FileMode::Input,
                "scratch" | "replace" | "new" | "unknown" => FileMode::Output,
                _ => FileMode::Output,
            };
        }
    }
    FileMode::Output
}

fn bind_fortran_param_declarations(params: &mut [Param], body: &mut Vec<Statement>) {
    let mut rewritten = Vec::with_capacity(body.len());
    for mut statement in body.drain(..) {
        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                let mut remaining = Vec::new();
                for declaration in declarations.drain(..) {
                    let mut matched_param = false;
                    if let BindingPattern::Ident(name) = &declaration.pattern {
                        if let Some(param) = params.iter_mut().find(|param| param.name.eq_ignore_ascii_case(name)) {
                            let declaration_type_hint = declaration.type_hint.as_ref().map(|type_hint| {
                                if declaration.array_bounds.is_some() && !type_hint.trim_end().ends_with("()") {
                                    format!("{}()", type_hint.trim())
                                } else {
                                    type_hint.clone()
                                }
                            });
                            if param.type_hint.is_none()
                                || (declaration.array_bounds.is_some()
                                    && param.type_hint.as_deref().is_some_and(|type_hint| !type_hint.trim_end().ends_with("()")))
                            {
                                param.type_hint = declaration_type_hint;
                            }
                            if matches!(param.pass_by, PassBy::Value)
                                && (declaration.array_bounds.is_some()
                                || declaration.type_hint.as_deref().and_then(parse_derived_type_name).is_some()
                            ) {
                                param.pass_by = PassBy::Const;
                            }
                            matched_param = true;
                        }
                    }
                    if !matched_param {
                        remaining.push(declaration);
                    }
                }

                if !remaining.is_empty() {
                    *declarations = remaining;
                    rewritten.push(statement);
                }
            }
            _ => rewritten.push(statement),
        }
    }
    *body = rewritten;
}

fn apply_fortran_param_declaration_modes(params: &mut [Param], rest: &[Pair<Rule>]) {
    for pair in rest {
        apply_fortran_param_declaration_modes_from_pair(params, pair.clone());
    }
}

fn apply_fortran_param_declaration_modes_from_pair(params: &mut [Param], pair: Pair<Rule>) {
    if pair.as_rule() == Rule::var_declaration {
        let mut inner = pair.clone().into_inner();
        let type_hint = inner.next().map(|child| child.as_str().trim().to_string());
        let mut intent_mode = None;

        for child in inner {
            if child.as_rule() != Rule::var_attributes {
                continue;
            }

            for attr in child.into_inner().filter(|attr| meaningful(attr)) {
                if attr.as_rule() != Rule::var_attribute {
                    continue;
                }

                if let Some(mode) = parse_fortran_intent_mode(attr.as_str()) {
                    intent_mode = Some(mode);
                }
            }
        }

        if type_hint.is_some() || intent_mode.is_some() {
            for declaration in pair.clone().into_inner().filter(|child| meaningful(child)) {
                if declaration.as_rule() != Rule::var_declarator_list {
                    continue;
                }

                for declarator in declaration.into_inner().filter(|child| meaningful(child)) {
                    if declarator.as_rule() != Rule::var_declarator {
                        continue;
                    }

                    let mut parts = declarator.into_inner().filter(|child| meaningful(child));
                    let Some(name) = parts.next().map(|child| child.as_str().to_string()) else {
                        continue;
                    };

                    let Some(param) = params.iter_mut().find(|param| param.name.eq_ignore_ascii_case(&name)) else {
                        continue;
                    };

                    if param.type_hint.is_none() {
                        param.type_hint = type_hint.clone();
                    }

                    if let Some(mode) = intent_mode {
                        param.pass_by = mode;
                    }
                }
            }
        }
    }

    if pair.as_rule() == Rule::procedure_decl {
        let mut interface_name: Option<String> = None;
        let mut intent_mode = None;

        for child in pair.clone().into_inner().filter(|child| meaningful(child)) {
            match child.as_rule() {
                Rule::identifier if interface_name.is_none() => {
                    interface_name = Some(child.as_str().trim().to_string());
                }
                Rule::proc_attribute => {
                    if let Some(mode) = parse_fortran_intent_mode(child.as_str()) {
                        intent_mode = Some(mode);
                    }
                }
                _ => {}
            }
        }

        let type_hint = interface_name
            .as_ref()
            .map(|iface| format!("procedure({iface})"))
            .unwrap_or_else(|| "procedure".to_string());

        for declaration in pair.clone().into_inner().filter(|child| meaningful(child)) {
            if declaration.as_rule() != Rule::proc_decl_item {
                continue;
            }

            let Some(name) = declaration
                .into_inner()
                .filter(|child| meaningful(child))
                .find(|child| child.as_rule() == Rule::identifier)
                .map(|child| child.as_str().to_string())
            else {
                continue;
            };

            let Some(param) = params.iter_mut().find(|param| param.name.eq_ignore_ascii_case(&name)) else {
                continue;
            };

            if param.type_hint.is_none() {
                param.type_hint = Some(type_hint.clone());
            }

            if let Some(mode) = intent_mode {
                param.pass_by = mode;
            }
        }
    }

    for child in pair.into_inner().filter(|child| meaningful(child)) {
        apply_fortran_param_declaration_modes_from_pair(params, child);
    }
}

fn parse_fortran_intent_mode(attr_text: &str) -> Option<PassBy> {
    let attr = attr_text.trim().to_ascii_lowercase();
    if !attr.starts_with("intent(") {
        return None;
    }

    let inner = attr.strip_prefix("intent(")?.strip_suffix(')')?.trim();
    match inner {
        "in" => Some(PassBy::Const),
        "out" => Some(PassBy::Out),
        "inout" => Some(PassBy::Ref),
        _ => None,
    }
}

fn promote_mutated_fortran_params(params: &mut [Param], body: &[Statement]) {
    for param in params.iter_mut() {
        if param.pass_by != PassBy::Const {
            continue;
        }
        if body.iter().any(|statement| statement_mutates_fortran_param(statement, &param.name)) {
            param.pass_by = PassBy::Ref;
        }
    }
}

fn statement_mutates_fortran_param(statement: &Statement, param_name: &str) -> bool {
    match &statement.kind {
        StmtKind::Assign { targets, .. } => targets.iter().any(|target| expr_targets_fortran_param(target, param_name)),
        StmtKind::Block(stmts)
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::With { body: stmts, .. }
        | StmtKind::Using { body: stmts, .. }
        | StmtKind::Lock { body: stmts, .. } => stmts.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)),
        StmtKind::If { then_body, elifs, else_body, .. } => {
            then_body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || elifs.iter().any(|(_, body)| body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
                || else_body.as_ref().is_some_and(|body| body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
        }
        StmtKind::While { body, else_body, .. }
        | StmtKind::ForIn { body, else_body, .. } => {
            body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || else_body.as_ref().is_some_and(|stmts| stmts.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
        }
        StmtKind::For { init, body, .. } => {
            init.as_ref().is_some_and(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name))
        }
        StmtKind::Switch { cases, default, .. } => {
            cases.iter().any(|case| case.body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
                || default.as_ref().is_some_and(|body| body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name))
                || catches.iter().any(|catch| catch.body.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
                || else_body.as_ref().is_some_and(|stmts| stmts.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
                || finally.as_ref().is_some_and(|stmts| stmts.iter().any(|stmt| statement_mutates_fortran_param(stmt, param_name)))
        }
        _ => false,
    }
}

fn expr_targets_fortran_param(expr: &Expression, param_name: &str) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => name.eq_ignore_ascii_case(param_name),
        ExprKind::Index { object, .. } | ExprKind::Member { object, .. } => {
            expr_targets_fortran_param(object, param_name)
        }
        _ => false,
    }
}

fn normalize_fortran_function_result(
    function_name: &str,
    result_name: Option<&str>,
    return_type: &mut Option<String>,
    body: &mut Vec<Statement>,
) {
    let result_var = result_name.unwrap_or(function_name).to_string();

    if return_type.is_none() {
        *return_type = find_function_result_type(body, &result_var);
    }

    rewrite_function_returns(body, &result_var);

    let needs_final_return = !matches!(body.last().map(|stmt| &stmt.kind), Some(StmtKind::Return(_)));
    if needs_final_return {
        body.push(Statement::new(StmtKind::Return(Some(Expression::ident(&result_var)))));
    }
}

fn find_function_result_type(body: &[Statement], result_var: &str) -> Option<String> {
    body.iter().find_map(|statement| {
        let StmtKind::VarDecl { declarations, .. } = &statement.kind else {
            return None;
        };
        declarations.iter().find_map(|declaration| {
            let BindingPattern::Ident(name) = &declaration.pattern else {
                return None;
            };
            if name.eq_ignore_ascii_case(result_var) {
                declaration.type_hint.as_ref().map(|type_hint| {
                    if declaration.array_bounds.is_some() && !type_hint.trim_end().ends_with("()") {
                        format!("{}()", type_hint.trim())
                    } else {
                        type_hint.clone()
                    }
                })
            } else {
                None
            }
        })
    })
}

fn rewrite_function_returns(body: &mut [Statement], result_var: &str) {
    for statement in body.iter_mut() {
        match &mut statement.kind {
            StmtKind::Return(value) if value.as_ref().is_none_or(|expr| matches!(expr.kind, ExprKind::Lit(Literal::Null))) => {
                *value = Some(Expression::ident(result_var));
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => rewrite_function_returns(stmts, result_var),
            StmtKind::If { then_body, elifs, else_body, .. } => {
                rewrite_function_returns(then_body, result_var);
                for (_, elif_body) in elifs {
                    rewrite_function_returns(elif_body, result_var);
                }
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                rewrite_function_returns(stmts, result_var);
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                if let Some(init) = init {
                    rewrite_function_returns(std::slice::from_mut(init.as_mut()), result_var);
                }
                rewrite_function_returns(stmts, result_var);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    rewrite_function_returns(&mut case.body, result_var);
                }
                if let Some(default) = default {
                    rewrite_function_returns(default, result_var);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                rewrite_function_returns(try_body, result_var);
                for catch in catches {
                    rewrite_function_returns(&mut catch.body, result_var);
                }
                if let Some(else_body) = else_body {
                    rewrite_function_returns(else_body, result_var);
                }
                if let Some(finally) = finally {
                    rewrite_function_returns(finally, result_var);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_semantics(params: &[Param], body: &mut Vec<Statement>) {
    let mut arrays = HashSet::new();
    let mut callables = HashSet::new();
    let mut array_fields = HashSet::new();
    for param in params {
        if param.type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()")) {
            arrays.insert(param.name.to_ascii_lowercase());
        }
        if param.type_hint.as_deref().is_some_and(is_fortran_string_type_hint) {
            arrays.insert(param.name.to_ascii_lowercase());
        }
        if param.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
            callables.insert(param.name.to_ascii_lowercase());
        }
    }
    collect_fortran_callable_names(body, &mut callables);
    collect_fortran_array_field_names(body, &mut array_fields);
    lower_fortran_array_semantics_with_env(body, &mut arrays, &mut callables, &array_fields);
}

fn repair_remaining_fortran_array_calls(body: &mut [Statement]) {
    let mut arrays = HashSet::new();
    let mut callables = HashSet::new();
    let mut array_fields = HashSet::new();
    collect_fortran_callable_names(body, &mut callables);
    collect_fortran_array_field_names(body, &mut array_fields);
    repair_remaining_fortran_array_calls_with_env(body, &mut arrays, &mut callables, &array_fields);
}

fn repair_remaining_fortran_array_calls_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    callables: &mut HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_remaining_fortran_array_calls_in_statement(statement, arrays, callables, array_fields);

        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration.array_bounds.is_some()
                    || declaration.init.as_ref().is_some_and(is_array_initializer_expr)
                    || declaration.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                {
                    arrays.insert(name.to_ascii_lowercase());
                }
                if declaration.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                    callables.insert(name.to_ascii_lowercase());
                }
            }
        }

        match &mut statement.kind {
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested = arrays.clone();
                let mut nested_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(stmts, &mut nested, &mut nested_callables, array_fields);
            }
            StmtKind::ModuleDecl { members, .. } => {
                repair_remaining_fortran_array_calls_in_members(members, arrays, callables, array_fields);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = arrays.clone();
                let mut then_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(then_body, &mut then_arrays, &mut then_callables, array_fields);
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    let mut elif_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(elif_body, &mut elif_arrays, &mut elif_callables, array_fields);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(stmts, &mut loop_arrays, &mut loop_callables, array_fields);
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                if let Some(init) = init {
                    repair_remaining_fortran_array_calls_with_env(std::slice::from_mut(init.as_mut()), &mut loop_arrays, &mut loop_callables, array_fields);
                }
                repair_remaining_fortran_array_calls_with_env(stmts, &mut loop_arrays, &mut loop_callables, array_fields);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    let mut case_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(&mut case.body, &mut case_arrays, &mut case_callables, array_fields);
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    let mut default_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(default, &mut default_arrays, &mut default_callables, array_fields);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = arrays.clone();
                let mut try_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(try_body, &mut try_arrays, &mut try_callables, array_fields);
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    let mut catch_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(&mut catch.body, &mut catch_arrays, &mut catch_callables, array_fields);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    let mut finally_callables = callables.clone();
                    repair_remaining_fortran_array_calls_with_env(finally, &mut finally_arrays, &mut finally_callables, array_fields);
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                let mut fn_arrays = arrays.clone();
                let mut fn_callables = callables.clone();
                for param in params.iter() {
                    if param.type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                    {
                        fn_arrays.insert(param.name.to_ascii_lowercase());
                    }
                    if param.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                        fn_callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                collect_fortran_callable_names(body, &mut fn_callables);
                repair_remaining_fortran_array_calls_with_env(body, &mut fn_arrays, &mut fn_callables, array_fields);
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                repair_remaining_fortran_array_calls_in_members(members, arrays, callables, array_fields);
            }
            _ => {}
        }
    }
}

fn repair_remaining_fortran_array_calls_in_members(
    members: &mut [ClassMember],
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { params, body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut method_arrays = arrays.clone();
                let mut method_callables = callables.clone();
                for param in params.iter() {
                    if param.type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                    {
                        method_arrays.insert(param.name.to_ascii_lowercase());
                    }
                    if param.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                        method_callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                repair_remaining_fortran_array_calls_with_env(body, &mut method_arrays, &mut method_callables, array_fields);
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = arrays.clone();
                let mut nested_callables = callables.clone();
                repair_remaining_fortran_array_calls_with_env(std::slice::from_mut(stmt.as_mut()), &mut nested_arrays, &mut nested_callables, array_fields);
            }
            _ => {}
        }
    }
}

fn rewrite_remaining_fortran_array_calls_in_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            if !is_fortran_allocator_intrinsic_expr(expr) {
                rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_remaining_fortran_array_calls_in_expr(target, arrays, callables, array_fields);
            }
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(target, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::Return(Some(expr)) => rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields),
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields);
            }
            if let Some(cause) = cause {
                rewrite_remaining_fortran_array_calls_in_expr(cause, arrays, callables, array_fields);
            }
        }
        StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. }
        | StmtKind::Switch { expr: cond, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
            }
            if let Some(update) = update {
                rewrite_remaining_fortran_array_calls_in_expr(update, arrays, callables, array_fields);
            }
        }
        _ => {}
    }
}

fn rewrite_remaining_fortran_array_calls_in_expr(
    expr: &mut Expression,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(left, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(right, arrays, callables, array_fields);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner) => rewrite_remaining_fortran_array_calls_in_expr(inner, arrays, callables, array_fields),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_remaining_fortran_array_calls_in_expr(cond, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(then, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(else_, arrays, callables, array_fields);
        }
        ExprKind::Member { object, .. } => rewrite_remaining_fortran_array_calls_in_expr(object, arrays, callables, array_fields),
        ExprKind::Index { object, index, .. } => {
            rewrite_remaining_fortran_array_calls_in_expr(object, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(index, arrays, callables, array_fields);
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(lower, arrays, callables, array_fields);
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(upper, arrays, callables, array_fields);
            }
            if let Some(step) = step.as_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(step, arrays, callables, array_fields);
            }
        }
        ExprKind::Call { callee, args, optional } => {
            rewrite_remaining_fortran_array_calls_in_expr(callee, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(&mut arg.value, arrays, callables, array_fields);
            }
            if !args.is_empty()
                && !*optional
                && !is_known_fortran_callable(callee, callables)
                && (is_known_fortran_array(callee, arrays, array_fields)
                    || matches!(&callee.kind, ExprKind::Index { .. }))
            {
                expr.kind = build_fortran_index_chain(callee.as_ref().clone(), args);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_remaining_fortran_array_calls_in_expr(class, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_remaining_fortran_array_calls_in_expr(&mut arg.value, arrays, callables, array_fields);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_remaining_fortran_array_calls_in_expr(target, arrays, callables, array_fields);
            rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
        }
        ExprKind::Object(items) => {
            for item in items {
                match item {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_remaining_fortran_array_calls_in_expr(key, arrays, callables, array_fields);
                        rewrite_remaining_fortran_array_calls_in_expr(value, arrays, callables, array_fields);
                    }
                    ObjectProperty::Spread(expr) => rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields),
                    ObjectProperty::Shorthand(_) | ObjectProperty::Method { .. } | ObjectProperty::Accessor { .. } => {}
                }
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = item.key.as_mut() {
                    rewrite_remaining_fortran_array_calls_in_expr(key, arrays, callables, array_fields);
                }
                rewrite_remaining_fortran_array_calls_in_expr(&mut item.value, arrays, callables, array_fields);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    rewrite_remaining_fortran_array_calls_in_expr(expr, arrays, callables, array_fields);
                }
            }
        }
        _ => {}
    }
}

fn collect_fortran_array_field_names(body: &[Statement], array_fields: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_names_in_members(members, array_fields);
            }
            StmtKind::ModuleDecl { members, .. } => collect_fortran_array_field_names_in_members(members, array_fields),
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_array_field_names(body, array_fields),
            _ => {}
        }
    }
}

fn collect_fortran_callable_names(body: &[Statement], callables: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    if declaration.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                        callables.insert(name.to_ascii_lowercase());
                    }
                }
            }
            StmtKind::FunctionDecl { params, body, .. } => {
                for param in params {
                    if param.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                        callables.insert(param.name.to_ascii_lowercase());
                    }
                }
                collect_fortran_callable_names(body, callables);
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                            collect_fortran_callable_names(std::slice::from_ref(stmt.as_ref()), callables);
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_callable_names(body, callables),
            _ => {}
        }
    }
}

fn collect_fortran_array_function_names(body: &[Statement], array_functions: &mut HashSet<String>) {
    for statement in body {
        match &statement.kind {
            StmtKind::FunctionDecl { name, return_type, body, .. } => {
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    array_functions.insert(name.to_ascii_lowercase());
                }
                collect_fortran_array_function_names(body, array_functions);
            }
            StmtKind::InterfaceDecl { members, .. } => {
                for member in members {
                    if let InterfaceMember::Method {
                        name,
                        return_type,
                        signature_source,
                        ..
                    } = member
                    {
                        if return_type
                            .as_deref()
                            .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        {
                            array_functions.insert(name.to_ascii_lowercase());
                        }
                        if signature_source
                            .as_ref()
                            .is_some_and(|source| array_functions.contains(&source.to_ascii_lowercase()))
                        {
                            array_functions.insert(name.to_ascii_lowercase());
                        }
                    }
                }
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_function_names_in_members(members, array_functions);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_function_names_in_members(members, array_functions);
            }
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_array_function_names(body, array_functions),
            _ => {}
        }
    }
}

fn collect_fortran_array_function_names_in_members(
    members: &[ClassMember],
    array_functions: &mut HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                collect_fortran_array_function_names(std::slice::from_ref(stmt.as_ref()), array_functions);
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_sizes(body: &[Statement], array_field_sizes: &mut HashMap<String, Expression>) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_sizes_in_members(members, array_field_sizes);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_field_sizes_in_members(members, array_field_sizes);
            }
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_array_field_sizes(body, array_field_sizes),
            _ => {}
        }
    }
}

fn collect_fortran_array_field_sizes_in_members(
    members: &[ClassMember],
    array_field_sizes: &mut HashMap<String, Expression>,
) {
    for member in members {
        match member {
            ClassMember::Field { name, array_bounds, init, .. } => {
                if let Some(size) = array_bounds
                    .as_deref()
                    .and_then(bounds_total_size_expr)
                    .or_else(|| init.as_ref().and_then(array_init_size_expr))
                {
                    array_field_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_sizes(std::slice::from_ref(stmt.as_ref()), array_field_sizes);
            }
            _ => {}
        }
    }
}

fn collect_fortran_array_field_names_in_members(
    members: &[ClassMember],
    array_fields: &mut HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Field { name, type_hint, array_bounds, .. } => {
                if array_bounds.is_some()
                    || type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                    || type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                {
                    array_fields.insert(name.to_ascii_lowercase());
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_names(std::slice::from_ref(stmt.as_ref()), array_fields);
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_semantics_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    callables: &mut HashSet<String>,
    array_fields: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_array_subscripts_in_statement(statement, arrays, callables, array_fields);

        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration.array_bounds.is_some()
                    || declaration.init.as_ref().is_some_and(is_array_initializer_expr)
                    || declaration.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                {
                    arrays.insert(name.to_ascii_lowercase());
                }
                if declaration.type_hint.as_deref().is_some_and(is_fortran_callable_type_hint) {
                    callables.insert(name.to_ascii_lowercase());
                }
            }
        }

        match &mut statement.kind {
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = arrays.clone();
                let mut nested_callables = callables.clone();
                lower_fortran_array_semantics_with_env(stmts, &mut nested, &mut nested_callables, array_fields);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = arrays.clone();
                let mut then_callables = callables.clone();
                lower_fortran_array_semantics_with_env(then_body, &mut then_arrays, &mut then_callables, array_fields);
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    let mut elif_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(elif_body, &mut elif_arrays, &mut elif_callables, array_fields);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                lower_fortran_array_semantics_with_env(stmts, &mut loop_arrays, &mut loop_callables, array_fields);
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = arrays.clone();
                let mut loop_callables = callables.clone();
                if let Some(init) = init {
                    lower_fortran_array_semantics_with_env(std::slice::from_mut(init.as_mut()), &mut loop_arrays, &mut loop_callables, array_fields);
                }
                lower_fortran_array_semantics_with_env(stmts, &mut loop_arrays, &mut loop_callables, array_fields);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    let mut case_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(&mut case.body, &mut case_arrays, &mut case_callables, array_fields);
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    let mut default_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(default, &mut default_arrays, &mut default_callables, array_fields);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = arrays.clone();
                let mut try_callables = callables.clone();
                lower_fortran_array_semantics_with_env(try_body, &mut try_arrays, &mut try_callables, array_fields);
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    let mut catch_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(&mut catch.body, &mut catch_arrays, &mut catch_callables, array_fields);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    let mut else_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(else_body, &mut else_arrays, &mut else_callables, array_fields);
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    let mut finally_callables = callables.clone();
                    lower_fortran_array_semantics_with_env(finally, &mut finally_arrays, &mut finally_callables, array_fields);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_scalar_array_assignments(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_ranks = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_field_ranks = HashMap::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_ranks(body, &mut array_field_ranks);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_scalar_array_assignments_with_env(
        body,
        &mut array_sizes,
        &mut array_ranks,
        &array_field_sizes,
        &array_field_ranks,
        &array_functions,
    );
}

fn lower_fortran_array_assignments(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_assignments_with_env(body, &mut array_sizes, &array_field_sizes, &array_fields, &array_functions);
}

fn lower_fortran_array_call_arguments(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_call_arguments_with_env(body, &mut array_sizes, &array_field_sizes, &array_fields, &array_functions);
}

fn lower_fortran_array_call_arguments_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_call_argument_statement(statement, array_sizes, array_field_sizes, array_fields, array_functions);

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::FunctionDecl { body: function_body, .. } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(function_body, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_call_arguments_in_members(members, array_sizes, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(stmts, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(then_body, &mut then_arrays, array_field_sizes, array_fields, array_functions);
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(elif_body, &mut elif_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(stmts, &mut loop_arrays, array_field_sizes, array_fields, array_functions);
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_call_arguments_with_env(std::slice::from_mut(init.as_mut()), &mut loop_arrays, array_field_sizes, array_fields, array_functions);
                }
                lower_fortran_array_call_arguments_with_env(stmts, &mut loop_arrays, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(&mut case.body, &mut case_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(default, &mut default_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(try_body, &mut try_arrays, array_field_sizes, array_fields, array_functions);
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(&mut catch.body, &mut catch_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_call_arguments_with_env(finally, &mut finally_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_call_arguments_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(body, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            ClassMember::NestedType(stmt) => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_call_arguments_with_env(std::slice::from_mut(stmt.as_mut()), &mut nested, array_field_sizes, array_fields, array_functions);
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_call_argument_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    let (callee, args, optional, rebuild) = match &statement.kind {
        StmtKind::Expr(expr) => {
            if is_fortran_allocator_intrinsic_expr(expr) {
                return;
            }
            let ExprKind::Call { callee, args, optional } = &expr.kind else {
                return;
            };
            (callee, args, optional, None)
        }
        StmtKind::Assign { targets, value } => {
            let ExprKind::Call { callee, args, optional } = &value.kind else {
                return;
            };
            (callee, args, optional, Some(targets.clone()))
        }
        _ => return,
    };
    if *optional {
        return;
    }

    let mut setup = Vec::new();
    let mut lowered_args = args.clone();
    let mut temp_index = 0usize;

    for arg in lowered_args.iter_mut() {
        if !expr_is_known_fortran_array(&arg.value, array_sizes, array_field_sizes, array_functions)
            || matches!(arg.value.kind, ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Array(_))
        {
            continue;
        }
        let Some(size) = resolve_fortran_array_expr_size(&arg.value, array_sizes, array_field_sizes) else {
            continue;
        };
        let temp_name = format!("__fortran_array_arg_{temp_index}");
        temp_index += 1;
        setup.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(temp_name.clone()),
                type_hint: None,
                init: Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("Array")),
                    args: vec![Argument::positional(size.clone()), Argument::positional(Expression::int(0))],
                    optional: false,
                })),
                array_bounds: Some(vec![size.clone()]),
                with_events: false,
            }],
            kind: VarDeclKind::Dim,
        }));
        setup.push(build_fortran_array_materialization_statement(
            &temp_name,
            size,
            &arg.value,
            array_sizes,
            array_fields,
            array_functions,
        ));
        arg.value = Expression::ident(&temp_name);
    }

    if setup.is_empty() {
        return;
    }

    let lowered_call = Expression::new(ExprKind::Call {
        callee: callee.clone(),
        args: lowered_args,
        optional: false,
    });
    setup.push(match rebuild {
        Some(targets) => Statement::new(StmtKind::Assign {
            targets,
            value: lowered_call,
        }),
        None => Statement::new(StmtKind::Expr(lowered_call)),
    });
    *statement = Statement::new(StmtKind::Block(setup));
}

fn resolve_fortran_array_expr_size(
    expr: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Ident(name) => array_sizes.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Member { field, .. } => array_field_sizes.get(&field.to_ascii_lowercase()).cloned(),
        ExprKind::Array(items) => Some(Expression::int(items.len() as i64)),
        ExprKind::Binary { left, right, .. } => resolve_fortran_array_expr_size(left, array_sizes, array_field_sizes)
            .or_else(|| resolve_fortran_array_expr_size(right, array_sizes, array_field_sizes)),
        ExprKind::Unary { expr: inner, .. } => resolve_fortran_array_expr_size(inner, array_sizes, array_field_sizes),
        ExprKind::Ternary { cond, then, else_ } => resolve_fortran_array_expr_size(cond, array_sizes, array_field_sizes)
            .or_else(|| resolve_fortran_array_expr_size(then, array_sizes, array_field_sizes))
            .or_else(|| resolve_fortran_array_expr_size(else_, array_sizes, array_field_sizes)),
        ExprKind::Slice { lower, upper, .. } => {
            let lower = lower.as_deref().cloned().unwrap_or_else(|| Expression::int(0));
            upper.as_deref().cloned().map(|upper| Expression::new(ExprKind::Binary {
                left: Box::new(upper),
                op: BinOp::Sub,
                right: Box::new(lower),
            }))
        }
        ExprKind::Index { object, index, .. } => match &index.kind {
            ExprKind::Slice { .. } => fortran_slice_extent(expr),
            _ => resolve_fortran_array_expr_size(object, array_sizes, array_field_sizes),
        },
        _ => None,
    }
}

fn build_fortran_array_materialization_statement(
    temp_name: &str,
    size: Expression,
    value: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Statement {
    let loop_var = "__fortran_array_arg_index";
    let loop_expr = Expression::ident(loop_var);
    Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(loop_var)],
            value: Expression::int(0),
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::ident(loop_var)),
            op: BinOp::Lt,
            right: Box::new(size),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(loop_var)),
            value: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::ident(loop_var)),
                op: BinOp::Add,
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: vec![Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(temp_name)),
                index: Box::new(loop_expr.clone()),
                null_safe: false,
            })],
            value: lower_fortran_array_materialization_value(
                value,
                &loop_expr,
                array_sizes,
                array_fields,
                array_functions,
            ),
        })],
    })
}

fn lower_fortran_array_materialization_value(
    expr: &Expression,
    loop_index: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if array_sizes.contains_key(&name.to_ascii_lowercase()) => Expression::new(ExprKind::Index {
            object: Box::new(expr.clone()),
            index: Box::new(loop_index.clone()),
            null_safe: false,
        }),
        ExprKind::Member { field, .. } if array_fields.contains(&field.to_ascii_lowercase()) => Expression::new(ExprKind::Index {
            object: Box::new(expr.clone()),
            index: Box::new(loop_index.clone()),
            null_safe: false,
        }),
        ExprKind::Array(_) | ExprKind::Slice { .. } => Expression::new(ExprKind::Index {
            object: Box::new(expr.clone()),
            index: Box::new(loop_index.clone()),
            null_safe: false,
        }),
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array") || array_functions.contains(&name.to_ascii_lowercase()))
                || matches!(&callee.kind, ExprKind::Member { field, .. } if matches!(field.to_ascii_lowercase().as_str(), "map" | "filter" | "flatmap")) =>
        {
            Expression::new(ExprKind::Index {
                object: Box::new(expr.clone()),
                index: Box::new(loop_index.clone()),
                null_safe: false,
            })
        }
        ExprKind::Index { object, index, null_safe } => Expression::new(ExprKind::Index {
            object: Box::new(lower_fortran_array_materialization_value(
                object,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            index: Box::new(lower_fortran_array_materialization_value(
                index,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            null_safe: *null_safe,
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(lower_fortran_array_materialization_value(
                left,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            right: Box::new(lower_fortran_array_materialization_value(
                right,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(lower_fortran_array_materialization_value(
                inner,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(lower_fortran_array_materialization_value(
                cond,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            then: Box::new(lower_fortran_array_materialization_value(
                then,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
            else_: Box::new(lower_fortran_array_materialization_value(
                else_,
                loop_index,
                array_sizes,
                array_fields,
                array_functions,
            )),
        }),
        _ => expr.clone(),
    }
}

fn lower_fortran_array_assignments_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_assignment_statement(statement, array_sizes, array_field_sizes, array_fields, array_functions);

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::FunctionDecl { body: function_body, .. } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(function_body, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_assignments_in_members(members, array_sizes, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(stmts, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(then_body, &mut then_arrays, array_field_sizes, array_fields, array_functions);
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(elif_body, &mut elif_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(stmts, &mut loop_arrays, array_field_sizes, array_fields, array_functions);
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_assignments_with_env(std::slice::from_mut(init.as_mut()), &mut loop_arrays, array_field_sizes, array_fields, array_functions);
                }
                lower_fortran_array_assignments_with_env(stmts, &mut loop_arrays, array_field_sizes, array_fields, array_functions);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(&mut case.body, &mut case_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(default, &mut default_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_assignments_with_env(try_body, &mut try_arrays, array_field_sizes, array_fields, array_functions);
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(&mut catch.body, &mut catch_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(else_body, &mut else_arrays, array_field_sizes, array_fields, array_functions);
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_assignments_with_env(finally, &mut finally_arrays, array_field_sizes, array_fields, array_functions);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_assignments_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { body, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(body, &mut nested, array_field_sizes, array_fields, array_functions);
            }
            ClassMember::NestedType(stmt) => {
                let mut nested = array_sizes.clone();
                lower_fortran_array_assignments_with_env(std::slice::from_mut(stmt.as_mut()), &mut nested, array_field_sizes, array_fields, array_functions);
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_assignment_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    let StmtKind::Assign { targets, value } = &statement.kind else {
        return;
    };
    let [target] = targets.as_slice() else {
        return;
    };

    let Some(extent) = fortran_array_assignment_extent(
        target,
        value,
        array_sizes,
        array_field_sizes,
    ) else {
        return;
    };

    let should_lower = contains_fortran_slice(target)
        || (expr_is_known_fortran_array(value, array_sizes, array_field_sizes, array_functions)
            && !matches!(value.kind, ExprKind::Ident(_) | ExprKind::Member { .. }));
    if !should_lower {
        return;
    }

    let loop_var = "__fortran_array_index";
    let loop_expr = Expression::ident(loop_var);
    let lowered_target = lower_fortran_array_assignment_target(target, &loop_expr);
    let lowered_value = lower_fortran_array_assignment_value(
        value,
        &loop_expr,
        array_sizes,
        array_field_sizes,
        array_fields,
        array_functions,
    );

    *statement = Statement::new(StmtKind::Block(vec![Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(loop_var)],
            value: Expression::int(0),
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            left: Box::new(Expression::ident(loop_var)),
            op: BinOp::Lt,
            right: Box::new(extent),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(loop_var)),
            value: Box::new(Expression::new(ExprKind::Binary {
                left: Box::new(Expression::ident(loop_var)),
                op: BinOp::Add,
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: vec![Statement::new(StmtKind::Assign {
            targets: vec![lowered_target],
            value: lowered_value,
        })],
    })]));
}

fn fortran_array_assignment_extent(
    target: &Expression,
    value: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    if let Some(extent) = fortran_slice_extent(target) {
        return Some(extent);
    }
    resolve_fortran_array_target_size(target, array_sizes, array_field_sizes)
        .or_else(|| resolve_fortran_array_expr_size(value, array_sizes, array_field_sizes))
}

fn fortran_slice_extent(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Index { object, index, .. } => match &index.kind {
            ExprKind::Slice { lower, upper, .. } => {
                let lower = lower.as_deref().cloned().unwrap_or_else(|| Expression::int(0));
                upper.as_deref().cloned().map(|upper| Expression::new(ExprKind::Binary {
                    left: Box::new(upper),
                    op: BinOp::Sub,
                    right: Box::new(lower),
                }))
            }
            _ => fortran_slice_extent(object),
        },
        _ => None,
    }
}

fn contains_fortran_slice(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Index { object, index, .. } => {
            matches!(index.kind, ExprKind::Slice { .. }) || contains_fortran_slice(object)
        }
        ExprKind::Member { object, .. } => contains_fortran_slice(object),
        _ => false,
    }
}

fn slice_target_is_known_fortran_array(
    expr: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => array_sizes.contains_key(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_field_sizes.contains_key(&field.to_ascii_lowercase()),
        ExprKind::Index { object, index, .. } => {
            matches!(index.kind, ExprKind::Slice { .. })
                && slice_target_is_known_fortran_array(object, array_sizes, array_field_sizes)
        }
        _ => false,
    }
}

fn lower_fortran_array_assignment_target(target: &Expression, loop_index: &Expression) -> Expression {
    match &target.kind {
        ExprKind::Index { object, index, null_safe } => match &index.kind {
            ExprKind::Slice { lower, .. } => {
                let base_index = lower.as_deref().cloned().unwrap_or_else(|| Expression::int(0));
                Expression::new(ExprKind::Index {
                    object: object.clone(),
                    index: Box::new(Expression::new(ExprKind::Binary {
                        left: Box::new(base_index),
                        op: BinOp::Add,
                        right: Box::new(loop_index.clone()),
                    })),
                    null_safe: *null_safe,
                })
            }
            _ => Expression::new(ExprKind::Index {
                object: Box::new(lower_fortran_array_assignment_target(object, loop_index)),
                index: index.clone(),
                null_safe: *null_safe,
            }),
        },
        _ => Expression::new(ExprKind::Index {
            object: Box::new(target.clone()),
            index: Box::new(loop_index.clone()),
            null_safe: false,
        }),
    }
}

fn lower_fortran_array_assignment_value(
    expr: &Expression,
    loop_index: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Slice { .. } | ExprKind::Array(_) | ExprKind::Call { .. } => {
            if matches!(expr.kind, ExprKind::Array(_))
                || matches!(&expr.kind, ExprKind::Member { field, .. } if array_fields.contains(&field.to_ascii_lowercase()))
                || expr_is_known_fortran_array(expr, array_sizes, array_field_sizes, array_functions)
            {
                lower_fortran_array_assignment_target(expr, loop_index)
            } else {
                expr.clone()
            }
        }
        ExprKind::Index { object, index, null_safe } => Expression::new(ExprKind::Index {
            object: Box::new(lower_fortran_array_assignment_value(
                object,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
            index: Box::new(lower_fortran_array_assignment_value(
                index,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
            null_safe: *null_safe,
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(lower_fortran_array_assignment_value(
                left,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
            right: Box::new(lower_fortran_array_assignment_value(
                right,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(lower_fortran_array_assignment_value(
                inner,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(lower_fortran_array_assignment_value(
                cond,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
            then: Box::new(lower_fortran_array_assignment_value(
                then,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
            else_: Box::new(lower_fortran_array_assignment_value(
                else_,
                loop_index,
                array_sizes,
                array_field_sizes,
                array_fields,
                array_functions,
            )),
        }),
        _ => expr.clone(),
    }
}

fn lower_fortran_array_return_calls(body: &mut [Statement]) {
    let mut array_sizes = HashMap::new();
    let mut array_field_sizes = HashMap::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_sizes(body, &mut array_field_sizes);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_return_calls_with_env(
        body,
        &mut array_sizes,
        &array_field_sizes,
        &array_functions,
        &HashSet::new(),
    );
}

fn lower_fortran_array_return_calls_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
    procedure_params: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_return_call_statement(
            statement,
            array_sizes,
            array_field_sizes,
            array_functions,
            procedure_params,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                }
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_return_calls_in_members(
                    members,
                    array_sizes,
                    array_field_sizes,
                    array_functions,
                );
            }
            StmtKind::FunctionDecl {
                params,
                body: function_body,
                return_type,
                is_sub,
                ..
            } => {
                let mut nested_arrays = array_sizes.clone();
                let nested_procedure_params = collect_fortran_array_procedure_params(params, array_functions);
                lower_fortran_array_return_calls_with_env(
                    function_body,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &nested_procedure_params,
                );

                if !*is_sub
                    && return_type
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    let result_type = return_type.take();
                    params.push(Param {
                        name: FORTRAN_ARRAY_RESULT_PARAM.to_string(),
                        type_hint: result_type,
                        default: None,
                        pass_by: PassBy::Out,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                    *is_sub = true;
                    rewrite_fortran_array_return_statements(function_body);
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    then_body,
                    &mut then_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        elif_body,
                        &mut elif_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = array_sizes.clone();
                if let Some(init) = init {
                    lower_fortran_array_return_calls_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                lower_fortran_array_return_calls_with_env(
                    stmts,
                    &mut loop_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        default,
                        &mut default_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    try_body,
                    &mut try_arrays,
                    array_field_sizes,
                    array_functions,
                    procedure_params,
                );
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        else_body,
                        &mut else_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    lower_fortran_array_return_calls_with_env(
                        finally,
                        &mut finally_arrays,
                        array_field_sizes,
                        array_functions,
                        procedure_params,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_return_calls_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { params, body, return_type, is_sub, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut nested_arrays = array_sizes.clone();
                let procedure_params = collect_fortran_array_procedure_params(params, array_functions);
                lower_fortran_array_return_calls_with_env(
                    body,
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &procedure_params,
                );
                if !*is_sub
                    && return_type
                        .as_deref()
                        .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    let result_type = return_type.take();
                    params.push(Param {
                        name: FORTRAN_ARRAY_RESULT_PARAM.to_string(),
                        type_hint: result_type,
                        default: None,
                        pass_by: PassBy::Out,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false,
                    });
                    *is_sub = true;
                    rewrite_fortran_array_return_statements(body);
                }
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = array_sizes.clone();
                lower_fortran_array_return_calls_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested_arrays,
                    array_field_sizes,
                    array_functions,
                    &HashSet::new(),
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_return_call_statement(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
    procedure_params: &HashSet<String>,
) {
    let StmtKind::Assign { targets, value } = &statement.kind else {
        return;
    };
    let [target] = targets.as_slice() else {
        return;
    };
    if resolve_fortran_array_target_size(target, array_sizes, array_field_sizes).is_none() {
        return;
    }
    let ExprKind::Call { callee, args, optional } = &value.kind else {
        return;
    };
    if *optional {
        return;
    }
    let ExprKind::Ident(name) = &callee.kind else {
        return;
    };
    let lower = name.to_ascii_lowercase();
    if !array_functions.contains(&lower) && !procedure_params.contains(&lower) {
        return;
    }

    let mut lowered_args = args.clone();
    lowered_args.push(Argument {
        value: target.clone(),
        name: None,
        by_ref: true,
        spread: false,
    });
    *statement = Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: callee.clone(),
        args: lowered_args,
        optional: false,
    })));
}

fn rewrite_fortran_array_return_statements(body: &mut [Statement]) {
    for statement in body.iter_mut() {
        match &mut statement.kind {
            StmtKind::Return(Some(expr)) => {
                *statement = Statement::new(StmtKind::Block(vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(FORTRAN_ARRAY_RESULT_PARAM)],
                        value: expr.clone(),
                    }),
                    Statement::new(StmtKind::Return(None)),
                ]));
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => rewrite_fortran_array_return_statements(stmts),
            StmtKind::If { then_body, elifs, else_body, .. } => {
                rewrite_fortran_array_return_statements(then_body);
                for (_, elif_body) in elifs {
                    rewrite_fortran_array_return_statements(elif_body);
                }
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                rewrite_fortran_array_return_statements(stmts);
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                if let Some(init) = init {
                    rewrite_fortran_array_return_statements(std::slice::from_mut(init.as_mut()));
                }
                rewrite_fortran_array_return_statements(stmts);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    rewrite_fortran_array_return_statements(&mut case.body);
                }
                if let Some(default) = default {
                    rewrite_fortran_array_return_statements(default);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                rewrite_fortran_array_return_statements(try_body);
                for catch in catches {
                    rewrite_fortran_array_return_statements(&mut catch.body);
                }
                if let Some(else_body) = else_body {
                    rewrite_fortran_array_return_statements(else_body);
                }
                if let Some(finally) = finally {
                    rewrite_fortran_array_return_statements(finally);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_scalar_array_assignments_with_env(
    body: &mut [Statement],
    array_sizes: &mut HashMap<String, Expression>,
    array_ranks: &mut HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_scalar_array_assignment(
            statement,
            array_sizes,
            array_ranks,
            array_field_sizes,
            array_field_ranks,
            array_functions,
        );

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(size) = declaration
                        .array_bounds
                        .as_deref()
                        .and_then(bounds_total_size_expr)
                        .or_else(|| declaration.init.as_ref().and_then(array_init_size_expr))
                    else {
                        continue;
                    };
                    array_sizes.insert(name.to_ascii_lowercase(), size);
                    if let Some(rank) = declaration.array_bounds.as_ref().map(Vec::len).filter(|rank| *rank > 0) {
                        array_ranks.insert(name.to_ascii_lowercase(), rank);
                    }
                }
            }
            StmtKind::FunctionDecl { body: function_body, .. } => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    function_body,
                    &mut nested,
                    &mut nested_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut nested,
                    &mut nested_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = array_sizes.clone();
                let mut then_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    then_body,
                    &mut then_arrays,
                    &mut then_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                for (_, elif_body) in elifs {
                    let mut elif_arrays = array_sizes.clone();
                    let mut elif_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        elif_body,
                        &mut elif_arrays,
                        &mut elif_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = array_sizes.clone();
                let mut loop_ranks = array_ranks.clone();
                if let Some(init) = init {
                    lower_fortran_scalar_array_assignments_with_env(
                        std::slice::from_mut(init.as_mut()),
                        &mut loop_arrays,
                        &mut loop_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                lower_fortran_scalar_array_assignments_with_env(
                    stmts,
                    &mut loop_arrays,
                    &mut loop_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = array_sizes.clone();
                    let mut case_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        &mut case.body,
                        &mut case_arrays,
                        &mut case_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(default) = default {
                    let mut default_arrays = array_sizes.clone();
                    let mut default_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        default,
                        &mut default_arrays,
                        &mut default_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = array_sizes.clone();
                let mut try_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    try_body,
                    &mut try_arrays,
                    &mut try_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
                for catch in catches {
                    let mut catch_arrays = array_sizes.clone();
                    let mut catch_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        &mut catch.body,
                        &mut catch_arrays,
                        &mut catch_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = array_sizes.clone();
                    let mut else_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        else_body,
                        &mut else_arrays,
                        &mut else_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = array_sizes.clone();
                    let mut finally_ranks = array_ranks.clone();
                    lower_fortran_scalar_array_assignments_with_env(
                        finally,
                        &mut finally_arrays,
                        &mut finally_ranks,
                        array_field_sizes,
                        array_field_ranks,
                        array_functions,
                    );
                }
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_scalar_array_assignments_in_members(
                    members,
                    array_sizes,
                    array_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn lower_fortran_scalar_array_assignments_in_members(
    members: &mut [ClassMember],
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                let mut nested = array_sizes.clone();
                let mut nested_ranks = array_ranks.clone();
                lower_fortran_scalar_array_assignments_with_env(
                    std::slice::from_mut(stmt.as_mut()),
                    &mut nested,
                    &mut nested_ranks,
                    array_field_sizes,
                    array_field_ranks,
                    array_functions,
                );
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_scalar_array_assignment(
    statement: &mut Statement,
    array_sizes: &HashMap<String, Expression>,
    array_ranks: &HashMap<String, usize>,
    array_field_sizes: &HashMap<String, Expression>,
    array_field_ranks: &HashMap<String, usize>,
    array_functions: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Assign { targets, value } => {
            if expr_is_known_fortran_array(value, array_sizes, array_field_sizes, array_functions) {
                return;
            }
            if !should_broadcast_fortran_array_value(value) {
                return;
            }
            if let Some(rank) = targets
                .iter()
                .find_map(|target| resolve_fortran_array_target_rank(target, array_ranks, array_field_ranks))
            {
                if rank > 1 {
                    if let Some(target) = targets.first() {
                        *value = build_fortran_nested_array_broadcast(target.clone(), rank, value.clone(), 0);
                        return;
                    }
                }
            }
            let Some(size) = targets
                .iter()
                .find_map(|target| resolve_fortran_array_target_size(target, array_sizes, array_field_sizes))
            else {
                return;
            };
            *value = build_fortran_array_fill(size, value.clone());
        }
        _ => {}
    }
}

fn should_broadcast_fortran_array_value(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::Member { .. } => true,
        ExprKind::Unary { expr, .. } => should_broadcast_fortran_array_value(expr),
        ExprKind::Binary { left, right, .. } => {
            should_broadcast_fortran_array_value(left)
                && should_broadcast_fortran_array_value(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            should_broadcast_fortran_array_value(cond)
                && should_broadcast_fortran_array_value(then)
                && should_broadcast_fortran_array_value(else_)
        }
        _ => false,
    }
}

fn resolve_fortran_array_target_size(
    target: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
) -> Option<Expression> {
    match &target.kind {
        ExprKind::Ident(name) => array_sizes.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Member { field, .. } => array_field_sizes.get(&field.to_ascii_lowercase()).cloned(),
        _ => None,
    }
}

fn resolve_fortran_array_target_rank(
    target: &Expression,
    array_ranks: &HashMap<String, usize>,
    array_field_ranks: &HashMap<String, usize>,
) -> Option<usize> {
    match &target.kind {
        ExprKind::Ident(name) => array_ranks.get(&name.to_ascii_lowercase()).copied(),
        ExprKind::Member { field, .. } => array_field_ranks.get(&field.to_ascii_lowercase()).copied(),
        _ => None,
    }
}

fn expr_is_known_fortran_array(
    expr: &Expression,
    array_sizes: &HashMap<String, Expression>,
    array_field_sizes: &HashMap<String, Expression>,
    array_functions: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => array_sizes.contains_key(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_field_sizes.contains_key(&field.to_ascii_lowercase()),
        ExprKind::Array(_) => true,
        ExprKind::Binary { left, right, .. } => {
            expr_is_known_fortran_array(left, array_sizes, array_field_sizes, array_functions)
                || expr_is_known_fortran_array(right, array_sizes, array_field_sizes, array_functions)
        }
        ExprKind::Unary { expr: inner, .. } => {
            expr_is_known_fortran_array(inner, array_sizes, array_field_sizes, array_functions)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_is_known_fortran_array(cond, array_sizes, array_field_sizes, array_functions)
                || expr_is_known_fortran_array(then, array_sizes, array_field_sizes, array_functions)
                || expr_is_known_fortran_array(else_, array_sizes, array_field_sizes, array_functions)
        }
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                name.eq_ignore_ascii_case("Array")
                    || array_functions.contains(&name.to_ascii_lowercase())
            }
            ExprKind::Member { field, .. } => matches!(field.to_ascii_lowercase().as_str(), "map" | "filter" | "flatmap"),
            _ => false,
        },
        ExprKind::Slice { .. } => true,
        _ => false,
    }
}

fn lower_fortran_array_expressions(body: &mut [Statement]) {
    let mut arrays = HashSet::new();
    let mut array_fields = HashSet::new();
    let mut array_functions = HashSet::new();
    collect_fortran_array_field_names(body, &mut array_fields);
    collect_fortran_array_function_names(body, &mut array_functions);
    lower_fortran_array_expressions_with_env(body, &mut arrays, &array_fields, &array_functions);
}

fn lower_fortran_array_expressions_with_env(
    body: &mut [Statement],
    arrays: &mut HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for statement in body.iter_mut() {
        rewrite_fortran_array_expressions_in_statement(statement, arrays, array_fields, array_functions);

        if let StmtKind::VarDecl { declarations, .. } = &statement.kind {
            for declaration in declarations {
                let BindingPattern::Ident(name) = &declaration.pattern else {
                    continue;
                };
                if declaration.array_bounds.is_some()
                    || declaration.init.as_ref().is_some_and(is_array_initializer_expr)
                    || declaration.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                {
                    arrays.insert(name.to_ascii_lowercase());
                }
            }
        }

        match &mut statement.kind {
            StmtKind::FunctionDecl { params, body, name, return_type, .. } => {
                let mut nested = arrays.clone();
                for param in params {
                    if param.type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                    {
                        nested.insert(param.name.to_ascii_lowercase());
                    }
                }
                let mut nested_functions = array_functions.clone();
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    nested_functions.insert(name.to_ascii_lowercase());
                }
                lower_fortran_array_expressions_with_env(body, &mut nested, array_fields, &nested_functions);
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. }
            | StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. }
            | StmtKind::NamespaceDecl { body: stmts, .. } => {
                let mut nested = arrays.clone();
                lower_fortran_array_expressions_with_env(stmts, &mut nested, array_fields, array_functions);
            }
            StmtKind::ModuleDecl { members, .. }
            | StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. } => {
                lower_fortran_array_expressions_in_members(members, arrays, array_fields, array_functions);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_arrays = arrays.clone();
                lower_fortran_array_expressions_with_env(then_body, &mut then_arrays, array_fields, array_functions);
                for (_, elif_body) in elifs {
                    let mut elif_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(elif_body, &mut elif_arrays, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(else_body, &mut else_arrays, array_fields, array_functions);
                }
            }
            StmtKind::While { body: stmts, else_body, .. }
            | StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_arrays = arrays.clone();
                lower_fortran_array_expressions_with_env(stmts, &mut loop_arrays, array_fields, array_functions);
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(else_body, &mut else_arrays, array_fields, array_functions);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_arrays = arrays.clone();
                if let Some(init) = init {
                    lower_fortran_array_expressions_with_env(std::slice::from_mut(init.as_mut()), &mut loop_arrays, array_fields, array_functions);
                }
                lower_fortran_array_expressions_with_env(stmts, &mut loop_arrays, array_fields, array_functions);
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(&mut case.body, &mut case_arrays, array_fields, array_functions);
                }
                if let Some(default) = default {
                    let mut default_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(default, &mut default_arrays, array_fields, array_functions);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_arrays = arrays.clone();
                lower_fortran_array_expressions_with_env(try_body, &mut try_arrays, array_fields, array_functions);
                for catch in catches {
                    let mut catch_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(&mut catch.body, &mut catch_arrays, array_fields, array_functions);
                }
                if let Some(else_body) = else_body {
                    let mut else_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(else_body, &mut else_arrays, array_fields, array_functions);
                }
                if let Some(finally) = finally {
                    let mut finally_arrays = arrays.clone();
                    lower_fortran_array_expressions_with_env(finally, &mut finally_arrays, array_fields, array_functions);
                }
            }
            _ => {}
        }
    }
}

fn lower_fortran_array_expressions_in_members(
    members: &mut [ClassMember],
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { params, body, name, return_type, .. } = &mut stmt.kind else {
                    continue;
                };
                let mut method_arrays = arrays.clone();
                for param in params.iter() {
                    if param.type_hint.as_deref().is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                        || param.type_hint.as_deref().is_some_and(is_fortran_string_type_hint)
                    {
                        method_arrays.insert(param.name.to_ascii_lowercase());
                    }
                }
                let mut method_functions = array_functions.clone();
                if return_type
                    .as_deref()
                    .is_some_and(|type_hint| type_hint.trim_end().ends_with("()"))
                {
                    method_functions.insert(name.to_ascii_lowercase());
                }
                lower_fortran_array_expressions_with_env(body, &mut method_arrays, array_fields, &method_functions);
            }
            ClassMember::NestedType(stmt) => {
                let mut nested_arrays = arrays.clone();
                lower_fortran_array_expressions_with_env(std::slice::from_mut(stmt.as_mut()), &mut nested_arrays, array_fields, array_functions);
            }
            _ => {}
        }
    }
}

fn rewrite_fortran_array_expressions_in_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            if !is_fortran_allocator_intrinsic_expr(expr) {
                rewrite_fortran_array_expressions_in_expr(expr, arrays, array_fields, array_functions);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_fortran_array_expressions_in_expr(target, arrays, array_fields, array_functions);
            }
            rewrite_fortran_array_expressions_in_expr(value, arrays, array_fields, array_functions);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_fortran_array_expressions_in_expr(target, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(value, arrays, array_fields, array_functions);
        }
        StmtKind::Return(Some(expr)) => rewrite_fortran_array_expressions_in_expr(expr, arrays, array_fields, array_functions),
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_fortran_array_expressions_in_expr(expr, arrays, array_fields, array_functions);
            }
            if let Some(cause) = cause {
                rewrite_fortran_array_expressions_in_expr(cause, arrays, array_fields, array_functions);
            }
        }
        StmtKind::If { cond, .. }
        | StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. } => {
            rewrite_fortran_array_expressions_in_expr(cond, arrays, array_fields, array_functions);
        }
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_fortran_array_expressions_in_expr(cond, arrays, array_fields, array_functions);
            }
            if let Some(update) = update {
                rewrite_fortran_array_expressions_in_expr(update, arrays, array_fields, array_functions);
            }
        }
        StmtKind::ForIn { iter, .. }
        | StmtKind::Using { resource: iter, .. }
        | StmtKind::Lock { expr: iter, .. }
        | StmtKind::Switch { expr: iter, .. } => {
            rewrite_fortran_array_expressions_in_expr(iter, arrays, array_fields, array_functions);
        }
        _ => {}
    }
}

fn rewrite_fortran_array_expressions_in_expr(
    expr: &mut Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            rewrite_fortran_array_expressions_in_expr(left, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(right, arrays, array_fields, array_functions);
            if let Some(lowered) = lower_fortran_array_binary_expr(*op, left.as_ref(), right.as_ref(), arrays, array_fields, array_functions) {
                *expr = lowered;
            }
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner) => rewrite_fortran_array_expressions_in_expr(inner, arrays, array_fields, array_functions),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_fortran_array_expressions_in_expr(cond, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(then, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(else_, arrays, array_fields, array_functions);
        }
        ExprKind::Member { object, .. } => rewrite_fortran_array_expressions_in_expr(object, arrays, array_fields, array_functions),
        ExprKind::Index { object, index, .. } => {
            rewrite_fortran_array_expressions_in_expr(object, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(index, arrays, array_fields, array_functions);
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_fortran_array_expressions_in_expr(lower, arrays, array_fields, array_functions);
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_fortran_array_expressions_in_expr(upper, arrays, array_fields, array_functions);
            }
            if let Some(step) = step.as_mut() {
                rewrite_fortran_array_expressions_in_expr(step, arrays, array_fields, array_functions);
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_fortran_array_expressions_in_expr(callee, arrays, array_fields, array_functions);
            for arg in args.iter_mut() {
                rewrite_fortran_array_expressions_in_expr(&mut arg.value, arrays, array_fields, array_functions);
            }
            if let Some(lowered) = lower_fortran_array_intrinsic_expr(callee.as_ref(), args, arrays, array_fields, array_functions) {
                *expr = lowered;
            }
        }
        ExprKind::New { class, args } => {
            rewrite_fortran_array_expressions_in_expr(class, arrays, array_fields, array_functions);
            for arg in args.iter_mut() {
                rewrite_fortran_array_expressions_in_expr(&mut arg.value, arrays, array_fields, array_functions);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_fortran_array_expressions_in_expr(target, arrays, array_fields, array_functions);
            rewrite_fortran_array_expressions_in_expr(value, arrays, array_fields, array_functions);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = item.key.as_mut() {
                    rewrite_fortran_array_expressions_in_expr(key, arrays, array_fields, array_functions);
                }
                rewrite_fortran_array_expressions_in_expr(&mut item.value, arrays, array_fields, array_functions);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_fortran_array_expressions_in_expr(item, arrays, array_fields, array_functions);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_fortran_array_expressions_in_expr(key, arrays, array_fields, array_functions);
                        rewrite_fortran_array_expressions_in_expr(value, arrays, array_fields, array_functions);
                    }
                    ObjectProperty::Spread(expr) => rewrite_fortran_array_expressions_in_expr(expr, arrays, array_fields, array_functions),
                    ObjectProperty::Shorthand(_) | ObjectProperty::Method { .. } | ObjectProperty::Accessor { .. } => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    rewrite_fortran_array_expressions_in_expr(expr, arrays, array_fields, array_functions);
                }
            }
        }
        _ => {}
    }
}

fn lower_fortran_array_binary_expr(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Option<Expression> {
    if !matches!(op, BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Div | BinOp::Pow) {
        return None;
    }
    let left_is_array = is_known_fortran_array_expr(left, arrays, array_fields, array_functions);
    let right_is_array = is_known_fortran_array_expr(right, arrays, array_fields, array_functions);
    if !left_is_array && !right_is_array {
        return None;
    }

    let item_name = "__fortran_array_item";
    let index_name = "__fortran_array_index";
    let lambda_body = if left_is_array && right_is_array {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(Expression::ident(item_name)),
            right: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(right.clone()),
                index: Box::new(Expression::ident(index_name)),
                null_safe: false,
            })),
        })
    } else if left_is_array {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(Expression::ident(item_name)),
            right: Box::new(right.clone()),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left.clone()),
            right: Box::new(Expression::ident(item_name)),
        })
    };

    Some(build_fortran_array_map(
        if left_is_array { left.clone() } else { right.clone() },
        lambda_body,
        left_is_array && right_is_array,
        item_name,
        index_name,
    ))
}

fn lower_fortran_array_intrinsic_expr(
    callee: &Expression,
    args: &[Argument],
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if args.len() != 1 || args[0].name.is_some() {
        return None;
    }
    let array_expr = &args[0].value;
    if !is_known_fortran_array_expr(array_expr, arrays, array_fields, array_functions) {
        return None;
    }

    let lowered = name.to_ascii_lowercase();
    let target = match lowered.as_str() {
        "minval" => "min",
        "maxval" => "max",
        _ => return None,
    };
    let acc_name = format!("__fortran_{}_acc", target);
    let item_name = format!("__fortran_{}_item", target);
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_expr.clone()),
            field: "reduce".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params: vec![
                Param {
                    name: acc_name.clone(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
                Param {
                    name: item_name.clone(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
            ],
            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(target)),
                args: vec![
                    Argument::positional(Expression::ident(&acc_name)),
                    Argument::positional(Expression::ident(&item_name)),
                ],
                optional: false,
            }))),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    }))
}

fn build_fortran_array_map(
    array_expr: Expression,
    body: Expression,
    include_index: bool,
    item_name: &str,
    index_name: &str,
) -> Expression {
    let mut params = vec![Param {
        name: item_name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }];
    if include_index {
        params.push(Param {
            name: index_name.to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(array_expr),
            field: "map".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Expr(Box::new(body)),
            is_async: false,
            captures: Vec::new(),
        }))],
        optional: false,
    })
}

fn is_known_fortran_array_expr(
    expr: &Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
    array_functions: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => arrays.contains(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_fields.contains(&field.to_ascii_lowercase()),
        ExprKind::Array(_) | ExprKind::Slice { .. } => true,
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => {
                matches!(name.to_ascii_lowercase().as_str(), "array")
                    || array_functions.contains(&name.to_ascii_lowercase())
            }
            ExprKind::Member { field, .. } => matches!(field.to_ascii_lowercase().as_str(), "map" | "filter" | "flatmap"),
            _ => false,
        },
        _ => false,
    }
}

fn build_fortran_array_fill(size: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![Argument::positional(size), Argument::positional(value)],
        optional: false,
    })
}

fn build_fortran_nested_array_broadcast(
    array_expr: Expression,
    rank: usize,
    value: Expression,
    depth: usize,
) -> Expression {
    if rank <= 1 {
        return build_fortran_array_map(
            array_expr,
            value,
            false,
            &format!("__fortran_broadcast_item_{depth}"),
            &format!("__fortran_broadcast_index_{depth}"),
        );
    }

    let item_name = format!("__fortran_broadcast_item_{depth}");
    build_fortran_array_map(
        array_expr,
        build_fortran_nested_array_broadcast(Expression::ident(&item_name), rank - 1, value, depth + 1),
        false,
        &item_name,
        &format!("__fortran_broadcast_index_{depth}"),
    )
}

fn collect_fortran_array_field_ranks(body: &[Statement], array_field_ranks: &mut HashMap<String, usize>) {
    for statement in body {
        match &statement.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                collect_fortran_array_field_ranks_in_members(members, array_field_ranks);
            }
            StmtKind::ModuleDecl { members, .. } => {
                collect_fortran_array_field_ranks_in_members(members, array_field_ranks);
            }
            StmtKind::NamespaceDecl { body, .. } => collect_fortran_array_field_ranks(body, array_field_ranks),
            _ => {}
        }
    }
}

fn collect_fortran_array_field_ranks_in_members(
    members: &[ClassMember],
    array_field_ranks: &mut HashMap<String, usize>,
) {
    for member in members {
        match member {
            ClassMember::Field { name, array_bounds, .. } => {
                if let Some(rank) = array_bounds.as_ref().map(Vec::len).filter(|rank| *rank > 0) {
                    array_field_ranks.insert(name.to_ascii_lowercase(), rank);
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_fortran_array_field_ranks(std::slice::from_ref(stmt.as_ref()), array_field_ranks);
            }
            _ => {}
        }
    }
}

fn bounds_total_size_expr(bounds: &[Expression]) -> Option<Expression> {
    let mut iter = bounds.iter().cloned();
    let first = iter.next()?;
    Some(iter.fold(first, |acc, bound| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Mul,
            left: Box::new(acc),
            right: Box::new(bound),
        })
    }))
}

fn array_init_size_expr(init: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &init.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !name.eq_ignore_ascii_case("Array") {
        return None;
    }
    args.first().map(|arg| arg.value.clone())
}

fn rewrite_array_subscripts_in_statement(
    statement: &mut Statement,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut statement.kind {
        StmtKind::Expr(expr) => {
            let preserve_intrinsic_args = matches!(
                &expr.kind,
                ExprKind::Call { callee, .. }
                    if matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("allocate") || name.eq_ignore_ascii_case("deallocate"))
            );
            if !preserve_intrinsic_args {
                rewrite_array_subscripts_in_expr(expr, arrays, callables, array_fields);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_array_subscripts_in_expr(target, arrays, callables, array_fields);
            }
            rewrite_array_subscripts_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_array_subscripts_in_expr(target, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(value, arrays, callables, array_fields);
        }
        StmtKind::Return(Some(expr)) => rewrite_array_subscripts_in_expr(expr, arrays, callables, array_fields),
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_array_subscripts_in_expr(expr, arrays, callables, array_fields);
            }
            if let Some(cause) = cause {
                rewrite_array_subscripts_in_expr(cause, arrays, callables, array_fields);
            }
        }
        StmtKind::If { cond, .. }
        | StmtKind::While { cond, .. }
        | StmtKind::DoWhile { cond, .. } => rewrite_array_subscripts_in_expr(cond, arrays, callables, array_fields),
        StmtKind::For { cond, update, .. } => {
            if let Some(cond) = cond {
                rewrite_array_subscripts_in_expr(cond, arrays, callables, array_fields);
            }
            if let Some(update) = update {
                rewrite_array_subscripts_in_expr(update, arrays, callables, array_fields);
            }
        }
        StmtKind::ForIn { iter, .. }
        | StmtKind::Using { resource: iter, .. }
        | StmtKind::Lock { expr: iter, .. }
        | StmtKind::Switch { expr: iter, .. } => rewrite_array_subscripts_in_expr(iter, arrays, callables, array_fields),
        _ => {}
    }
}

fn rewrite_array_subscripts_in_expr(
    expr: &mut Expression,
    arrays: &HashSet<String>,
    callables: &HashSet<String>,
    array_fields: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_array_subscripts_in_expr(left, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(right, arrays, callables, array_fields);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::TypeOf(inner) => rewrite_array_subscripts_in_expr(inner, arrays, callables, array_fields),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_array_subscripts_in_expr(cond, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(then, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(else_, arrays, callables, array_fields);
        }
        ExprKind::Member { object, .. } => rewrite_array_subscripts_in_expr(object, arrays, callables, array_fields),
        ExprKind::Index { object, index, .. } => {
            rewrite_array_subscripts_in_expr(object, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(index, arrays, callables, array_fields);
            *index = Box::new(normalize_array_index_operand(index.as_ref().clone(), FORTRAN_ARRAY_INDEXING));
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_array_subscripts_in_expr(lower, arrays, callables, array_fields);
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_array_subscripts_in_expr(upper, arrays, callables, array_fields);
            }
            if let Some(step) = step.as_mut() {
                rewrite_array_subscripts_in_expr(step, arrays, callables, array_fields);
            }
        }
        ExprKind::Call { callee, args, optional } => {
            rewrite_array_subscripts_in_expr(callee, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_array_subscripts_in_expr(&mut arg.value, arrays, callables, array_fields);
            }
            if !args.is_empty()
                && !*optional
                && !is_known_fortran_callable(callee, callables)
                && (is_known_fortran_array(callee, arrays, array_fields)
                    || matches!(&callee.kind, ExprKind::Index { .. }))
            {
                expr.kind = build_fortran_index_chain(callee.as_ref().clone(), args);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_array_subscripts_in_expr(class, arrays, callables, array_fields);
            for arg in args.iter_mut() {
                rewrite_array_subscripts_in_expr(&mut arg.value, arrays, callables, array_fields);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_array_subscripts_in_expr(target, arrays, callables, array_fields);
            rewrite_array_subscripts_in_expr(value, arrays, callables, array_fields);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    rewrite_array_subscripts_in_expr(key, arrays, callables, array_fields);
                }
                rewrite_array_subscripts_in_expr(&mut item.value, arrays, callables, array_fields);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_array_subscripts_in_expr(item, arrays, callables, array_fields);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_array_subscripts_in_expr(key, arrays, callables, array_fields);
                        rewrite_array_subscripts_in_expr(value, arrays, callables, array_fields);
                    }
                    ObjectProperty::Spread(expr) => rewrite_array_subscripts_in_expr(expr, arrays, callables, array_fields),
                    ObjectProperty::Computed { key, value } => {
                        rewrite_array_subscripts_in_expr(key, arrays, callables, array_fields);
                        rewrite_array_subscripts_in_expr(value, arrays, callables, array_fields);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => rewrite_array_subscripts_in_expr(expr, arrays, callables, array_fields),
                    InterpolPart::Text(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn build_fortran_index_chain(object: Expression, args: &[Argument]) -> ExprKind {
    let mut current = object;
    for arg in args {
        current = Expression::new(ExprKind::Index {
            object: Box::new(current),
            index: Box::new(normalize_array_index_operand(arg.value.clone(), FORTRAN_ARRAY_INDEXING)),
            null_safe: false,
        });
    }
    current.kind
}

fn is_known_fortran_array(
    expr: &Expression,
    arrays: &HashSet<String>,
    array_fields: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => arrays.contains(&name.to_ascii_lowercase()),
        ExprKind::Member { field, .. } => array_fields.contains(&field.to_ascii_lowercase()),
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => matches!(name.to_ascii_lowercase().as_str(), "str_split" | "str_getcsv" | "array"),
            _ => false,
        },
        ExprKind::Index { .. } => true,
        _ => false,
    }
}

fn is_known_fortran_callable(expr: &Expression, callables: &HashSet<String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => callables.contains(&name.to_ascii_lowercase()),
        _ => false,
    }
}

fn is_fortran_callable_type_hint(type_hint: &str) -> bool {
    let lower = type_hint.trim().to_ascii_lowercase();
    lower.starts_with("procedure(") || lower == "procedure"
}

fn fortran_callable_signature_name(type_hint: &str) -> Option<String> {
    type_hint
        .trim()
        .strip_prefix("procedure(")
        .and_then(|rest| rest.strip_suffix(')'))
        .map(|name| name.trim().to_ascii_lowercase())
}

fn collect_fortran_array_procedure_params(
    params: &[Param],
    array_functions: &HashSet<String>,
) -> HashSet<String> {
    params
        .iter()
        .filter_map(|param| {
            if param.type_hint.is_none() && matches!(param.pass_by, PassBy::Value) {
                return Some(param.name.to_ascii_lowercase());
            }

            param.type_hint
                .as_deref()
                .and_then(fortran_callable_signature_name)
                .filter(|signature_name| array_functions.contains(signature_name))
                .map(|_| param.name.to_ascii_lowercase())
        })
        .collect()
}

fn is_fortran_string_type_hint(type_hint: &str) -> bool {
    let lower = type_hint.trim().to_ascii_lowercase();
    lower == "character" || lower.starts_with("character(") || lower.starts_with("character*")
}

fn is_array_initializer_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Array(_) => true,
        ExprKind::Call { callee, .. } => matches!(&callee.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array")),
        _ => false,
    }
}

fn lower_fortran_body_intrinsics(params: &[Param], body: &mut Vec<Statement>) {
    let mut type_env = HashMap::new();
    for param in params {
        if let Some(type_hint) = &param.type_hint {
            type_env.insert(param.name.to_ascii_lowercase(), type_hint.clone());
        }
    }
    lower_fortran_body_intrinsics_with_env(body, &mut type_env);
}

fn lower_fortran_body_intrinsics_with_env(body: &mut [Statement], type_env: &mut HashMap<String, String>) {
    for statement in body.iter_mut() {
        if let Some(lowered) = lower_body_intrinsic_statement(statement, type_env) {
            *statement = lowered;
        }

        match &mut statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    let BindingPattern::Ident(name) = &declaration.pattern else {
                        continue;
                    };
                    let Some(type_hint) = &declaration.type_hint else {
                        continue;
                    };
                    type_env.insert(name.to_ascii_lowercase(), type_hint.clone());
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::DoWhile { body: stmts, .. } => {
                let mut nested_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut nested_env);
            }
            StmtKind::If { then_body, elifs, else_body, .. } => {
                let mut then_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(then_body, &mut then_env);
                for (_, elif_body) in elifs {
                    let mut elif_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(elif_body, &mut elif_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::While { body: stmts, else_body, .. } => {
                let mut loop_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::For { init, body: stmts, .. } => {
                let mut loop_env = type_env.clone();
                if let Some(init) = init {
                    lower_fortran_body_intrinsics_with_env(std::slice::from_mut(init.as_mut()), &mut loop_env);
                }
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
            }
            StmtKind::ForIn { body: stmts, else_body, .. } => {
                let mut loop_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut loop_env);
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    let mut case_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(&mut case.body, &mut case_env);
                }
                if let Some(default) = default {
                    let mut default_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(default, &mut default_env);
                }
            }
            StmtKind::Try { body: try_body, catches, else_body, finally } => {
                let mut try_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(try_body, &mut try_env);
                for catch in catches {
                    let mut catch_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(&mut catch.body, &mut catch_env);
                }
                if let Some(else_body) = else_body {
                    let mut else_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(else_body, &mut else_env);
                }
                if let Some(finally) = finally {
                    let mut finally_env = type_env.clone();
                    lower_fortran_body_intrinsics_with_env(finally, &mut finally_env);
                }
            }
            StmtKind::With { body: stmts, .. }
            | StmtKind::Using { body: stmts, .. }
            | StmtKind::Lock { body: stmts, .. } => {
                let mut nested_env = type_env.clone();
                lower_fortran_body_intrinsics_with_env(stmts, &mut nested_env);
            }
            _ => {}
        }
    }
}

fn lower_body_intrinsic_statement(statement: &Statement, type_env: &HashMap<String, String>) -> Option<Statement> {
    let StmtKind::Expr(expr) = &statement.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    None
}

fn lower_allocate_statement(args: &[Argument], type_env: &HashMap<String, String>) -> Option<Statement> {
    let mut assigns = Vec::new();
    for arg in args {
        let (target, value) = lower_allocate_target(&arg.value, type_env)?;
        assigns.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
        }));
    }
    Some(Statement::new(StmtKind::Block(assigns)))
}

fn lower_allocate_target(expr: &Expression, type_env: &HashMap<String, String>) -> Option<(Expression, Expression)> {
    match &expr.kind {
        ExprKind::Ident(name) => {
            let type_hint = type_env.get(&name.to_ascii_lowercase())?;
            let class_name = parse_derived_type_name(type_hint)?;
            Some((expr.clone(), Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(&class_name)),
                args: Vec::new(),
            })))
        }
        ExprKind::Call { callee: _, args, .. } => {
            let target = allocation_target_expr(expr);
            Some((target, Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("Array")),
                args: vec![
                    args.first().cloned().unwrap_or_else(|| Argument::positional(Expression::int(0))),
                    Argument::positional(Expression::null()),
                ],
                optional: false,
            })))
        }
        _ => None,
    }
}

fn allocation_target_expr(expr: &Expression) -> Expression {
    match &expr.kind {
        ExprKind::Call { callee, .. } => callee.as_ref().clone(),
        _ => expr.clone(),
    }
}

fn lower_intrinsic_expr_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let lowered = name.to_ascii_lowercase();
    match lowered.as_str() {
        "achar" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("char")),
            args: args.to_vec(),
            optional: false,
        })),
        "iachar" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("ichar")),
            args: args.to_vec(),
            optional: false,
        })),
        "associated" | "allocated" if args.len() == 1 => Some(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(args[0].value.clone()),
            right: Box::new(Expression::null()),
        })),
        "len" if args.len() == 1 => Some(Expression::new(ExprKind::Member {
            object: Box::new(args[0].value.clone()),
            field: "length".to_string(),
            null_safe: false,
        })),
        "product" if args.len() == 1 && args.iter().all(|arg| arg.name.is_none()) => {
            let acc_name = "__fortran_product_acc";
            let item_name = "__fortran_product_item";
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(args[0].value.clone()),
                    field: "reduce".to_string(),
                    null_safe: false,
                })),
                args: vec![
                    Argument::positional(Expression::new(ExprKind::Lambda {
                        params: vec![
                            Param {
                                name: acc_name.to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            },
                            Param {
                                name: item_name.to_string(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            },
                        ],
                        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Mul,
                            left: Box::new(Expression::ident(acc_name)),
                            right: Box::new(Expression::ident(item_name)),
                        }))),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    Argument::positional(Expression::int(1)),
                ],
                optional: false,
            }))
        }
        "verify" if args.len() >= 2 => {
            let mut lowered_args = args.to_vec();
            lowered_args[0] = Argument::positional(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("trim")),
                args: vec![Argument::positional(args[0].value.clone())],
                optional: false,
            }));
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("verify")),
                args: lowered_args,
                optional: false,
            }))
        }
        _ => None,
    }
}

fn walk_binop(pair: Pair<Rule>) -> Result<Expression, String> {
    let rule = pair.as_rule();
    // Capture the source slice BEFORE `into_inner()` consumes the
    // pair — `unary`'s inline `-`/`+` literals don't survive as
    // child pairs, so we have to read them from the rule's own span.
    let raw = pair.as_str().to_string();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    // Unary not — `not_op ~ comparison`. inner has [not_op, operand].
    if rule == Rule::logical_not && inner.len() == 2 {
        return Ok(Expression::new(ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(walk_expr(inner.remove(1))?) }));
    }
    // Unary minus/plus — `unary = "-" primary | "+" primary | primary`.
    // Pest doesn't emit inline-literal sign tokens as child pairs,
    // so `inner` is always [primary_expr]. Recover the sign from the
    // rule's leading source character.
    if rule == Rule::unary {
        let trimmed = raw.trim_start();
        if trimmed.starts_with('-') {
            let operand = walk_expr(inner.remove(0))?;
            return Ok(Expression::new(ExprKind::Unary {
                op: UnaryOp::Neg, expr: Box::new(operand),
            }));
        }
        // `+` is a no-op; the `primary` form is the bare value.
        return walk_expr(inner.remove(0));
    }
    // `power = { concat ~ ("**" ~ concat)? }` — inline `**` literal,
    // so inner has either [base] or [base, exponent]. Apply Pow when 2.
    if rule == Rule::power && inner.len() == 2 {
        let base = walk_expr(inner.remove(0))?;
        let exp  = walk_expr(inner.remove(0))?;
        return Ok(Expression::new(ExprKind::Binary {
            left: Box::new(base),
            op: BinOp::Pow,
            right: Box::new(exp),
        }));
    }
    // `concat = { unary ~ ("//" ~ unary)* }` — inline `//` literal,
    // so inner has [u1, u2, u3, ...] without operator pairs. Fold
    // left as Concat (BinOp::Concat).
    if rule == Rule::concat && inner.len() >= 2 {
        let mut result = walk_expr(inner.remove(0))?;
        for next in inner.into_iter() {
            let right = walk_expr(next)?;
            result = Expression::new(ExprKind::Binary {
                left: Box::new(result),
                op: BinOp::Concat,
                right: Box::new(right),
            });
        }
        return Ok(result);
    }
    if inner.len() == 1 { return walk_expr(inner.remove(0)); }
    if inner.len() >= 3 {
        let mut result = walk_expr(inner.remove(0))?;
        let mut i = 0;
        while i + 1 < inner.len() {
            let op = to_binop(&inner[i]);
            let right = walk_expr(inner[i + 1].clone())?;
            result = Expression::new(ExprKind::Binary { left: Box::new(result), op, right: Box::new(right) });
            i += 2;
        }
        return Ok(result);
    }
    if inner.is_empty() { return Ok(Expression::new(ExprKind::Lit(Literal::Null))); }
    walk_expr(inner.remove(0))
}

fn to_binop(pair: &Pair<Rule>) -> BinOp {
    // Pest's `add_op` / `mul_op` / etc. spans can include trailing
    // whitespace before the next operand (the grammar rule is
    // `add_op ~ multiplication`; pest's WHITESPACE-implicit consume
    // can land inside the op's span). Trim before matching so
    // `"- "` still maps to `Sub`, not falling through to the Add
    // default.
    match pair.as_str().to_lowercase().trim() {
        "+" => BinOp::Add, "-" => BinOp::Sub, "*" => BinOp::Mul, "/" => BinOp::Div,
        "**" => BinOp::Pow, "//" => BinOp::Concat,
        "==" | ".eq." => BinOp::Eq, "/=" | ".ne." => BinOp::NotEq,
        "<" | ".lt." => BinOp::Lt, ">" | ".gt." => BinOp::Gt,
        "<=" | ".le." => BinOp::LtEq, ">=" | ".ge." => BinOp::GtEq,
        ".and." => BinOp::And, ".or." => BinOp::Or,
        ".eqv." => BinOp::Eq, ".neqv." => BinOp::NotEq,
        _ => BinOp::Add,
    }
}

fn meaningful(pair: &Pair<Rule>) -> bool {
    !matches!(pair.as_rule(), Rule::NEWLINE | Rule::EOI)
}

fn is_expr_rule(r: Rule) -> bool {
    matches!(r, Rule::expression | Rule::logical_or | Rule::logical_and | Rule::comparison
        | Rule::addition | Rule::multiplication | Rule::power | Rule::concat | Rule::unary
        | Rule::primary_expr | Rule::literal | Rule::number_literal | Rule::string_literal
        | Rule::identifier | Rule::function_call_or_subscript | Rule::logical_literal | Rule::logical_not)
}

fn to_class_member(stmt: Statement) -> ClassMember {
    match stmt.kind {
        StmtKind::ClassDecl { .. }
        | StmtKind::StructDecl { .. }
        | StmtKind::EnumDecl { .. }
        | StmtKind::InterfaceDecl { .. }
        | StmtKind::ModuleDecl { .. } => ClassMember::NestedType(Box::new(stmt)),
        StmtKind::FunctionDecl { .. } => ClassMember::Method(Box::new(stmt)),
        StmtKind::VarDecl { ref declarations, .. } => {
            if let Some(d) = declarations.first() {
                if let BindingPattern::Ident(name) = &d.pattern {
                    let field_type_hint = d.type_hint.as_ref().map(|type_hint| {
                        if d.array_bounds.is_some() && !type_hint.trim_end().ends_with("()") {
                            format!("{}()", type_hint.trim())
                        } else {
                            type_hint.clone()
                        }
                    });
                    return ClassMember::Field {
                        name: name.clone(), type_hint: field_type_hint, init: d.init.clone(),
                        modifiers: Modifiers::default(), with_events: false, array_bounds: d.array_bounds.clone(),
                    };
                }
            }
            ClassMember::Method(Box::new(stmt))
        }
        _ => ClassMember::Method(Box::new(stmt)),
    }
}
