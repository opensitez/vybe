//! PHP walker — pest `Pair<Rule>` → `vybex::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//! Once this returns a `Module`, the rest of the compilation pipeline
//! (compile_class / compile_expression / etc.) is shared with every
//! other vybex language and works without any PHP-specific knowledge.
//!
//! ## Notes on PHP semantics that the walker normalises
//!
//! - **`$variable`** vs **bare identifier**. PHP distinguishes them at
//!   the lexical level: `$x` is a variable, `x` is a function name or a
//!   constant. The walker emits `Ident` for both kinds — for variables
//!   we strip the leading `$` so the canonical AST identifier matches
//!   what every other language uses.
//!
//! - **`echo` and `print`** become `StmtKind::Echo(...)` directly, which
//!   the compiler routes through `compiler_common::io::emit_print`.
//!
//! - **`$obj->method()`** is `Call { callee: Member { object, field } }`
//!   exactly like JS `obj.method()`. PHP `?->` becomes `Member {
//!   null_safe: true }`.
//!
//! - **`Class::method()` / `Class::CONST`** uses `ExprKind::StaticAccess`
//!   which the compiler treats as a struct_get on the class global.
//!
//! - **`use Foo\Bar;`** and **`namespace Foo\Bar;`** are parsed but
//!   discarded — the compiler treats every name as a flat global. PHP
//!   namespaces are mostly cosmetic for our purposes.
//!
//! - **Type hints** are parsed and discarded. We don't type-check.
//!
//! - **Promoted constructor parameters** (PHP 8): `public int $foo` in
//!   the constructor parameter list. The walker emits the param AND
//!   synthesises a property + an assignment in the body so the
//!   downstream compiler doesn't need to know about the promotion.
//!
//! - **`<?php` open tag**: stripped at the grammar level (`open_tag` is
//!   silent). User scripts may or may not have it.

use pest::Parser;
use pest::iterators::Pair;
use crate::ast::*;
use super::{PhpParser, Rule};

/// Returns true for `kw_*` token rules. Pest preserves atomic rule
/// nodes as siblings inside their parent rule's parse tree, so without
/// this filter the keyword tokens leak into walker positional indexing
/// (e.g. `if (...)` would land `kw_if` as the first child of
/// `if_statement` and walk_if would try to parse it as an expression).
fn is_kw(r: Rule) -> bool {
    matches!(r,
        Rule::kw_if | Rule::kw_elseif | Rule::kw_else_if | Rule::kw_else
        | Rule::kw_while | Rule::kw_do | Rule::kw_for | Rule::kw_foreach
        | Rule::kw_as | Rule::kw_switch | Rule::kw_case | Rule::kw_default
        | Rule::kw_break | Rule::kw_continue | Rule::kw_return
        | Rule::kw_function | Rule::kw_class | Rule::kw_extends
        | Rule::kw_implements | Rule::kw_interface | Rule::kw_trait
        | Rule::kw_enum | Rule::kw_new | Rule::kw_clone | Rule::kw_echo
        | Rule::kw_print | Rule::kw_null | Rule::kw_true | Rule::kw_false
        | Rule::kw_instanceof | Rule::kw_throw | Rule::kw_try | Rule::kw_catch
        | Rule::kw_finally | Rule::kw_static | Rule::kw_public | Rule::kw_private
        | Rule::kw_protected | Rule::kw_abstract | Rule::kw_final | Rule::kw_const
        | Rule::kw_match | Rule::kw_fn | Rule::kw_use | Rule::kw_namespace
        | Rule::kw_yield_from | Rule::kw_yield | Rule::kw_list | Rule::kw_global
        | Rule::kw_readonly | Rule::kw_and | Rule::kw_or | Rule::kw_xor
        | Rule::kw_self | Rule::kw_parent | Rule::kw_isset | Rule::kw_empty
        | Rule::kw_unset | Rule::kw_endif | Rule::kw_endwhile | Rule::kw_endfor
        | Rule::kw_endforeach | Rule::kw_endswitch
    )
}

/// `pair.into_inner()` with `kw_*` siblings stripped — use this in any
/// walker that does positional indexing on a rule body that includes
/// keywords.
fn inner_nokw(pair: Pair<Rule>) -> std::vec::IntoIter<Pair<Rule>> {
    let kept: Vec<Pair<Rule>> = pair.into_inner().filter(|p| !is_kw(p.as_rule())).collect();
    kept.into_iter()
}

pub fn parse(source: &str) -> Result<Module, String> {
    let mut pairs = PhpParser::parse(Rule::program, source)
        .map_err(|e| format!("PHP parse error: {}", e))?;
    let program = pairs.next().ok_or("empty parse")?;

    let mut body = Vec::new();
    for pair in program.into_inner() {
        if matches!(pair.as_rule(), Rule::EOI) { continue; }
        if let Some(stmt) = walk_statement(pair)? {
            body.push(stmt);
        }
    }

    Ok(Module {
        name: String::new(),
        language: Lang::PHP,
        body,
        imports: Vec::new(),
    })
}

// ─── Statements ────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,

        Rule::block_statement => {
            let inner = pair.into_inner();
            let mut stmts = Vec::new();
            for s in inner {
                if let Some(st) = walk_statement(s)? {
                    stmts.push(st);
                }
            }
            StmtKind::Block(stmts)
        }

        Rule::echo_statement | Rule::print_statement => {
            let exprs: Result<Vec<Expression>, String> = pair
                .into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .collect();
            StmtKind::Echo(exprs?)
        }

        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().unwrap())?;
            StmtKind::Expr(expr)
        }

        Rule::const_statement => {
            // const NAME = expr;
            let mut inner = inner_nokw(pair);
            let name = inner.next().unwrap().as_str().to_string();
            let value = walk_expression(inner.next().unwrap())?;
            StmtKind::VarDecl {
                kind: VarDeclKind::Const,
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: None,
                    init: Some(value),
                    array_bounds: None,
                    with_events: false,
                }],
            }
        }

        Rule::global_statement => {
            // global $a, $b;  → ScopeDecl { Global, names }
            let names: Vec<String> = pair.into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::variable))
                .map(|p| strip_dollar(p.as_str()).to_string())
                .collect();
            StmtKind::ScopeDecl { kind: ScopeDeclKind::Global, names }
        }

        Rule::namespace_statement => {
            // namespace Foo\Bar; or namespace Foo\Bar { ... }
            // We honour the form but flatten the body — PHP namespace
            // resolution is otherwise cosmetic for our compilation.
            let mut name = String::new();
            let mut body = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::qualified_name => name = p.as_str().to_string(),
                    Rule::block_statement => {
                        for s in p.into_inner() {
                            if let Some(st) = walk_statement(s)? {
                                body.push(st);
                            }
                        }
                    }
                    _ => {}
                }
            }
            // For the bare `namespace Foo;` form, just discard — return Empty.
            if body.is_empty() {
                StmtKind::Empty
            } else {
                StmtKind::NamespaceDecl { name, body }
            }
        }

        Rule::use_statement => {
            // `use Foo\Bar;` / `use function Foo\bar;` — discard.
            // PHP `use` is for namespace resolution, which we flatten.
            StmtKind::Empty
        }

        Rule::function_declaration => walk_function_decl(pair)?,

        Rule::class_declaration => walk_class_decl(pair)?,

        Rule::interface_declaration => walk_interface_decl(pair)?,

        Rule::trait_declaration => walk_trait_decl(pair)?,

        Rule::enum_declaration => walk_enum_decl(pair)?,

        Rule::if_statement => walk_if(pair)?,

        Rule::while_statement => {
            let mut inner = inner_nokw(pair);
            let cond = walk_expression(inner.next().unwrap())?;
            let body = walk_statement_into_body(inner.next().unwrap())?;
            StmtKind::While { cond, body, else_body: None }
        }

        Rule::do_while_statement => {
            let mut inner = inner_nokw(pair);
            let body = walk_statement_into_body(inner.next().unwrap())?;
            let cond = walk_expression(inner.next().unwrap())?;
            StmtKind::DoWhile { body, cond, until: false }
        }

        Rule::for_statement => walk_for(pair)?,

        Rule::foreach_statement => walk_foreach(pair)?,

        Rule::switch_statement => walk_switch(pair)?,

        Rule::return_statement => {
            let expr = pair.into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            StmtKind::Return(expr)
        }

        Rule::break_statement => {
            let level = pair.into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            // PHP break/continue can take an integer level
            let target = match level {
                Some(Expression { kind: ExprKind::Lit(Literal::Int(n)), .. }) =>
                    BreakTarget::Level(n as u32),
                Some(_) => BreakTarget::Implicit,
                None => BreakTarget::Implicit,
            };
            StmtKind::Break(target)
        }

        Rule::continue_statement => {
            let level = pair.into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .transpose()?;
            let target = match level {
                Some(Expression { kind: ExprKind::Lit(Literal::Int(n)), .. }) =>
                    ContinueTarget::Level(n as u32),
                _ => ContinueTarget::Implicit,
            };
            StmtKind::Continue(target)
        }

        Rule::throw_statement => {
            let expr = walk_expression(inner_nokw(pair).next().unwrap())?;
            StmtKind::Throw { expr: Some(expr), cause: None }
        }

        Rule::try_statement => walk_try(pair)?,

        // Skip pest-internal end-of-input markers
        Rule::EOI => return Ok(None),

        other => return Err(format!("walker: unhandled statement rule {:?}", other)),
    };

    Ok(Some(Statement::with_span(kind, span)))
}

/// Walk a `statement` rule into a `Vec<Statement>` (a body). If the
/// rule is a block, return its contents; otherwise wrap the single
/// statement in a one-element Vec.
fn walk_statement_into_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    if matches!(pair.as_rule(), Rule::block_statement) {
        let mut stmts = Vec::new();
        for s in pair.into_inner() {
            if let Some(st) = walk_statement(s)? {
                stmts.push(st);
            }
        }
        Ok(stmts)
    } else {
        match walk_statement(pair)? {
            Some(s) => Ok(vec![s]),
            None => Ok(Vec::new()),
        }
    }
}

// ─── Control flow ──────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let cond = walk_expression(inner.next().unwrap())?;
    let then_body = walk_statement_into_body(inner.next().unwrap())?;
    let mut elifs = Vec::new();
    let mut else_body = None;
    for p in inner {
        match p.as_rule() {
            Rule::elseif_clause => {
                let mut e = inner_nokw(p);
                let c = walk_expression(e.next().unwrap())?;
                let b = walk_statement_into_body(e.next().unwrap())?;
                elifs.push((c, b));
            }
            Rule::else_clause => {
                let s = inner_nokw(p).next().unwrap();
                else_body = Some(walk_statement_into_body(s)?);
            }
            _ => {}
        }
    }
    Ok(StmtKind::If { cond, then_body, elifs, else_body })
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // for_statement = { kw_for ~ "(" ~ for_init? ~ ";" ~ expression? ~ ";"
    //                   ~ for_update? ~ ")" ~ statement }
    let mut init: Option<Vec<Expression>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Vec<Expression>> = None;
    let mut body_stmt: Option<Pair<Rule>> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_init => {
                let exprs: Result<Vec<_>, _> = p.into_inner().map(walk_expression).collect();
                init = Some(exprs?);
            }
            Rule::expression => {
                cond = Some(walk_expression(p)?);
            }
            Rule::for_update => {
                let exprs: Result<Vec<_>, _> = p.into_inner().map(walk_expression).collect();
                update = Some(exprs?);
            }
            _ => {
                body_stmt = Some(p);
            }
        }
    }
    let body = walk_statement_into_body(body_stmt.ok_or("for: missing body")?)?;

    // Compose multi-init / multi-update into a `Sequence` expression
    // wrapped in an Expr statement, since the common AST's `For` only
    // takes a single init Box<Statement> and a single update Expression.
    let init_stmt = init.map(|exprs| {
        let stmt_kind = if exprs.len() == 1 {
            StmtKind::Expr(exprs.into_iter().next().unwrap())
        } else {
            StmtKind::Expr(Expression::new(ExprKind::Sequence(exprs)))
        };
        Box::new(Statement::new(stmt_kind))
    });
    let update_expr = update.map(|exprs| {
        if exprs.len() == 1 {
            exprs.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Sequence(exprs))
        }
    });

    Ok(StmtKind::For {
        init: init_stmt,
        cond,
        update: update_expr,
        body,
    })
}

fn walk_foreach(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // foreach_statement = { kw_foreach ~ "(" ~ expression ~ kw_as
    //                       ~ foreach_target ~ ")" ~ statement }
    let mut iter: Option<Expression> = None;
    let mut target_pair: Option<Pair<Rule>> = None;
    let mut body_stmt: Option<Pair<Rule>> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => iter = Some(walk_expression(p)?),
            Rule::foreach_target => target_pair = Some(p),
            _ => {
                if matches!(p.as_rule(),
                    Rule::block_statement | Rule::expression_statement | Rule::if_statement |
                    Rule::while_statement | Rule::do_while_statement | Rule::for_statement |
                    Rule::foreach_statement | Rule::switch_statement | Rule::return_statement |
                    Rule::break_statement | Rule::continue_statement | Rule::throw_statement |
                    Rule::try_statement | Rule::echo_statement | Rule::print_statement |
                    Rule::empty_statement | Rule::function_declaration | Rule::class_declaration
                ) {
                    body_stmt = Some(p);
                }
            }
        }
    }

    let target = target_pair.ok_or("foreach: missing target")?;
    // foreach_target = { variable "=>" "&"? variable | "&"? variable }
    let mut tparts = target.into_inner();
    let first = tparts.next().ok_or("foreach: empty target")?;
    let second = tparts.next();

    let (key, var) = if let Some(second_var) = second {
        // key => value form
        let k = strip_dollar(first.as_str()).to_string();
        let v = strip_dollar(second_var.as_str()).to_string();
        (Some(k), v)
    } else {
        let v = strip_dollar(first.as_str()).to_string();
        (None, v)
    };

    let body = walk_statement_into_body(body_stmt.ok_or("foreach: missing body")?)?;
    Ok(StmtKind::ForIn {
        var,
        key,
        iter: iter.ok_or("foreach: missing iterable")?,
        body,
        of: true, // PHP foreach iterates values, like JS for...of
        else_body: None,
        is_async: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let expr = walk_expression(inner.next().unwrap())?;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;
    for p in inner {
        if !matches!(p.as_rule(), Rule::switch_case) { continue; }
        // switch_case = { (kw_case ~ expression | kw_default) ~ ":" ~ statement* }
        // After filtering kw_case/kw_default, the remaining children are
        // [expression?] + [statements...]. We detect "default" by checking
        // the source string since both the kw_* tokens are filtered.
        let case_src = p.as_str();
        let is_default = case_src.trim_start().to_lowercase().starts_with("default");
        let mut case_inner = inner_nokw(p);
        let mut case_value: Option<Expression> = None;
        if !is_default {
            if let Some(e) = case_inner.next() {
                if matches!(e.as_rule(), Rule::expression) {
                    case_value = Some(walk_expression(e)?);
                }
            }
        }
        let body: Result<Vec<Statement>, String> = case_inner
            .filter_map(|p| {
                walk_statement(p).transpose()
            })
            .collect();
        let body = body?;
        if is_default {
            default = Some(body);
        } else {
            cases.push(SwitchCase {
                conditions: vec![CaseCondition::Value(case_value.unwrap_or_else(Expression::null))],
                body,
            });
        }
    }
    Ok(StmtKind::Switch { expr, cases, default })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = inner_nokw(pair);
    let block = inner.next().unwrap();
    let body = walk_statement_into_body(block)?;
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;
    for p in inner {
        match p.as_rule() {
            Rule::catch_clause => {
                let mut cat = inner_nokw(p);
                let catch_type = cat.next().unwrap();
                let types: Vec<String> = catch_type.into_inner()
                    .map(|q| q.as_str().to_string())
                    .collect();
                let mut var: Option<String> = None;
                let mut catch_body_pair: Option<Pair<Rule>> = None;
                for sub in cat {
                    match sub.as_rule() {
                        Rule::variable => var = Some(strip_dollar(sub.as_str()).to_string()),
                        Rule::block_statement => catch_body_pair = Some(sub),
                        _ => {}
                    }
                }
                let catch_body = walk_statement_into_body(
                    catch_body_pair.ok_or("catch: missing body")?
                )?;
                catches.push(CatchClause {
                    types,
                    var_name: var,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                });
            }
            Rule::finally_clause => {
                let body = walk_statement_into_body(inner_nokw(p).next().unwrap())?;
                finally = Some(body);
            }
            _ => {}
        }
    }
    Ok(StmtKind::Try { body, catches, else_body: None, finally })
}

// ─── Function & class declarations ────────────────────────────────────────

/// Recursively scan a function body for `yield` / `yield from` expressions.
/// Does NOT descend into nested function/closure/class bodies — those are
/// their own generator scope.
fn body_contains_yield(stmts: &[Statement]) -> bool {
    fn ey(e: &Expression) -> bool {
        match &e.kind {
            ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
            // Scope boundaries — separate generator context
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => false,
            // Leaves
            ExprKind::Lit(_) | ExprKind::Ident(_) | ExprKind::This | ExprKind::Super
            | ExprKind::AddressOf(_) | ExprKind::Destructure(_) => false,
            // Unary wrappers
            ExprKind::Unary { expr: i, .. } | ExprKind::IsType { expr: i, .. }
            | ExprKind::Cast { expr: i, .. } | ExprKind::TypeOf(i)
            | ExprKind::Spread(i) | ExprKind::Await(i) | ExprKind::Void(i)
            | ExprKind::Delete(i) => ey(i),
            // Binary / two-child
            ExprKind::Binary { left: a, right: b, .. }
            | ExprKind::NullCoalesce { left: a, right: b }
            | ExprKind::Assign { target: a, value: b }
            | ExprKind::Walrus { target: a, value: b }
            | ExprKind::Range { start: a, end: b, .. } => ey(a) || ey(b),
            ExprKind::StaticAccess { class: a, member: b } => ey(a) || ey(b),
            ExprKind::Ternary { cond, then, else_ } => ey(cond) || ey(then) || ey(else_),
            ExprKind::Member { object, .. } => ey(object),
            ExprKind::Index { object, index } => ey(object) || ey(index),
            ExprKind::Call { callee, args, .. } => ey(callee) || args.iter().any(|a| ey(&a.value)),
            ExprKind::New { class, args } => ey(class) || args.iter().any(|a| ey(&a.value)),
            ExprKind::SuperCall { args, .. } => args.iter().any(|a| ey(&a.value)),
            ExprKind::Array(elems) => elems.iter().any(|el| ey(&el.value) || el.key.as_ref().map_or(false, |k| ey(k))),
            ExprKind::Tuple(es) | ExprKind::Set(es) | ExprKind::Sequence(es) => es.iter().any(|x| ey(x)),
            ExprKind::Object(props) => props.iter().any(|p| match p {
                ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => ey(key) || ey(value),
                ObjectProperty::Spread(x) => ey(x),
                _ => false,
            }),
            ExprKind::Interpolation(parts) => parts.iter().any(|p| match p {
                InterpolPart::Expr(x) | InterpolPart::Formatted(x, _) => ey(x),
                _ => false,
            }),
            ExprKind::Match { subject, arms } => {
                ey(subject) || arms.iter().any(|a| {
                    a.conditions.as_ref().map_or(false, |cs| cs.iter().any(|c| ey(c)))
                    || ey(&a.body)
                })
            }
            ExprKind::Comprehension { element, generators, .. } => {
                ey(element) || generators.iter().any(|g| ey(&g.iter) || g.conditions.iter().any(|c| ey(c)))
            }
            ExprKind::Slice { lower, upper, step } => {
                [lower, upper, step].iter().any(|o| o.as_ref().map_or(false, |x| ey(x)))
            }
        }
    }
    fn sy(s: &Statement) -> bool {
        match &s.kind {
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => false,
            StmtKind::Expr(e) => ey(e),
            StmtKind::Block(ss) => ss.iter().any(|s| sy(s)),
            StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|d| d.init.as_ref().map_or(false, |e| ey(e))),
            StmtKind::Return(e) => e.as_ref().map_or(false, |e| ey(e)),
            StmtKind::If { cond, then_body, elifs, else_body } => {
                ey(cond) || then_body.iter().any(|s| sy(s))
                || elifs.iter().any(|(c, b)| ey(c) || b.iter().any(|s| sy(s)))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::While { cond, body, else_body } => {
                ey(cond) || body.iter().any(|s| sy(s))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::DoWhile { body, cond, .. } => body.iter().any(|s| sy(s)) || ey(cond),
            StmtKind::For { init, cond, update, body } => {
                init.as_ref().map_or(false, |s| sy(s))
                || cond.as_ref().map_or(false, |e| ey(e))
                || update.as_ref().map_or(false, |e| ey(e))
                || body.iter().any(|s| sy(s))
            }
            StmtKind::ForIn { iter, body, else_body, .. } => {
                ey(iter) || body.iter().any(|s| sy(s))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Switch { expr, cases, default } => {
                ey(expr) || cases.iter().any(|c| c.body.iter().any(|s| sy(s)))
                || default.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                body.iter().any(|s| sy(s))
                || catches.iter().any(|c| c.body.iter().any(|s| sy(s)))
                || else_body.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
                || finally.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Assign { targets, value } => targets.iter().any(|e| ey(e)) || ey(value),
            StmtKind::CompoundAssign { target, value, .. } => ey(target) || ey(value),
            StmtKind::Throw { expr, cause } => {
                expr.as_ref().map_or(false, |e| ey(e)) || cause.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::Labeled { body, .. } => sy(body),
            StmtKind::Echo(es) | StmtKind::Delete(es) => es.iter().any(|e| ey(e)),
            StmtKind::Export { declaration, default, .. } => {
                declaration.as_ref().map_or(false, |s| sy(s))
                || default.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::With { body, .. } | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. } | StmtKind::NamespaceDecl { body, .. } => body.iter().any(|s| sy(s)),
            StmtKind::MatchStatement { subject, cases } => {
                ey(subject) || cases.iter().any(|c| {
                    c.guard.as_ref().map_or(false, |e| ey(e)) || c.body.iter().any(|s| sy(s))
                })
            }
            StmtKind::Assert { test, msg } => ey(test) || msg.as_ref().map_or(false, |e| ey(e)),
            _ => false,
        }
    }
    stmts.iter().any(|s| sy(s))
}

fn walk_function_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params: Vec<Param> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::return_type_annotation => {
                return_type = Some(p.as_str().trim_start_matches(':').trim().to_string());
            }
            Rule::block_statement => {
                body = walk_statement_into_body(p)?;
            }
            _ => {}
        }
    }

    let is_generator = body_contains_yield(&body);
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator,
        is_sub: false,
    })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::param) { continue; }
        out.push(walk_param(p)?);
    }
    Ok(out)
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;
    let pass_by = PassBy::Value; // We don't track by-ref yet — PHP `&` is rare in modern code.
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_modifier => { /* discard — promoted ctor params handled in walk_class */ }
            Rule::type_annotation => type_hint = Some(p.as_str().to_string()),
            Rule::variable => name = strip_dollar(p.as_str()).to_string(),
            Rule::expression => default = Some(walk_expression(p)?),
            _ => {}
        }
    }
    let is_optional = default.is_some();
    Ok(Param {
        name,
        type_hint,
        default,
        pass_by,
        is_rest: false,
        is_kwargs: false,
        is_optional,
        is_nullable: false,
    })
}

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut modifiers = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifier => {
                let s = p.as_str().to_lowercase();
                match s.as_str() {
                    "abstract" => modifiers.is_abstract = true,
                    "final" => modifiers.is_sealed = true,
                    _ => {}
                }
            }
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::qualified_name => {
                if parents.is_empty() {
                    parents.push(p.as_str().to_string());
                } else {
                    interfaces.push(p.as_str().to_string());
                }
            }
            Rule::use_trait | Rule::class_constant | Rule::property_declaration
                | Rule::method_declaration | Rule::empty_statement => {
                if let Some(member) = walk_class_member(p)? {
                    members.push(member);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl { name, parents, interfaces, members, modifiers })
}

fn walk_interface_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // For now we walk interfaces as InterfaceDecl. Member methods become
    // signature-only entries. Constants become Const members on the
    // interface, which the compiler treats as static fields.
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<InterfaceMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::qualified_name => parents.push(p.as_str().to_string()),
            Rule::method_declaration => {
                // Signature only — discard body.
                let mut method_name = String::new();
                let mut params: Vec<Param> = Vec::new();
                let mut return_type: Option<String> = None;
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::identifier => method_name = m.as_str().to_string(),
                        Rule::param_list => params = walk_params(m)?,
                        Rule::return_type_annotation => {
                            return_type = Some(m.as_str().trim_start_matches(':').trim().to_string());
                        }
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Method {
                    name: method_name,
                    params,
                    return_type,
                    is_sub: false,
                });
            }
            _ => {}
        }
    }

    Ok(StmtKind::InterfaceDecl { name, parents, members })
}

fn walk_trait_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // Treat a trait as a regular class for compilation purposes —
    // `use TraitName` inside another class flattens via the same parent
    // chain.
    let mut name = String::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::use_trait | Rule::class_constant | Rule::property_declaration
                | Rule::method_declaration | Rule::empty_statement => {
                if let Some(member) = walk_class_member(p)? {
                    members.push(member);
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
    })
}

fn walk_enum_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    // PHP enums compile to a class with static constants. The walker
    // converts each enum case into a Const class member with the case
    // name as the key.
    let mut name = String::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::enum_case => {
                let mut case_name = String::new();
                let mut value: Option<Expression> = None;
                for c in p.into_inner() {
                    match c.as_rule() {
                        Rule::identifier => case_name = c.as_str().to_string(),
                        Rule::expression => value = Some(walk_expression(c)?),
                        _ => {}
                    }
                }
                let value = value.unwrap_or_else(|| Expression::string(&case_name));
                members.push(ClassMember::Const {
                    name: case_name,
                    type_hint: None,
                    value,
                    visibility: Visibility::Public,
                });
            }
            Rule::class_constant | Rule::method_declaration | Rule::use_trait => {
                if let Some(m) = walk_class_member(p)? { members.push(m); }
            }
            _ => {}
        }
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
    })
}

fn walk_class_member(pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::empty_statement => Ok(None),
        Rule::use_trait => {
            // `use TraitName;` inside a class — for now, no-op. Trait
            // method copy-in happens via the dotnet/dotnet-style
            // inheritance chain at compile time, which we don't model
            // here yet. Future: synthesize parent-trait method bindings.
            Ok(None)
        }
        Rule::class_constant => {
            let mut name = String::new();
            let mut value: Option<Expression> = None;
            let mut visibility = Visibility::Public;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier => visibility = parse_visibility(p.as_str(), visibility),
                    Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
                    Rule::expression => value = Some(walk_expression(p)?),
                    _ => {}
                }
            }
            Ok(Some(ClassMember::Const {
                name,
                type_hint: None,
                value: value.unwrap_or_else(Expression::null),
                visibility,
            }))
        }
        Rule::property_declaration => {
            let mut name = String::new();
            let mut type_hint: Option<String> = None;
            let mut init: Option<Expression> = None;
            let mut modifiers = Modifiers::default();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier => apply_member_modifier(&mut modifiers, p.as_str()),
                    Rule::type_annotation => type_hint = Some(p.as_str().to_string()),
                    Rule::variable => name = strip_dollar(p.as_str()).to_string(),
                    Rule::expression => init = Some(walk_expression(p)?),
                    _ => {}
                }
            }
            Ok(Some(ClassMember::Field {
                name,
                type_hint,
                init,
                modifiers,
                with_events: false,
                array_bounds: None,
            }))
        }
        Rule::method_declaration => {
            let mut method_name = String::new();
            let mut params: Vec<Param> = Vec::new();
            let mut body: Vec<Statement> = Vec::new();
            let mut return_type: Option<String> = None;
            let mut modifiers = Modifiers::default();
            let mut has_body = false;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::member_modifier => apply_member_modifier(&mut modifiers, p.as_str()),
                    Rule::identifier => method_name = p.as_str().to_string(),
                    Rule::param_list => params = walk_params(p)?,
                    Rule::return_type_annotation => {
                        return_type = Some(p.as_str().trim_start_matches(':').trim().to_string());
                    }
                    Rule::block_statement => {
                        body = walk_statement_into_body(p)?;
                        has_body = true;
                    }
                    _ => {}
                }
            }

            // PHP `__construct` becomes a Constructor class member so
            // the compiler-side child-class flow recognises it.
            if method_name == "__construct" {
                return Ok(Some(ClassMember::Constructor {
                    params,
                    body,
                    base_args: None,
                    visibility: modifiers.visibility,
                }));
            }

            // Build a Method wrapping a FunctionDecl.
            let method_body = if has_body { body } else { Vec::new() };
            let is_generator = body_contains_yield(&method_body);
            let stmt = Statement::new(StmtKind::FunctionDecl {
                name: method_name,
                params,
                return_type,
                body: method_body,
                modifiers,
                handles: Vec::new(),
                is_async: false,
                is_generator,
                is_sub: false,
            });
            Ok(Some(ClassMember::Method(Box::new(stmt))))
        }
        _ => Ok(None),
    }
}

fn apply_member_modifier(mods: &mut Modifiers, kw: &str) {
    let lower = kw.to_lowercase();
    match lower.as_str() {
        "public" => mods.visibility = Visibility::Public,
        "private" => mods.visibility = Visibility::Private,
        "protected" => mods.visibility = Visibility::Protected,
        "static" => mods.is_static = true,
        "abstract" => mods.is_abstract = true,
        "final" => mods.is_not_overridable = true,
        "readonly" => mods.is_readonly = true,
        _ => {}
    }
}

fn parse_visibility(s: &str, default: Visibility) -> Visibility {
    match s.to_lowercase().as_str() {
        "public" => Visibility::Public,
        "private" => Visibility::Private,
        "protected" => Visibility::Protected,
        _ => default,
    }
}

// ─── Expressions ──────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    let kind = match rule {
        Rule::expression => return walk_expression(pair.into_inner().next().unwrap()),
        Rule::assignment_expression => return walk_assignment(pair),
        Rule::yield_expression => return walk_yield(pair),
        Rule::logical_or_expression
            | Rule::null_coalesce_expression
            | Rule::logic_or_expression
            | Rule::logic_and_expression
            | Rule::bit_or_expression
            | Rule::equality_expression
            | Rule::comparison_expression
            | Rule::shift_expression
            | Rule::additive_expression
            | Rule::multiplicative_expression
                => return walk_left_assoc_binary(pair),
        Rule::ternary_expression => return walk_ternary(pair),
        Rule::unary_expression => return walk_unary(pair),
        Rule::cast_expression => return walk_cast(pair),
        Rule::postfix_expression => return walk_postfix(pair),

        Rule::primary_expression => return walk_expression(pair.into_inner().next().unwrap()),
        Rule::parenthesized_expression => return walk_expression(pair.into_inner().next().unwrap()),

        Rule::literal => return walk_literal(pair),
        Rule::number_lit => return Ok(Expression::with_span(walk_number(&pair).kind, span)),
        Rule::string_lit => return Ok(Expression::with_span(walk_string(&pair).kind, span)),

        Rule::variable => ExprKind::Ident(strip_dollar(pair.as_str()).to_string()),
        Rule::identifier => ExprKind::Ident(pair.as_str().to_string()),
        Rule::qualified_name => {
            // Preserve the qualified path so the compiler can resolve it
            // against the profile's `host_packages` list. A qualified
            // name whose first segment matches a host package (e.g.
            // `\Vybe\Http\Response\set_status`) becomes a Component
            // Model host call at compile time. Anything else still
            // resolves by last-segment (user namespaces are flattened
            // for now — worth revisiting when we model user namespaces).
            let s = pair.as_str().trim_start_matches('\\');
            if s.contains('\\') {
                // Store the full qualified name with backslashes preserved
                // as an identifier. The compiler's `try_compile_builtin`
                // path detects backslashes and routes to host calls.
                ExprKind::Ident(s.to_string())
            } else {
                ExprKind::Ident(s.to_string())
            }
        }

        Rule::kw_self => ExprKind::This,  // PHP `self::` inside a method ≈ `this`
        Rule::kw_parent => ExprKind::Super,
        Rule::kw_static => ExprKind::Ident("static".to_string()),

        Rule::new_expression => return walk_new(pair),
        Rule::clone_expression => {
            let inner = inner_nokw(pair).next().unwrap();
            let arg = walk_expression(inner)?;
            // Translate to a call: __clone(arg)
            ExprKind::Call {
                callee: Box::new(Expression::ident("__clone")),
                args: vec![Argument::positional(arg)],
                optional: false,
            }
        }
        Rule::match_expression => return walk_match(pair),
        Rule::isset_expression => {
            let args: Result<Vec<_>, _> = pair.into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .collect();
            let args: Vec<Argument> = args?.into_iter().map(Argument::positional).collect();
            ExprKind::Call {
                callee: Box::new(Expression::ident("isset")),
                args,
                optional: false,
            }
        }
        Rule::empty_expression => {
            let arg = walk_expression(inner_nokw(pair).next().unwrap())?;
            ExprKind::Call {
                callee: Box::new(Expression::ident("empty")),
                args: vec![Argument::positional(arg)],
                optional: false,
            }
        }
        Rule::unset_expression => {
            // PHP `unset($a, $b)` — emit as a call to a builtin so the
            // compiler can route through compiler_common's delete path.
            // The expression-level `Delete` AST node only takes a
            // single Box<Expression>, so we wrap multi-arg unset() as
            // a Call instead.
            let exprs: Result<Vec<_>, _> = pair.into_inner()
                .filter(|p| matches!(p.as_rule(), Rule::expression))
                .map(walk_expression)
                .collect();
            let args: Vec<Argument> = exprs?.into_iter().map(Argument::positional).collect();
            ExprKind::Call {
                callee: Box::new(Expression::ident("unset")),
                args,
                optional: false,
            }
        }
        Rule::list_expression => {
            // list($a, $b, $c) — destructure target. Walk each element
            // into a Destructure pattern.
            let mut elems = Vec::new();
            for p in pair.into_inner() {
                if matches!(p.as_rule(), Rule::list_element) {
                    let inner = p.into_inner().next();
                    if let Some(e) = inner {
                        let expr = walk_expression(e)?;
                        if let ExprKind::Ident(name) = expr.kind {
                            elems.push(ArrayPatternElem::Pattern(BindingPattern::Ident(name), None));
                        } else {
                            elems.push(ArrayPatternElem::Hole);
                        }
                    } else {
                        elems.push(ArrayPatternElem::Hole);
                    }
                }
            }
            ExprKind::Destructure(DestructurePattern::Array(elems))
        }
        Rule::array_expression | Rule::short_array_expression => return walk_array(pair),
        Rule::closure_expression => return walk_closure(pair),
        Rule::arrow_function => return walk_arrow_function(pair),

        other => return Err(format!("walker: unhandled expression rule {:?}", other)),
    };

    Ok(Expression::with_span(kind, span))
}

fn walk_left_assoc_binary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner().peekable();
    let mut left = walk_expression(inner.next().unwrap())?;
    while let Some(op_pair) = inner.next() {
        // op_pair is the operator alternation (eq_op, cmp_op, …) OR a
        // direct binary operator string. Skip if it's actually an
        // operand (left-associative chain has alternating operand/op).
        let op_str = op_pair.as_str().to_string();
        // The next pair should be the right operand.
        let right_pair = match inner.next() {
            Some(p) => p,
            None => break,
        };
        let right = walk_expression(right_pair)?;
        let op = parse_binop(&op_str);
        left = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(left),
                right: Box::new(right),
            },
            span.clone(),
        );
    }
    Ok(left)
}

fn parse_binop(s: &str) -> BinOp {
    match s {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "**" => BinOp::Pow,
        "." => BinOp::Concat,
        "==" => BinOp::Eq,
        "===" => BinOp::StrictEq,
        "!=" | "<>" => BinOp::NotEq,
        "!==" => BinOp::StrictNotEq,
        "<" => BinOp::Lt,
        ">" => BinOp::Gt,
        "<=" => BinOp::LtEq,
        ">=" => BinOp::GtEq,
        "<=>" => BinOp::Spaceship,
        "&&" | "and" | "AND" => BinOp::And,
        "||" | "or" | "OR" => BinOp::Or,
        "xor" | "XOR" => BinOp::Xor,
        "|" => BinOp::BitOr,
        "&" => BinOp::BitAnd,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "instanceof" | "INSTANCEOF" => BinOp::InstanceOf,
        _ => BinOp::Add, // fallback — safer than panic
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let lhs_pair = inner.next().unwrap();
    // If lhs_pair is yield_expression, just walk it through.
    if matches!(lhs_pair.as_rule(), Rule::yield_expression) {
        return walk_expression(lhs_pair);
    }
    let lhs = walk_expression(lhs_pair)?;
    if let Some(op_pair) = inner.next() {
        let op = op_pair.as_str();
        let rhs = walk_expression(inner.next().unwrap())?;
        let kind = match op {
            "=" => ExprKind::Assign {
                target: Box::new(lhs),
                value: Box::new(rhs),
            },
            other => {
                let cop = parse_compound_op(other);
                // CompoundAssign is a stmt-level node; expression-level
                // compound assignments synthesize `target = target OP rhs`.
                let combined = Expression::with_span(
                    ExprKind::Binary {
                        op: compound_to_binop(cop),
                        left: Box::new(lhs.clone()),
                        right: Box::new(rhs),
                    },
                    span.clone(),
                );
                ExprKind::Assign {
                    target: Box::new(lhs),
                    value: Box::new(combined),
                }
            }
        };
        Ok(Expression::with_span(kind, span))
    } else {
        Ok(lhs)
    }
}

fn parse_compound_op(s: &str) -> CompoundOp {
    match s {
        "+=" => CompoundOp::Add,
        "-=" => CompoundOp::Sub,
        "*=" => CompoundOp::Mul,
        "/=" => CompoundOp::Div,
        "%=" => CompoundOp::Mod,
        ".=" => CompoundOp::Concat,
        "**=" => CompoundOp::Pow,
        "<<=" => CompoundOp::Shl,
        ">>=" => CompoundOp::Shr,
        "&=" => CompoundOp::BitAnd,
        "|=" => CompoundOp::BitOr,
        "^=" => CompoundOp::BitXor,
        "&&=" => CompoundOp::And,
        "||=" => CompoundOp::Or,
        "??=" => CompoundOp::NullCoalesce,
        _ => CompoundOp::Add,
    }
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::Concat => BinOp::Concat,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::UShr => BinOp::UShr,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::IDiv => BinOp::IDiv,
    }
}

fn walk_yield(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    // Detect `yield from` from the source slice — kw_yield/kw_yield_from
    // are filtered out alongside the rest of the keyword tokens.
    let yield_from = pair.as_str().trim_start().to_lowercase().starts_with("yield from");
    let mut inner = inner_nokw(pair);
    if yield_from {
        let val = walk_expression(inner.next().unwrap())?;
        return Ok(Expression::with_span(
            ExprKind::YieldFrom(Box::new(val)),
            span,
        ));
    }
    // bare `yield`, `yield expr`, or `yield key => value`
    let val = inner.next().map(walk_expression).transpose()?;
    Ok(Expression::with_span(
        ExprKind::Yield(val.map(Box::new)),
        span,
    ))
}

fn walk_ternary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().unwrap())?;
    let next = inner.next();
    if next.is_none() {
        return Ok(cond);
    }
    let mut next = next.unwrap();
    // Two forms:
    //   `cond ? then : else`
    //   `cond ?: else` (Elvis — short ternary)
    if matches!(next.as_rule(), Rule::expression) {
        // We have a `then` branch.
        let then_expr = walk_expression(next)?;
        let else_expr = walk_expression(inner.next().unwrap())?;
        return Ok(Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then_expr),
                else_: Box::new(else_expr),
            },
            span,
        ));
    } else {
        // Elvis: cond ?: else  →  cond ?? else (semantically close enough)
        // Actually PHP's ?: returns cond if truthy, else the right side.
        // We model as: cond ? cond : else
        next = inner.next().unwrap_or(next);
        let else_expr = walk_expression(next)?;
        Ok(Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond.clone()),
                then: Box::new(cond),
                else_: Box::new(else_expr),
            },
            span,
        ))
    }
}

fn walk_unary(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    if matches!(first.as_rule(), Rule::unary_op) {
        let op = parse_unary_op(first.as_str());
        let expr = walk_expression(inner.next().unwrap())?;
        Ok(Expression::with_span(
            ExprKind::Unary { op, expr: Box::new(expr) },
            span,
        ))
    } else {
        walk_expression(first)
    }
}

fn parse_unary_op(s: &str) -> UnaryOp {
    match s {
        "!" => UnaryOp::Not,
        "~" => UnaryOp::BitNot,
        "-" => UnaryOp::Neg,
        "+" => UnaryOp::Pos,
        "++" => UnaryOp::PreInc,
        "--" => UnaryOp::PreDec,
        "@" => UnaryOp::Pos, // PHP error suppression — semantically a no-op for us
        _ => UnaryOp::Pos,
    }
}

fn walk_cast(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cast_kw = inner.next().unwrap().as_str().to_string();
    let expr = walk_expression(inner.next().unwrap())?;
    Ok(Expression::with_span(
        ExprKind::Cast { expr: Box::new(expr), type_name: cast_kw },
        span,
    ))
}

fn walk_postfix(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let mut expr = walk_expression(inner.next().unwrap())?;
    for op_pair in inner {
        expr = apply_postfix(expr, op_pair, &span)?;
    }
    Ok(expr)
}

fn apply_postfix(receiver: Expression, op: Pair<Rule>, span: &Span) -> Result<Expression, String> {
    // The grammar wraps all variants in a non-silent `postfix_op` rule, so
    // pest yields a `postfix_op` pair whose single child is the actual
    // op rule (`method_call_op`, `inc_dec_op`, etc.). Unwrap once so the
    // match below sees the real rule; otherwise every postfix silently
    // falls through to the `_ => Ok(receiver)` arm (dropping `$i++`,
    // `$obj->foo(...)`, `$arr[0]`, …).
    let op = if matches!(op.as_rule(), Rule::postfix_op) {
        op.into_inner().next().ok_or("empty postfix_op")?
    } else {
        op
    };
    let rule = op.as_rule();
    match rule {
        Rule::method_call_op => {
            // The grammar emits: ("?->"|"->") ~ member_name ~ "(" ~ arg_list? ~ ")"
            // The literal "->" / "?->" appears as a non-rule token, so
            // pest does NOT yield it as a child pair. Detect null-safe
            // from the outer pair's source text instead of trying to
            // read it from inner pairs.
            let null_safe = op.as_str().trim_start().starts_with("?->");
            let mut name_pair: Option<Pair<Rule>> = None;
            let mut arg_list_pair: Option<Pair<Rule>> = None;
            for p in op.into_inner() {
                match p.as_rule() {
                    Rule::member_name => name_pair = Some(p),
                    Rule::arg_list => arg_list_pair = Some(p),
                    _ => {}
                }
            }
            let name = name_pair
                .ok_or("method_call_op: missing name")?
                .into_inner().next().unwrap().as_str().to_string();
            let member = Expression::with_span(
                ExprKind::Member {
                    object: Box::new(receiver),
                    field: name,
                    null_safe,
                },
                span.clone(),
            );
            let args = arg_list_pair.map(walk_args).transpose()?.unwrap_or_default();
            Ok(Expression::with_span(
                ExprKind::Call { callee: Box::new(member), args, optional: null_safe },
                span.clone(),
            ))
        }
        Rule::property_access_op => {
            // Grammar: ("?->"|"->") ~ member_name. The arrow is a
            // literal token (pest does not yield it as a child pair),
            // so the only inner rule pair is `member_name`. Read
            // null_safe from the outer pair's source text.
            let null_safe = op.as_str().trim_start().starts_with("?->");
            let name_pair = op.into_inner().next()
                .ok_or("property_access_op: missing name")?;
            let name = name_pair.into_inner().next().unwrap().as_str().to_string();
            Ok(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(receiver),
                    field: name,
                    null_safe,
                },
                span.clone(),
            ))
        }
        Rule::static_access_op => {
            let mut inner = op.into_inner();
            let name_pair = inner.next().unwrap();
            let name = name_pair.into_inner().next().unwrap().as_str().to_string();
            Ok(Expression::with_span(
                ExprKind::StaticAccess {
                    class: Box::new(receiver),
                    member: Box::new(Expression::ident(&name)),
                },
                span.clone(),
            ))
        }
        Rule::array_index_op => {
            let mut inner = op.into_inner();
            let index = if let Some(i) = inner.next() {
                walk_expression(i)?
            } else {
                Expression::null()
            };
            Ok(Expression::with_span(
                ExprKind::Index {
                    object: Box::new(receiver),
                    index: Box::new(index),
                },
                span.clone(),
            ))
        }
        Rule::call_op => {
            let mut inner = op.into_inner();
            let args = if let Some(al) = inner.next() {
                walk_args(al)?
            } else {
                Vec::new()
            };
            // Normalize PHP-specific argument conventions to the common
            // AST's canonical order BEFORE the compiler sees them. Once
            // in the common AST, PHP and JS calls should be
            // indistinguishable — the compiler emits a single canonical
            // host call regardless of surface syntax.
            let args = canonicalize_php_call_args(&receiver, args);
            Ok(Expression::with_span(
                ExprKind::Call { callee: Box::new(receiver), args, optional: false },
                span.clone(),
            ))
        }
        Rule::inc_dec_op => {
            let op = if op.as_str() == "++" { UnaryOp::PostInc } else { UnaryOp::PostDec };
            Ok(Expression::with_span(
                ExprKind::Unary { op, expr: Box::new(receiver) },
                span.clone(),
            ))
        }
        _ => Ok(receiver),
    }
}

fn walk_args(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    let mut out = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::argument) { continue; }
        // Detect spread/by_ref by inspecting the source slice — pest
        // doesn't capture the literal `...` / `&` as named rules.
        let raw = p.as_str();
        let spread = raw.trim_start().starts_with("...");
        let by_ref = raw.trim_start().starts_with('&');

        let mut name: Option<String> = None;
        let mut value: Option<Expression> = None;
        for sub in p.into_inner() {
            match sub.as_rule() {
                Rule::identifier => name = Some(sub.as_str().to_string()),
                Rule::expression => value = Some(walk_expression(sub)?),
                _ => {}
            }
        }
        if let Some(v) = value {
            out.push(Argument {
                name,
                value: v,
                by_ref,
                spread,
            });
        }
    }
    Ok(out)
}

fn walk_new(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    // new_expression = { kw_new ~ (qualified_name | variable | "(" expr ")")
    //                    ~ ("(" arg_list? ")")? }
    // After filtering kw_new, the first child is the class designator and
    // the optional second is the arg_list.
    let mut class: Option<Expression> = None;
    let mut args: Vec<Argument> = Vec::new();
    for p in inner_nokw(pair) {
        match p.as_rule() {
            Rule::arg_list => args = walk_args(p)?,
            _ => {
                if class.is_none() {
                    class = Some(walk_expression(p)?);
                }
            }
        }
    }
    Ok(Expression::with_span(
        ExprKind::New {
            class: Box::new(class.ok_or("new: missing class designator")?),
            args,
        },
        span,
    ))
}

fn walk_match(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = inner_nokw(pair);
    let subject = walk_expression(inner.next().unwrap())?;
    let mut arms: Vec<MatchArm> = Vec::new();
    for p in inner {
        if !matches!(p.as_rule(), Rule::match_arm) { continue; }
        // match_arm = { (kw_default | match_conditions) ~ "=>" ~ expression }
        // Use the source slice to detect default-arms because kw_default
        // is filtered out alongside other keyword tokens.
        let arm_src = p.as_str().trim_start();
        let is_default = arm_src.to_lowercase().starts_with("default");
        let mut conditions: Option<Vec<Expression>> = None;
        let mut body: Option<Expression> = None;
        for sub in inner_nokw(p) {
            match sub.as_rule() {
                Rule::match_conditions => {
                    let exprs: Result<Vec<_>, _> = sub.into_inner().map(walk_expression).collect();
                    conditions = Some(exprs?);
                }
                Rule::expression => body = Some(walk_expression(sub)?),
                _ => {}
            }
        }
        if is_default { conditions = None; }
        arms.push(MatchArm {
            conditions,
            body: body.unwrap_or_else(Expression::null),
        });
    }
    Ok(Expression::with_span(
        ExprKind::Match { subject: Box::new(subject), arms },
        span,
    ))
}

fn walk_array(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut elems = Vec::new();
    for p in pair.into_inner() {
        if !matches!(p.as_rule(), Rule::array_element) { continue; }
        let mut sub_iter = p.into_inner();
        let first = sub_iter.next();
        let second = sub_iter.next();
        match (first, second) {
            (Some(first), Some(second)) => {
                // key => value
                let key = walk_expression(first)?;
                let value = walk_expression(second)?;
                elems.push(ArrayElement {
                    key: Some(key),
                    value,
                    spread: false,
                    by_ref: false,
                });
            }
            (Some(first), None) => {
                let value = walk_expression(first)?;
                elems.push(ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                });
            }
            _ => {}
        }
    }
    Ok(Expression::with_span(ExprKind::Array(elems), span))
}

fn walk_closure(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut params: Vec<Param> = Vec::new();
    let mut captures: Vec<String> = Vec::new();
    let mut body: Vec<Statement> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::closure_use => {
                for v in p.into_inner() {
                    if matches!(v.as_rule(), Rule::closure_use_var) {
                        if let Some(var) = v.into_inner().find(|q| matches!(q.as_rule(), Rule::variable)) {
                            captures.push(strip_dollar(var.as_str()).to_string());
                        }
                    }
                }
            }
            Rule::block_statement => body = walk_statement_into_body(p)?,
            _ => {}
        }
    }
    Ok(Expression::with_span(
        ExprKind::Lambda {
            params,
            body: LambdaBody::Block(body),
            is_async: false,
            captures,
        },
        span,
    ))
}

fn walk_arrow_function(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut params: Vec<Param> = Vec::new();
    let mut body_expr: Option<Expression> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::expression => body_expr = Some(walk_expression(p)?),
            _ => {}
        }
    }
    Ok(Expression::with_span(
        ExprKind::Lambda {
            params,
            body: LambdaBody::Expr(Box::new(body_expr.unwrap_or_else(Expression::null))),
            is_async: false,
            captures: Vec::new(),
        },
        span,
    ))
}

// ─── Literals ─────────────────────────────────────────────────────────────

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next().unwrap();
    let kind = match inner.as_rule() {
        Rule::number_lit => walk_number(&inner).kind,
        Rule::string_lit => walk_string(&inner).kind,
        Rule::kw_null => ExprKind::Lit(Literal::Null),
        Rule::kw_true => ExprKind::Lit(Literal::Bool(true)),
        Rule::kw_false => ExprKind::Lit(Literal::Bool(false)),
        _ => ExprKind::Lit(Literal::Null),
    };
    Ok(Expression::with_span(kind, span))
}

fn walk_number(pair: &Pair<Rule>) -> Expression {
    let s = pair.as_str();
    let kind = if let Some(rest) = s.strip_prefix("0x").or_else(|| s.strip_prefix("0X")) {
        i64::from_str_radix(rest, 16).map(Literal::Int).map(ExprKind::Lit).unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0b").or_else(|| s.strip_prefix("0B")) {
        i64::from_str_radix(rest, 2).map(Literal::Int).map(ExprKind::Lit).unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if let Some(rest) = s.strip_prefix("0o").or_else(|| s.strip_prefix("0O")) {
        i64::from_str_radix(rest, 8).map(Literal::Int).map(ExprKind::Lit).unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else if s.contains('.') || s.contains('e') || s.contains('E') {
        s.parse::<f64>().map(Literal::Float).map(ExprKind::Lit).unwrap_or(ExprKind::Lit(Literal::Int(0)))
    } else {
        s.parse::<i64>().map(Literal::Int).map(ExprKind::Lit).unwrap_or(ExprKind::Lit(Literal::Int(0)))
    };
    Expression::new(kind)
}

fn walk_string(pair: &Pair<Rule>) -> Expression {
    let raw = pair.as_str();
    let body = &raw[1..raw.len() - 1];

    if raw.starts_with('\'') {
        // Single-quoted: literal, only \' and \\ escapes. No
        // interpolation in PHP.
        let mut out = String::with_capacity(body.len());
        let mut chars = body.chars().peekable();
        while let Some(c) = chars.next() {
            if c == '\\' {
                if let Some(&next) = chars.peek() {
                    if next == '\'' || next == '\\' {
                        out.push(chars.next().unwrap());
                        continue;
                    }
                }
            }
            out.push(c);
        }
        return Expression::new(ExprKind::Lit(Literal::Str(out)));
    }

    // Double-quoted: PHP interpolation. Scan for `$var`, `$var[key]`,
    // `$var->prop`, `{$expr}` and split the body into InterpolParts.
    // Empty or interp-free strings collapse back to a plain literal so
    // the compiler's string path stays fast.
    let parts = parse_php_interpolation(body);
    if parts.len() == 1 {
        if let InterpolPart::Text(s) = &parts[0] {
            return Expression::new(ExprKind::Lit(Literal::Str(s.clone())));
        }
    }
    if parts.is_empty() {
        return Expression::new(ExprKind::Lit(Literal::Str(String::new())));
    }
    Expression::new(ExprKind::Interpolation(parts))
}

/// Scan a double-quoted PHP string body into `InterpolPart`s, handling:
///   - escape sequences (`\n`, `\t`, `\"`, `\\`, `\$`, …)
///   - `$var`, `$var_with_underscores`
///   - `$arr[key]` — PHP-classic "unquoted key is a string" rule; digit
///     keys become int literals
///   - `$obj->prop`
///   - `{$arbitrary_expr}` — balanced brace matching; inner text parsed
///     by re-entering the PHP expression rule
fn parse_php_interpolation(body: &str) -> Vec<InterpolPart> {
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut text = String::new();
    let mut chars = body.chars().peekable();

    let flush = |parts: &mut Vec<InterpolPart>, text: &mut String| {
        if !text.is_empty() {
            parts.push(InterpolPart::Text(std::mem::take(text)));
        }
    };

    while let Some(c) = chars.next() {
        // Escapes — must run before $ detection so `\$name` stays literal.
        if c == '\\' {
            match chars.next() {
                Some('n') => text.push('\n'),
                Some('t') => text.push('\t'),
                Some('r') => text.push('\r'),
                Some('"') => text.push('"'),
                Some('\\') => text.push('\\'),
                Some('$') => text.push('$'),
                Some('{') => text.push('{'),
                Some('0') => text.push('\0'),
                Some(other) => { text.push('\\'); text.push(other); }
                None => text.push('\\'),
            }
            continue;
        }

        // `{$...}` complex form — balanced brace scan, re-parse inner.
        if c == '{' && chars.peek() == Some(&'$') {
            flush(&mut parts, &mut text);
            chars.next(); // consume $
            let mut expr_src = String::from("$");
            let mut depth: i32 = 1;
            let mut in_str: Option<char> = None;
            while let Some(&nc) = chars.peek() {
                chars.next();
                if let Some(q) = in_str {
                    expr_src.push(nc);
                    if nc == '\\' {
                        if let Some(&esc) = chars.peek() {
                            expr_src.push(esc);
                            chars.next();
                        }
                        continue;
                    }
                    if nc == q { in_str = None; }
                    continue;
                }
                if nc == '"' || nc == '\'' { in_str = Some(nc); expr_src.push(nc); continue; }
                if nc == '{' { depth += 1; expr_src.push(nc); continue; }
                if nc == '}' {
                    depth -= 1;
                    if depth == 0 { break; }
                    expr_src.push(nc);
                    continue;
                }
                expr_src.push(nc);
            }
            match parse_interpol_expression(&expr_src) {
                Ok(expr) => parts.push(InterpolPart::Expr(expr)),
                // Fall back to literal so we don't lose user content on
                // parse failure.
                Err(_) => parts.push(InterpolPart::Text(format!("{{{}}}", expr_src))),
            }
            continue;
        }

        // `$identifier` — possibly followed by `[key]` or `->prop`.
        if c == '$' {
            let peek = chars.peek().copied();
            if matches!(peek, Some(c) if c.is_ascii_alphabetic() || c == '_') {
                flush(&mut parts, &mut text);
                let mut name = String::new();
                while let Some(&nc) = chars.peek() {
                    if nc.is_ascii_alphanumeric() || nc == '_' {
                        name.push(nc);
                        chars.next();
                    } else { break; }
                }
                let mut expr = Expression::new(ExprKind::Ident(name));

                // `$var[key]` — simple unquoted key; per PHP's rule,
                // identifiers are string literals, digit-runs are ints.
                if chars.peek() == Some(&'[') {
                    chars.next(); // consume [
                    let mut key_text = String::new();
                    while let Some(&nc) = chars.peek() {
                        if nc == ']' { chars.next(); break; }
                        key_text.push(nc);
                        chars.next();
                    }
                    let key_trimmed = key_text.trim();
                    let key_expr = if let Ok(n) = key_trimmed.parse::<i64>() {
                        Expression::new(ExprKind::Lit(Literal::Int(n)))
                    } else {
                        // PHP quirk: `$a[$b]` inside string is a variable
                        // if starts with `$`, else unquoted string.
                        if let Some(inner) = key_trimmed.strip_prefix('$') {
                            Expression::new(ExprKind::Ident(inner.to_string()))
                        } else {
                            Expression::new(ExprKind::Lit(Literal::Str(key_trimmed.to_string())))
                        }
                    };
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(key_expr),
                    });
                } else if chars.peek() == Some(&'-') {
                    // Look ahead for `->`. If absent, `-` is literal.
                    let mut save = chars.clone();
                    save.next();
                    if save.peek() == Some(&'>') {
                        chars.next(); // -
                        chars.next(); // >
                        let mut prop = String::new();
                        while let Some(&nc) = chars.peek() {
                            if nc.is_ascii_alphanumeric() || nc == '_' {
                                prop.push(nc);
                                chars.next();
                            } else { break; }
                        }
                        if !prop.is_empty() {
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field: prop,
                                null_safe: false,
                            });
                        }
                    }
                }

                parts.push(InterpolPart::Expr(expr));
                continue;
            }
            // Lone `$` before non-identifier — literal dollar.
            text.push(c);
            continue;
        }

        text.push(c);
    }

    flush(&mut parts, &mut text);
    parts
}

/// Re-enter the PHP pest grammar on a `{$...}` inner expression.
fn parse_interpol_expression(src: &str) -> Result<Expression, String> {
    use pest::Parser;
    let mut pairs = super::PhpParser::parse(super::Rule::expression, src)
        .map_err(|e| format!("interpolation expr parse failed: {}", e))?;
    let pair = pairs.next().ok_or_else(|| "empty interpolation expression".to_string())?;
    walk_expression(pair)
}

// ─── Helpers ──────────────────────────────────────────────────────────────

/// Normalize PHP function-call argument order into the canonical common-AST
/// convention, which matches the JS / Component-Model shape that the
/// compiler emits. PHP builtins whose signature differs from JS need
/// their args rewritten at the walker layer so the downstream compiler
/// sees ONE canonical shape per operation.
///
/// Entries in the match table:
///   ("php_name", &[arg_indices...]) — each arg_indices entry selects
///   which position in the original PHP call the canonical form takes.
///   E.g. `("array_key_exists", &[1, 0])` means the canonical
///   (container, key) order pulls arg 1 first, arg 0 second.
fn canonicalize_php_call_args(callee: &Expression, args: Vec<Argument>) -> Vec<Argument> {
    let name = match &callee.kind {
        ExprKind::Ident(n) => n.as_str(),
        _ => return args,
    };
    let order: &[usize] = match name {
        // PHP: array_key_exists($key, $arr). Canonical: hasOwn($arr, $key).
        "array_key_exists" | "key_exists" => &[1, 0],
        // PHP: in_array($needle, $haystack). Canonical: includes($arr, $needle).
        // Note: `in_array` has an optional 3rd arg (strict); pass through.
        "in_array" => &[1, 0, 2],
        _ => return args,
    };
    if args.len() < order.iter().filter(|&&i| i < args.len()).count() {
        return args;
    }
    let mut out = Vec::with_capacity(order.len());
    for &i in order {
        if let Some(a) = args.get(i).cloned() {
            out.push(a);
        }
    }
    out
}

fn strip_dollar(s: &str) -> &str {
    s.strip_prefix('$').unwrap_or(s)
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (start_line, start_col) = s.start_pos().line_col();
    let (end_line, end_col) = s.end_pos().line_col();
    Span {
        start_line: start_line as u32,
        start_col: start_col as u32,
        end_line: end_line as u32,
        end_col: end_col as u32,
    }
}
