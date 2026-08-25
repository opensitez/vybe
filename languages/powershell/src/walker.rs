//! PowerShell walker.
//!
//! Transforms Pest parse trees into the shared JS-shaped AST used by compiler
//! primitives.

use super::Rule;
use pest::Parser;
use pest::iterators::Pair;
use vybe_ast::*;

use std::collections::HashMap;
use std::collections::VecDeque;

pub fn parse(source: &str) -> Result<Module, String> {
    // Every registry this walk keeps, created here and dropped when `parse`
    // returns — including on the `?` paths below.
    let mut __w_owned = PsWalker::default();
    let __w = &mut __w_owned;
    let source = normalize_source(source);
    let mut pairs = super::PowerShellParser::parse(Rule::program, &source)
        .map_err(|e| format!("Parse error: {e}"))?;

    let pair = pairs
        .next()
        .ok_or_else(|| "No PowerShell parse root".to_string())?;

    // Which method names this script's own classes declare, collected BEFORE
    // the walk so a rewrite never captures a call meant for a user method.
    collect_declared_methods(__w, pair.clone());

    let mut body = Vec::new();

    for child in pair.into_inner() {
        if let Some(stmt) = parse_statement(__w, child)? {
            body.push(stmt);
        }
    }

    Ok(Module {
        canon: Default::default(),
        name: "main".into(),
        language: Lang::Unknown,
        body: apply_aliases(apply_traps(body)),
        imports: Vec::new(),
        directives: vybe_ast::Directives {
            // PowerShell is case-insensitive throughout, ASCII — `$Path` and
            // `$path` are one variable, `Get-Item` and `get-item` one command.
            variable_case: Some(vybe_ast::CaseMatch::Folded),
            callable_case: Some(vybe_ast::CaseMatch::Folded),
            case_alphabet: Some(vybe_ast::CaseAlphabet::Ascii),
            ..Default::default()
        },
    })
}

fn normalize_source(source: &str) -> String {
    let src = source.trim_start_matches('\u{feff}').replace("\r\n", "\n");
    let mut out = String::new();
    let mut lines = src.lines().peekable();

    while let Some(raw) = lines.next() {
        let trimmed = raw.trim_end();
        if trimmed.ends_with('`') {
            out.push_str(trimmed.trim_end_matches('`'));
            if lines.peek().is_some() {
                out.push(' ');
            }
        } else {
            out.push_str(trimmed);
            if lines.peek().is_some() {
                out.push('\n');
            }
        }
    }

    out
}

fn parse_statement(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let mut queue = VecDeque::new();
    queue.push_back(pair);

    while let Some(pair) = queue.pop_front() {
        match pair.as_rule() {
            Rule::statement => {
                let text = pair.as_str().trim();
                if looks_like_command_invocation(text) && text.contains(' ') {
                    if let Some(expr) = parse_command_line(__w, text) {
                        return Ok(Some(Statement::new(StmtKind::Expr(expr))));
                    }
                }

                let children: Vec<Pair<Rule>> = pair.into_inner().collect();
                let mut tokens = Vec::new();
                for child in &children {
                    collect_command_tokens(child, &mut tokens);
                }

                if let Some(expr) = parse_command_tokens_as_expr(__w, &tokens) {
                    return Ok(Some(Statement::new(StmtKind::Expr(expr))));
                }

                for child in children {
                    queue.push_back(child);
                }
            }
            Rule::COMMENT => return Ok(None),
            Rule::namespace_decl => return Ok(Some(parse_namespace_decl(__w, pair)?)),
            Rule::class_decl => return Ok(Some(parse_class_decl(__w, pair)?)),
            Rule::enum_decl => return Ok(Some(parse_enum_decl(__w, pair)?)),
            Rule::function_decl => return Ok(Some(parse_function_decl(__w, pair)?)),
            Rule::if_stmt => return Ok(Some(parse_if_stmt(__w, pair)?)),
            Rule::switch_stmt => return Ok(Some(parse_switch_stmt(__w, pair)?)),
            Rule::foreach_stmt => return Ok(Some(parse_foreach_stmt(__w, pair)?)),
            Rule::for_stmt => return Ok(Some(parse_for_stmt(__w, pair)?)),
            Rule::while_stmt => return Ok(Some(parse_while_stmt(__w, pair)?)),
            Rule::do_while_stmt => return Ok(Some(parse_do_while_stmt(__w, pair)?)),
            Rule::try_stmt => return Ok(Some(parse_try_stmt(__w, pair)?)),
            Rule::return_stmt => return Ok(Some(parse_return_stmt(__w, pair))),
            Rule::throw_stmt => return Ok(Some(parse_throw_stmt(__w, pair))),
            Rule::break_stmt => return Ok(Some(parse_break_stmt(pair))),
            Rule::continue_stmt => return Ok(Some(parse_continue_stmt(pair))),
            Rule::param_stmt => return Ok(None),
            // A bare label carries no runtime behaviour of its own.
            Rule::label_decl => return Ok(None),
            Rule::labeled_stmt => return parse_labeled_stmt(__w, pair),
            Rule::trap_stmt => return parse_trap_stmt(__w, pair).map(Some),
            Rule::named_block => {
                // `begin`/`process`/`end` bodies run in declaration order.
                let stmts = parse_block_statements(__w, 
                    pair.into_inner()
                        .find(|c| c.as_rule() == Rule::block)
                        .ok_or_else(|| "named block without body".to_string())?,
                )?;
                return Ok(Some(Statement::new(StmtKind::Block(stmts))));
            }
            // `[CmdletBinding()]`, `[OutputType(…)]` — declaration metadata that
            // carries no runtime behaviour.
            Rule::attribute_stmt => return Ok(None),
            Rule::using_stmt => return Ok(None),
            Rule::assignment_stmt => return Ok(Some(parse_assignment_statement(__w, pair))),
            Rule::increment_stmt => return Ok(Some(parse_increment_statement(__w, pair))),
            Rule::expr_stmt => {
                let expr = pair
                    .into_inner()
                    .next()
                    .map(|__x| walk_expr(__w, __x))
                    .unwrap_or_else(Expression::null);
                return Ok(Some(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::command_stmt => return Ok(Some(parse_command_statement(__w, pair))),
            _ => {
                if let Some(stmt) = parse_statement_fallback(__w, pair)? {
                    return Ok(Some(stmt));
                }
            }
        }
    }

    Ok(None)
}

fn parse_statement_fallback(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let body = pair.as_str().trim();
    if body.is_empty() {
        return Ok(None);
    }
    if pair.as_rule() == Rule::expression {
        return Ok(Some(Statement::new(StmtKind::Expr(walk_expr(__w, pair)))));
    }
    let _ = body;
    Ok(None)
}

/// `:outer while (…) { … }` — attach the label to the loop it introduces.
fn parse_labeled_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let mut label = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::label_decl => {
                label = Some(child.as_str().trim_start_matches(':').to_string());
            }
            _ => {
                let stmt = match parse_statement(__w, child)? {
                    Some(stmt) => stmt,
                    None => continue,
                };
                return Ok(Some(match label.take() {
                    Some(label) => Statement::new(StmtKind::Labeled {
                        label,
                        body: Box::new(stmt),
                    }),
                    None => stmt,
                }));
            }
        }
    }
    Ok(None)
}

/// `trap { … }` is a statement-level handler for anything thrown after it, so
/// it lowers to a try/catch wrapping the rest of the enclosing block. Here it
/// becomes the catch half; the walker cannot see the remainder, so the body is
/// emitted as a catch-all handler.
/// The variable a `trap` handler binds, and the marker that identifies a parsed
/// `trap` before [`apply_traps`] turns it into real handlers.
const TRAP_VAR: &str = "__trap";

fn parse_trap_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::block {
            body = parse_block_statements(__w, child)?;
        }
    }

    // `continue` in a trap means "resume at the next statement" — which is
    // already what falling out of the handler does under the lowering below.
    // `break` means "give up and rethrow". Neither is the LOOP keyword, so
    // leaving them in place would make the shared compiler read them as one.
    let body = body
        .into_iter()
        .map(|stmt| match stmt.kind {
            StmtKind::Break { .. } => Statement::new(StmtKind::Throw {
                expr: Some(Expression::ident(TRAP_VAR)),
                cause: None,
            }),
            _ => stmt,
        })
        .filter(|stmt| !matches!(stmt.kind, StmtKind::Continue { .. }))
        .collect();

    // A marker, not the final shape: an empty `try` guarding nothing. Which
    // statements it guards is not known until the enclosing list is complete,
    // so `apply_traps` rewrites it there.
    Ok(Statement::new(StmtKind::Try {
        body: Vec::new(),
        catches: vec![CatchClause {
            types: Vec::new(),
            var_name: Some(TRAP_VAR.to_string()),
            stack_var: None,
            body,
            when_clause: None,
        }],
        else_body: None,
        finally: None,
    }))
}

/// `trap { … }` guards every statement that FOLLOWS it in its own block, and
/// execution resumes at the statement after the one that threw. So each of
/// those statements gets its own `try`/`catch` rather than one `try` around the
/// remainder — wrapping the remainder once would skip everything after the
/// throw, which is what `catch` means and precisely not what `trap` means.
fn apply_traps(stmts: Vec<Statement>) -> Vec<Statement> {
    let Some(at) = stmts.iter().position(is_trap_marker) else {
        return stmts;
    };

    let mut out: Vec<Statement> = Vec::with_capacity(stmts.len());
    let mut rest = stmts;
    let tail = rest.split_off(at);
    out.append(&mut rest);

    let mut tail = tail.into_iter();
    let handler = match tail.next().map(|s| s.kind) {
        Some(StmtKind::Try { catches, .. }) => catches,
        _ => return out,
    };

    // Recurse first, so a second `trap` later in the same block installs its own
    // handler before this one wraps the statements around it.
    for stmt in apply_traps(tail.collect()) {
        out.push(Statement::new(StmtKind::Try {
            body: vec![stmt],
            catches: handler.clone(),
            else_body: None,
            finally: None,
        }));
    }
    out
}

fn drop_trailing_break(mut body: Vec<Statement>) -> Vec<Statement> {
    if matches!(body.last().map(|s| &s.kind), Some(StmtKind::Break { .. })) {
        body.pop();
    }
    body
}

fn is_trap_marker(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Try { body, catches, .. } => {
            body.is_empty()
                && catches.len() == 1
                && catches[0].var_name.as_deref() == Some(TRAP_VAR)
        }
        _ => false,
    }
}

/// The name of the accumulator [`accumulate_outputs`] introduces.
const OUT_ACC: &str = "__ps_output";

/// PowerShell's SUCCESS STREAM: every value a function body produces and does
/// not consume is output, and the caller receives all of them. `function F { 'a';
/// return 'b' }` returns two values, not one.
///
/// This is not output buffering — the stream carries live objects, not rendered
/// bytes (`$r[0].X` must reach a property). So it is a value mechanism: collect
/// the emitted values and hand back the collection, unwrapped to a scalar when
/// there is exactly one, `$null` when there are none.
///
/// Applied ONLY to a body with two or more emit sites. A body with one or none
/// already lowers correctly through the plain trailing-return path, and leaving
/// it alone keeps this transformation off every function that does not need it.
fn accumulate_outputs(body: Vec<Statement>) -> Option<Vec<Statement>> {
    if body.iter().filter(|s| emits_value(s)).count() < 2 {
        return None;
    }

    let mut out = Vec::with_capacity(body.len() + 2);
    out.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(OUT_ACC)],
        value: Expression::new(ExprKind::Array(Vec::new())),
        by_ref: false,
    }));

    let mut returned = false;
    for stmt in body {
        match stmt.kind {
            StmtKind::Expr(expr) if expression_emits(&expr) => {
                out.push(Statement::new(StmtKind::Expr(collect_call(expr))));
            }
            StmtKind::Return(Some(expr)) => {
                out.push(Statement::new(StmtKind::Expr(collect_call(expr))));
                out.push(Statement::new(StmtKind::Return(Some(unwrap_output()))));
                returned = true;
                break;
            }
            _ => out.push(stmt),
        }
    }

    if !returned {
        out.push(Statement::new(StmtKind::Return(Some(unwrap_output()))));
    }
    Some(out)
}

fn collect_call(value: Expression) -> Expression {
    method_call_expr(Expression::ident(OUT_ACC), "Add", vec![value])
}

/// `$out.Count -eq 0 ? $null : ($out.Count -eq 1 ? $out[0] : $out)` — PowerShell
/// hands back a scalar for a single value and `$null` for none, and only wraps
/// in an array from two upward.
fn unwrap_output() -> Expression {
    let count = || {
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(OUT_ACC)),
            field: "Count".to_string(),
            null_safe: false,
        })
    };
    let count_is = |n: i64| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(count()),
            right: Box::new(Expression::int(n)),
        })
    };

    Expression::new(ExprKind::Ternary {
        cond: Box::new(count_is(0)),
        then: Box::new(Expression::null()),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(count_is(1)),
            then: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(OUT_ACC)),
                index: Box::new(Expression::int(0)),
                null_safe: false,
            })),
            else_: Box::new(Expression::ident(OUT_ACC)),
        })),
    })
}

fn emits_value(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => expression_emits(expr),
        StmtKind::Return(Some(_)) => true,
        _ => false,
    }
}

/// Whether evaluating this as a statement contributes to the success stream.
/// `Write-Host` and friends write to the HOST — a different stream, which the
/// caller never captures — and an assignment consumes its own value.
fn expression_emits(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Assign { .. } => false,
        ExprKind::Call { callee, .. } => match &callee.kind {
            ExprKind::Ident(name) => !writes_to_host(name),
            _ => true,
        },
        _ => true,
    }
}

fn writes_to_host(name: &str) -> bool {
    matches!(
        name.to_lowercase().as_str(),
        "write-host"
            | "write-verbose"
            | "write-warning"
            | "write-debug"
            | "write-information"
            | "write-error"
            | "write-progress"
            | "out-null"
            | "out-host"
            | "set-alias"
            | "new-alias"
            | "remove-alias"
            | "set-variable"
            | "start-sleep"
    )
}

/// `Set-Alias hi Write-Output` makes `hi 'x'` mean `Write-Output 'x'`. The
/// alias is a NAME rewrite, so it is resolved here rather than becoming a
/// runtime lookup table: collect every alias the script defines, then retarget
/// the calls that use one. `Set-Alias` itself becomes a no-op — the mapping it
/// carried has already been applied.
fn apply_aliases(mut body: Vec<Statement>) -> Vec<Statement> {
    let mut aliases: HashMap<String, String> = HashMap::new();
    for stmt in &body {
        collect_aliases(stmt, &mut aliases);
    }
    if aliases.is_empty() {
        return body;
    }

    for stmt in &mut body {
        stmt.walk_exprs_mut(&mut |expr| {
            let ExprKind::Call { callee, .. } = &mut expr.kind else {
                return;
            };
            let ExprKind::Ident(name) = &callee.kind else {
                return;
            };
            if let Some(target) = aliases.get(&name.to_lowercase()) {
                *callee = Box::new(Expression::ident(target));
            }
        });
    }
    body
}

fn collect_aliases(stmt: &Statement, out: &mut HashMap<String, String>) {
    let mut visit = |expr: &mut Expression| {
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return;
        };
        let ExprKind::Ident(name) = &callee.kind else {
            return;
        };
        if !matches!(name.to_lowercase().as_str(), "set-alias" | "new-alias") {
            return;
        }
        let named = |key: &str| {
            args.iter()
                .find(|a| {
                    a.name
                        .as_deref()
                        .is_some_and(|n| n.eq_ignore_ascii_case(key))
                })
                .map(|a| a.value.clone())
        };
        let positional: Vec<&Argument> = args.iter().filter(|a| a.name.is_none()).collect();
        let alias = named("Name").or_else(|| positional.first().map(|a| a.value.clone()));
        let target = named("Value").or_else(|| positional.get(1).map(|a| a.value.clone()));
        if let (Some(alias), Some(target)) = (alias, target) {
            if let (Some(a), Some(t)) = (literal_text(&alias), literal_text(&target)) {
                out.insert(a.to_lowercase(), t);
            }
        }
    };
    // `walk_exprs_mut` on a clone: the collection pass only READS, and a
    // shared read-only walker does not exist.
    let mut probe = stmt.clone();
    probe.walk_exprs_mut(&mut visit);
}

/// `enum Color { Red; Green = 5; Blue }`. A member without an explicit value
/// takes the previous member's value plus one, starting at 0 — the same rule
/// .NET uses, filled in here so the shared `EnumDecl` sees every value.
fn parse_enum_decl(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut members: Vec<EnumMember> = Vec::new();
    let mut next_value: i64 = 0;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::IDENT if name.is_empty() => name = child.as_str().trim().to_string(),
            Rule::enum_member => {
                let mut member_name = String::new();
                let mut value = None;
                for part in child.into_inner() {
                    match part.as_rule() {
                        Rule::IDENT if member_name.is_empty() => {
                            member_name = part.as_str().trim().to_string();
                        }
                        Rule::expression | Rule::ternary_expr => {
                            value = Some(walk_expr(__w, part));
                        }
                        _ => {}
                    }
                }
                if member_name.is_empty() {
                    continue;
                }
                if let Some(ExprKind::Lit(Literal::Int(n))) = value.as_ref().map(|e| &e.kind) {
                    next_value = *n;
                }
                let value = value.unwrap_or_else(|| Expression::int(next_value));
                next_value += 1;
                members.push(EnumMember {
                    name: member_name,
                    value: Some(value),
                    constructor_args: Vec::new(),
                });
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::EnumDecl {
        name,
        members,
        visibility: Visibility::Public,
        is_flags: false,
        backing_type: None,
        interfaces: Vec::new(),
        body_members: Vec::new(),
        decorators: Vec::new(),
    }))
}

fn parse_namespace_decl(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::namespace_name | Rule::DOTTED_NAME => {
                name = child.as_str().trim().to_string();
            }
            Rule::block => body = parse_block_statements(__w, child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::NamespaceDecl { name, body }))
}

fn parse_block_statements(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for child in pair.into_inner() {
        if let Some(stmt) = parse_statement(__w, child)? {
            body.push(stmt);
        }
    }
    Ok(apply_traps(body))
}

fn parse_class_decl(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let interfaces = Vec::new();
    let mut members = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::IDENT => {
                name = child.as_str().trim().to_string();
            }
            // Walk the DOTTED_NAME children, not the rule's raw text — the text
            // still carries the leading `:` and would yield a parent named
            // ": A" that resolves to nothing.
            Rule::class_heritage => {
                for base in child.into_inner() {
                    let name = base.as_str().trim();
                    if !name.is_empty() {
                        parents.push(name.to_string());
                    }
                }
            }
            Rule::class_body => {
                members = parse_class_body(__w, child)?;
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: ClassModifiers::default(),
        decorators: Vec::new(),
    }))
}

fn parse_class_body(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for child in pair.into_inner() {
        if let Some(member) = parse_class_member(__w, child)? {
            members.push(member);
        }
    }
    Ok(members)
}

/// `[string] Speak() { … }` — a method, identified by its declared return type.
fn parse_ps_method(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::class_modifier => {
                if child.as_str().eq_ignore_ascii_case("static") {
                    is_static = true;
                }
            }
            Rule::member_name => name = child.as_str().to_string(),
            Rule::function_params => params = parse_function_params(__w, child),
            Rule::block => {
                body = implicit_return(parse_block_with_function_params(__w, child, &mut params)?)
            }
            _ => {}
        }
    }

    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            body,
            modifiers,
            is_async: false,
            is_generator: false,
            is_sub: false,
            return_type: None,
            handles: Vec::new(),
        },
    ))))
}

/// `Animal([string]$name) { … }` — a constructor: class-named, no return type.
fn parse_ps_constructor(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::IDENT => name = Some(child.as_str().to_string()),
            Rule::function_params => params = parse_function_params(__w, child),
            Rule::ctor_base => {
                let args = child
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::arg_list)
                    .map(|__x| walk_arg_list(__w, __x))
                    .unwrap_or_default();
                base_args = Some(args);
            }
            Rule::block => body = parse_block_with_function_params(__w, child, &mut params)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name,
        params,
        body,
        base_args,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

/// `[string]$Name = 'x'` — a field, with an optional type and initialiser.
fn parse_ps_property(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut type_hint = None;
    let mut name = String::new();
    let mut init = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::class_modifier => {
                if child.as_str().eq_ignore_ascii_case("static") {
                    is_static = true;
                }
            }
            Rule::type_literal => {
                let t = type_literal_name(child.as_str()).trim();
                if !t.is_empty() {
                    type_hint = Some(t.to_string());
                }
            }
            Rule::var_ref => {
                name = scope_qualified_name(child.as_str().trim_start_matches('$')).to_string();
            }
            _ => init = Some(walk_expr(__w, child)),
        }
    }

    let mut modifiers = Modifiers::default();
    modifiers.is_static = is_static;

    Ok(ClassMember::Field {
        name,
        type_hint,
        init,
        modifiers,
        with_events: false,
        array_bounds: None,
        storage: None,
    })
}

fn parse_class_member(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Option<ClassMember>, String> {
    match pair.as_rule() {
        Rule::ps_method => parse_ps_method(__w, pair).map(Some),
        Rule::ps_constructor => parse_ps_constructor(__w, pair).map(Some),
        Rule::ps_property => parse_ps_property(__w, pair).map(Some),
        Rule::class_function_decl => {
            let statement = parse_function_decl(__w, pair)?;
            Ok(Some(ClassMember::Method(Box::new(statement))))
        }
        Rule::constructor_decl => parse_constructor_decl(__w, pair).map(Some),
        _ => Ok(None),
    }
}

fn parse_constructor_decl(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::function_params => params = parse_function_params(__w, child),
            Rule::block => body = parse_block_with_function_params(__w, child, &mut params)?,
            _ => {}
        }
    }

    Ok(ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

fn parse_function_decl(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::function_name => {
                name = child.as_str().trim().to_string();
            }
            Rule::function_params => params = parse_function_params(__w, child),
            Rule::block => {
                let locals = function_local_names(&child);
                let wants_args = mentions_args(child.as_str());
                // `return_last_of_branches`, not `implicit_return`: a function
                // ending in `if ($x) { 'yes' }` outputs `'yes'`, so the trailing
                // expression of each BRANCH is the return value, not just a
                // trailing expression statement.
                let parsed = parse_block_with_function_params(__w, child, &mut params)?;
                body = match accumulate_outputs(parsed.clone()) {
                    Some(accumulated) => accumulated,
                    None => return_last_of_branches(parsed),
                };
                // `$args` is the automatic variable holding every argument the
                // declared parameters did not take — a rest parameter, which is
                // a shape the compiler already has.
                if wants_args && !params.iter().any(|p| p.name.eq_ignore_ascii_case("args")) {
                    params.push(Param {
                        name: "args".to_string(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: true,
                        is_kwargs: false,
                        is_optional: true,
                        is_nullable: false,
                    });
                }
                body = declare_function_locals(locals, &params, body);
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

/// PowerShell function scoping: a function READS the caller's variables, but an
/// ASSIGNMENT always creates a local — `function F { $x = 99 }` never touches
/// the caller's `$x`. Without this the shared `emit_var_set` falls through to
/// the module global and the function mutates its caller.
///
/// Expressed as an explicit declaration rather than a scope policy because the
/// rule is per-NAME, not per-scope: `ScopeDeclKind::Closed` would also cut off
/// reads, which PowerShell allows. Declaring exactly the assigned names leaves
/// every other name resolving outward as before.
///
/// `$script:x` / `$global:x` name the outer storage on purpose and are excluded.
fn function_local_names(block: &Pair<Rule>) -> Vec<String> {
    let mut out = Vec::new();
    collect_assigned_locals(block.clone(), &mut out);
    out
}

fn collect_assigned_locals(pair: Pair<Rule>, out: &mut Vec<String>) {
    match pair.as_rule() {
        // A nested function owns its own locals; hoisting them here would make
        // the inner function write the outer one's frame.
        Rule::function_decl => return,
        Rule::assignment_stmt | Rule::increment_stmt => {
            for child in pair.clone().into_inner() {
                if child.as_rule() != Rule::lvalue {
                    continue;
                }
                // Only a BARE `$name` is a local. An indexed or member target
                // (`$h['k'] = 1`) mutates something that already exists, and
                // declaring the base would shadow it with an empty local.
                let mut parts = child.into_inner();
                let Some(first) = parts.next() else { continue };
                if first.as_rule() != Rule::var_ref || parts.next().is_some() {
                    continue;
                }
                let raw = first.as_str().trim().trim_start_matches('$');
                let name = raw.trim_matches(|c| c == '{' || c == '}');
                if name.contains(':') {
                    continue;
                }
                if !name.is_empty() && !out.iter().any(|n| n == name) {
                    out.push(name.to_string());
                }
            }
        }
        _ => {}
    }

    for child in pair.into_inner() {
        collect_assigned_locals(child, out);
    }
}

fn declare_function_locals(
    names: Vec<String>,
    params: &[Param],
    body: Vec<Statement>,
) -> Vec<Statement> {
    let declarations: Vec<VarDeclarator> = names
        .into_iter()
        .filter(|n| !params.iter().any(|p| p.name.eq_ignore_ascii_case(n)))
        .map(|name| VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint: None,
            init: None,
            array_bounds: None,
            with_events: false,
        })
        .collect();

    if declarations.is_empty() {
        return body;
    }

    let mut out = Vec::with_capacity(body.len() + 1);
    out.push(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Var,
    }));
    out.extend(body);
    out
}

fn parse_function_params(__w: &mut PsWalker, pair: Pair<Rule>) -> Vec<Param> {
    let mut out = Vec::new();
    let mut param_nodes = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::function_param_list || child.as_rule() == Rule::function_param {
            param_nodes.push(child);
        }
    }

    if param_nodes.is_empty() {
        return out;
    }

    for raw in param_nodes.into_iter().flat_map(|node| node.into_inner()) {
        if raw.as_rule() != Rule::function_param {
            continue;
        }

        let mut name = String::new();
        let mut default = None;
        let mut type_hint = None;

        for piece in raw.into_inner() {
            match piece.as_rule() {
                Rule::type_literal => {
                    let inner = type_literal_name(piece.as_str()).trim();
                    if !inner.is_empty() {
                        type_hint = Some(inner.to_string());
                    }
                }
                Rule::var_ref | Rule::IDENT => {
                    name = scope_qualified_name(piece.as_str().trim_start_matches('$')).to_string();
                }
                // The grammar spells a default as `ternary_expr`, not
                // `expression` — matching only the latter dropped every
                // `param($name = "World")` default on the floor.
                Rule::expression | Rule::ternary_expr => {
                    default = Some(walk_expr(__w, piece));
                }
                _ => {}
            }
        }

        if !name.is_empty() {
            // Keep `default` readable for metadata checks before moving into `Param`.
            let is_optional = default.as_ref().is_some();
            out.push(Param {
                name,
                type_hint: type_hint.map(Into::into),
                default,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional,
                is_nullable: false,
            });
        }
    }
    out
}

fn parse_param_stmt(__w: &mut PsWalker, pair: Pair<Rule>, out: &mut Vec<Param>) {
    // Hand `parse_function_params` the ENCLOSING pair, not the
    // `function_param_list` itself: it descends two levels, so passing the list
    // skipped straight past every `function_param`.
    out.extend(parse_function_params(__w, pair));
}

fn parse_block_with_function_params(__w: &mut PsWalker, 
    pair: Pair<Rule>,
    params: &mut Vec<Param>,
) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();

    // `statement` is a SILENT rule, so a block's children are the concrete
    // statement rules (`assignment_stmt`, `command_stmt`, …) and never
    // `Rule::statement` itself. Matching on that name here skipped every
    // statement and left function and constructor bodies empty.
    for child in pair.into_inner() {
        if child.as_rule() == Rule::param_stmt {
            parse_param_stmt(__w, child, params);
            continue;
        }

        if let Some(stmt) = parse_statement(__w, child)? {
            body.push(stmt);
        }
    }

    Ok(body)
}

fn parse_if_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::condition_expr => cond = Some(walk_condition(__w, child)),
            Rule::block => then_body = parse_block_statements(__w, child)?,
            Rule::elseif_stmt => {
                let mut branch_cond = None;
                let mut branch_body = Vec::new();
                for part in child.into_inner() {
                    if part.as_rule() == Rule::condition_expr {
                        branch_cond = Some(walk_condition(__w, part));
                    } else if part.as_rule() == Rule::block {
                        branch_body = parse_block_statements(__w, part)?;
                    }
                }
                if let Some(branch_cond) = branch_cond {
                    elifs.push((branch_cond, branch_body));
                }
            }
            Rule::else_stmt => {
                if let Some(block) = child.into_inner().find(|c| c.as_rule() == Rule::block) {
                    else_body = Some(parse_block_statements(__w, block)?);
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(Expression::null),
        then_body,
        elifs,
        else_body,
    }))
}

/// `condition_expr` wraps a parenthesised expression — unwrap to the inner
/// `expression` pair rather than re-parsing the text.
fn walk_condition(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    pair.into_inner()
        .find(|c| matches!(c.as_rule(), Rule::expression | Rule::command_pipeline))
        .map(|__x| walk_expr(__w, __x))
        .unwrap_or_else(Expression::null)
}

fn parse_foreach_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut iter_var = String::new();
    let mut iter = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::var_ref => {
                iter_var = scope_qualified_name(child.as_str().trim_start_matches('$')).to_string()
            }
            Rule::expression | Rule::command_pipeline => iter = Some(walk_expr(__w, child)),
            Rule::block => body = parse_block_statements(__w, child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::ForIn {
        var: iter_var,
        key: None,
        iter: iter.unwrap_or_else(Expression::null),
        body,
        of: true,
        else_body: None,
        is_async: false,
    }))
}

fn parse_switch_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut expr = None;
    let mut cases = Vec::new();
    let mut default = None;
    let mut matcher = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            // `switch -Regex (…)` / `-Wildcard (…)` change how each case
            // CONDITION is tested against the subject: pattern match, not
            // equality.
            Rule::switch_flag => {
                matcher = match child
                    .as_str()
                    .trim_start_matches('-')
                    .to_lowercase()
                    .as_str()
                {
                    "regex" => Some(false),
                    "wildcard" => Some(true),
                    _ => matcher,
                };
            }
            Rule::expression => {
                expr = Some(walk_expr(__w, child));
            }
            Rule::switch_case => {
                let mut branch_body = None;
                let mut branch_conditions = Vec::new();
                let is_default = child
                    .as_str()
                    .trim_start()
                    .to_lowercase()
                    .starts_with("default");

                for part in child.into_inner() {
                    match part.as_rule() {
                        Rule::switch_case_body => {
                            branch_body = Some(parse_block_statements(__w, part)?);
                        }
                        Rule::switch_default_case => {
                            for inner in part.into_inner() {
                                if inner.as_rule() == Rule::switch_case_body {
                                    branch_body = Some(parse_block_statements(__w, inner)?);
                                    break;
                                }
                            }
                        }
                        Rule::switch_case_value => {
                            for inner in part.into_inner() {
                                match inner.as_rule() {
                                    Rule::expression => {
                                        branch_conditions
                                            .push(CaseCondition::Value(walk_expr(__w, inner)));
                                    }
                                    Rule::switch_case_body => {
                                        branch_body = Some(parse_block_statements(__w, inner)?);
                                    }
                                    _ => {}
                                }
                            }
                        }
                        _ => {}
                    }
                }

                if is_default {
                    default = branch_body.or(Some(Vec::new()));
                    continue;
                }

                if let Some(body) = branch_body {
                    cases.push(SwitchCase {
                        conditions: branch_conditions,
                        body,
                    });
                }
            }
            _ => {}
        }
    }

    let subject = expr.unwrap_or_else(Expression::null);

    // A pattern switch is not a `Switch` at all — the shared node compares each
    // condition for EQUALITY. Lowered to the if/elseif chain it stands for, so
    // the pattern test is the same `Regex.IsMatch` `-match` / `-like` use.
    if let Some(is_wildcard) = matcher {
        return Ok(pattern_switch(subject, cases, default, is_wildcard));
    }

    Ok(Statement::new(StmtKind::Switch {
        expr: subject,
        cases,
        default,
    }))
}

fn pattern_switch(
    subject: Expression,
    cases: Vec<SwitchCase>,
    default: Option<Vec<Statement>>,
    is_wildcard: bool,
) -> Statement {
    let test = |cond: CaseCondition| -> Expression {
        let pattern = match cond {
            CaseCondition::Value(p) => p,
            _ => Expression::null(),
        };
        let pattern = if is_wildcard {
            glob_to_regex(pattern)
        } else {
            pattern
        };
        dotnet_static_call(
            "System.Text.RegularExpressions.Regex",
            "IsMatch",
            vec![subject.clone(), pattern],
        )
    };

    let mut arms: Vec<(Expression, Vec<Statement>)> = Vec::new();
    for case in cases {
        let mut conds = case.conditions.into_iter().map(test);
        let Some(first) = conds.next() else { continue };
        let cond = conds.fold(first, |acc, next| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(acc),
                right: Box::new(next),
            })
        });
        arms.push((cond, case.body));
    }

    let mut arms = arms.into_iter();
    let Some((cond, then_body)) = arms.next() else {
        return Statement::new(StmtKind::Block(default.unwrap_or_default()));
    };

    Statement::new(StmtKind::If {
        cond,
        then_body,
        elifs: arms.collect(),
        else_body: default,
    })
}

fn parse_for_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut init = None;
    let mut cond = None;
    let mut update = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::for_init => {
                if let Some(inner) = child.into_inner().next() {
                    init = parse_statement(__w, inner)?;
                }
            }
            Rule::for_cond => cond = Some(walk_expr(__w, child)),
            Rule::for_update => {
                if let Some(inner) = child.into_inner().next() {
                    update = parse_statement(__w, inner)?;
                }
            }
            Rule::block => body = parse_block_statements(__w, child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::For {
        init: init.map(Box::new),
        cond,
        update: update.map(statement_to_expression),
        body,
    }))
}

/// `StmtKind::For.update` is an `Expression`, but the update slot parses as a
/// statement (`$i++`, `$i = $i + 1`). Re-shape assignment statements into the
/// equivalent assignment expressions the shared loop lowering expects.
fn statement_to_expression(stmt: Statement) -> Expression {
    match stmt.kind {
        StmtKind::Expr(expr) => expr,
        StmtKind::Assign {
            mut targets, value, ..
        } => {
            let target = if targets.is_empty() {
                Expression::null()
            } else {
                targets.remove(0)
            };
            Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            })
        }
        StmtKind::CompoundAssign { target, op, value } => {
            let bin = compound_to_binop(op);
            binary_assign(target, bin, value)
        }
        _ => Expression::null(),
    }
}

/// A statement used where a VALUE is wanted: `$r = switch ($v) { 1 { 'one' } }`.
/// PowerShell hands back whatever the taken branch output, so the statement
/// becomes the body of a zero-argument lambda that is called immediately and
/// every branch's trailing expression becomes a `return`. Nothing new is
/// introduced — `Lambda` + `Call` are shared nodes the compiler already emits.
fn statement_value_expr(stmt: Statement) -> Expression {
    let body = return_last_of_branches(vec![stmt]);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

/// `implicit_return` reaching INTO branches: the value of an `if` / `switch` /
/// `try` is the last expression of whichever branch ran, so each branch body
/// needs its own trailing `return`, not just the outermost statement list.
fn return_last_of_branches(body: Vec<Statement>) -> Vec<Statement> {
    let mut body = implicit_return(body);
    let Some(last) = body.pop() else {
        return body;
    };
    let rewritten = match last.kind {
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => Statement::new(StmtKind::If {
            cond,
            then_body: return_last_of_branches(then_body),
            elifs: elifs
                .into_iter()
                .map(|(c, b)| (c, return_last_of_branches(b)))
                .collect(),
            else_body: else_body.map(return_last_of_branches),
        }),
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => Statement::new(StmtKind::Switch {
            expr,
            cases: cases
                .into_iter()
                .map(|c| SwitchCase {
                    conditions: c.conditions,
                    // `1 { 'one'; break }` — the `break` stops the switch from
                    // testing later conditions, it is not the case's VALUE.
                    // Left in place it would be the last statement and the
                    // trailing `return` would land on it, so the case yielded
                    // nothing.
                    body: return_last_of_branches(drop_trailing_break(c.body)),
                })
                .collect(),
            default: default.map(return_last_of_branches),
        }),
        StmtKind::Try {
            body: try_body,
            catches,
            else_body,
            finally,
        } => Statement::new(StmtKind::Try {
            body: return_last_of_branches(try_body),
            catches: catches
                .into_iter()
                .map(|mut c| {
                    c.body = return_last_of_branches(c.body);
                    c
                })
                .collect(),
            else_body: else_body.map(return_last_of_branches),
            // `finally` runs for its effects; its last expression is not the
            // statement's value.
            finally,
        }),
        _ => last,
    };
    body.push(rewritten);
    body
}

fn compound_to_binop(op: CompoundOp) -> BinOp {
    match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        CompoundOp::Mul => BinOp::Mul,
        CompoundOp::Div => BinOp::Div,
        CompoundOp::Mod => BinOp::Mod,
        CompoundOp::Pow => BinOp::Pow,
        CompoundOp::BitAnd => BinOp::BitAnd,
        CompoundOp::BitOr => BinOp::BitOr,
        CompoundOp::BitXor => BinOp::BitXor,
        CompoundOp::Shl => BinOp::Shl,
        CompoundOp::Shr => BinOp::Shr,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        _ => BinOp::Add,
    }
}

fn binary_assign(target: Expression, op: BinOp, value: Expression) -> Expression {
    Expression::new(ExprKind::Assign {
        target: Box::new(target.clone()),
        value: Box::new(Expression::new(ExprKind::Binary {
            op,
            left: Box::new(target),
            right: Box::new(value),
        })),
    })
}

fn parse_while_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut cond = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::condition_expr => cond = Some(walk_condition(__w, child)),
            Rule::block => body = parse_block_statements(__w, child)?,
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::While {
        cond: cond.unwrap_or_else(Expression::null),
        body,
        else_body: None,
    }))
}

fn parse_do_while_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    let mut cond = None;
    let mut until = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::block => body = parse_block_statements(__w, child)?,
            Rule::do_while_while | Rule::do_while_until => {
                until = child.as_rule() == Rule::do_while_until;
                for inner in child.into_inner() {
                    if inner.as_rule() == Rule::condition_expr {
                        cond = Some(walk_condition(__w, inner));
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::DoWhile {
        body,
        cond: cond.unwrap_or_else(Expression::null),
        until,
    }))
}

fn parse_try_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::block => body = parse_block_statements(__w, child)?,
            Rule::catch_clause => catches.push(parse_catch_clause(__w, child)?),
            // The `block` INSIDE the clause, not the clause. `parse_block_statements`
            // walks its argument's children as statements, and a `finally_clause`'s
            // children are `kw_finally` and `block` — neither is a statement, so
            // every `finally` body came out empty and simply never ran.
            Rule::finally_clause => {
                finally = Some(
                    match child.into_inner().find(|p| p.as_rule() == Rule::block) {
                        Some(block) => parse_block_statements(__w, block)?,
                        None => Vec::new(),
                    },
                );
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    }))
}

fn parse_catch_clause(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<CatchClause, String> {
    let mut var_name = None;
    let mut body = Vec::new();

    for child in pair.into_inner() {
        if child.as_rule() == Rule::var_ref {
            let raw = child.as_str().trim().trim_start_matches('$');
            if !raw.is_empty() {
                var_name = Some(raw.to_string());
            }
        }
        if child.as_rule() == Rule::block {
            body = parse_block_statements(__w, child)?;
        }
    }

    Ok(CatchClause {
        types: Vec::new(),
        var_name,
        stack_var: None,
        body,
        when_clause: None,
    })
}

fn parse_return_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Statement {
    let expr = pair
        .into_inner()
        .find(|c| matches!(c.as_rule(), Rule::expression | Rule::command_pipeline))
        .map(|__x| walk_expr(__w, __x));
    Statement::new(StmtKind::Return(expr))
}

fn parse_throw_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Statement {
    let expr = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::expression)
        .map(|__x| walk_expr(__w, __x));
    Statement::new(StmtKind::Throw { expr, cause: None })
}

fn parse_break_stmt(pair: Pair<Rule>) -> Statement {
    let level = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::integer)
        .and_then(|c| c.as_str().parse::<u32>().ok());
    Statement::new(StmtKind::Break(match level {
        Some(level) => BreakTarget::Level(level),
        None => BreakTarget::Implicit,
    }))
}

fn parse_continue_stmt(pair: Pair<Rule>) -> Statement {
    let level = pair
        .into_inner()
        .find(|c| c.as_rule() == Rule::integer)
        .and_then(|c| c.as_str().parse::<u32>().ok());
    Statement::new(StmtKind::Continue(match level {
        Some(level) => ContinueTarget::Level(level),
        None => ContinueTarget::Implicit,
    }))
}

fn parse_command_statement(__w: &mut PsWalker, pair: Pair<Rule>) -> Statement {
    let text = pair.as_str().trim().to_string();
    if let Some(stmt) = parse_exit_command_statement(__w, &text) {
        return stmt;
    }
    // Prefer the PEST TREE over the text tokenizer. The text path splits on
    // spaces without tracking parens, so `Write-Host (I 5)` came apart into
    // `(I` and `5)`; the grammar already groups that correctly.
    let expr = parse_pipeline(__w, pair.clone())
        .or_else(|| parse_command_line(__w, &text))
        .unwrap_or_else(|| expr_from_text(__w, &text));
    Statement::new(StmtKind::Expr(expr))
}

fn parse_exit_command_statement(__w: &mut PsWalker, text: &str) -> Option<Statement> {
    let tokens = split_command_tokens(text);
    if tokens.is_empty() || !tokens[0].eq_ignore_ascii_case("exit") {
        return None;
    }
    let status = if tokens.len() > 1 {
        Some(expr_from_text(__w, &tokens[1..].join(" ")))
    } else {
        None
    };
    Some(Statement::new(StmtKind::Exit { status }))
}

fn parse_assignment_statement(__w: &mut PsWalker, pair: Pair<Rule>) -> Statement {
    let mut targets: Vec<Expression> = Vec::new();
    let mut op = "=".to_string();
    let mut rhs = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            // Several when the source destructures: `$a, $b = 1, 2`.
            Rule::lvalue => {
                targets.push(walk_lvalue(__w, child));
            }
            Rule::assignment_op => {
                op = child.as_str().to_string();
            }
            // `rhs_value` is a silent rule: the RHS arrives as either an
            // `expression` or a `command_pipeline` (`$x = Get-Item | …`).
            Rule::expression | Rule::command_pipeline => {
                // A lone bare word on the right of `=` is a COMMAND, not a name:
                // `$r = Test-Local` calls the function. The `expression` branch
                // of `rhs_value` matches first (a bare word is a valid primary),
                // so `command_pipeline` never sees it and `$r` was bound to the
                // function object itself.
                rhs = Some(match lone_bare_word(&child) {
                    Some(word) if !is_literal_word(&word) => Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(&word)),
                        args: Vec::new(),
                        optional: false,
                    }),
                    _ => walk_expr(__w, child),
                });
            }
            // `$r = switch (…) { … }` — the RHS is a statement used as a value.
            Rule::value_stmt => {
                rhs = child
                    .into_inner()
                    .next()
                    .and_then(|s| parse_statement(__w, s).ok().flatten())
                    .map(statement_value_expr);
            }
            _ => {}
        }
    }

    let value = rhs.unwrap_or_else(Expression::null);
    // Compound assignment has exactly one target; only `=` destructures.
    let target = targets.first().cloned().unwrap_or_else(Expression::null);
    // Several `targets` in a shared `Assign` means CHAINED assignment (`a = b =
    // c`) — every target takes the same value. `$a, $b = 1, 2` gives each target
    // one ELEMENT, which is `Destructure`, the node Python builds for `x, y = …`.
    let targets = if targets.len() > 1 {
        vec![Expression::new(ExprKind::Destructure(
            DestructurePattern::Array(
                targets
                    .into_iter()
                    .map(|t| {
                        let name = match &t.kind {
                            ExprKind::Ident(name) => name.clone(),
                            _ => String::new(),
                        };
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    })
                    .collect(),
            ),
        ))]
    } else {
        targets
    };
    let kind = match op.as_str() {
        "=" => StmtKind::Assign {
            targets,
            value,
            by_ref: false,
        },
        // `$a += $b` spelled out as `$a = $a + $b`. `+` is the operator whose
        // meaning depends on the left operand's runtime type (array append /
        // string concat / arithmetic), and only the BINARY path consults
        // `[builtin_slots.array] add` — `compile_compound_op` has its own arm
        // that goes straight to the numeric add, so `$arr += $x` trapped in
        // `toF64`. Growing an array with `+=` is idiomatic PowerShell.
        "+=" => StmtKind::Assign {
            targets: vec![target.clone()],
            value: Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(target),
                right: Box::new(value),
            }),
            by_ref: false,
        },
        "-=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Sub,
            value,
        },
        "*=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Mul,
            value,
        },
        "/=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Div,
            value,
        },
        "%=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Mod,
            value,
        },
        "**=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Pow,
            value,
        },
        "&=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitAnd,
            value,
        },
        "|=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitOr,
            value,
        },
        "^=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::BitXor,
            value,
        },
        "<<=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Shl,
            value,
        },
        ">>=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Shr,
            value,
        },
        "&&=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::And,
            value,
        },
        "||=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::Or,
            value,
        },
        "??=" => StmtKind::CompoundAssign {
            target,
            op: CompoundOp::NullCoalesce,
            value,
        },
        _ => StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(target),
            value: Box::new(value),
        })),
    };

    Statement::new(kind)
}

fn parse_increment_statement(__w: &mut PsWalker, pair: Pair<Rule>) -> Statement {
    let mut target = None;
    let mut is_inc = true;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::lvalue => target = Some(walk_lvalue(__w, child)),
            Rule::increment_op => is_inc = child.as_str() == "++",
            _ => {}
        }
    }

    let Some(target) = target else {
        return Statement::new(StmtKind::Expr(Expression::null()));
    };

    Statement::new(StmtKind::CompoundAssign {
        target,
        op: if is_inc {
            CompoundOp::Add
        } else {
            CompoundOp::Sub
        },
        value: Expression::int(1),
    })
}

/// An assignment target: a variable, optionally followed by `.member` /
/// `[index]` steps. Same node shapes the expression walker produces.
fn walk_lvalue(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    // A leading `[int]` is a type constraint on the target, not part of its
    // identity — the shared compiler infers from the assigned value.
    let mut inner = pair.into_inner().peekable();
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_literal) {
        inner.next();
    }

    let mut expr = match inner.next() {
        Some(p) if p.as_rule() == Rule::var_ref => walk_var_ref(p.as_str()),
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };

    for step in inner {
        match step.as_rule() {
            Rule::member_get => {
                let name = step
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().trim_start_matches('$').to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
            }
            Rule::index_get => {
                let index = step
                    .into_inner()
                    .next()
                    .map(|__x| walk_expr(__w, __x))
                    .unwrap_or_else(Expression::null);
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(fold_literal_key(index)),
                    null_safe: false,
                });
            }
            _ => {}
        }
    }

    expr
}

fn parse_command_line(__w: &mut PsWalker, text: &str) -> Option<Expression> {
    let mut segment_iter = split_command_segments(text).into_iter();
    let first_segment = segment_iter.next()?;
    let first_tokens = split_command_tokens(&first_segment);
    if first_tokens.is_empty() {
        return None;
    }

    let (head, args) = parse_command_parts(__w, &first_tokens)?;
    let mut expr = Expression::new(ExprKind::Call {
        callee: Box::new(head),
        args,
        optional: false,
    });

    for segment in segment_iter {
        let segment_tokens = split_command_tokens(&segment);
        if segment_tokens.is_empty() {
            continue;
        }
        let (next, mut next_args) = parse_command_parts(__w, &segment_tokens)?;
        let mut chained = vec![Argument::positional(expr)];
        chained.append(&mut next_args);
        expr = Expression::new(ExprKind::Call {
            callee: Box::new(next),
            args: chained,
            optional: false,
        });
    }

    Some(expr)
}

fn parse_pipeline(__w: &mut PsWalker, pair: Pair<Rule>) -> Option<Expression> {
    let mut work = std::collections::VecDeque::new();
    work.push_back(pair);
    let mut segments: Vec<Pair<Rule>> = Vec::new();

    while let Some(node) = work.pop_front() {
        match node.as_rule() {
            Rule::command_stmt | Rule::pipeline_expr => {
                for child in node.into_inner() {
                    work.push_back(child);
                }
            }
            Rule::command_segment => segments.push(node),
            _ => {}
        }
    }

    let mut segments = segments.into_iter();
    let first = segments.next()?;
    let head_text = first.as_str().trim().to_string();
    let (head, args) = parse_command_segment(__w, first)?;
    // A pipeline can START with a value: `$o | Add-Member …` feeds `$o` in, it
    // does not invoke it. Everything else goes through `build_command_call`,
    // not a bare `Call` — a cmdlet reached as a STATEMENT must get the same
    // `normalize_cmdlet` rewrite it gets as an expression, or `Set-Variable
    // -Name x -Value 1` compiles to a call to a function never defined.
    let mut expr = if args.is_empty() && pipeline_head_is_value(&head_text) {
        head
    } else {
        build_command_call(head, args)
    };

    for segment in segments {
        let (next, mut next_args) = parse_command_segment(__w, segment)?;
        // The same collection-cmdlet rewrite the value-position builder does.
        if let Some(folded) = pipeline_stage_as_method(&expr, &next, &next_args) {
            expr = folded;
            continue;
        }
        let mut chained = vec![Argument::positional(expr)];
        chained.append(&mut next_args);
        // Also through `build_command_call`. A cmdlet does not stop being one
        // because it is downstream of a `|`: `$o | Add-Member …` needs the same
        // rewrite as the head of the pipeline, and without it compiled to a
        // call to a function that was never defined.
        expr = build_command_call(next, chained);
    }

    Some(expr)
}

/// True when a pipeline's first segment is a VALUE being fed in rather than a
/// command to invoke.
///
/// A pipeline can start with either — `Get-Item | …` invokes, `$o | …` and
/// `1..3 | …` feed. The test used to be "starts with `$`", so every other value
/// form was compiled as a call to a function named after its own source text:
/// `1..3 | ForEach-Object { $_ }` trapped with "f64 is not callable" and
/// `@("x","y") | Sort-Object` with "undefined is not callable".
///
/// A command name is a bare word or a path; none of these openers can begin one.
fn pipeline_head_is_value(text: &str) -> bool {
    let text = text.trim_start();
    match text.chars().next() {
        Some('$') | Some('(') | Some('"') | Some('\'') | Some('[') => true,
        // `@(…)` and `@{…}` are values; `@name` is a splat, which is an argument
        // rather than a pipeline head.
        Some('@') => matches!(text.chars().nth(1), Some('(') | Some('{')),
        Some(c) => c.is_ascii_digit(),
        None => false,
    }
}

fn parse_command_tokens_as_expr(__w: &mut PsWalker, tokens: &[String]) -> Option<Expression> {
    let (callee, args) = parse_command_parts(__w, tokens)?;
    Some(build_command_call_in(callee, args, true))
}

/// Build the call for a command invocation, giving `normalize_cmdlet` first
/// refusal so .NET-shaped cmdlets never need a builtin of their own.
fn build_command_call(callee: Expression, args: Vec<Argument>) -> Expression {
    build_command_call_in(callee, args, false)
}

/// `in_value_position` distinguishes `Write-Output $x` used as a STATEMENT —
/// where the host renders it — from the same call used as a VALUE, where it
/// passes `$x` down the pipeline and prints nothing. One cmdlet, two lowerings,
/// and only the walker knows which position it is in.
fn build_command_call_in(
    callee: Expression,
    args: Vec<Argument>,
    in_value_position: bool,
) -> Expression {
    if let ExprKind::Ident(name) = &callee.kind {
        if in_value_position {
            if let Some(expr) = passthrough_cmdlet(name, &args) {
                return expr;
            }
        }
        if let Some(expr) = normalize_cmdlet(name, &args) {
            return expr;
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn parse_command_segment(__w: &mut PsWalker, pair: Pair<Rule>) -> Option<(Expression, Vec<Argument>)> {
    let mut tokens = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::command_head | Rule::command_arg | Rule::command_token => {
                tokens.push(extract_command_token_text(child));
            }
            _ => {}
        }
    }
    if tokens.is_empty() {
        return None;
    }

    parse_command_parts(__w, &tokens)
}

fn parse_command_parts(__w: &mut PsWalker, tokens: &[String]) -> Option<(Expression, Vec<Argument>)> {
    let callee = parse_command_head(__w, &tokens[0]);
    let mut args = Vec::new();
    let mut i = 1;
    while i < tokens.len() {
        let token = tokens[i].as_str();
        if token.starts_with('-') && token.len() > 1 {
            let flag = &token[1..];
            if let Some((key, value)) = flag.split_once(':') {
                args.push(Argument {
                    value: parse_atom(__w, value),
                    name: Some(key.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 1;
                continue;
            }
            if let Some((key, value)) = flag.split_once('=') {
                args.push(Argument {
                    value: parse_atom(__w, value),
                    name: Some(key.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 1;
                continue;
            }
            if let Some(next) = tokens.get(i + 1) {
                // `-x -5` passes -5 TO `-x`; a leading `-` only starts another
                // parameter when what follows is a name, not a number.
                if (next.starts_with('-') && !is_negative_number(next)) || next.trim().is_empty() {
                    args.push(Argument {
                        value: Expression::bool(true),
                        name: Some(flag.to_string()),
                        by_ref: false,
                        spread: false,
                    });
                    i += 1;
                    continue;
                }
                args.push(Argument {
                    value: parse_atom(__w, next),
                    name: Some(flag.to_string()),
                    by_ref: false,
                    spread: false,
                });
                i += 2;
                continue;
            }
            args.push(Argument {
                value: Expression::bool(true),
                name: Some(flag.to_string()),
                by_ref: false,
                spread: false,
            });
            i += 1;
            continue;
        }
        // `Cmd @params` SPLATS: the collection or hashtable in `$params` supplies
        // the arguments, positionally or by name depending on its runtime shape.
        // Marked `spread` — the same flag Python's `*args` AND `**kwargs` both
        // set, because the shape is a runtime question either way.
        if let Some(name) = splat_token_name(token) {
            args.push(Argument {
                value: Expression::ident(name),
                name: None,
                by_ref: false,
                spread: true,
            });
            i += 1;
            continue;
        }
        args.push(Argument {
            value: parse_atom(__w, token),
            name: None,
            by_ref: false,
            spread: false,
        });
        i += 1;
    }
    Some((callee, args))
}

/// A hashtable key is CASE-INSENSITIVE in PowerShell: `$h['Name']` and
/// `$h['name']` name one entry. The compiler already folds keys at construction
/// and on member access (this profile is `case_sensitive = false`), but the
/// index path passed the literal through unchanged, so `$h['Name']` missed the
/// `name` it had just stored. Folding a literal key here matches the storage the
/// compiler chose. Only a LITERAL is folded — a computed key is whatever it
/// evaluates to.
fn fold_literal_key(index: Expression) -> Expression {
    match &index.kind {
        ExprKind::Lit(Literal::Str(s)) if s.chars().any(|c| c.is_uppercase()) => {
            Expression::string(&s.to_lowercase())
        }
        _ => index,
    }
}

/// `-5` / `-3.2` — a negative numeric literal, not a parameter name.
fn is_negative_number(token: &str) -> bool {
    let rest = token.trim().trim_start_matches('-');
    !rest.is_empty()
        && rest.starts_with(|c: char| c.is_ascii_digit() || c == '.')
        && rest.chars().all(|c| c.is_ascii_digit() || c == '.')
}

/// `@name` as a command argument is a splat. `@(…)` and `@{…}` are an array and
/// a hashtable literal and are NOT — the character only splats before a name.
fn splat_token_name(token: &str) -> Option<&str> {
    let rest = token.strip_prefix('@')?;
    let ok = !rest.is_empty() && rest.chars().all(|c| c.is_ascii_alphanumeric() || c == '_');
    ok.then_some(rest)
}

fn parse_command_head(__w: &mut PsWalker, raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    if looks_like_command_name(text) || is_path_like_command(text) {
        return Expression::ident(text);
    }
    // A bare word in HEAD position is a command name, not the string argument
    // `parse_atom` would produce — a string callee can never be callable.
    if is_bare_command_word(text) {
        return Expression::ident(text);
    }
    parse_atom(__w, text)
}

fn is_path_like_command(text: &str) -> bool {
    let text = text.trim();
    text.starts_with('&')
        || text.ends_with(".ps1")
        || text.ends_with(".psm1")
        || text.ends_with(".psd1")
        || text.contains('/')
        || text.contains('\\')
}

fn split_command_segments(input: &str) -> Vec<String> {
    let mut segments = Vec::new();
    let mut current = String::new();
    let mut quote = None;
    let mut depth_paren = 0usize;
    let mut depth_bracket = 0usize;
    let mut depth_brace = 0usize;
    let mut escaped = false;
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.next() {
        if escaped {
            current.push(ch);
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            current.push(ch);
            continue;
        }

        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
        } else if ch == '(' {
            depth_paren += 1;
            current.push(ch);
        } else if ch == ')' {
            depth_paren = depth_paren.saturating_sub(1);
            current.push(ch);
        } else if ch == '[' {
            depth_bracket += 1;
            current.push(ch);
        } else if ch == ']' {
            depth_bracket = depth_bracket.saturating_sub(1);
            current.push(ch);
        } else if ch == '{' {
            depth_brace += 1;
            current.push(ch);
        } else if ch == '}' {
            depth_brace = depth_brace.saturating_sub(1);
            current.push(ch);
        } else if ch == '|' && depth_paren == 0 && depth_bracket == 0 && depth_brace == 0 {
            if !current.trim().is_empty() {
                segments.push(current.trim().to_string());
            } else {
                segments.push(String::new());
            }
            current.clear();
            continue;
        } else if ch == '"' || ch == '\'' {
            quote = Some(ch);
            current.push(ch);
        } else {
            current.push(ch);
        }
    }

    if !current.trim().is_empty() {
        segments.push(current.trim().to_string());
    }

    segments
}

fn extract_command_token_text(pair: Pair<Rule>) -> String {
    let raw = pair.as_str().trim().to_string();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::command_token {
            return child.as_str().trim().to_string();
        }
    }
    raw
}

fn collect_command_tokens(pair: &Pair<Rule>, out: &mut Vec<String>) {
    if pair.as_rule() == Rule::command_token {
        out.push(pair.as_str().trim().to_string());
        return;
    }

    for child in pair.clone().into_inner() {
        collect_command_tokens(&child, out);
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Expression walking
//
// The pest expression grammar is the ONE place operator precedence is defined.
// Nothing below re-splits source text on operator spellings — a text fragment
// captured in command-token position is re-parsed through `Rule::expr_entry`,
// i.e. through the same grammar, so there is a single source of truth.
// ────────────────────────────────────────────────────────────────────────────

/// Parse a text fragment in **expression mode** — a bare word is an identifier.
fn expr_from_text(__w: &mut PsWalker, raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    parse_expr_fragment(__w, text).unwrap_or_else(|| Expression::ident(text))
}

/// Parse a token in **command mode** — a bare word is a string argument, the
/// way PowerShell treats `Write-Host FAIL`.
fn parse_atom(__w: &mut PsWalker, raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }
    if is_bare_command_word(text) {
        return Expression::string(text);
    }
    parse_expr_fragment(__w, text).unwrap_or_else(|| Expression::string(text))
}

/// A token that carries no expression syntax at all — in command mode this is
/// a literal string argument.
fn is_bare_command_word(text: &str) -> bool {
    !text.is_empty()
        && !text.starts_with('$')
        && !text.starts_with('@')
        && !text.starts_with('[')
        && !text.starts_with('(')
        && !text.starts_with('"')
        && !text.starts_with('\'')
        && !text.starts_with('&')
        && !is_numeric_like(text)
        && text
            .chars()
            .all(|c| c.is_alphanumeric() || matches!(c, '_' | '-' | '.' | ':' | '\\' | '/' | '*'))
        && text.chars().any(|c| c.is_alphabetic())
}

fn parse_expr_fragment(__w: &mut PsWalker, text: &str) -> Option<Expression> {
    let pairs = super::PowerShellParser::parse(Rule::expr_entry, text).ok()?;
    let root = pairs.into_iter().next()?;
    let expr = root
        .into_inner()
        .find(|c| c.as_rule() == Rule::expression)?;
    Some(walk_expr(__w, expr))
}

/// Walk one expression pair into the shared AST.
fn walk_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    match pair.as_rule() {
        // `(Get-Date)` is a command INVOCATION, not a read of a name. Only a
        // lone bare word qualifies — `($x)` and `(1 + 2)` stay expressions.
        Rule::paren_expr => match lone_bare_word(&pair) {
            Some(name) => Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&name)),
                args: Vec::new(),
                optional: false,
            }),
            None => first_inner_expr(__w, pair),
        },

        Rule::expression | Rule::expr_statement | Rule::for_cond | Rule::condition_expr => {
            first_inner_expr(__w, pair)
        }

        Rule::comma_expr => walk_comma_expr(__w, pair),
        Rule::coalesce => walk_binary_chain(__w, pair),
        Rule::ternary_expr => walk_ternary(__w, pair),

        Rule::logical_or
        | Rule::logical_and
        | Rule::comparison
        | Rule::additive
        | Rule::multiplicative
        | Rule::power => walk_binary_chain(__w, pair),

        // The argument list of `-replace` / `-split` / `-join`. Reuses the comma
        // walk so a single operand stays scalar and several become an `Array`
        // that `spread_operands` unpacks back into arguments.
        Rule::here_string_double => walk_here_string(__w, pair.as_str(), true),
        Rule::here_string_single => walk_here_string(__w, pair.as_str(), false),

        Rule::cmp_list => walk_comma_expr(__w, pair),
        Rule::format_expr => walk_format(__w, pair),
        Rule::range_expr => walk_range(__w, pair),
        Rule::unary => walk_unary(__w, pair),
        Rule::cast_expr => walk_cast(__w, pair),
        Rule::postfix => walk_postfix(__w, pair),

        Rule::command_pipeline => walk_command_pipeline(__w, pair),
        Rule::command_segment => walk_command_segment_expr(__w, pair),
        Rule::array_expr => walk_array_expr(__w, pair),
        Rule::hash_literal => walk_hash_literal(__w, pair),
        Rule::sub_expr => walk_sub_expr(__w, pair),
        Rule::script_block_expr => walk_script_block_expr(__w, pair),

        Rule::number => walk_number(pair.as_str()),
        Rule::quoted_string => {
            let raw = pair.as_str();
            parse_double_quoted_string(__w, raw.get(1..raw.len().saturating_sub(1)).unwrap_or(""))
        }
        Rule::single_quoted_string => {
            let raw = pair.as_str();
            let inner = raw.get(1..raw.len().saturating_sub(1)).unwrap_or("");
            Expression::string(&inner.replace("''", "'"))
        }
        Rule::var_ref => walk_var_ref(pair.as_str()),
        Rule::type_literal => type_literal_expr(type_literal_name(pair.as_str())),
        Rule::bare_word => walk_bare_word(pair.as_str()),
        // `@args` reads the variable; the SPREAD is applied by the call site.
        Rule::splat_ref => Expression::ident(pair.as_str().trim_start_matches('@')),

        _ => first_inner_expr(__w, pair),
    }
}

/// The single `bare_word` a group bottoms out in, if the group contains nothing
/// else — no operators, no arguments, no member access. Walks the single-child
/// chain and gives up the moment a level has siblings.
fn lone_bare_word(pair: &Pair<Rule>) -> Option<String> {
    let mut current = pair.clone();
    loop {
        if current.as_rule() == Rule::bare_word {
            return Some(current.as_str().to_string());
        }
        let mut inner = current.clone().into_inner();
        let only = inner.next()?;
        if inner.next().is_some() {
            return None;
        }
        current = only;
    }
}

/// Whether a function body reads `$args`. Scanned from the source text because
/// the answer is needed to build the parameter list, before the body walks.
fn mentions_args(text: &str) -> bool {
    let bytes = text.as_bytes();
    let mut at = 0;
    while let Some(hit) = text[at..].find('$') {
        let start = at + hit + 1;
        let end = (start + 4).min(bytes.len());
        if text[start..end].eq_ignore_ascii_case("args")
            && !bytes
                .get(end)
                .is_some_and(|c| c.is_ascii_alphanumeric() || *c == b'_')
        {
            return true;
        }
        at = start;
    }
    false
}

/// The bare words that are VALUES rather than command names.
fn is_literal_word(word: &str) -> bool {
    matches!(word.to_lowercase().as_str(), "true" | "false" | "null")
}

fn first_inner_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    pair.into_inner()
        .next()
        .map(|__x| walk_expr(__w, __x))
        .unwrap_or_else(Expression::null)
}

/// `1,2,3` builds an array; a single element passes straight through.
fn walk_comma_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut items: Vec<Expression> = pair.into_inner().map(|__x| walk_expr(__w, __x)).collect();
    match items.len() {
        0 => Expression::null(),
        1 => items.remove(0),
        _ => Expression::new(ExprKind::Array(
            items
                .into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )),
    }
}

fn walk_ternary(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let cond = match inner.next() {
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };
    let Some(then_pair) = inner.next() else {
        return cond;
    };
    let then_expr = walk_expr(__w, then_pair);
    let else_expr = inner.next().map(|__x| walk_expr(__w, __x)).unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(else_expr),
    })
}

/// Left-associative fold over `operand (op operand)*`.
/// The right operand of a list-taking operator, as the argument list it stands
/// for. A single operand stays a single argument.
fn spread_operands(right: Expression) -> Vec<Expression> {
    match right.kind {
        ExprKind::Array(elems) => elems.into_iter().map(|e| e.value).collect(),
        _ => vec![right],
    }
}

fn walk_binary_chain(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let mut left = match inner.next() {
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };
    while let Some(op_pair) = inner.next() {
        let Some(rhs_pair) = inner.next() else { break };
        let right = walk_expr(__w, rhs_pair);
        left = build_binary(op_pair.as_str(), left, right);
    }
    left
}

/// `"{0} {1}" -f $a, $b` — .NET composite formatting. Lowered to
/// `String.Format(fmt, …)` so the shared dotnet surface owns the format
/// pictures (`{0:N2}`, `{0,-8}`) rather than this walker.
fn walk_format(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let Some(head) = inner.next() else {
        return Expression::null();
    };
    let fmt = walk_expr(__w, head);

    let mut args = vec![Argument::positional(fmt)];
    let mut saw_op = false;
    for child in inner {
        match child.as_rule() {
            Rule::format_op => saw_op = true,
            Rule::format_args => {
                args.extend(
                    child
                        .into_inner()
                        .map(|a| Argument::positional(walk_expr(__w, a))),
                );
            }
            _ => {}
        }
    }

    if !saw_op {
        return args.remove(0).value;
    }

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("String")),
            field: "Format".to_string(),
            null_safe: false,
        })),
        args,
        optional: false,
    })
}

fn walk_range(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let start = match inner.next() {
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };
    match inner.next() {
        Some(end) => Expression::new(ExprKind::Range {
            start: Box::new(start),
            end: Box::new(walk_expr(__w, end)),
            inclusive: true,
        }),
        None => start,
    }
}

fn walk_unary(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let Some(first) = inner.next() else {
        return Expression::null();
    };
    if first.as_rule() != Rule::unary_op {
        return walk_expr(__w, first);
    }
    let op_text = first.as_str().trim().to_lowercase();
    let operand = inner.next().map(|__x| walk_expr(__w, __x)).unwrap_or_else(Expression::null);

    // `-join` / `-split` in unary position are PowerShell's collection forms;
    // keep them as ordinary method calls so shared dispatch owns them.
    match op_text.as_str() {
        "-join" => return method_call_expr(operand, "join", vec![Expression::string("")]),
        "-split" => return method_call_expr(operand, "split", vec![Expression::string(" ")]),
        _ => {}
    }

    // Fold a negated numeric literal into the literal itself, so `$a[-1]` hands
    // the shared indexing a real negative constant rather than a runtime
    // `Unary{Neg, 1}` it cannot reason about.
    if op_text == "-" {
        match &operand.kind {
            ExprKind::Lit(Literal::Int(n)) => return Expression::int(-n),
            ExprKind::Lit(Literal::Float(f)) => return Expression::float(-f),
            _ => {}
        }
    }

    let op = match op_text.as_str() {
        "++" => UnaryOp::PreInc,
        "--" => UnaryOp::PreDec,
        "!" | "-not" => UnaryOp::Not,
        "-bnot" => UnaryOp::BitNot,
        "+" => UnaryOp::Pos,
        _ => UnaryOp::Neg,
    };
    Expression::new(ExprKind::Unary {
        op,
        expr: Box::new(operand),
    })
}

fn walk_cast(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let type_name = inner
        .next()
        .map(|p| type_literal_name(p.as_str()).to_string())
        .unwrap_or_default();
    let expr = inner.next().map(|__x| walk_expr(__w, __x)).unwrap_or_else(Expression::null);

    // `[bool]$x` is a TRUTHINESS conversion, and the shared `Cast` lowering has
    // no bool arm — it would hand the value straight back, so `[bool]0` stayed
    // `0`. `!!$x` reaches the profile's own truthiness rule and yields a real
    // boolean, so the conversion is expressed with nodes the compiler already
    // owns rather than a new cast target.
    if matches!(
        type_name.to_lowercase().as_str(),
        "bool" | "boolean" | "system.boolean"
    ) {
        return negate(negate(expr));
    }

    // `[PSCustomObject]@{ … }` builds an object with those properties — which
    // is what the hashtable literal already walked to. The cast selects a
    // .NET wrapper type that has no counterpart here, so it is the identity.
    if matches!(
        type_name.to_lowercase().as_str(),
        "pscustomobject" | "psobject" | "system.management.automation.pscustomobject"
    ) {
        return expr;
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name,
    })
}

fn walk_postfix(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut inner = pair.into_inner();
    let mut expr = match inner.next() {
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };

    for op in inner {
        match op.as_rule() {
            Rule::member_get => {
                let name = op
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().trim_start_matches('$').to_string())
                    .unwrap_or_default();
                // `$h.Keys` / `$h.Values` carry NO parentheses — they are
                // property reads, so `[value_methods]` (a call-site table) can
                // never see them. Compiled as a plain member they became a
                // struct field read of a field no `@{…}` object has, which is
                // why `($h.Keys -join '|')` answered the empty string while
                // `$h.Count` (a different path) answered correctly.
                //
                // Rewritten to a call, the profile's object-shaped binding
                // applies. An object carrying its own literal `Keys` property
                // loses to this, which is the right trade: a PowerShell
                // hashtable is overwhelmingly the receiver that spells it.
                if let Some(private) = hashtable_only_property(__w, &name) {
                    expr = method_call_expr(expr, private, Vec::new());
                    continue;
                }
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: normalize_member_name(&name),
                    null_safe: false,
                });
            }
            Rule::static_member => {
                let name = op
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
            }
            Rule::method_call | Rule::static_call => {
                let is_static = op.as_rule() == Rule::static_call;
                let mut parts = op.into_inner();
                let name = parts
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_default();
                let args = parts.next().map(|__x| walk_arg_list(__w, __x)).unwrap_or_default();

                // Rename the dictionary-only methods the runtime collection
                // registry would otherwise claim. `runtime_collection_scope`
                // makes `scope_declares_member_arity` answer yes for
                // `ContainsKey/1` (System.Collections.Hashtable declares it),
                // which routes the call to runtime `__type` dispatch — and a
                // PowerShell `@{…}` is a bare object carrying no `__type`, so
                // the lookup reached undefined and the call trapped.
                //
                // Only names NO list receiver shares are renamed here.
                // `Add` / `Remove` / `Clear` / `Contains` are spelled the same
                // on an ArrayList, so renaming them by spelling alone would
                // break `New-Object System.Collections.ArrayList` — those need
                // a receiver-shape test, not a rewrite.
                // `$h.ContainsValue($v)` is membership over the VALUES, so the
                // receiver of the test is the values array — not `$h`. A
                // profile rename cannot express that (the binding's receiver is
                // always the object), so this becomes the `-contains` node the
                // shared `[builtin_slots.array] contains` slot already lowers.
                // `BinOp::In` takes the needle on the left, matching how
                // `-contains` itself is built above.
                if !is_static && args.len() == 1 && name.eq_ignore_ascii_case("ContainsValue") {
                    let values = method_call_expr(expr, "__ps_ht_values", Vec::new());
                    let mut args = args;
                    expr = Expression::new(ExprKind::Binary {
                        op: BinOp::In,
                        left: Box::new(args.remove(0)),
                        right: Box::new(values),
                    });
                    continue;
                }

                let name = match hashtable_only_method(__w, &name, is_static, args.len()) {
                    Some(private) => private.to_string(),
                    None => name,
                };

                // `[Animal]::new(…)` is construction, not a static call — hand
                // the shared compiler the `New` node it already knows.
                if is_static && name.eq_ignore_ascii_case("new") {
                    if let Some(class) = type_name_of(&expr) {
                        expr = Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident(&class)),
                            args: args
                                .into_iter()
                                .map(|value| Argument {
                                    value,
                                    name: None,
                                    by_ref: false,
                                    spread: false,
                                })
                                .collect(),
                        });
                        continue;
                    }
                }

                // `$sb.Invoke(…)` / `.InvokeReturnAsIs(…)` CALL the script
                // block. There is no `Invoke` member on a lambda to look up, so
                // this is the call itself — the same thing `& $sb` already
                // compiles to.
                if !is_static
                    && matches!(name.to_lowercase().as_str(), "invoke" | "invokereturnasis")
                {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: args.into_iter().map(Argument::positional).collect(),
                        optional: false,
                    });
                    continue;
                }

                expr = method_call_expr(expr, &name, args);
            }
            Rule::index_get => {
                let index = op
                    .into_inner()
                    .next()
                    .map(|__x| walk_expr(__w, __x))
                    .unwrap_or_else(Expression::null);
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(fold_literal_key(index)),
                    null_safe: false,
                });
            }
            _ => {}
        }
    }

    expr
}

fn walk_arg_list(__w: &mut PsWalker, pair: Pair<Rule>) -> Vec<Expression> {
    pair.into_inner().map(|__x| walk_expr(__w, __x)).collect()
}

// The PowerShell-private spelling for a method the runtime collection registry
// claims but a bare `@{…}` object cannot answer, or `None` to leave the name
// alone. Every name here is dictionary-only: no list/set receiver spells it the
// same way, so the rewrite cannot capture a call meant for another type.
// Method names declared by classes in the script being walked, folded to
// lowercase. Read by the hashtable rewrites so a user method always wins.
/// Every registry the powershell walk keeps, owned by one `parse` call.
///
/// Was a process-global static: the method names one script declared stayed
/// visible to the next program compiled on this thread.
#[derive(Default)]
pub(crate) struct PsWalker {
    declared_methods: std::collections::HashSet<String>,
}


/// Record every method name the script's classes declare.
///
/// Without this the rewrites below decide by SPELLING alone, and a user class
/// that happens to spell a hashtable method loses its own body: `class
/// Calculator { [int]Add([int]$a, [int]$b) }` made `$calc.Add(6, 7)` return
/// null. Measured as a regression in `classes/class_method`, which is why the
/// rewrite table asks this first.
fn collect_declared_methods(__w: &mut PsWalker, root: Pair<Rule>) {
    let mut names = std::collections::HashSet::new();
    let mut work = VecDeque::new();
    work.push_back(root);
    while let Some(node) = work.pop_front() {
        match node.as_rule() {
            Rule::ps_method => {
                if let Some(n) = node
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::member_name)
                {
                    names.insert(n.as_str().to_lowercase());
                }
            }
            Rule::class_function_decl => {
                if let Some(n) = node
                    .clone()
                    .into_inner()
                    .find(|c| c.as_rule() == Rule::function_name)
                {
                    names.insert(n.as_str().to_lowercase());
                }
            }
            _ => {}
        }
        for child in node.into_inner() {
            work.push_back(child);
        }
    }
    __w.declared_methods = names;
}

/// True when a class in this script declares `name` as a method.
fn is_declared_method(__w: &mut PsWalker, name: &str) -> bool {
    let folded = name.to_lowercase();
    __w.declared_methods.contains(&folded)
}

/// The PowerShell-private call that replaces a hashtable PROPERTY read.
fn hashtable_only_property(__w: &mut PsWalker, name: &str) -> Option<&'static str> {
    if is_declared_method(__w, name) {
        return None;
    }
    match name.to_lowercase().as_str() {
        "keys" => Some("__ps_ht_keys"),
        "values" => Some("__ps_ht_values"),
        _ => None,
    }
}

fn hashtable_only_method(__w: &mut PsWalker, name: &str, is_static: bool, argc: usize) -> Option<&'static str> {
    if is_static || is_declared_method(__w, name) {
        return None;
    }
    match (name.to_lowercase().as_str(), argc) {
        ("containskey", 1) => Some("__ps_ht_haskey"),
        // `Add` is spelled by a list (`Add(x)`) and by user classes too. Arity
        // separates the list — one element vs a key AND a value — and
        // `is_declared_method` above separates the user class, which is what
        // `classes/class_method` regressed on when only arity guarded it.
        ("add", 2) => Some("__ps_ht_add"),
        // NO `Remove` / `Clear` / `Contains`. Each names several receivers at
        // ONE arity — a hashtable, an ArrayList, and (for `Remove`) a string,
        // where `'abc'.Remove(1)` is substring arithmetic. Arity cannot split
        // them, so they need a receiver-shape test at runtime rather than a
        // rewrite. Renaming them by spelling would fix the hashtable tests by
        // cementing a wrong lowering for the other two receivers.
        _ => None,
    }
}

fn method_call_expr(receiver: Expression, name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(receiver),
            field: name.to_string(),
            null_safe: false,
        })),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

/// `a | b | c` — each stage's value becomes argument 0 of the next callee, so
/// the shared call resolution in `primitives/calls.rs` keeps full control.
fn walk_command_pipeline(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut stages = pair.into_inner();
    let mut expr = match stages.next() {
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };

    for stage in stages {
        let (callee, mut args) = match stage.as_rule() {
            Rule::command_segment => match parse_command_segment(__w, stage) {
                Some(parts) => parts,
                None => continue,
            },
            _ => (walk_expr(__w, stage), Vec::new()),
        };

        if let Some(folded) = pipeline_stage_as_method(&expr, &callee, &args) {
            expr = folded;
            continue;
        }
        args.insert(
            0,
            Argument {
                value: expr,
                name: None,
                by_ref: false,
                spread: false,
            },
        );
        expr = Expression::new(ExprKind::Call {
            callee: Box::new(callee),
            args,
            optional: false,
        });
    }

    expr
}

/// Fold one pipeline stage onto the value accumulated so far, when that stage is
/// one of the core collection cmdlets. Rewriting them to the equivalent method
/// call lets the existing `[array_methods]` dispatch handle them — no cmdlet
/// builtins, no emitter arm.
///
/// Shared by BOTH pipeline builders. A pipeline reaches the walker two ways —
/// `walk_command_pipeline` for one in value position (`$r = $a | Sort-Object`)
/// and `parse_pipeline` for one used as a statement (`$a | Sort-Object`) — and
/// only the first applied this rewrite. The statement path emitted a call to
/// `Sort-Object`, a function nothing defines, so EVERY bare-statement pipeline
/// trapped with "undefined is not callable" while the assigned form worked.
fn pipeline_stage_as_method(
    upstream: &Expression,
    callee: &Expression,
    args: &[Argument],
) -> Option<Expression> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    let method = pipeline_cmdlet_method(name)?;
    if method.is_empty() {
        // `… | Out-Null` discards the pipeline value.
        return Some(Expression::null());
    }
    let positional: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| a.value.clone())
        .collect();
    Some(method_call_expr(upstream.clone(), method, positional))
}

/// The method a pipeline cmdlet is equivalent to, including its `?` / `%`
/// aliases. `""` means the stage drops the value (`Out-Null`).
fn pipeline_cmdlet_method(name: &str) -> Option<&'static str> {
    match name.to_lowercase().as_str() {
        "where-object" | "where" | "?" => Some("Where"),
        "foreach-object" | "foreach" | "%" => Some("ForEach"),
        "sort-object" | "sort" => Some("Sort"),
        "measure-object" | "measure" => Some("Count"),
        "out-null" => Some(""),
        _ => None,
    }
}

/// Cmdlets that are really .NET calls in disguise. Rewriting them here routes
/// them through the dotnet tree-mount that already resolves `System.*` — no
/// cmdlet builtins, no host functions, no emitter arm.
///
/// `New-Object` is construction; the file cmdlets are `System.IO.File` statics.
/// Cmdlets that, used as a VALUE, are the identity on their arguments.
/// `@(Write-Output (1..5))` is five elements, not one call that printed.
fn passthrough_cmdlet(name: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(name.to_lowercase().as_str(), "write-output" | "echo") {
        return None;
    }
    let values: Vec<&Argument> = args.iter().filter(|a| a.name.is_none()).collect();
    match values.as_slice() {
        [] => Some(Expression::null()),
        [only] => Some(only.value.clone()),
        many => Some(Expression::new(ExprKind::Array(
            many.iter()
                .map(|a| ArrayElement {
                    key: None,
                    value: a.value.clone(),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ))),
    }
}

fn normalize_cmdlet(name: &str, args: &[Argument]) -> Option<Expression> {
    let positional: Vec<&Argument> = args.iter().filter(|a| a.name.is_none()).collect();
    let named = |key: &str| {
        args.iter()
            .find(|a| {
                a.name
                    .as_deref()
                    .is_some_and(|n| n.eq_ignore_ascii_case(key))
            })
            .map(|a| a.value.clone())
    };

    match name.to_lowercase().as_str() {
        // `New-Object System.Text.StringBuilder` / `-TypeName … -ArgumentList …`
        "new-object" => {
            let type_expr =
                named("TypeName").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let type_name = literal_text(&type_expr)?;

            let mut ctor_args: Vec<Argument> = Vec::new();
            if let Some(list) = named("ArgumentList") {
                match list.kind {
                    ExprKind::Array(elems) => {
                        ctor_args.extend(elems.into_iter().map(|e| Argument::positional(e.value)))
                    }
                    _ => ctor_args.push(Argument::positional(list)),
                }
            } else {
                for a in positional.iter().skip(1) {
                    ctor_args.push((*a).clone());
                }
            }

            Some(Expression::new(ExprKind::New {
                class: Box::new(type_literal_expr(&type_name)),
                args: ctor_args,
            }))
        }

        "test-path" => Some(dotnet_static_call(
            "System.IO.File",
            "Exists",
            vec![named("Path").or_else(|| positional.first().map(|a| a.value.clone()))?],
        )),
        "get-content" => Some(dotnet_static_call(
            "System.IO.File",
            "ReadAllText",
            vec![named("Path").or_else(|| positional.first().map(|a| a.value.clone()))?],
        )),
        "set-content" => {
            let path = named("Path").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let value = named("Value").or_else(|| positional.get(1).map(|a| a.value.clone()))?;
            Some(dotnet_static_call(
                "System.IO.File",
                "WriteAllText",
                vec![path, value],
            ))
        }
        "remove-item" => Some(dotnet_static_call(
            "System.IO.File",
            "Delete",
            vec![named("Path").or_else(|| positional.first().map(|a| a.value.clone()))?],
        )),
        // `$o | Add-Member -MemberType NoteProperty -Name X -Value 1` attaches a
        // property. The pipeline hands the object in as argument 0, so this is
        // the assignment `$o.X = 1` written the long way.
        "add-member" => {
            let target = positional.first().map(|a| a.value.clone())?;
            let name = literal_text(
                &named("Name").or_else(|| positional.get(1).map(|a| a.value.clone()))?,
            )?;
            let value = named("Value").or_else(|| positional.get(2).map(|a| a.value.clone()))?;
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(target),
                    field: normalize_member_name(&name),
                    null_safe: false,
                })),
                value: Box::new(value),
            }))
        }

        "join-path" => {
            let head = named("Path").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let tail = named("ChildPath").or_else(|| positional.get(1).map(|a| a.value.clone()))?;
            Some(dotnet_static_call(
                "System.IO.Path",
                "Combine",
                vec![head, tail],
            ))
        }

        // `Set-Variable -Name x -Value 1` IS an assignment to `$x` — the same
        // storage `$x = 1` writes. Normalizing it here keeps one mechanism.
        // `-Option ReadOnly|Constant` is deliberately NOT modelled: enforcing it
        // would mean statically tracking which names were declared read-only and
        // compiling a later write to `throw`, which is test-shaped rather than a
        // semantic the compiler can honestly claim.
        "set-variable" => {
            let name_expr =
                named("Name").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let target = literal_text(&name_expr)?;
            let value = named("Value")
                .or_else(|| positional.get(1).map(|a| a.value.clone()))
                .unwrap_or_else(Expression::null);
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident(scope_qualified_name(&target))),
                value: Box::new(value),
            }))
        }
        "get-variable" => {
            let name_expr =
                named("Name").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let target = literal_text(&name_expr)?;
            Some(Expression::ident(scope_qualified_name(&target)))
        }

        // Sleep takes seconds by default; the shared threading primitive is in
        // milliseconds, which is also what `-Milliseconds` already supplies.
        "start-sleep" => {
            let ms = match named("Milliseconds") {
                Some(v) => v,
                None => {
                    let secs =
                        named("Seconds").or_else(|| positional.first().map(|a| a.value.clone()))?;
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(secs),
                        right: Box::new(Expression::int(1000)),
                    })
                }
            };
            Some(dotnet_static_call(
                "System.Threading.Thread",
                "Sleep",
                vec![ms],
            ))
        }

        _ => None,
    }
}

/// `[System.IO.File]::Method(args)` as a member chain the resolver understands.
fn dotnet_static_call(type_name: &str, method: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(type_literal_expr(type_name)),
            field: method.to_string(),
            null_safe: false,
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

/// The text of a literal string / identifier argument, for cmdlet arguments
/// that name a TYPE rather than carry a value.
fn literal_text(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        ExprKind::Ident(name) => Some(name.clone()),
        _ => None,
    }
}

/// A bare command invocation used where an expression is expected: `(hi 'PASS')`.
fn walk_command_segment_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    match parse_command_segment(__w, pair) {
        Some((callee, args)) => build_command_call_in(callee, args, true),
        None => Expression::null(),
    }
}

/// `@( … )` is the ARRAY SUBEXPRESSION operator: it guarantees an array and
/// flattens one level, so `@(1..5)` is five elements and `@($arr)` is `$arr`'s
/// elements — not one element holding a collection. Each element is therefore
/// spread rather than nested; a scalar element spreads to itself.
fn walk_array_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    // ONE argument, not one per element. `@( … )` collects the output of the
    // whole body — it does not flatten each comma-separated element in turn.
    // `@(1, @(2,3))` is TWO elements whose second is an array, because the
    // commas build one array and `@()` guarantees that array; flattening
    // per-element would splice the inner one and give four.
    let elements: Vec<Expression> = pair.into_inner().map(|__x| walk_expr(__w, __x)).collect();
    let body = match <[Expression; 1]>::try_from(elements) {
        Ok([only]) => only,
        Err(many) => Expression::new(ExprKind::Array(
            many.into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )),
    };

    // Whether that one value flattens is a question about its RUNTIME shape —
    // a collection contributes its elements, a scalar contributes itself — so
    // `__ps_array` answers it rather than the syntax.
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__ps_array")),
        args: vec![Argument::positional(body)],
        optional: false,
    })
}

fn walk_hash_literal(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut props = Vec::new();
    for entry in pair.into_inner() {
        if entry.as_rule() != Rule::hash_entry {
            continue;
        }
        let mut parts = entry.into_inner();
        let Some(key_pair) = parts.next() else {
            continue;
        };
        let key = hash_key_text(key_pair);
        let value = parts.next().map(|__x| walk_expr(__w, __x)).unwrap_or_else(Expression::null);
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&key),
            value,
        });
    }
    Expression::new(ExprKind::Object(props))
}

fn hash_key_text(pair: Pair<Rule>) -> String {
    let inner = pair.into_inner().next();
    match inner {
        Some(p) => {
            let raw = p.as_str();
            match p.as_rule() {
                Rule::quoted_string | Rule::single_quoted_string => raw
                    .get(1..raw.len().saturating_sub(1))
                    .unwrap_or("")
                    .to_string(),
                _ => raw.to_string(),
            }
        }
        None => String::new(),
    }
}

/// `$( … )` — a statement list whose value is the last expression.
fn walk_sub_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    // `sub_body` is silent, so a lone-expression `$( … )` arrives as a single
    // `expression` child and the statement path never sees it.
    let mut inner = pair.clone().into_inner();
    if let (Some(only), None) = (inner.next(), inner.next()) {
        if only.as_rule() == Rule::expression {
            return walk_expr(__w, only);
        }
    }
    let stmts = collect_statements(__w, pair);
    last_expression_of(stmts)
}

/// A script block is a lambda. PowerShell binds the current pipeline item to
/// `$_`, so unless the block declares its own `param(…)` the walker gives it a
/// single implicit `_` parameter — that is what makes `{ $_ -gt 2 }` receive the
/// element the shared HOF dispatch passes in.
fn walk_script_block_expr(__w: &mut PsWalker, pair: Pair<Rule>) -> Expression {
    let mut params = Vec::new();
    let mut body = Vec::new();
    // A script block scopes exactly like a function: `& { $x = 3 }` leaves the
    // caller's `$x` alone.
    let locals = function_local_names(&pair);

    for child in pair.into_inner() {
        if child.as_rule() == Rule::param_stmt {
            parse_param_stmt(__w, child, &mut params);
            continue;
        }
        if let Ok(Some(stmt)) = parse_statement(__w, child) {
            body.push(stmt);
        }
    }

    if params.is_empty() {
        params.push(Param {
            name: "_".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: false,
        });
    }

    let body = declare_function_locals(locals, &params, return_last_of_branches(body));

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

/// PowerShell yields the value of a trailing expression rather than requiring
/// `return`. Rewriting that last statement here keeps the shared compiler
/// generic — the same normalization Ruby's walker does.
fn implicit_return(mut body: Vec<Statement>) -> Vec<Statement> {
    let Some(last) = body.pop() else {
        return body;
    };
    let replaced = match last.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Return(Some(expr))),
        _ => last,
    };
    body.push(replaced);
    body
}

fn collect_statements(__w: &mut PsWalker, pair: Pair<Rule>) -> Vec<Statement> {
    let mut out = Vec::new();
    for child in pair.into_inner() {
        if let Ok(Some(stmt)) = parse_statement(__w, child) {
            out.push(stmt);
        }
    }
    out
}

fn last_expression_of(stmts: Vec<Statement>) -> Expression {
    match stmts.into_iter().last() {
        Some(stmt) => match stmt.kind {
            StmtKind::Expr(expr) => expr,
            _ => Expression::null(),
        },
        None => Expression::null(),
    }
}

fn walk_number(raw: &str) -> Expression {
    let text = raw.trim();
    let lower = text.to_lowercase();

    if let Some(hex) = lower.strip_prefix("0x") {
        if let Ok(v) = i64::from_str_radix(hex, 16) {
            return Expression::int(v);
        }
    }

    if let Some(bin) = lower.strip_prefix("0b") {
        if let Ok(v) = i64::from_str_radix(bin, 2) {
            return Expression::int(v);
        }
    }

    // Size suffixes: 1kb, 2mb, …
    for (suffix, factor) in [
        ("kb", 1024_i64),
        ("mb", 1024 * 1024),
        ("gb", 1024 * 1024 * 1024),
        ("tb", 1024_i64 * 1024 * 1024 * 1024),
        ("pb", 1024_i64 * 1024 * 1024 * 1024 * 1024),
    ] {
        if let Some(head) = lower.strip_suffix(suffix) {
            if let Ok(v) = head.parse::<f64>() {
                return Expression::int((v * factor as f64) as i64);
            }
        }
    }

    let stripped = lower.trim_end_matches(['l', 'd']);
    if let Ok(v) = stripped.parse::<i64>() {
        return Expression::int(v);
    }
    if let Ok(v) = stripped.parse::<f64>() {
        return Expression::float(v);
    }
    Expression::int(0)
}

fn walk_var_ref(raw: &str) -> Expression {
    let name = raw
        .trim_start_matches('$')
        .trim_matches(|c| c == '{' || c == '}');
    match name.to_lowercase().as_str() {
        "true" => Expression::bool(true),
        "false" => Expression::bool(false),
        "null" => Expression::null(),
        "this" => Expression::new(ExprKind::This),
        _ => Expression::ident(scope_qualified_name(name)),
    }
}

/// PowerShell scope modifiers (`$script:x`, `$global:x`, `$local:x`, `$env:x`)
/// are not part of the variable's identity — `$script:count` and `$count` name
/// the same storage at script scope. Drop the modifier so both spellings
/// resolve to one binding; keep `env:` since it names a different store.
fn scope_qualified_name(name: &str) -> &str {
    match name.split_once(':') {
        Some((scope, rest))
            if matches!(
                scope.to_lowercase().as_str(),
                "script" | "global" | "local" | "private" | "using" | "variable"
            ) && !rest.is_empty() =>
        {
            rest
        }
        _ => name,
    }
}

fn walk_bare_word(raw: &str) -> Expression {
    match raw.to_lowercase().as_str() {
        "true" => Expression::bool(true),
        "false" => Expression::bool(false),
        "null" => Expression::null(),
        _ => Expression::ident(raw),
    }
}

/// The class name behind a `[Type]` literal, once the walker has turned it into
/// a string. The FULL dotted name is kept: `[System.Text.StringBuilder]` has to
/// stay whole so the shared namespace resolver can match it through the dotnet
/// tree-mount. Truncating to the last segment would discard exactly what
/// `resolve_profile_namespace_chain` matches on.
fn type_name_of(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) if !name.trim().is_empty() => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", type_name_of(object)?, field))
        }
        _ => None,
    }
}

/// `[System.Math]` becomes the member chain `System.Math`, the same shape the
/// C# walker produces for a dotted `System.*` name. That is what lets the shared
/// namespace resolver reach it through the dotnet tree-mount — a plain string
/// object could never resolve.
fn type_literal_expr(name: &str) -> Expression {
    let name = name.trim();
    if name.is_empty() {
        return Expression::null();
    }
    let mut segments = name.split('.').filter(|s| !s.is_empty());
    let Some(first) = segments.next() else {
        return Expression::null();
    };
    let mut expr = Expression::ident(first);
    for seg in segments {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: seg.to_string(),
            null_safe: false,
        });
    }
    expr
}

/// PowerShell's `.Count` / `.Length` are the same idea as JS `.length`, so the
/// walker rewrites them the way PHP's `count($a)` becomes `a.length`. Anything
/// else keeps its spelling and resolves through the profile.
fn normalize_member_name(name: &str) -> String {
    match name.to_lowercase().as_str() {
        "count" | "length" => "length".to_string(),
        _ => name.to_string(),
    }
}

fn type_literal_name(raw: &str) -> &str {
    raw.trim().trim_start_matches('[').trim_end_matches(']')
}

/// Map a PowerShell operator spelling onto the shared `BinOp`.
fn build_binary(op_raw: &str, left: Expression, right: Expression) -> Expression {
    let op_text = op_raw.trim().to_lowercase();
    let symbolic = matches!(op_text.as_str(), "+" | "-" | "*" | "/" | "%" | "**");

    let word = if symbolic {
        op_text.clone()
    } else {
        let stripped = op_text.trim_start_matches('-');
        // `-ceq` / `-ieq` are the case-sensitive / case-insensitive forms of
        // `-eq`; they select comparison casing, not a different operator.
        match stripped
            .strip_prefix('c')
            .or_else(|| stripped.strip_prefix('i'))
        {
            Some(rest) if is_comparison_word(rest) => rest.to_string(),
            _ => stripped.to_string(),
        }
    };

    // `-contains` / `-notcontains` take the collection on the LEFT, the inverse
    // of `-in` / `-notin`. Swapping the operands is what makes both spellings
    // reach the same shared `In` lowering.
    // `-notin` / `-notcontains` become `Not(In(…))` rather than `NotIn`: the
    // negation then goes through the same materialization the other comparison
    // operators use, so the result is the boolean `$false`, not `1`.
    // `-match` is a REGEX test and `-like` a wildcard one. `BinOp::Like` lowers
    // to `ecma:regexp.test`, which takes (pattern, string) while the operands
    // arrive as (string, pattern) — the arm even documents itself as unreachable
    // because VB rewrites before it. So rewrite here too, onto `Regex.IsMatch`
    // in the dotnet tree, which takes the input first.
    if matches!(word.as_str(), "match" | "notmatch" | "like" | "notlike") {
        let pattern = if word.starts_with("like") || word == "notlike" {
            glob_to_regex(right)
        } else {
            right
        };
        let test = dotnet_static_call(
            "System.Text.RegularExpressions.Regex",
            "IsMatch",
            vec![left, pattern],
        );
        return if word.starts_with("not") {
            negate(test)
        } else {
            test
        };
    }

    // `-is` / `-isnot` are TYPE TESTS. `BinOp::Is` is reference equality
    // (Python's `is`), so it answered `$x -is [int]` by comparing 42 to the
    // resolved type — false for every operand, primitive or class alike.
    if word == "is" || word == "isnot" {
        let test = type_test_expr(left, &right);
        return if word == "isnot" { negate(test) } else { test };
    }

    match word.as_str() {
        "notcontains" => {
            return negate(Expression::new(ExprKind::Binary {
                op: BinOp::In,
                left: Box::new(right),
                right: Box::new(left),
            }));
        }
        "notin" => {
            return negate(Expression::new(ExprKind::Binary {
                op: BinOp::In,
                left: Box::new(left),
                right: Box::new(right),
            }));
        }
        _ => {}
    }

    let (op, left, right) = match word.as_str() {
        "contains" => (BinOp::In, right, left),
        other => {
            let op = match other {
                "+" => BinOp::Add,
                "-" => BinOp::Sub,
                "*" => BinOp::Mul,
                "/" => BinOp::Div,
                "%" => BinOp::Mod,
                "**" => BinOp::Pow,
                "eq" => BinOp::Eq,
                "ne" => BinOp::NotEq,
                "gt" => BinOp::Gt,
                "ge" => BinOp::GtEq,
                "lt" => BinOp::Lt,
                "le" => BinOp::LtEq,
                "and" => BinOp::And,
                "or" => BinOp::Or,
                "xor" => BinOp::Xor,
                "in" => BinOp::In,
                "notin" => BinOp::NotIn,
                "is" => BinOp::Is,
                "isnot" => BinOp::IsNot,
                "like" | "match" => BinOp::Like,
                "band" => BinOp::BitAnd,
                "bor" => BinOp::BitOr,
                "bxor" => BinOp::BitXor,
                "shl" => BinOp::Shl,
                "shr" => BinOp::Shr,
                "??" => BinOp::NullCoalesce,
                // `-split` / `-replace` / `-join` carry an argument LIST, which
                // arrives here as an `Array` when the source spelled more than
                // one. Spread it: `$s -replace 'a', 'b'` is `replace(a, b)`,
                // not `replace(['a','b'])`.
                // `-split` takes a REGEX, not a literal: `'a.b.c' -split '\.'`
                // splits on the dot. `strings.split` is a literal split and was
                // matching the two characters `\.`, so the result was one
                // element. `Regex.Split` is the operator's actual meaning, and
                // it is the same dotnet route `-match` and `-like` take.
                "split" => {
                    let mut args = vec![left];
                    args.extend(spread_operands(right));
                    return dotnet_static_call(
                        "System.Text.RegularExpressions.Regex",
                        "Split",
                        args,
                    );
                }
                // `-replace` is regex too, and replaces EVERY match. The string
                // `replace` is literal and first-only, so `'the cat sat on the
                // mat' -replace '[cm]at','dog'` changed nothing.
                "replace" => {
                    let mut args = vec![left];
                    args.extend(spread_operands(right));
                    return dotnet_static_call(
                        "System.Text.RegularExpressions.Regex",
                        "Replace",
                        args,
                    );
                }
                "join" => {
                    return method_call_expr(left, "join", spread_operands(right));
                }
                "notlike" | "notmatch" => {
                    return Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(Expression::new(ExprKind::Binary {
                            op: BinOp::Like,
                            left: Box::new(left),
                            right: Box::new(right),
                        })),
                    });
                }
                // Unreachable for a valid operator: every spelling `cmp_word`
                // admits is mapped above. Keeping `Eq` here would turn a future
                // grammar addition into a silently wrong comparison, so fall
                // back to concatenation, which shows up as a wrong VALUE rather
                // than a plausible-looking boolean.
                _ => BinOp::Add,
            };
            (op, left, right)
        }
    };

    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn negate(expr: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(expr),
    })
}

/// `-like`'s wildcard pattern as the equivalent anchored regex, so both
/// operators reach one matcher. Only a literal pattern can be translated; a
/// computed one is passed through and matches as a regex, which is wrong but
/// visible, rather than silently never matching.
fn glob_to_regex(pattern: Expression) -> Expression {
    let ExprKind::Lit(Literal::Str(glob)) = &pattern.kind else {
        return pattern;
    };

    let mut out = String::from("^");
    for ch in glob.chars() {
        match ch {
            '*' => out.push_str(".*"),
            '?' => out.push('.'),
            // `[a-c]` is a character class in both syntaxes.
            '[' | ']' => out.push(ch),
            c if "\\.+^$(){}|".contains(c) => {
                out.push('\\');
                out.push(c);
            }
            c => out.push(c),
        }
    }
    out.push('$');
    Expression::string(&out)
}

/// `$x -is [T]`. A BUILT-IN spelling is a value-kind question and answers
/// through `typeof`; anything else names a type and answers through
/// `instanceof`. The two cannot share a mechanism: `42` has no prototype chain
/// to walk, and a user class has no `typeof` tag of its own.
fn type_test_expr(value: Expression, type_expr: &Expression) -> Expression {
    let name = dotted_name_of(type_expr).unwrap_or_default();
    let leaf = name.rsplit('.').next().unwrap_or(&name).to_lowercase();

    let tag = match leaf.trim_end_matches("[]") {
        "int" | "int16" | "int32" | "int64" | "long" | "short" | "byte" | "sbyte" | "uint"
        | "uint16" | "uint32" | "uint64" | "ushort" | "double" | "single" | "float" | "decimal" => {
            Some("number")
        }
        "string" | "char" => Some("string"),
        "bool" | "boolean" => Some("boolean"),
        // A PowerShell hashtable is an ordinary object here — `@{a=1}` walks to
        // `ExprKind::Object`, not a Map — so the value-kind tag answers it.
        "hashtable" | "dictionary" | "ordereddictionary" | "ordered" | "psobject"
        | "pscustomobject" => Some("object"),
        _ => None,
    };

    if let Some(tag) = tag {
        return Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(value)))),
            right: Box::new(Expression::string(tag)),
        });
    }

    // `[array]` is the runtime's own Array; every other name is taken as
    // written so a user class reaches its own prototype. Double-negated
    // because `InstanceOf` yields a truthy value rather than a boolean, and
    // `$r -ne $true` in the tests compares against the real `$true`.
    let ctor = if leaf.trim_end_matches("[]") == "array" || leaf.ends_with("[]") {
        "Array".to_string()
    } else {
        name
    };

    negate(negate(Expression::new(ExprKind::Binary {
        op: BinOp::InstanceOf,
        left: Box::new(value),
        right: Box::new(type_literal_expr(&ctor)),
    })))
}

/// The dotted name a type-literal member chain spells, if it is one.
fn dotted_name_of(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            Some(format!("{}.{}", dotted_name_of(object)?, field))
        }
        _ => None,
    }
}

fn is_comparison_word(word: &str) -> bool {
    matches!(
        word,
        "eq" | "ne"
            | "gt"
            | "ge"
            | "lt"
            | "le"
            | "like"
            | "notlike"
            | "match"
            | "notmatch"
            | "contains"
            | "notcontains"
            | "in"
            | "notin"
            | "replace"
            | "split"
            | "join"
    )
}

/// `@"…"@` / `@'…'@`. The body starts on the line AFTER the opener and ends at
/// the newline before the closer, so neither delimiter line is content. The
/// double-quoted form interpolates; the single-quoted form is literal, and
/// neither treats its own quote character as a delimiter — that is the whole
/// point of a here-string.
fn walk_here_string(__w: &mut PsWalker, raw: &str, interpolating: bool) -> Expression {
    // Past `@"` / `@'` and its line, up to the newline before `"@` / `'@`.
    let body = raw
        .get(2..raw.len().saturating_sub(2))
        .unwrap_or("")
        .strip_prefix('\n')
        .or_else(|| {
            raw.get(2..raw.len().saturating_sub(2))
                .unwrap_or("")
                .split_once('\n')
                .map(|(_, rest)| rest)
        })
        .unwrap_or("");
    let body = body.strip_suffix('\n').unwrap_or(body);

    if interpolating {
        parse_double_quoted_string(__w, body)
    } else {
        Expression::string(body)
    }
}

fn parse_double_quoted_string(__w: &mut PsWalker, text: &str) -> Expression {
    let chars: Vec<char> = text.chars().collect();
    let mut parts: Vec<InterpolPart> = Vec::new();
    let mut literal = String::new();
    let mut i = 0;

    while i < chars.len() {
        let ch = chars[i];

        if ch == '`' {
            if let Some(next) = chars.get(i + 1) {
                if let Some(mapped) = parse_powershell_escape(*next) {
                    literal.push(mapped);
                } else {
                    literal.push(*next);
                }
                i += 2;
            } else {
                i += 1;
            }
            continue;
        }

        // `""` inside a double-quoted string is ONE literal quote — the same
        // doubling `''` uses inside a single-quoted one. The grammar matches the
        // pair so it cannot end the string; this is where it collapses.
        if ch == '"' && chars.get(i + 1) == Some(&'"') {
            literal.push('"');
            i += 2;
            continue;
        }

        if ch != '$' {
            literal.push(ch);
            i += 1;
            continue;
        }

        if i + 1 >= chars.len() {
            literal.push(ch);
            i += 1;
            continue;
        }

        if !literal.is_empty() {
            parts.push(InterpolPart::Text(std::mem::take(&mut literal)));
            literal.clear();
        }

        let next = chars[i + 1];
        match next {
            '$' | '`' => {
                literal.push('$');
                i += 2;
            }

            '{' => {
                if let Some(close) = find_matching_in_chars(&chars, i + 1, '{', '}') {
                    let inner = chars[(i + 2)..close].iter().collect::<String>();
                    parts.push(InterpolPart::Expr(parse_script_expression(__w, &inner)));
                    i = close + 1;
                } else {
                    literal.push('$');
                    literal.push('{');
                    i += 2;
                }
            }

            '(' => {
                if let Some(close) = find_matching_in_chars(&chars, i + 1, '(', ')') {
                    let inner = chars[(i + 2)..close].iter().collect::<String>();
                    parts.push(InterpolPart::Expr(parse_script_expression(__w, &inner)));
                    i = close + 1;
                } else {
                    literal.push('$');
                    i += 1;
                }
            }

            c if is_ps_variable_char(c) => {
                let mut end = i + 1;
                while end < chars.len() && is_ps_variable_char(chars[end]) {
                    end += 1;
                }

                let name = chars[(i + 1)..end].iter().collect::<String>();
                parts.push(InterpolPart::Expr(parse_script_expression(__w, &format!(
                    "${}",
                    name
                ))));
                i = end;
            }

            _ => {
                literal.push('$');
                i += 1;
            }
        }
    }

    if !literal.is_empty() {
        parts.push(InterpolPart::Text(std::mem::take(&mut literal)));
    }

    match parts.as_slice() {
        [] => Expression::string(""),
        [InterpolPart::Text(s)] => Expression::string(s),
        // A lone `"$x"` is NOT `$x` — it is `$x` converted to a string. Handing
        // the expression back unwrapped made `"$n" -eq '5'` compare a number to
        // a string and `"$arr"` compare an array to one. One part still goes
        // through `Interpolation` so the string conversion happens.
        _ => Expression::new(ExprKind::Interpolation(parts)),
    }
}

/// The body of a `$( … )` / `${ … }` interpolation. PowerShell allows a whole
/// statement list here, so fall back to parsing statements and taking the last
/// expression when it is not a single expression.
fn parse_script_expression(__w: &mut PsWalker, raw: &str) -> Expression {
    let text = raw.trim();
    if text.is_empty() {
        return Expression::null();
    }

    if let Some(expr) = parse_expr_fragment(__w, text) {
        return expr;
    }

    match super::PowerShellParser::parse(Rule::program, text) {
        Ok(mut pairs) => match pairs.next() {
            Some(root) => statements_as_value(collect_statements(__w, root)),
            None => Expression::null(),
        },
        Err(_) => Expression::string(text),
    }
}

/// The value of a `$( … )` whose body is a statement rather than an expression.
///
/// `last_expression_of` alone answers null for a BRANCH — an `if` is not an
/// `Expr` statement, so `"v=$(if ($true) { 'yes' })"` interpolated nothing.
/// A branch's value is the last expression of whichever arm ran, which is
/// exactly what `statement_value_expr` builds and what `$x = if (…) { … }`
/// already goes through.
fn statements_as_value(mut stmts: Vec<Statement>) -> Expression {
    match stmts.pop() {
        Some(stmt) => match stmt.kind {
            StmtKind::Expr(expr) => expr,
            StmtKind::If { .. } | StmtKind::Switch { .. } | StmtKind::Try { .. } => {
                statement_value_expr(stmt)
            }
            _ => Expression::null(),
        },
        None => Expression::null(),
    }
}

fn is_numeric_like(raw: &str) -> bool {
    let text = raw.trim();
    if text.is_empty() {
        return false;
    }

    text.parse::<i64>().is_ok() || text.parse::<f64>().is_ok() || text == "$null"
}

fn looks_like_command_invocation(raw: &str) -> bool {
    let text = raw.trim();
    if text.is_empty() || is_numeric_like(text) {
        return false;
    }

    if text.contains(' ') {
        if let Some(head) = text.split_whitespace().next() {
            return looks_like_command_name(head) || is_path_like_command(head);
        }
        return false;
    }

    looks_like_command_name(text) || is_path_like_command(text)
}

fn parse_powershell_escape(next: char) -> Option<char> {
    match next {
        '`' => Some('`'),
        '"' => Some('"'),
        '\'' => Some('\''),
        '$' => Some('$'),
        'n' => Some('\n'),
        'r' => Some('\r'),
        't' => Some('\t'),
        '0' => Some('\0'),
        _ => None,
    }
}

fn looks_like_command_name(raw: &str) -> bool {
    let mut chars = raw.chars();
    let Some(first) = chars.next() else {
        return false;
    };

    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }

    chars.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '-')
}

fn is_ps_variable_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '-' || c == ':'
}

fn find_matching_in_chars(
    chars: &[char],
    open_idx: usize,
    open: char,
    close: char,
) -> Option<usize> {
    let mut in_single = false;
    let mut in_double = false;
    let mut escaped = false;
    let mut depth = 0usize;

    for (idx, ch) in chars.iter().enumerate() {
        let ch = *ch;

        if escaped {
            escaped = false;
            continue;
        }

        if ch == '`' {
            escaped = true;
            continue;
        }

        if in_single {
            if ch == '\'' {
                in_single = false;
            }
            continue;
        }

        if in_double {
            if ch == '"' {
                in_double = false;
                continue;
            }

            if ch == open {
                depth += 1;
            } else if ch == close && depth > 0 {
                depth -= 1;
                if depth == 0 {
                    return Some(idx);
                }
            }

            continue;
        }

        if ch == '\'' {
            in_single = true;
            continue;
        }

        if ch == '"' {
            in_double = true;
            continue;
        }

        if idx < open_idx {
            continue;
        }

        if idx == open_idx {
            if ch == open {
                depth = 1;
            }
            continue;
        }

        if ch == open {
            depth += 1;
        } else if ch == close {
            if depth == 0 {
                return None;
            }
            depth -= 1;
            if depth == 0 {
                return Some(idx);
            }
        }
    }

    None
}

fn split_command_tokens(input: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    let mut quote = None;
    let mut chars = input.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch == '\\' {
            if let Some(next) = chars.next() {
                current.push(next);
            }
            continue;
        }
        if let Some(q) = quote {
            current.push(ch);
            if ch == q {
                quote = None;
            }
            continue;
        } else if ch == '"' || ch == '\'' {
            quote = Some(ch);
            current.push(ch);
            continue;
        }

        if ch.is_whitespace() && quote.is_none() {
            if !current.trim().is_empty() {
                tokens.push(current.trim().to_string());
                current.clear();
            }
            continue;
        }
        current.push(ch);
    }

    if !current.trim().is_empty() {
        tokens.push(current.trim().to_string());
    }

    tokens
}

#[allow(dead_code)]
trait IdentEmpty {
    fn as_ident_empty(&self) -> bool;
}

impl IdentEmpty for Expression {
    fn as_ident_empty(&self) -> bool {
        match &self.kind {
            ExprKind::Ident(v) => v.is_empty(),
            _ => false,
        }
    }
}
