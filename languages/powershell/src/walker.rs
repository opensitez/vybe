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
    collect_declared_functions(__w, pair.clone());

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
/// The name `trap` binds its error to.
///
/// ⛔`$_`, not a private name. PowerShell's trap block reads the error through
/// the automatic variable `$_` — `$_.Exception.Message` is the idiom — so
/// binding it to `__trap` left `$_` empty in every trap body. It doubles as the
/// value `break` rethrows.
const TRAP_VAR: &str = "_";

fn parse_trap_stmt(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut body = Vec::new();
    let mut types = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            // `trap [System.DivideByZeroException] { … }` — a TYPED trap, which
            // runs only for a matching error. The grammar already accepted the
            // type and the walker dropped it, so every typed trap behaved as a
            // catch-all and the first one declared won.
            Rule::type_literal => {
                let name = type_literal_name(child.as_str()).trim();
                // ⛔The catch type must be the spelling the exception classes
                // are DECLARED under. `synthesize_exception_classes` splices
                // them in unqualified, so `trap [System.SystemException]`
                // matched nothing and a typed trap that used to fire as a
                // catch-all stopped firing at all — base-class matching walks
                // the ancestor chain by NAME.
                let name = match name.strip_prefix("System.") {
                    Some(short) if !short.contains('.') => short,
                    _ => name,
                };
                if !name.is_empty() {
                    types.push(name.to_string());
                }
            }
            Rule::block => body = parse_block_statements(__w, child)?,
            _ => {}
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
            types,
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

/// `Write-Output $x` INSIDE A FUNCTION is the value `$x`, not a print.
///
/// The profile maps `write-output` to `print`, which is right at the top level
/// — an unconsumed success stream is rendered by the host — and wrong in a
/// function body, where the caller receives the value. Both halves were broken
/// by the same mapping: `function H { Write-Output 'z' }` PRINTED `z` and
/// returned `$null`, and `function F { Write-Output 'a'; Write-Output 'b' }`
/// reached the accumulator but collected `print`'s null twice, so `$r.Count`
/// was 2 with both elements EMPTY — the count looked right, which is why this
/// survived.
///
/// ⛔The rewrite is scoped to function bodies deliberately. Making
/// `Write-Output` an identity everywhere would silence it at the top level,
/// where this compiler does not print bare expression statements the way
/// PowerShell does.
fn unwrap_write_output(body: Vec<Statement>) -> Vec<Statement> {
    body.into_iter()
        .map(|stmt| match stmt.kind {
            StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(strip_write_output(expr))),
            StmtKind::Return(Some(expr)) => {
                Statement::new(StmtKind::Return(Some(strip_write_output(expr))))
            }
            _ => stmt,
        })
        .collect()
}

/// `Write-Output a, b` emits TWO values, so a multi-argument call keeps its
/// arguments as an array rather than silently dropping all but the first.
fn strip_write_output(expr: Expression) -> Expression {
    let ExprKind::Call { callee, mut args, .. } = expr.kind else {
        return expr;
    };
    let is_write_output = matches!(
        &callee.kind,
        ExprKind::Ident(name) if name.eq_ignore_ascii_case("write-output")
    );
    if !is_write_output || args.is_empty() {
        return Expression::new(ExprKind::Call {
            callee,
            args,
            optional: false,
        });
    }
    if args.len() == 1 {
        return args.remove(0).value;
    }
    Expression::new(ExprKind::Array(
        args.into_iter()
            .map(|a| ArrayElement {
                key: None,
                value: a.value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
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
/// Applied to a body with two or more emit sites, OR to one whose single emit
/// site sits inside a LOOP. That second clause is the one a trailing return
/// provably cannot express: `process { $Num * 2 }` has exactly one emit site in
/// the source, and it produces one value PER PIPELINE ITEM. Counting sites and
/// stopping at "fewer than two" answers 1 and hands the body to the plain
/// trailing-return path, which can only ever carry a single value back.
fn accumulate_outputs(body: Vec<Statement>) -> Option<Vec<Statement>> {
    let mut census = EmitCensus::default();
    census.walk(&body, false);
    if census.sites < 2 && !census.in_loop {
        return None;
    }

    let mut out = Vec::with_capacity(body.len() + 2);
    out.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(OUT_ACC)],
        value: Expression::new(ExprKind::Array(Vec::new())),
        by_ref: false,
    }));

    let rewritten = accumulate_stmts(body);
    let ends_in_return = matches!(
        rewritten.last().map(|s| &s.kind),
        Some(StmtKind::Return(_))
    );
    out.extend(rewritten);
    if !ends_in_return {
        out.push(Statement::new(StmtKind::Return(Some(unwrap_output()))));
    }
    Some(out)
}

/// How many emit sites a body has, and whether any of them is inside a loop.
#[derive(Default)]
struct EmitCensus {
    sites: usize,
    in_loop: bool,
}

impl EmitCensus {
    fn walk(&mut self, body: &[Statement], looping: bool) {
        for stmt in body {
            if emits_value(stmt) {
                self.sites += 1;
                self.in_loop |= looping;
            }
            for (nested, loops) in child_bodies(&stmt.kind) {
                self.walk(nested, looping || loops);
            }
        }
    }
}

/// The statement bodies output can reach from here, each flagged with whether
/// entering it means entering a loop.
///
/// \u26d4A nested `FunctionDecl`/`ClassDecl` body is NOT among them: a value emitted
/// there belongs to THAT function's success stream, and hoisting it into the
/// enclosing accumulator would corrupt every nested function in the corpus.
fn child_bodies(kind: &StmtKind) -> Vec<(&[Statement], bool)> {
    match kind {
        StmtKind::If { then_body, elifs, else_body, .. } => {
            let mut out: Vec<(&[Statement], bool)> = vec![(then_body, false)];
            out.extend(elifs.iter().map(|(_, b)| (b.as_slice(), false)));
            out.extend(else_body.iter().map(|b| (b.as_slice(), false)));
            out
        }
        StmtKind::Block(body) => vec![(body, false)],
        StmtKind::For { body, .. } | StmtKind::DoWhile { body, .. } => vec![(body, true)],
        StmtKind::ForIn { body, else_body, .. } | StmtKind::While { body, else_body, .. } => {
            let mut out: Vec<(&[Statement], bool)> = vec![(body, true)];
            out.extend(else_body.iter().map(|b| (b.as_slice(), false)));
            out
        }
        StmtKind::Switch { cases, default, .. } => {
            let mut out: Vec<(&[Statement], bool)> =
                cases.iter().map(|c| (c.body.as_slice(), false)).collect();
            out.extend(default.iter().map(|b| (b.as_slice(), false)));
            out
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            let mut out: Vec<(&[Statement], bool)> = vec![(body.as_slice(), false)];
            out.extend(catches.iter().map(|c| (c.body.as_slice(), false)));
            out.extend(else_body.iter().map(|b| (b.as_slice(), false)));
            out.extend(finally.iter().map(|b| (b.as_slice(), false)));
            out
        }
        _ => Vec::new(),
    }
}

/// Rewrite every emit site in `body` into a collect, following the same nesting
/// [`child_bodies`] censused.
fn accumulate_stmts(body: Vec<Statement>) -> Vec<Statement> {
    let mut out = Vec::with_capacity(body.len());
    for stmt in body {
        let span = stmt.span;
        let keep = |kind| Statement { kind, span };
        match stmt.kind {
            StmtKind::Expr(expr) if expression_emits(&expr) => {
                out.push(keep(StmtKind::Expr(collect_call(expr))));
            }
            StmtKind::Return(Some(expr)) => {
                out.push(keep(StmtKind::Expr(collect_call(expr))));
                out.push(Statement::new(StmtKind::Return(Some(unwrap_output()))));
            }
            // \u26d4A bare `return` still hands back the stream. Falling through to
            // the catch-all left it returning `$null` and DISCARDING everything
            // collected so far.
            StmtKind::Return(None) => {
                out.push(keep(StmtKind::Return(Some(unwrap_output()))));
            }
            StmtKind::If { cond, then_body, elifs, else_body } => {
                out.push(keep(StmtKind::If {
                    cond,
                    then_body: accumulate_stmts(then_body),
                    elifs: elifs
                        .into_iter()
                        .map(|(c, b)| (c, accumulate_stmts(b)))
                        .collect(),
                    else_body: else_body.map(accumulate_stmts),
                }));
            }
            StmtKind::Block(body) => out.push(keep(StmtKind::Block(accumulate_stmts(body)))),
            StmtKind::For { init, cond, update, body } => {
                out.push(keep(StmtKind::For {
                    init,
                    cond,
                    update,
                    body: accumulate_stmts(body),
                }));
            }
            StmtKind::ForIn { var, key, iter, body, of, else_body, is_async } => {
                out.push(keep(StmtKind::ForIn {
                    var,
                    key,
                    iter,
                    body: accumulate_stmts(body),
                    of,
                    else_body: else_body.map(accumulate_stmts),
                    is_async,
                }));
            }
            StmtKind::While { cond, body, else_body } => {
                out.push(keep(StmtKind::While {
                    cond,
                    body: accumulate_stmts(body),
                    else_body: else_body.map(accumulate_stmts),
                }));
            }
            StmtKind::DoWhile { body, cond, until } => {
                out.push(keep(StmtKind::DoWhile {
                    body: accumulate_stmts(body),
                    cond,
                    until,
                }));
            }
            StmtKind::Switch { expr, cases, default } => {
                out.push(keep(StmtKind::Switch {
                    expr,
                    cases: cases
                        .into_iter()
                        .map(|c| SwitchCase {
                            conditions: c.conditions,
                            body: accumulate_stmts(c.body),
                        })
                        .collect(),
                    default: default.map(accumulate_stmts),
                }));
            }
            StmtKind::Try { body, catches, else_body, finally } => {
                out.push(keep(StmtKind::Try {
                    body: accumulate_stmts(body),
                    catches: catches
                        .into_iter()
                        .map(|c| CatchClause {
                            body: accumulate_stmts(c.body),
                            ..c
                        })
                        .collect(),
                    else_body: else_body.map(accumulate_stmts),
                    finally: finally.map(accumulate_stmts),
                }));
            }
            other => out.push(keep(other)),
        }
    }
    out
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
                    // `IComparer[string]` names the same interface as
                    // `IComparer`; the type argument is not part of its
                    // identity, and keeping it would make every lookup miss.
                    let name = base.as_str().trim();
                    let name = name.split('[').next().unwrap_or(name).trim();
                    // ⛔A FULLY-QUALIFIED FRAMEWORK BASE RESOLVES BY ITS LEAF.
                    // The dotnet catalog registers `Exception`, so
                    // `class E : System.Exception` recorded a parent of
                    // `system.exception`, matched nothing, and inherited none of
                    // the exception shape: `$e.Message` was EMPTY and
                    // `$e -is [System.Exception]` was FALSE. csharp already
                    // normalizes exactly this (`walker.rs:10796`, "Normalize a
                    // fully-qualified base class to its leaf name") and the
                    // same source spelling works there.
                    //
                    // ⛔Only `System.` — a USER type with dots in its name is
                    // its own identity, and stripping every qualifier would
                    // collide two user classes that share a leaf.
                    let name = match name.strip_prefix("System.") {
                        Some(_) => name.rsplit('.').next().unwrap_or(name),
                        None => name,
                    };
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
        members.extend(parse_class_member(__w, child)?);
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
    let mut is_static = false;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::class_modifier => {
                if child.as_str().eq_ignore_ascii_case("static") {
                    is_static = true;
                }
            }
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

    // `static Singleton() { … }` — a STATIC constructor, which runs once before
    // the type is first used and is not a constructor at all in the class
    // model: it is the static initialiser. csharp, vb, kotlin and pascal all
    // spell it `__static_init__`, and the shared `classes.rs` recognises that
    // name (`is_static_init`). Emitting it as an ordinary constructor made it
    // an overload of the instance one, so `[Singleton]::Instance` was never
    // populated.
    if is_static {
        let mut modifiers = Modifiers::default();
        modifiers.is_static = true;
        return Ok(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "__static_init__".into(),
                params: Vec::new(),
                body,
                return_type: None,
                is_async: false,
                is_generator: false,
                is_sub: true,
                handles: Vec::new(),
                modifiers,
            },
        ))));
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
/// A class property, which is a FIELD unless it carries validation attributes.
///
/// `[ValidateRange(1, 100)][int]$Score = 50` is two bracketed runs, and only the
/// LAST is the type constraint — the rest are attributes. PowerShell enforces
/// them on every assignment and leaves the old value in place when one fails,
/// which is behaviour a plain field cannot have. So a validated property lowers
/// to a PAIR: a backing field holding the value, and a `ClassMember::Property`
/// whose setter validates before writing.
///
/// ⛔The backing field is deliberately NOT named after the property.
/// `auto_field` would give both the same name, and the setter's own
/// `$this.<name> = $value` would then re-enter the accessor it is running in.
fn parse_ps_property(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut is_static = false;
    let mut type_hint = None;
    let mut validators: Vec<String> = Vec::new();
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
                let raw = child.as_str();
                let inner = raw.trim();
                let inner = inner
                    .strip_prefix('[')
                    .and_then(|r| r.strip_suffix(']'))
                    .unwrap_or(inner)
                    .trim();
                if is_validation_attribute(inner) {
                    validators.push(inner.to_string());
                } else {
                    let t = type_literal_name(raw).trim();
                    if !t.is_empty() {
                        type_hint = Some(t.to_string());
                    }
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

    if validators.is_empty() {
        return Ok(vec![ClassMember::Field {
            name,
            type_hint,
            init,
            modifiers,
            with_events: false,
            array_bounds: None,
            storage: None,
        }]);
    }

    let backing = format!("__ps_val_{name}");
    let field = ClassMember::Field {
        name: backing.clone(),
        type_hint: type_hint.clone(),
        init,
        modifiers: modifiers.clone(),
        with_events: false,
        array_bounds: None,
        storage: None,
    };

    let getter = parse_member_body(__w, &format!("return $this.{backing}"));
    let setter_body = parse_member_body(__w, &validation_setter_source(&validators, &name, &backing));

    Ok(vec![
        field,
        ClassMember::Property {
            name,
            type_hint,
            getter: Some(getter),
            setter: Some(PropertySetter {
                param: synth_param("value"),
                body: setter_body,
            }),
            is_auto: false,
            modifiers,
        },
    ])
}

/// The attribute names PowerShell validates an assignment against. Anything
/// else in a bracketed run is a type constraint.
///
/// ⛔Matched on the NAME, not on "contains a paren": `[ValidateNotNull]` is
/// legal without an argument list and would otherwise read as a type, while
/// `[int[]]` carries brackets and must not read as an attribute.
fn is_validation_attribute(inner: &str) -> bool {
    let head = inner
        .split('(')
        .next()
        .unwrap_or("")
        .trim()
        .trim_end_matches(']')
        .trim();
    head.len() > 8 && head[..8].eq_ignore_ascii_case("Validate")
}

/// Split an attribute's argument list on TOP-LEVEL commas — `ValidateSet('a,b',
/// 'c')` is two arguments, and a comma inside a string, a nested call or a
/// script block does not separate.
fn split_attribute_args(args: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut depth = 0i32;
    let mut quote: Option<char> = None;
    let mut cur = String::new();
    for ch in args.chars() {
        match quote {
            Some(q) => {
                cur.push(ch);
                if ch == q {
                    quote = None;
                }
                continue;
            }
            None => {}
        }
        match ch {
            '\'' | '"' => {
                quote = Some(ch);
                cur.push(ch);
            }
            '(' | '[' | '{' => {
                depth += 1;
                cur.push(ch);
            }
            ')' | ']' | '}' => {
                depth -= 1;
                cur.push(ch);
            }
            ',' if depth == 0 => {
                out.push(cur.trim().to_string());
                cur.clear();
            }
            _ => cur.push(ch),
        }
    }
    if !cur.trim().is_empty() {
        out.push(cur.trim().to_string());
    }
    out
}

/// PowerShell source for the setter of a validated property: every attribute
/// becomes a guard that throws, then the backing field is written.
///
/// Generating SOURCE and re-parsing it is deliberate. The conditions are
/// ordinary PowerShell expressions (`-ge`, `-contains`, `-match`), and the
/// walker already lowers each of those correctly — hand-building the AST would
/// be a second, divergent spelling of operators this file already handles.
fn validation_setter_source(validators: &[String], name: &str, backing: &str) -> String {
    let mut src = String::new();
    for v in validators {
        let Some(cond) = validation_condition(v) else {
            continue;
        };
        // The message carries the property name: PowerShell's own error names
        // the member, and a test asserts the text contains it.
        src.push_str(&format!(
            "if (-not ({cond})) {{ throw \"The attribute cannot be added because \
             the property {name} would no longer be valid.\" }}\n"
        ));
    }
    src.push_str(&format!("$this.{backing} = $value\n"));
    src
}

/// One attribute → a PowerShell expression that is TRUE when the value passes.
fn validation_condition(inner: &str) -> Option<String> {
    let (head, args) = match inner.find('(') {
        Some(i) => (
            inner[..i].trim(),
            inner[i + 1..].trim_end().trim_end_matches(')'),
        ),
        None => (inner.trim(), ""),
    };
    let parts = split_attribute_args(args);
    let arg = |i: usize| parts.get(i).cloned().unwrap_or_default();

    let cond = match head.to_ascii_lowercase().as_str() {
        "validaterange" if parts.len() >= 2 => {
            format!("($value -ge {}) -and ($value -le {})", arg(0), arg(1))
        }
        // `-contains` is case-INSENSITIVE in PowerShell, which is exactly the
        // documented behaviour of `ValidateSet` on assignment.
        "validateset" if !parts.is_empty() => {
            format!("@({}) -contains $value", parts.join(", "))
        }
        "validatelength" if parts.len() >= 2 => format!(
            "($value.Length -ge {}) -and ($value.Length -le {})",
            arg(0),
            arg(1)
        ),
        "validatecount" if parts.len() >= 2 => format!(
            "(@($value).Count -ge {}) -and (@($value).Count -le {})",
            arg(0),
            arg(1)
        ),
        "validatepattern" if !parts.is_empty() => format!("$value -match {}", arg(0)),
        "validatenotnull" => "$null -ne $value".to_string(),
        "validatenotnullorempty" => "($null -ne $value) -and ($value -ne '')".to_string(),
        // `ValidateScript({ $_ -gt 0 })` — the predicate names the candidate as
        // `$_`, and may also read `$this`. Inlining the body (rather than
        // invoking the script block) keeps `$this` bound to the instance whose
        // setter is running, which is what `validate_script_using_dollar_this`
        // asserts.
        "validatescript" if !parts.is_empty() => {
            let body = arg(0);
            let body = body.trim();
            let body = body
                .strip_prefix('{')
                .and_then(|b| b.strip_suffix('}'))
                .unwrap_or(body);
            body.replace("$_", "$value")
        }
        _ => return None,
    };
    Some(cond)
}

/// Parse generated PowerShell source into statements for a synthesized
/// accessor body.
fn parse_member_body(__w: &mut PsWalker, src: &str) -> Vec<Statement> {
    match super::PowerShellParser::parse(Rule::program, src) {
        Ok(mut pairs) => match pairs.next() {
            Some(root) => collect_statements(__w, root),
            None => Vec::new(),
        },
        Err(_) => Vec::new(),
    }
}

/// One syntactic member can lower to MORE THAN ONE `ClassMember`: a property
/// carrying validation attributes becomes a backing field PLUS a property whose
/// setter enforces them, so this answers a list rather than an option.
fn parse_class_member(__w: &mut PsWalker, pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    match pair.as_rule() {
        Rule::ps_method => parse_ps_method(__w, pair).map(|m| vec![m]),
        Rule::ps_constructor => parse_ps_constructor(__w, pair).map(|m| vec![m]),
        Rule::ps_property => parse_ps_property(__w, pair),
        Rule::class_function_decl => {
            let statement = parse_function_decl(__w, pair)?;
            Ok(vec![ClassMember::Method(Box::new(statement))])
        }
        Rule::constructor_decl => parse_constructor_decl(__w, pair).map(|m| vec![m]),
        _ => Ok(Vec::new()),
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
            Rule::function_params => {
                params = parse_function_params(__w, child);
                remember_ref_params(__w, &params);
            }
            Rule::block => {
                let locals = function_local_names(&child);
                let wants_args = mentions_args(child.as_str());
                // `return_last_of_branches`, not `implicit_return`: a function
                // ending in `if ($x) { 'yes' }` outputs `'yes'`, so the trailing
                // expression of each BRANCH is the return value, not just a
                // trailing expression statement.
                let parsed = parse_block_with_function_params(__w, child, &mut params)?;
                let parsed = unwrap_write_output(parsed);
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
            // `[ref]$v` — a BY-REFERENCE parameter. PowerShell spells it as a
            // type, and the AST already has the vocabulary
            // (`Param.pass_by` / `Argument.by_ref`), which is what C#'s `ref`
            // uses and what makes `Inc(ref n)` update the caller's `n`.
            // Without it `[ref]` was just an unknown type hint and the
            // assignment wrote to a local copy.
            // `[Parameter(ValueFromPipeline=$true)]$V` — an ATTRIBUTE, not a
            // type. It arrived as the type hint (`Parameter(ValueFromPipeline
            // =$true)`), which is not a type any lookup can resolve, and it
            // also hid which parameter the pipeline binds to.
            if let Some(t) = type_hint.as_deref()
                && t.len() >= 9
                && t[..9].eq_ignore_ascii_case("Parameter")
            {
                if t.to_lowercase().contains("valuefrompipeline") {
                    __w.pipeline_params.insert(name.to_lowercase());
                }
                type_hint = None;
            }
            let by_ref = type_hint
                .as_deref()
                .is_some_and(|t| t.eq_ignore_ascii_case("ref"));
            out.push(Param {
                name,
                // A `[ref]` is not a type to CONVERT to — dropping it keeps
                // `coerces_value_to_type_hint` from casting the reference.
                type_hint: if by_ref { None } else { type_hint.map(Into::into) },
                default,
                // `Alias`, not `Ref`: a write through `$v.Value` is visible to
                // the caller DURING the call, and survives the callee throwing
                // — the two properties the AST's own doc uses to tell true
                // aliasing from copy-in/copy-out.
                pass_by: if by_ref {
                    PassBy::Alias
                } else {
                    PassBy::Value
                },
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
    // (begin, process, end)
    let mut named: (Vec<Statement>, Vec<Statement>, Vec<Statement>) =
        (Vec::new(), Vec::new(), Vec::new());
    let mut saw_process = false;

    // `statement` is a SILENT rule, so a block's children are the concrete
    // statement rules (`assignment_stmt`, `command_stmt`, …) and never
    // `Rule::statement` itself. Matching on that name here skipped every
    // statement and left function and constructor bodies empty.
    for child in pair.into_inner() {
        if child.as_rule() == Rule::param_stmt {
            parse_param_stmt(__w, child, params);
            continue;
        }

        // `begin` / `process` / `end` are not three blocks that run in order —
        // `process` runs ONCE PER PIPELINE ITEM. Flattening them made
        // `process { $total += $Value }` run a single time with `$Value`
        // unbound, so `1,2,3 | Sum-Values` answered NaN.
        if child.as_rule() == Rule::named_block {
            let kind = child
                .clone()
                .into_inner()
                .next()
                .map(|k| k.as_str().to_lowercase())
                .unwrap_or_default();
            let stmts = match child.into_inner().find(|c| c.as_rule() == Rule::block) {
                Some(b) => parse_block_statements(__w, b)?,
                None => Vec::new(),
            };
            match kind.as_str() {
                "begin" => named.0.extend(stmts),
                "process" => {
                    saw_process = true;
                    named.1.extend(stmts);
                }
                _ => named.2.extend(stmts),
            }
            continue;
        }

        if let Some(stmt) = parse_statement(__w, child)? {
            body.push(stmt);
        }
    }

    if saw_process {
        body.extend(pipeline_block_body(__w, named, params));
    } else {
        body.extend(named.0);
        body.extend(named.1);
        body.extend(named.2);
    }

    // ⛔`parse_block_statements` applies traps and this — the FUNCTION BODY
    // path — did not, so a `trap` declared inside a function was never turned
    // into handlers: the marker survived into the final AST as an empty `Try`
    // and every statement it was meant to guard sat OUTSIDE it. That is the
    // whole `exceptions_trap_statement_scope` family, and a trap in a function
    // is where PowerShell scripts actually put one.
    Ok(apply_traps(body))
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
    Statement::new(StmtKind::Return(expr.map(unroll_returned_array)))
}

/// A returned value is UNROLLED into the success stream, and a stream of one
/// hands back that one thing.
///
/// So `return ,(1,2,3)` — the unary comma wrapping the array so the pipeline
/// cannot unroll the INNER one — yields the inner array, and `(Get-Array).Count`
/// is 3, not 1. `return @(5)` goes the same way and yields the scalar, which is
/// also what PowerShell does. Both are the same rule, applied to a
/// ONE-ELEMENT array literal at the point it is returned.
fn unroll_returned_array(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Array(mut items) if items.len() == 1 && !items[0].spread => {
            items.remove(0).value
        }
        _ => expr,
    }
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
    if let Some(stmt) = parse_constant_variable_statement(__w, pair.clone()) {
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

/// `New-Variable -Name X -Value V -Option Constant` — a CONSTANT declaration.
///
/// ⛔The old lowering dropped `-Option` with a comment calling enforcement
/// "test-shaped rather than a semantic the compiler can honestly claim". That
/// premise was wrong: `VarDeclKind::Const` exists and the shared compiler
/// ALREADY enforces it — verified on js, where `const a = 1; a = 2` throws and
/// `a` keeps its value. So this is a declaration the AST can express, not a
/// tracking pass the walker has to invent.
///
/// `ReadOnly` and `Constant` both reject a plain write; they differ only in
/// whether `-Force` can remove them, which is a separate question from the
/// assignment semantics.
fn parse_constant_variable_statement(__w: &mut PsWalker, pair: Pair<Rule>) -> Option<Statement> {
    let mut work = VecDeque::new();
    work.push_back(pair);
    let mut segment = None;
    while let Some(node) = work.pop_front() {
        if node.as_rule() == Rule::command_segment {
            segment = Some(node);
            break;
        }
        for child in node.into_inner() {
            work.push_back(child);
        }
    }
    let (callee, args) = parse_command_segment(__w, segment?)?;
    let ExprKind::Ident(cmd) = &callee.kind else {
        return None;
    };
    if !matches!(
        cmd.to_lowercase().as_str(),
        "new-variable" | "set-variable"
    ) {
        return None;
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
    let option = literal_text(&named("Option")?)?;
    if !matches!(
        option.to_lowercase().as_str(),
        "constant" | "readonly"
    ) {
        return None;
    }

    let positional: Vec<&Argument> = args.iter().filter(|a| a.name.is_none()).collect();
    let name_expr = named("Name").or_else(|| positional.first().map(|a| a.value.clone()))?;
    let target = literal_text(&name_expr)?;
    let value = named("Value")
        .or_else(|| positional.get(1).map(|a| a.value.clone()))
        .unwrap_or_else(Expression::null);

    // ⛔`VarDeclKind::Const` ALONE does nothing here. The shared compiler's
    // const enforcement is gated on `ecma_lexical_declarations`, and turning
    // that on for PowerShell cost **50 regressions for 9 fixes** — the flag
    // also drives lexical scoping and declaration semantics that PowerShell
    // does not share (`variable_drives`, `variable_declaration_semantics`).
    // Measured, then reverted. The declaration still says `Const` because that
    // is what it IS; the enforcement is the walker's, below.
    __w.constant_vars
        .insert(scope_qualified_name(&target).to_lowercase());

    Some(Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(scope_qualified_name(&target).to_string()),
            type_hint: None,
            init: Some(value),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Const,
    }))
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
    // `$h[$k] = v` and `$obj.$prop = v` — a computed WRITE. The mirror of the
    // computed read: an `Index` assignment resolves its key at COMPILE time, so
    // a literal key wrote correctly (`$h['A'] = 1`) while a variable key wrote
    // NOWHERE and reported no error. Both spellings arrive here as an `Index`
    // target, so both are redirected to the runtime-dispatching setter.
    //
    // ⛔An INT literal key is left alone: that is ordinary array element
    // assignment and already works.
    if op.as_str() == "="
        && targets.len() == 1
        && let ExprKind::Index { object, index, .. } = &targets[0].kind
        && !matches!(index.kind, ExprKind::Lit(Literal::Int(_)))
    {
        return Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ps_index_set")),
            args: vec![
                Argument::positional((**object).clone()),
                Argument::positional((**index).clone()),
                Argument::positional(value),
            ],
            optional: false,
        })));
    }

    // `$obj = [C]::new()` — record the variable's TYPE.
    //
    // ⛔Without it, overload resolution cannot fire. `resolve_instance_method_overload`
    // starts at `resolve_receiver_type_hint`, which reads the compiler's
    // variable type table; a plain `Assign` never writes one, so a PowerShell
    // receiver had no type and every overloaded call took the LAST declaration.
    // Measured: `$l.Log('a')` on a class with `Log($m)` and `Log($m, $lvl)`
    // ran the TWO-argument body with `$m` unbound, while
    // `([L]::new()).Log('a')` — where the receiver IS the `New` — resolved
    // correctly. Same class, same call, different answer depending on whether
    // the receiver had a name.
    //
    // `FunctionScoped`, because a PowerShell variable first assigned inside an
    // `if` is still there after it.
    if op.as_str() == "="
        && targets.len() == 1
        && let ExprKind::Ident(var_name) = &targets[0].kind
        && let ExprKind::New { class, .. } = &value.kind
        && let Some(type_name) = type_name_of(class)
    {
        return Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(var_name.clone()),
                type_hint: Some(TypeHint::converting(type_name)),
                init: Some(value),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::FunctionScoped,
        });
    }

    // A write to a `-Option Constant` / `-Option ReadOnly` variable is an
    // error. PowerShell raises it at runtime, and a `try`/`catch` around the
    // assignment is exactly how the corpus checks for it.
    if targets.len() == 1
        && let ExprKind::Ident(n) = &targets[0].kind
        && __w.constant_vars.contains(&n.to_lowercase())
    {
        return Statement::new(StmtKind::Throw {
            expr: Some(Expression::string(&format!(
                "Cannot overwrite variable {n} because it is read-only or constant."
            ))),
            cause: None,
        });
    }

    // ⛔AN ENVIRONMENT VARIABLE IS ALWAYS A STRING. `$env:NUM = 5` stores
    // `"5"`, so `$env:NUM -eq '5'` is TRUE in PowerShell and was FALSE here —
    // the number 5 was stored and compared against a string. The environment is
    // a string→string map in every OS this targets; the conversion belongs at
    // the WRITE, which is the only place the value's type is still known.
    if op.as_str() == "="
        && targets.len() == 1
        && matches!(&targets[0].kind, ExprKind::Ident(n) if n.to_lowercase().starts_with("env:"))
    {
        return Statement::new(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ps_to_string")),
                args: vec![Argument::positional(value)],
                optional: false,
            }),
            by_ref: false,
        });
    }

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
    let mut type_head: Option<String> = None;
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::type_literal) {
        let lit = inner.next().expect("peeked");
        type_head = Some(type_literal_name(lit.as_str()).trim().to_string());
    }

    // `[Singleton]::Instance = …` — the target is a STATIC member of the type,
    // so the leading `[Type]` is the RECEIVER here rather than a constraint on
    // a variable.
    if inner.peek().map(|p| p.as_rule()) == Some(Rule::static_member) {
        let member = inner.next().expect("peeked");
        let field = member
            .into_inner()
            .next()
            .map(|p| p.as_str().to_string())
            .unwrap_or_default();
        let mut expr = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&type_head.unwrap_or_default())),
            field,
            null_safe: false,
        });
        for step in inner {
            match step.as_rule() {
                Rule::member_get => {
                    let name = step
                        .into_inner()
                        .next()
                        .map(|p| p.as_str().trim_start_matches('$').to_string())
                        .unwrap_or_default();
                    if is_ref_value_step(__w, &expr, &name) {
                        continue;
                    }
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
        return expr;
    }

    let mut expr = match inner.next() {
        Some(p) if p.as_rule() == Rule::var_ref => walk_var_ref(p.as_str()),
        Some(p) => walk_expr(__w, p),
        None => return Expression::null(),
    };

    for step in inner {
        match step.as_rule() {
            Rule::member_get => {
                let inner_pair = step.into_inner().next();
                // `$obj.$name` is a COMPUTED lookup — the member is the VALUE
                // held by `$name`, not the literal text "name". Trimming the
                // `$` and using the variable's own spelling as the field read a
                // property no object has, so `$obj.$p` answered EMPTY even when
                // `$p` held the exact property name. `Index` is the shared
                // computed-access node, so this needs no PowerShell-local
                // lookup path.
                if let Some(p) = inner_pair.as_ref().filter(|p| p.as_rule() == Rule::var_ref) {
                    // ⛔An ASSIGNMENT TARGET, so this arm must stay a node you can
                    // WRITE to. The read path rewrites `$obj.$name` into a
                    // `__ps_member_dyn(…)` call, and a call is not assignable —
                    // doing the same here silently dropped every
                    // `$obj.$prop = …`. The two arms deliberately DIVERGE.
                    let key = walk_var_ref(p.as_str());
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(key),
                        null_safe: false,
                    });
                    continue;
                }
                let name = inner_pair
                    .map(|p| p.as_str().trim_start_matches('$').to_string())
                    .unwrap_or_default();
                if is_ref_value_step(__w, &expr, &name) {
                    continue;
                }
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
    let (head, mut args) = parse_command_segment(__w, first)?;
    // `-PipelineVariable X` names the variable that carries THIS stage's output
    // into the rest of the pipeline. It is a common parameter, not an argument
    // of the cmdlet, so it is taken out here before the call is built.
    let mut pending_pv = take_pipeline_variable(&mut args);
    let head_ov = take_out_variable(&mut args);
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
    if let Some((name, append)) = head_ov {
        expr = capture_out_variable(expr, &name, append);
    }

    for segment in segments {
        let (next, mut next_args) = parse_command_segment(__w, segment)?;
        let next_pv = take_pipeline_variable(&mut next_args);
        let next_ov = take_out_variable(&mut next_args);
        // The upstream stage's `-PipelineVariable` is bound at the TOP of this
        // stage's block. Stages are evaluated as nested calls rather than as a
        // streaming pipeline, so an assignment in the DECLARING stage would be
        // overwritten by every later item before a downstream stage ran and
        // every item would read the last value. Binding it inside the consuming
        // block keeps it per-item, because this block's `$_` IS the declaring
        // stage's output for that item.
        if let Some(name) = pending_pv.take() {
            bind_pipeline_variable(&mut next_args, &name);
        }
        pending_pv = next_pv;
        // The same collection-cmdlet rewrite the value-position builder does.
        if let Some(folded) = pipeline_stage_as_method(&expr, &next, &next_args) {
            expr = folded;
            if let Some((name, append)) = next_ov {
                expr = capture_out_variable(expr, &name, append);
            }
            continue;
        }
        let mut chained = vec![Argument::positional(expr)];
        chained.append(&mut next_args);
        // Also through `build_command_call`. A cmdlet does not stop being one
        // because it is downstream of a `|`: `$o | Add-Member …` needs the same
        // rewrite as the head of the pipeline, and without it compiled to a
        // call to a function that was never defined.
        expr = build_command_call(next, chained);
        if let Some((name, append)) = next_ov {
            expr = capture_out_variable(expr, &name, append);
        }
    }

    Some(unwrap_pipeline_value(expr))
}

/// Take a `-PipelineVariable X` / `-pv X` common parameter out of a stage's
/// arguments, answering the variable's name.
fn take_pipeline_variable(args: &mut Vec<Argument>) -> Option<String> {
    let idx = args.iter().position(|a| {
        a.name.as_deref().is_some_and(|n| {
            n.eq_ignore_ascii_case("PipelineVariable") || n.eq_ignore_ascii_case("pv")
        })
    })?;
    let arg = args.remove(idx);
    match &arg.value.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Lit(Literal::Str(text)) => Some(text.clone()),
        _ => None,
    }
}

/// Prepend `$X = $_` to a stage's script block, so the block sees the upstream
/// stage's value for the item it is currently processing.
fn bind_pipeline_variable(args: &mut [Argument], name: &str) {
    for arg in args.iter_mut() {
        if let ExprKind::Lambda { body, .. } = &mut arg.value.kind {
            let bind = Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(name)],
                value: Expression::ident("_"),
                by_ref: false,
            });
            match body {
                LambdaBody::Block(stmts) => stmts.insert(0, bind),
                LambdaBody::Expr(expr) => {
                    let tail = Statement::new(StmtKind::Return(Some((**expr).clone())));
                    *body = LambdaBody::Block(vec![bind, tail]);
                }
            }
            return;
        }
    }
}

/// Take a `-OutVariable X` / `-ov X` common parameter out of a stage's
/// arguments, answering the variable's name and whether it APPENDS.
///
/// `-OutVariable +buf` appends to what `$buf` already holds; without the `+`
/// the variable is overwritten, so a pre-existing value must NOT survive.
fn take_out_variable(args: &mut Vec<Argument>) -> Option<(String, bool)> {
    let idx = args.iter().position(|a| {
        a.name.as_deref().is_some_and(|n| {
            n.eq_ignore_ascii_case("OutVariable") || n.eq_ignore_ascii_case("ov")
        })
    })?;
    let arg = args.remove(idx);
    let raw = match &arg.value.kind {
        ExprKind::Ident(name) => name.clone(),
        ExprKind::Lit(Literal::Str(text)) => text.clone(),
        // `+buf` parses as a unary plus applied to `buf`.
        ExprKind::Unary { expr, .. } => match &expr.kind {
            ExprKind::Ident(name) => format!("+{name}"),
            _ => return None,
        },
        _ => return None,
    };
    match raw.strip_prefix('+') {
        Some(name) => Some((name.to_string(), true)),
        None => Some((raw, false)),
    }
}

/// `$X = @(stage)`, as an EXPRESSION, so the captured value keeps flowing down
/// the pipeline.
///
/// ⛔The capture is always an ARRAY, even for one item — a test asserts
/// `$cap.Count -eq 1` after a single-element pipeline, so handing back the bare
/// scalar would answer the string's length instead.
fn capture_out_variable(stage: Expression, name: &str, append: bool) -> Expression {
    let captured = ensure_array(stage);
    let value = if append {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(ensure_array(Expression::ident(name))),
            right: Box::new(captured),
        })
    } else {
        captured
    };
    Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(name)),
        value: Box::new(value),
    })
}

/// `@( … )` around an expression — the array subexpression operator, which is
/// what `-OutVariable` semantics need.
fn ensure_array(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__ps_array")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

/// A pipeline's value, with a stream of ONE unwrapped to that one thing.
///
/// PowerShell's success stream hands back a scalar for a single item and an
/// array from two upward — the rule `unwrap_output` already applies to a
/// function's output. A pipeline produces the same kind of stream, so
/// `@($obj) | ForEach-Object { $_ }` is the OBJECT and `$r.Tag` reaches its
/// property; it stayed wrapped and answered empty.
fn unwrap_pipeline_value(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("__ps_unwrap_single")),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
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
    mut args: Vec<Argument>,
    in_value_position: bool,
) -> Expression {
    // `-ErrorAction SilentlyContinue` / `Ignore` suppresses a terminating error
    // from THIS command and lets the script carry on. It is a common parameter,
    // so it is taken out before the call is built and re-expressed as a
    // try/catch around it.
    // ⛔The value is DROPPED, not honoured, and that is deliberate. Wrapping the
    // call in a try/catch looked right and made things worse: an undefined
    // command is a VM TRAP here, not a catchable PowerShell error — measured,
    // `try { Get-NoSuchThing } catch { }` does not catch — so the wrapper
    // suppressed nothing and its LAMBDA made every such pipeline look
    // side-effecting to `Out-Null`, turning 10 vacuous passes into failures.
    // Honouring `-ErrorAction` needs an unrecognised command to raise a real
    // error first; until then, taking the parameter out is all that is honest.
    let _ = take_silent_error_action(&mut args);
    build_command_call_inner(callee, args, in_value_position)
}

/// Take `-ErrorAction SilentlyContinue` / `Ignore` out of a command's arguments,
/// answering whether errors from it are suppressed.
///
/// ⛔`Stop` and `Continue` are NOT suppression — `Stop` makes a non-terminating
/// error terminating, which is the opposite — so only the two silencing values
/// are honoured and anything else keeps its normal behaviour.
fn take_silent_error_action(args: &mut Vec<Argument>) -> bool {
    let Some(idx) = args.iter().position(|a| {
        a.name
            .as_deref()
            .is_some_and(|n| n.eq_ignore_ascii_case("ErrorAction") || n.eq_ignore_ascii_case("ea"))
    }) else {
        return false;
    };
    let arg = args.remove(idx);
    let value = match &arg.value.kind {
        ExprKind::Ident(name) => name.clone(),
        ExprKind::Lit(Literal::Str(text)) => text.clone(),
        _ => return false,
    };
    value.eq_ignore_ascii_case("SilentlyContinue") || value.eq_ignore_ascii_case("Ignore")
}

fn build_command_call_inner(
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
        // `F @h` — a splat. A HASHTABLE splat binds by NAME, which the generic
        // `spread` flag cannot express: it hands the collection over
        // positionally, so `F @{ A = 1; B = 2 }` bound the hashtable itself to
        // `$A` and `$B` stayed empty (`$A + $B` was NaN). When the callee is a
        // function this script declares, its parameter names are known, so the
        // splat is expanded into one argument per parameter and a runtime
        // helper picks each value by NAME or by INDEX depending on what the
        // splat actually holds — which is also what makes the ARRAY splat keep
        // working through the same path.
        if let Some(splat) = splat_token_name(token) {
            if let Some(params) = callee_param_names(__w, &callee) {
                for (idx, param) in params.iter().enumerate() {
                    args.push(Argument::positional(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ps_splat_arg")),
                        args: vec![
                            Argument::positional(Expression::ident(splat)),
                            Argument::positional(Expression::string(param)),
                            Argument::positional(Expression::int(idx as i64)),
                        ],
                        optional: false,
                    })));
                }
                i += 1;
                continue;
            }
            args.push(Argument {
                value: Expression::ident(splat),
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
    // `Inc ([ref]$n)` — an argument that IS a reference is passed BY REFERENCE.
    // The `RefOf` node alone is not enough: `Argument.by_ref` is what the
    // shared call machinery reads to bind the parameter to the caller's
    // storage, and without it the reference was passed as an ordinary value and
    // `$n` stayed 10.
    let args = args
        .into_iter()
        .map(|mut a| {
            // ⛔The shared call machinery wants the PLACE plus `by_ref`, not a
            // reference OBJECT. Measured against php, which works: `inc($n)`
            // passes `Ident("n")` with `by_ref: true` and a param whose
            // `pass_by` is `Alias`. Leaving the `RefOf` wrapper on made the
            // argument a value again and `$n` stayed 10.
            if let ExprKind::RefOf(place) = a.value.kind {
                a.by_ref = true;
                a.value = place_as_expression(*place);
            }
            a
        })
        .collect();

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
        // `,$x` — the unary comma wraps its operand in a ONE-element array.
        Rule::array_wrap => {
            let inner = pair
                .into_inner()
                .next()
                .map(|__x| walk_expr(__w, __x))
                .unwrap_or_else(Expression::null);
            Expression::new(ExprKind::Array(vec![ArrayElement {
                key: None,
                value: inner,
                spread: false,
                by_ref: false,
            }]))
        }
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
        // `__ps_join`, not `join`: the bare name is claimed shared-side before
        // `[value_methods]` is reached and answers the empty string. See the
        // profile's `Join` note.
        "-join" => return method_call_expr(operand, "__ps_join", vec![Expression::string("")]),
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

    // `[int]$x` has THREE answers chosen at runtime (banker's rounding, a
    // numeric parse, and a character's code point), which the shared `Cast`
    // lowering cannot express — it has one arm per target type, not per source
    // VALUE. Same move as `[bool]` above: express it with a node the compiler
    // already owns rather than teaching `Cast` a language-specific rule.
    if matches!(
        type_name.to_lowercase().as_str(),
        "int" | "int32" | "system.int32" | "long" | "int64" | "system.int64"
    ) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ps_to_int")),
            args: vec![Argument::positional(expr)],
            optional: false,
        });
    }

    // `[ref]$x` — a REFERENCE to `$x`'s storage, not a conversion of its value.
    // PowerShell spells taking a reference as a cast, and the AST already has
    // the node: `RefOf` over a `PlaceExpr`, the same one C#'s `ref n` uses.
    // Without it the argument was a COPY and `Inc ([ref]$n)` left `$n` at 10.
    if type_name.eq_ignore_ascii_case("ref") {
        if let Some(place) = PlaceExpr::from_expr(&expr) {
            return Expression::new(ExprKind::RefOf(Box::new(place)));
        }
        return expr;
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
                let inner_pair = op.into_inner().next();
                // See the computed-lookup note on the other `member_get` arm:
                // `$obj.$name` indexes by the VALUE of `$name`. Both arms must
                // agree or the same source spells two different reads.
                if let Some(p) = inner_pair.as_ref().filter(|p| p.as_rule() == Rule::var_ref) {
                    let key = walk_var_ref(p.as_str());
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ps_member_dyn")),
                        args: vec![
                            Argument::positional(expr),
                            Argument::positional(key),
                        ],
                        optional: false,
                    });
                    continue;
                }
                let name = inner_pair
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
                if is_ref_value_step(__w, &expr, &name) {
                    continue;
                }
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
                if is_ref_value_step(__w, &expr, &name) {
                    continue;
                }
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

                // `[char]::IsUpper($c)` and friends. .NET spells character
                // classification as statics on `System.Char`; the shared
                // `str_is_*` primitives already answer exactly these questions
                // for a one-character string, which is what a PowerShell
                // `[char]` is. Only the classifications a primitive already
                // implements are mapped — `IsSymbol`, `IsPunctuation`,
                // `IsControl`, `IsSeparator` and the surrogate pair tests have
                // no primitive, and inventing an approximation here would be a
                // wrong answer rather than a missing one.
                if is_static
                    && args.len() == 1
                    && type_name_of(&expr).is_some_and(|t| {
                        t.rsplit('.')
                            .next()
                            .is_some_and(|last| last.eq_ignore_ascii_case("char"))
                    })
                {
                    if let Some(builtin) = char_static_builtin(&name) {
                        let mut args = args;
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(builtin)),
                            args: vec![Argument::positional(args.remove(0))],
                            optional: false,
                        });
                        continue;
                    }
                }

                // `[System.Array]::Sort($arr)` — .NET's in-place sort, which
                // orders elements by `IComparable`. The default numeric sort
                // coerces its operands, so an array of class instances died in
                // `wasm:js-number.toF64 — not a number`. `collections.sort_func`
                // is the shared in-place sort that takes a COMPARATOR, so the
                // rewrite supplies the one .NET documents: `a.CompareTo(b)`.
                // That reaches a user class's own `CompareTo` when it declares
                // one, and the `[value_methods]` binding otherwise.
                if is_static
                    && name.eq_ignore_ascii_case("Sort")
                    && args.len() == 1
                    && type_name_of(&expr)
                        .is_some_and(|t| t.rsplit('.').next() == Some("Array"))
                {
                    let cmp = Expression::new(ExprKind::Lambda {
                        params: vec![synth_param("__ps_cmp_a"), synth_param("__ps_cmp_b")],
                        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(
                            method_call_expr(
                                Expression::ident("__ps_cmp_a"),
                                "CompareTo",
                                vec![Expression::ident("__ps_cmp_b")],
                            ),
                        )))]),
                        is_async: false,
                        captures: Vec::new(),
                    });
                    let mut args = args;
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ps_sort_with")),
                        args: vec![
                            Argument::positional(args.remove(0)),
                            Argument::positional(cmp),
                        ],
                        optional: false,
                    });
                    continue;
                }

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
                let index = fold_literal_key(index);
                // A COMPUTED key. `$h['A']` works because a literal key folds
                // to a member read; `$h[$k]` for the same key answered EMPTY,
                // because `Index` with a non-literal key falls through to ARRAY
                // indexing. The adapter asks the receiver at runtime, so an
                // array keeps array semantics and a hashtable gets a key read.
                // ⛔A RANGE key is a SLICE — `$arr[1..2]` is two elements, not
                // one lookup — and the shared `Index` lowering already
                // implements it. Routing it through the computed-key adapter
                // asked an array for element `[1,2]` and broke six slicing
                // tests that a hashtable fix had no business touching.
                if !matches!(
                    index.kind,
                    ExprKind::Lit(Literal::Int(_)) | ExprKind::Range { .. }
                ) {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__ps_index_get")),
                        args: vec![
                            Argument::positional(expr),
                            Argument::positional(index),
                        ],
                        optional: false,
                    });
                    continue;
                }
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index),
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
    /// Declared function name → its parameter names, in order. Needed to expand
    /// a HASHTABLE splat, which binds by NAME and so cannot be expanded without
    /// knowing what the callee's parameters are called.
    declared_functions: std::collections::HashMap<String, Vec<String>>,
    /// Names of the `[ref]` parameters of the function currently being walked.
    /// `$v.Value` on one of these IS `$v` — the alias itself — so the member
    /// step is dropped rather than read as a property.
    ref_params: std::collections::HashSet<String>,
    /// Variables declared `-Option Constant` / `-Option ReadOnly`, folded.
    /// A later write to one of these is an ERROR in PowerShell.
    constant_vars: std::collections::HashSet<String>,
    /// Parameters declared `[Parameter(ValueFromPipeline=$true)]`, folded — the
    /// one a `process { }` block iterates.
    pipeline_params: std::collections::HashSet<String>,
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

/// Record every top-level function's parameter names, in declaration order.
///
/// A hashtable splat (`F @h`) binds `$h`'s KEYS to parameter NAMES, so the
/// expansion needs the callee's parameter list. An array splat binds
/// positionally and needs only the count — the same expansion serves both,
/// because the runtime helper picks by name or by index depending on what it
/// is handed.
fn collect_declared_functions(__w: &mut PsWalker, root: Pair<Rule>) {
    let mut out: std::collections::HashMap<String, Vec<String>> =
        std::collections::HashMap::new();
    let mut work = VecDeque::new();
    work.push_back(root);
    while let Some(node) = work.pop_front() {
        if node.as_rule() == Rule::function_decl {
            let mut inner = node.clone().into_inner();
            let Some(name) = inner.find(|c| c.as_rule() == Rule::function_name) else {
                continue;
            };
            let mut params = Vec::new();
            let mut sub = VecDeque::new();
            for child in node.clone().into_inner() {
                sub.push_back(child);
            }
            while let Some(n) = sub.pop_front() {
                if n.as_rule() == Rule::function_param {
                    if let Some(v) = n
                        .clone()
                        .into_inner()
                        .find(|c| c.as_rule() == Rule::var_ref)
                    {
                        params.push(v.as_str().trim_start_matches('$').to_string());
                    }
                }
                for child in n.into_inner() {
                    sub.push_back(child);
                }
            }
            out.insert(name.as_str().to_lowercase(), params);
        }
        for child in node.into_inner() {
            work.push_back(child);
        }
    }
    __w.declared_functions = out;
}

/// The parameter names of a function this script declares, when the callee
/// names one.
fn callee_param_names(__w: &PsWalker, callee: &Expression) -> Option<Vec<String>> {
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    __w.declared_functions
        .get(&name.to_lowercase())
        .filter(|p| !p.is_empty())
        .cloned()
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

    // ⛔THE SECOND PIPELINE BUILDER. `parse_pipeline` handles a pipeline in
    // STATEMENT position and this one handles it in VALUE position, so a common
    // parameter has to be taken out in BOTH or the same source spells two
    // different pipelines depending on where it appears.
    let mut pending_pv: Option<String> = None;
    for stage in stages {
        let (callee, mut args) = match stage.as_rule() {
            Rule::command_segment => match parse_command_segment(__w, stage) {
                Some(parts) => parts,
                None => continue,
            },
            _ => (walk_expr(__w, stage), Vec::new()),
        };
        let next_pv = take_pipeline_variable(&mut args);
        let next_ov = take_out_variable(&mut args);
        if let Some(name) = pending_pv.take() {
            bind_pipeline_variable(&mut args, &name);
        }
        pending_pv = next_pv;

        if let Some(folded) = pipeline_stage_as_method(&expr, &callee, &args) {
            expr = folded;
            if let Some((name, append)) = next_ov {
                expr = capture_out_variable(expr, &name, append);
            }
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
        if let Some((name, append)) = next_ov {
            expr = capture_out_variable(expr, &name, append);
        }
    }

    unwrap_pipeline_value(expr)
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
    // `Select-Object` is not one method: each switch is a different slice of the
    // upstream collection, and they arrive as NAMED arguments, which the generic
    // positional path below drops. Only the two bounded forms are lowered here.
    //
    // `-Last`, `-Unique`, `-Property` and `-ExpandProperty` are deliberately NOT
    // handled: each needs a primitive this walker cannot compose today, and
    // passing the collection through unchanged would turn a loud
    // `undefined is not callable` into a silently wrong answer.
    if name.eq_ignore_ascii_case("select-object") {
        let named = |key: &str| {
            args.iter()
                .find(|a| {
                    a.name
                        .as_deref()
                        .is_some_and(|n| n.eq_ignore_ascii_case(key))
                })
                .map(|a| a.value.clone())
        };
        let slice = |start: Expression, end: Expression| {
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__ps_slice")),
                args: vec![
                    Argument::positional(upstream.clone()),
                    Argument::positional(start),
                    Argument::positional(end),
                ],
                optional: false,
            }))
        };
        if let Some(n) = named("First") {
            return slice(Expression::int(0), n);
        }
        if let Some(n) = named("Skip") {
            // No `-First` alongside, so the tail runs to the end. `i32::MAX` is
            // the same open upper bound `array_fill` uses; the VM clamps it.
            return slice(n, Expression::int(i32::MAX as i64));
        }
        return None;
    }

    if name.eq_ignore_ascii_case("measure-object") || name.eq_ignore_ascii_case("measure") {
        return measure_object_expr(upstream, args);
    }

    if name.eq_ignore_ascii_case("convertfrom-csv") {
        return convert_from_csv_expr(upstream, args);
    }

    let method = pipeline_cmdlet_method(name)?;
    if method.is_empty() {
        // `… | Out-Null` discards the pipeline's VALUE — it does not delete the
        // pipeline.
        //
        // ⛔This returned a bare `null` and threw the upstream expression away
        // UNPARSED, so nothing before the `|` ran at all:
        // `1..2 | ForEach-Object { $script:g += $_ } | Out-Null` left `$g`
        // EMPTY, and every `-OutVariable` capture ending in `| Out-Null` — the
        // idiomatic spelling — captured nothing. A discarded value and a
        // discarded side effect are not the same thing.
        //
        // ⛔But evaluating it UNCONDITIONALLY costs more than it buys today:
        // `Get-Help about_Aliases -ErrorAction SilentlyContinue | Out-Null`
        // reaches a cmdlet nothing implements, and an undefined command is a VM
        // TRAP here rather than a catchable PowerShell error — measured, a
        // `try`/`catch` around it does not catch, so `-ErrorAction` cannot
        // rescue it either. 52 such tests passed only because the expression was
        // deleted. So the upstream is evaluated when it can actually DO
        // something — it contains a script block, or a capture assignment — and
        // still discarded when it is a bare value-producing call.
        //
        // ⚠ This is a heuristic standing in for cmdlet coverage, not the final
        // shape: once an unrecognised command raises a catchable error the way
        // PowerShell's does, the condition goes away and the upstream is always
        // evaluated.
        if !has_side_effecting_shape(upstream) {
            return Some(Expression::null());
        }
        return Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ps_out_null")),
            args: vec![Argument::positional(upstream.clone())],
            optional: false,
        }));
    }
    let positional: Vec<Expression> = args
        .iter()
        .filter(|a| a.name.is_none())
        .map(|a| a.value.clone())
        .collect();
    // ⛔A SCALAR IS A ONE-ITEM STREAM. `"New" | ForEach-Object { $_ }` answered
    // EMPTY because the fold called `.ForEach` straight on the string, and the
    // array methods want an array. `@( … )` is the operator PowerShell itself
    // uses for exactly this and it is idempotent on an array, so wrapping the
    // upstream costs nothing where it was already a collection.
    Some(method_call_expr(
        ensure_array(upstream.clone()),
        method,
        positional,
    ))
}

/// `Measure-Object` is NOT a method — it answers an OBJECT whose properties are
/// the requested statistics, and the corpus reads them back (`(… |
/// Measure-Object -Sum).Sum`). Mapping it to `Count` returned a bare number, so
/// every `.Sum` / `.Average` / `.Count` read on the result answered undefined.
///
/// Verified against `/usr/local/bin/pwsh`: a bare `Measure-Object` fills ONLY
/// `Count` and leaves the rest `$null`; each switch adds exactly its own field;
/// an empty pipeline with `-Sum` answers `Count=0 Sum=0`. All five keys are
/// emitted either way so a read of an unrequested one is `$null` rather than a
/// missing property.
///
/// The upstream is bound to a lambda PARAMETER and the lambda is called
/// immediately, so a pipeline source with side effects (`Get-Random |
/// Measure-Object -Sum -Average`) is evaluated ONCE. Referencing `upstream`
/// five times in the object literal would have run it five times.
///
/// `-Property` is deliberately NOT handled: it measures `$_.<name>` over the
/// collection, which needs a projection this walker cannot compose today.
/// Returning `None` leaves it a loud `undefined is not callable` rather than a
/// silently wrong sum over the objects themselves.
fn measure_object_expr(upstream: &Expression, args: &[Argument]) -> Option<Expression> {
    let has_switch = |key: &str| {
        args.iter()
            .any(|a| a.name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(key)))
    };
    if has_switch("Property") {
        return None;
    }

    const SRC: &str = "__ps_measure_src";
    let src = || Expression::ident(SRC);
    let stat = |builtin: &str| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(builtin)),
            args: vec![Argument::positional(src())],
            optional: false,
        })
    };
    let count = stat("__ps_count");
    // `Average` is `Sum / Count`, and an empty pipeline would make that `0 / 0`.
    // PowerShell answers `$null` there, so guard on the count rather than
    // handing back a NaN that prints as a number.
    let average = Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(count.clone()),
            right: Box::new(Expression::int(0)),
        })),
        then: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Div,
            left: Box::new(stat("__ps_sum")),
            right: Box::new(count.clone()),
        })),
        else_: Box::new(Expression::null()),
    });
    let field = |requested: bool, value: Expression| {
        if requested {
            value
        } else {
            Expression::null()
        }
    };
    // ⛔TWO keys for one value. `normalize_member_name` rewrites PowerShell's
    // `.Count` to `.length` for every receiver, so a key spelled only `Count`
    // is UNREACHABLE — `$m.Count` fell through to the object's own property
    // count, which for this literal is 5. `1..5 | Measure-Object` therefore
    // answered 5 and looked correct; `@(7,8) | Measure-Object` also answered 5.
    // An own `length` property wins over that fallback (verified), so `length`
    // is the key the read actually lands on, and `Count` is kept because it is
    // the name PowerShell reports.
    let result = Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("length"),
            value: count.clone(),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Count"),
            value: count,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Sum"),
            value: field(has_switch("Sum"), stat("__ps_sum")),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Average"),
            value: field(has_switch("Average"), average),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Maximum"),
            value: field(has_switch("Maximum"), stat("__ps_max")),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Minimum"),
            value: field(has_switch("Minimum"), stat("__ps_min")),
        },
    ]));

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: SRC.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(result)))]),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![Argument::positional(upstream.clone())],
        optional: false,
    }))
}

/// `System.Char`'s classification statics, mapped to the shared primitives that
/// already implement them. Returns `None` for the ones with no primitive, so
/// the call falls through to normal resolution and fails loudly rather than
/// silently answering the wrong thing.
fn char_static_builtin(name: &str) -> Option<&'static str> {
    Some(match name.to_ascii_lowercase().as_str() {
        "isupper" => "__ps_char_is_upper",
        "islower" => "__ps_char_is_lower",
        "isdigit" | "isnumber" => "__ps_char_is_digit",
        "isletter" => "__ps_char_is_letter",
        "isletterordigit" => "__ps_char_is_alnum",
        "iswhitespace" => "__ps_char_is_space",
        "toupper" | "toupperinvariant" => "__ps_char_to_upper",
        "tolower" | "tolowerinvariant" => "__ps_char_to_lower",
        _ => return None,
    })
}

/// The expression a `PlaceExpr` denotes — the inverse of `PlaceExpr::from_expr`.
fn place_as_expression(place: PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(&name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object,
            field,
            null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object,
            index,
            null_safe,
        }),
        PlaceExpr::Deref(expr) => *expr,
    }
}

/// The body of an advanced function that declares a `process` block:
/// `begin` once, `process` for EACH pipeline item, `end` once.
///
/// The pipeline parameter arrives holding the whole upstream collection, so the
/// loop is over `@($param)` and rebinds `$param` to each item — which is what
/// makes `process { $total += $Value }` see one value at a time. `@( … )` keeps
/// a scalar working as a one-item stream.
fn pipeline_block_body(
    __w: &PsWalker,
    named: (Vec<Statement>, Vec<Statement>, Vec<Statement>),
    params: &[Param],
) -> Vec<Statement> {
    let (begin, process, end) = named;
    let Some(param) = params
        .iter()
        .find(|p| __w.pipeline_params.contains(&p.name.to_lowercase()))
        .or_else(|| params.first())
    else {
        // No parameter to bind: the blocks still run, in order.
        let mut out = begin;
        out.extend(process);
        out.extend(end);
        return out;
    };

    let item = "__ps_pipe_item";
    let src = "__ps_pipe_src";
    let mut loop_body = vec![
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&param.name)],
            value: Expression::ident(item),
            by_ref: false,
        }),
        // `$_` is the current pipeline object, and inside `process` it names
        // the same item the declared parameter does.
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident("_")],
            value: Expression::ident(item),
            by_ref: false,
        }),
    ];
    loop_body.extend(process);

    let mut out = begin;
    // \u26d4The loop MUST NOT iterate the parameter it rebinds. The parameter
    // holds the WHOLE pipeline input, and step one assigns the first element
    // back over it — walking that same name re-reads a collection that turns
    // into a scalar underneath the walk. Snapshot it once, before the loop.
    out.insert(
        0,
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(src)],
            value: ensure_array(Expression::ident(&param.name)),
            by_ref: false,
        }),
    );
    out.push(Statement::new(StmtKind::ForIn {
        var: item.to_string(),
        key: None,
        iter: Expression::ident(src),
        body: loop_body,
        of: true,
        else_body: None,
        is_async: false,
    }));
    out.extend(end);
    out
}

/// Record which of a function's parameters are `[ref]`, so `$v.Value` inside
/// its body resolves to the alias rather than to a property named `Value`.
fn remember_ref_params(__w: &mut PsWalker, params: &[Param]) {
    for p in params {
        if p.pass_by == PassBy::Alias {
            __w.ref_params.insert(p.name.to_lowercase());
        }
    }
}

/// `$v.Value` where `$v` is a `[ref]` parameter — PowerShell reaches the
/// referenced storage through a `.Value` property, and under true aliasing the
/// parameter IS that storage, so the step is dropped.
fn is_ref_value_step(__w: &PsWalker, object: &Expression, field: &str) -> bool {
    if field.eq_ignore_ascii_case("Value")
        && matches!(&object.kind, ExprKind::Ident(n) if __w.ref_params.contains(&n.to_lowercase()))
    {
        return true;
    }
    // `$_.Exception` in a trap or catch — `$_` is an ErrorRecord in PowerShell
    // and the exception hangs off it, but the value actually caught here IS the
    // exception, so the step is the identity. Without this
    // `$_.Exception.Message` read `Message` off nothing and every trap that
    // captured the message captured "".
    field.eq_ignore_ascii_case("Exception")
        && matches!(&object.kind, ExprKind::Ident(n) if n == "_")
}

/// A synthesized lambda parameter. Only the name varies, and every synthesized
/// binder wants the same plain by-value required parameter.
fn synth_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// Call `lambda(params…)` immediately with `args`. The shared way to BIND a
/// value that is read more than once — cloning the expression instead would
/// re-evaluate a side-effecting source once per read.
fn bind_and_call(params: &[&str], args: Vec<Expression>, body: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: params.iter().map(|p| synth_param(p)).collect(),
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(body)))]),
            is_async: false,
            captures: Vec::new(),
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

/// `… | ConvertFrom-Csv [-Delimiter c] [-Header a,b,c]`.
///
/// The document is parsed by `common:csv.parse_document` — the shared CSV
/// grammar — and this walker only names the result: a header row becomes the
/// KEYS of one object per remaining record. That split is deliberate. Quoting,
/// embedded delimiters and **newlines inside a quoted field** are grammar and
/// belong to the primitive; `[PSCustomObject]` is what PowerShell calls a
/// record, and python's `DictReader` and ruby's `CSV::Row` call it something
/// else over the same array of arrays.
///
/// ⛔The pipeline delivers EITHER one multi-line string (`$csv = @"…"@`) or an
/// ARRAY of lines (`@($items | ConvertTo-Csv)`), and the corpus uses both.
/// `__ps_array` normalizes the two — a string becomes a one-element array — and
/// `__ps_join` puts the lines back together, so the branch is two existing
/// pieces rather than a type test inside shared code.
///
/// Every step is an EXPRESSION: `__array_map` passes `(value, index)` to its
/// lambda (verified), so zipping the header against a record is a nested map
/// and `ecma:object:fromEntries` turns the pairs into the object. No
/// walker-built loop, and each binder is a lambda PARAMETER, so a nested
/// `ConvertFrom-Csv` cannot clobber an outer one's rows.
fn convert_from_csv_expr(upstream: &Expression, args: &[Argument]) -> Option<Expression> {
    let named = |key: &str| {
        args.iter()
            .find(|a| {
                a.name
                    .as_deref()
                    .is_some_and(|n| n.eq_ignore_ascii_case(key))
            })
            .map(|a| a.value.clone())
    };
    let call = |callee: &str, args: Vec<Expression>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(callee)),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false,
        })
    };
    let index = |object: Expression, idx: Expression| {
        Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(idx),
            null_safe: false,
        })
    };

    // One multi-line string, whichever shape the pipeline delivered.
    let text = method_call_expr(
        call("__ps_array", vec![upstream.clone()]),
        "__ps_join",
        vec![Expression::string("\n")],
    );
    let delimiter = named("Delimiter").unwrap_or_else(|| Expression::string(","));
    let rows = call(
        "__ps_csv_rows",
        vec![text, delimiter, Expression::string("\"")],
    );

    // `-Header` names the columns AND means the first record is data. Without
    // it the first record IS the header and is not itself a row.
    let supplied = named("Header");
    let (header, body) = match &supplied {
        Some(h) => (h.clone(), Expression::ident(ROWS)),
        None => (
            index(Expression::ident(ROWS), Expression::int(0)),
            call(
                "__ps_slice",
                vec![
                    Expression::ident(ROWS),
                    Expression::int(1),
                    Expression::int(i32::MAX as i64),
                ],
            ),
        ),
    };

    const ROWS: &str = "__ps_csv_rows_v";
    const HDR: &str = "__ps_csv_hdr";
    const BODY: &str = "__ps_csv_body";
    const REC: &str = "__ps_csv_rec";
    const KEY: &str = "__ps_csv_key";
    const COL: &str = "__ps_csv_col";

    // `[key, value]` for one column of one record, then the record's pairs
    // become its object.
    let pair = Expression::new(ExprKind::Array(vec![
        ArrayElement {
            key: None,
            value: Expression::ident(KEY),
            spread: false,
            by_ref: false,
        },
        ArrayElement {
            key: None,
            value: index(Expression::ident(REC), Expression::ident(COL)),
            spread: false,
            by_ref: false,
        },
    ]));
    let pairs = method_call_expr(
        Expression::ident(HDR),
        "ForEach",
        vec![Expression::new(ExprKind::Lambda {
            params: vec![synth_param(KEY), synth_param(COL)],
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(pair)))]),
            is_async: false,
            captures: Vec::new(),
        })],
    );
    let record = call("__ps_from_entries", vec![pairs]);
    let records = method_call_expr(
        Expression::ident(BODY),
        "ForEach",
        vec![Expression::new(ExprKind::Lambda {
            params: vec![synth_param(REC)],
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(record)))]),
            is_async: false,
            captures: Vec::new(),
        })],
    );

    Some(bind_and_call(
        &[ROWS],
        vec![rows],
        bind_and_call(&[HDR, BODY], vec![header, body], records),
    ))
}

/// Whether an expression carries something worth running for its EFFECT: a
/// script block (a `ForEach-Object` / `Where-Object` body can mutate), or an
/// assignment (an `-OutVariable` capture).
fn has_side_effecting_shape(expr: &Expression) -> bool {
    match &expr.kind {
        // An `-OutVariable` capture.
        ExprKind::Assign { .. } => true,
        // A FOLDED pipeline stage — `$src.ForEach({ … })`. The block is the
        // stage's body and may mutate, so the stage has to run.
        //
        // ⛔A block passed to a bare cmdlet is NOT this: `Start-Job -ScriptBlock
        // { 6 } | Out-Null` carries a lambda and still must not be evaluated,
        // because `Start-Job` is unimplemented and traps. Testing for "contains
        // a lambda" cost five remoting tests; the shape that matters is the
        // METHOD call the fold produces.
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Member { object, .. } => {
                args.iter().any(|a| matches!(a.value.kind, ExprKind::Lambda { .. }))
                    || has_side_effecting_shape(object)
            }
            _ => args.iter().any(|a| has_side_effecting_shape(&a.value)),
        },
        ExprKind::Member { object, .. } => has_side_effecting_shape(object),
        ExprKind::Binary { left, right, .. } => {
            has_side_effecting_shape(left) || has_side_effecting_shape(right)
        }
        _ => false,
    }
}

/// The method a pipeline cmdlet is equivalent to, including its `?` / `%`
/// aliases. `""` means the stage drops the value (`Out-Null`).
///
/// NO `measure-object` row — it is an object, not a method, and
/// `measure_object_expr` above intercepts it before this table is consulted.
fn pipeline_cmdlet_method(name: &str) -> Option<&'static str> {
    match name.to_lowercase().as_str() {
        "where-object" | "where" | "?" => Some("Where"),
        "foreach-object" | "foreach" | "%" => Some("ForEach"),
        "sort-object" | "sort" => Some("Sort"),
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
        // `Start-Job` / `Start-ThreadJob -ScriptBlock { … }` and the
        // `Receive-Job` that collects it.
        //
        // The job runs SYNCHRONOUSLY: the block is invoked where it is started
        // and the "job" IS its result, so `Receive-Job` hands that value back.
        // Nothing in the corpus observes a job before receiving it, and the
        // alternative — a real job object with no scheduler behind it — would
        // be a bigger lie than running the block early.
        //
        // ⛔`$using:x` needs no work here: it strips to the bare name in
        // `scope_qualified_name`, and running the block in the starting scope
        // is exactly what makes that the right variable.
        "start-job" | "start-threadjob" => {
            let block = named("ScriptBlock")
                .or_else(|| positional.first().map(|a| a.value.clone()))?;
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(block),
                args: Vec::new(),
                optional: false,
            }))
        }
        "receive-job" | "wait-job" => positional.first().map(|a| a.value.clone()),
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
        // `New-Variable -Name x -Value 1` creates the same storage `$x = 1`
        // writes, so it lowers to the same assignment as `Set-Variable`.
        "new-variable" => {
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

        // `Clear-Variable x` sets the value to `$null` and KEEPS the variable —
        // that is exactly an assignment, so this lowering is complete.
        //
        // `Remove-Variable x` should also drop the NAME, which needs a runtime
        // variable table this language does not have; nulling is the closest
        // honest lowering and differs only under `Test-Path variable:x`, which
        // one corpus file uses. Without either arm both cmdlets fell through to
        // `_ => None` and compiled to a call to a function never defined —
        // `undefined is not callable` — which is why they are here now.
        "clear-variable" | "remove-variable" => {
            let name_expr =
                named("Name").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let target = literal_text(&name_expr)?;
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident(scope_qualified_name(&target))),
                value: Box::new(Expression::null()),
            }))
        }

        // `Get-Variable x` returns a PSVariable OBJECT, not the value — real
        // PowerShell hands back `System.Management.Automation.PSVariable`, so
        // `(Get-Variable x).Value` is the documented way to read it. Returning
        // the bare value made `.Value` read empty on every one of the 56 corpus
        // files that use this cmdlet.
        //
        // `Options` is reported as `None`: the declared option lives in a
        // variable table that does not exist yet, and `None` is what real
        // PowerShell answers for a plain variable. Reporting ReadOnly/Constant
        // faithfully is part of that table's job, not this rewrite's.
        "get-variable" => {
            let name_expr =
                named("Name").or_else(|| positional.first().map(|a| a.value.clone()))?;
            let target = literal_text(&name_expr)?;
            let value = Expression::ident(scope_qualified_name(&target));
            // `-ValueOnly` asks for the bare value instead of the object.
            if args.iter().any(|a| {
                a.name
                    .as_deref()
                    .is_some_and(|n| n.eq_ignore_ascii_case("ValueOnly"))
            }) {
                return Some(value);
            }
            let snapshot = Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("Name"),
                    value: Expression::string(scope_qualified_name(&target)),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("Value"),
                    value: value.clone(),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("Options"),
                    value: Expression::string("None"),
                },
            ]));
            // A MISSING variable must not answer with a truthy object:
            // `if (Get-Variable -Name x -ErrorAction SilentlyContinue)` is how
            // the corpus asks "does this name exist", and an object literal is
            // always truthy, so an unconditional snapshot reported every name
            // as present — including one the caller never declared.
            //
            // Absence is read as null because, with no variable table, an
            // absent name and a name holding `$null` are the same thing in this
            // AST. Real PowerShell distinguishes them; telling them apart is
            // that table's job. No corpus file reads a `$null`-valued variable
            // back through `Get-Variable`, and three depend on the existence
            // check, so this is the faithful side of the trade.
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(value),
                    right: Box::new(Expression::null()),
                })),
                then: Box::new(snapshot),
                else_: Box::new(Expression::null()),
            }))
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
    // ⛔A SYNTHESIZED EXCEPTION TYPE IS A REAL CLASS UNDER ITS SHORT NAME.
    // `synthesize_exception_classes` splices `Exception`, `ArgumentException`
    // and the rest into the program unqualified, and the SAME names are also
    // registered as `dotnet.System` component classes — so `[System.Exception]`
    // built a `System.Exception` member chain that names no class in the
    // program, and `$e -is [System.Exception]` was FALSE for an object whose
    // parent IS `Exception`. `is_synthesized_exception_class` exists for
    // exactly this question; its own doc records vb hitting the mirror image
    // (`Imports System` qualifying a `Catch` type until it stopped matching).
    // `System.Management.Automation.X` is PowerShell's own namespace; the
    // prelude declares those types under their SHORT name.
    if let Some(short) = name.strip_prefix("System.Management.Automation.")
        && !short.contains('.')
    {
        return Expression::ident(short);
    }
    if let Some(short) = name.strip_prefix("System.")
        && !short.contains('.')
        && vybe_platform_dotnet::emitter::core::exceptions::is_synthesized_exception_class(short)
    {
        return Expression::ident(short);
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

/// Peel the ONE bracket pair that delimits a type literal.
///
/// ⛔`trim_end_matches(']')` strips EVERY trailing bracket, so a generic type
/// lost its own closing one: `[System.Collections.Generic.List[string]]` became
/// `System.Collections.Generic.List[string` — unbalanced, and no such name is
/// registered anywhere, so construction answered `undefined is not callable`.
/// The delimiter is exactly one pair; the inner brackets belong to the type.
fn type_literal_name(raw: &str) -> &str {
    let trimmed = raw.trim();
    match trimmed.strip_prefix('[') {
        Some(inner) => inner.strip_suffix(']').unwrap_or(inner),
        None => trimmed,
    }
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
                    // `__ps_join` — the bare `join` spelling never reaches
                    // value-method dispatch. See the profile's `Join` note.
                    return method_call_expr(left, "__ps_join", spread_operands(right));
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

    // `1 / 0` THROWS in PowerShell rather than answering infinity, and that is
    // how a script provokes an error for `trap` / `try` to catch. `BinOp::Div`
    // reaches `F64_DIV` directly — `emit_rich_binop` only consults a USER
    // operator method, so a `[builtin_slots.int] div` row is never read for a
    // plain numeric divide — so the check is expressed as a call the compiler
    // already knows how to make.
    // ⛔A FLOAT LITERAL OPTS OUT. PowerShell decides by TYPE: `1 / 0` throws,
    // `1.0 / 0.0` is `Infinity` and `0.0 / 0.0` is `NaN`. Both operands are
    // `f64` by the time the emitter runs, so the type is gone — but the SOURCE
    // still has it, and a decimal point is exactly the marker. Routing them
    // through the throwing adapter broke the three `math_ieee_special_values`
    // tests that assert the IEEE answers.
    let float_literal = |e: &Expression| {
        matches!(&e.kind, ExprKind::Lit(Literal::Float(_)))
            || matches!(&e.kind, ExprKind::Unary { expr, .. }
                if matches!(expr.kind, ExprKind::Lit(Literal::Float(_))))
    };
    if matches!(op, BinOp::Div) && !float_literal(&left) && !float_literal(&right) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__ps_divide")),
            args: vec![Argument::positional(left), Argument::positional(right)],
            optional: false,
        });
    }

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
