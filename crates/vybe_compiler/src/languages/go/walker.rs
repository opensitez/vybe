//! Go walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//!
//! ## Go-specific normalisations
//!
//! - **Multiple return values**: Go functions can return multiple values.
//!   For simplicity we compile to returning a single array/tuple.
//! - **Short variable declaration** (`:=`): Maps to `VarDecl` with `Let`.
//! - **Methods**: Go methods on structs are compiled into `StructDecl`
//!   fragments with the receiver kept as the first explicit parameter.
//! - **Structs**: Mapped to `StructDecl` with fields.
//! - **Interfaces**: Mapped to `InterfaceDecl`.
//! - **`range`**: Mapped to `ForIn` with `of: true`.
//! - **`defer`**: Lowered to a per-function stack of zero-arg closures that
//!   drain from a synthesized `finally` block in LIFO order.
//! - **`go`**: Lowered to the shared thread/task emitter surface.
//! - **`fallthrough`**: Not yet supported in switch.
//! - **`select`**: Lowered as a compile-safe block for the dummy concurrency tests.
//! - **`chan` / `<-`**: Lowered into compile-safe object/array operations.
//! - **`nil`**: Mapped to `ExprKind::Lit(Literal::Null)`.
//! - **`make` / `new`**: `make` for slices/maps is rewritten to array/dict
//!   creation. `new(T)` becomes `&T{}` (pointer to zero value).
//! - **`append`**: Rewritten to slice concat so the updated slice value is preserved.
//! - **`len` / `cap`**: Builtin functions mapped to host calls.
//! - **`panic` / `recover`**: Mapped to throw/try-catch.
//! - **`_` blank identifier**: Ignored in assignments.

use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};
use crate::ast::*;
use super::{GoParser, Rule};

// ══════════════════════════════════════════════════════════════════════════════════════════
// Entry point
// ══════════════════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs = GoParser::parse(Rule::program, source)
        .map_err(|e| format!("Go parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut _package_name = String::new();

    for top in pairs {
        if top.as_rule() == Rule::EOI { continue; }
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            _ => {
                if let Some(stmt) = walk_top_level(top)? {
                    body.push(stmt);
                }
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI => continue,
                Rule::package_clause => {
                    _package_name = walk_package_clause(pair)?;
                }
                Rule::import_declarations => {
                    for imp in pair.into_inner() {
                        if imp.as_rule() == Rule::import_declaration {
                            imports.push(walk_import(imp)?);
                        }
                    }
                }
                _ => {
                    if let Some(stmt) = walk_top_level(pair)? {
                        body.push(stmt);
                    }
                }
            }
        }
    }

    Ok(normalize_go_module(Module {
        name: _package_name,
        language: Lang::Go,
        body,
        imports,
    }))
}

#[derive(Clone, Default)]
struct GoFunctionSignature {
    params: Vec<Option<String>>,
    return_type: Option<String>,
}

#[derive(Clone, Default)]
struct GoNormalizeEnv {
    value_types: HashMap<String, String>,
    fixed_arrays: HashMap<String, String>,
    slice_caps: HashMap<String, Expression>,
    type_names: HashSet<String>,
    return_type: Option<String>,
}

#[derive(Default)]
struct GoNormalizeState {
    next_temp: usize,
}

struct GoSignatureInfo {
    params: Vec<Param>,
    return_type: Option<String>,
    named_results: Vec<Param>,
}

fn normalize_go_module(mut module: Module) -> Module {
    module.body = merge_go_struct_decls(&module.body);
    let signatures = collect_go_function_signatures(&module.body);
    let globals = collect_go_global_fixed_arrays(&module.body, &signatures);
    let type_names = collect_go_type_names(&module.body);
    let mut state = GoNormalizeState::default();
    let mut env = GoNormalizeEnv {
        value_types: HashMap::new(),
        fixed_arrays: globals.clone(),
        slice_caps: HashMap::new(),
        type_names,
        return_type: None,
    };

    let mut normalized = Vec::with_capacity(module.body.len());
    for stmt in &module.body {
        normalized.extend(normalize_go_statement(stmt, &mut env, &signatures, &mut state));
    }
    module.body = normalized;
    module
}

fn collect_go_function_signatures(body: &[Statement]) -> HashMap<String, GoFunctionSignature> {
    let mut signatures = HashMap::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                ..
            } => {
                signatures.insert(
                    name.clone(),
                    GoFunctionSignature {
                        params: params.iter().map(|param| param.type_hint.clone()).collect(),
                        return_type: return_type.clone(),
                    },
                );
            }
            StmtKind::StructDecl { members, .. } => {
                for member in members {
                    if let ClassMember::Method(stmt) = member {
                        if let StmtKind::FunctionDecl {
                            name,
                            params,
                            return_type,
                            ..
                        } = &stmt.kind
                        {
                            signatures.insert(
                                name.clone(),
                                GoFunctionSignature {
                                    params: params.iter().map(|param| param.type_hint.clone()).collect(),
                                    return_type: return_type.clone(),
                                },
                            );
                        }
                    }
                }
            }
            _ => {}
        }
    }
    signatures
}

fn collect_go_type_names(body: &[Statement]) -> HashSet<String> {
    let mut type_names = HashSet::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. }
            | StmtKind::InterfaceDecl { name, .. }
            | StmtKind::EnumDecl { name, .. }
            | StmtKind::ClassDecl { name, .. } => {
                type_names.insert(name.clone());
            }
            StmtKind::VarDecl { declarations, kind } if *kind == VarDeclKind::Let => {
                for decl in declarations {
                    let BindingPattern::Ident(name) = &decl.pattern else {
                        continue;
                    };
                    if matches!(decl.init.as_ref().map(|expr| &expr.kind), Some(ExprKind::Lit(Literal::Str(_))))
                        && decl.type_hint.is_none()
                    {
                        type_names.insert(name.clone());
                    }
                }
            }
            _ => {}
        }
    }
    type_names
}

fn merge_go_struct_decls(body: &[Statement]) -> Vec<Statement> {
    let mut first_index: HashMap<String, usize> = HashMap::new();
    for (index, stmt) in body.iter().enumerate() {
        if let StmtKind::StructDecl { name, .. } = &stmt.kind {
            first_index.entry(name.clone()).or_insert(index);
        }
    }

    let mut emitted = std::collections::HashSet::new();
    let mut merged_body = Vec::with_capacity(body.len());

    for (index, stmt) in body.iter().enumerate() {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. } => {
                if first_index.get(name) != Some(&index) || !emitted.insert(name.clone()) {
                    continue;
                }

                let mut merged = stmt.clone();
                if let StmtKind::StructDecl {
                    interfaces,
                    members,
                    ..
                } = &mut merged.kind
                {
                    for later in body.iter().skip(index + 1) {
                        if let StmtKind::StructDecl {
                            name: later_name,
                            interfaces: later_interfaces,
                            members: later_members,
                            ..
                        } = &later.kind
                        {
                            if later_name == name {
                                members.extend(later_members.clone());
                                for interface in later_interfaces {
                                    if !interfaces.iter().any(|existing| existing == interface) {
                                        interfaces.push(interface.clone());
                                    }
                                }
                            }
                        }
                    }
                }

                merged_body.push(merged);
            }
            _ => merged_body.push(stmt.clone()),
        }
    }

    merged_body
}

fn collect_go_global_fixed_arrays(
    body: &[Statement],
    signatures: &HashMap<String, GoFunctionSignature>,
) -> HashMap<String, String> {
    let env = GoNormalizeEnv::default();
    let mut globals = HashMap::new();

    for stmt in body {
        if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
            for decl in declarations {
                if let Some((name, type_name)) = go_decl_fixed_array_binding(decl, &env, signatures) {
                    globals.insert(name, type_name);
                }
            }
        }
    }

    globals
}

fn normalize_go_block(
    stmts: &[Statement],
    base_env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut env = base_env.clone();
    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        normalized.extend(normalize_go_statement(stmt, &mut env, signatures, state));
    }
    normalized
}

fn normalize_go_function_body(
    stmts: &[Statement],
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut named_result: Option<Param> = None;
    let mut body_stmts = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        if named_result.is_none() {
            if let Some(param) = go_extract_named_result_marker(stmt) {
                env.value_types.insert(param.name.clone(), param.type_hint.clone().unwrap_or_else(|| "object".to_string()));
                named_result = Some(param);
                continue;
            }
        }
        body_stmts.push(stmt.clone());
    }

    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in &body_stmts {
        normalized.extend(normalize_go_statement(stmt, env, signatures, state));
    }

    let (normalized, final_return) = if let Some(param) = named_result {
        go_lower_named_result_body(normalized, &param, state)
    } else {
        (normalized, None)
    };

    lower_go_defer_body(normalized, env, signatures, state, final_return)
}

fn lower_go_defer_body(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    final_return: Option<Expression>,
) -> Vec<Statement> {
    let stack_name = fresh_go_temp(state, "__go_defer_stack");
    let (lowered_body, has_defer) = lower_go_defer_statements(body, env, signatures, state, &stack_name);
    if !has_defer {
        let mut body = lowered_body;
        if let Some(expr) = final_return {
            body.push(Statement::new(StmtKind::Return(Some(expr))));
        }
        return body;
    }

    let stack_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(stack_name.clone()),
            type_hint: None,
            init: Some(Expression::null()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let drain_name = fresh_go_temp(state, "__go_defer_fn");
    let drain_loop = Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::ident(&stack_name)),
            right: Box::new(Expression::null()),
        }),
        body: vec![
            go_defer_temp_decl(
                drain_name.clone(),
                None,
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "fn".to_string(),
                    null_safe: false,
                }),
            ),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&stack_name)],
                value: Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "next".to_string(),
                    null_safe: false,
                }),
            }),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&drain_name)),
                args: Vec::new(),
                optional: false,
            }))),
        ],
        else_body: None,
    });

    let mut body = vec![
        stack_decl,
        Statement::new(StmtKind::Try {
            body: lowered_body,
            catches: Vec::new(),
            else_body: None,
            finally: Some(vec![drain_loop]),
        }),
    ];
    if let Some(expr) = final_return {
        body.push(Statement::new(StmtKind::Return(Some(expr))));
    }
    body
}

fn go_extract_named_result_marker(stmt: &Statement) -> Option<Param> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_named_result") || args.len() != 2 {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let type_hint = go_type_name_from_expr(&args[1].value)?;
    Some(Param {
        name: name.clone(),
        type_hint: Some(type_hint),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

fn go_lower_named_result_body(
    body: Vec<Statement>,
    result: &Param,
    state: &mut GoNormalizeState,
) -> (Vec<Statement>, Option<Expression>) {
    let sentinel = fresh_go_temp(state, "__go_named_return");
    let catch_name = fresh_go_temp(state, "__go_named_return_exc");
    let result_name = result.name.clone();
    let result_type = result.type_hint.clone().unwrap_or_else(|| "object".to_string());

    let rewritten_body = go_rewrite_named_result_returns(body, &result_name, &sentinel);
    let result_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(result_name.clone()),
            type_hint: Some(result_type.clone()),
            init: Some(go_zero_value_expr(&result_type)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });
    let catch_body = vec![Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::ident(&catch_name)),
            right: Box::new(Expression::string(&sentinel)),
        }),
        then_body: vec![Statement::new(StmtKind::Throw {
            expr: Some(Expression::ident(&catch_name)),
            cause: None,
        })],
        elifs: Vec::new(),
        else_body: None,
    })];

    (
        vec![
            result_decl,
            Statement::new(StmtKind::Try {
                body: rewritten_body,
                catches: vec![CatchClause {
                    types: Vec::new(),
                    var_name: Some(catch_name),
                    stack_var: None,
                    body: catch_body,
                    when_clause: None,
                }],
                else_body: None,
                finally: None,
            }),
        ],
        Some(Expression::ident(&result_name)),
    )
}

fn go_rewrite_named_result_returns(
    body: Vec<Statement>,
    result_name: &str,
    sentinel: &str,
) -> Vec<Statement> {
    let mut rewritten = Vec::with_capacity(body.len());
    for stmt in body {
        rewritten.extend(go_rewrite_named_result_return_stmt(stmt, result_name, sentinel));
    }
    rewritten
}

fn go_rewrite_named_result_return_stmt(
    stmt: Statement,
    result_name: &str,
    sentinel: &str,
) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::Return(expr) => {
            let mut rewritten = Vec::new();
            if let Some(expr) = expr {
                rewritten.push(Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(result_name)],
                    value: expr,
                }));
            }
            rewritten.push(Statement::new(StmtKind::Throw {
                expr: Some(Expression::string(sentinel)),
                cause: None,
            }));
            rewritten
        }
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(go_rewrite_named_result_returns(
            body,
            result_name,
            sentinel,
        )))],
        StmtKind::If { cond, then_body, elifs, else_body } => vec![Statement::new(StmtKind::If {
            cond,
            then_body: go_rewrite_named_result_returns(then_body, result_name, sentinel),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| (cond, go_rewrite_named_result_returns(body, result_name, sentinel)))
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::For { init, cond, update, body } => vec![Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
        })],
        StmtKind::ForIn { var, key, iter, body, of, else_body, is_async } => vec![Statement::new(StmtKind::ForIn {
            var,
            key,
            iter,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            of,
            else_body: else_body.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
            is_async,
        })],
        StmtKind::While { cond, body, else_body } => vec![Statement::new(StmtKind::While {
            cond,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            else_body: else_body.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            cond,
            until,
        })],
        StmtKind::Switch { expr, cases, default } => vec![Statement::new(StmtKind::Switch {
            expr,
            cases: cases
                .into_iter()
                .map(|case| SwitchCase {
                    conditions: case.conditions,
                    body: go_rewrite_named_result_returns(case.body, result_name, sentinel),
                })
                .collect(),
            default: default.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::Try { body, catches, else_body, finally } => vec![Statement::new(StmtKind::Try {
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            catches: catches
                .into_iter()
                .map(|catch| CatchClause {
                    types: catch.types,
                    var_name: catch.var_name,
                    stack_var: catch.stack_var,
                    body: go_rewrite_named_result_returns(catch.body, result_name, sentinel),
                    when_clause: catch.when_clause,
                })
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
            finally: finally.map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        _ => vec![stmt],
    }
}

fn lower_go_defer_statements(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
) -> (Vec<Statement>, bool) {
    let mut lowered = Vec::with_capacity(body.len());
    let mut has_defer = false;

    for stmt in body {
        if let Some(expr) = go_extract_defer_expr(&stmt) {
            lowered.extend(go_lower_defer_stmt(expr, env, signatures, state, stack_name));
            has_defer = true;
            continue;
        }

        let (next_stmt, nested_has_defer) = lower_go_defer_statement(stmt, env, signatures, state, stack_name);
        lowered.push(next_stmt);
        has_defer |= nested_has_defer;
    }

    (lowered, has_defer)
}

fn lower_go_defer_statement(
    stmt: Statement,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
) -> (Statement, bool) {
    match stmt.kind {
        StmtKind::Block(body) => {
            let (body, has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            (Statement::new(StmtKind::Block(body)), has_defer)
        }
        StmtKind::If { cond, then_body, elifs, else_body } => {
            let (next_then, mut has_defer) = lower_go_defer_statements(then_body, env, signatures, state, stack_name);
            let mut next_elifs = Vec::with_capacity(elifs.len());
            for (elif_cond, elif_body) in elifs {
                let (next_body, nested_has_defer) = lower_go_defer_statements(elif_body, env, signatures, state, stack_name);
                next_elifs.push((elif_cond, next_body));
                has_defer |= nested_has_defer;
            }
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::If {
                    cond,
                    then_body: next_then,
                    elifs: next_elifs,
                    else_body: next_else,
                }),
                has_defer,
            )
        }
        StmtKind::For { init, cond, update, body } => {
            let (body, has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            (
                Statement::new(StmtKind::For { init, cond, update, body }),
                has_defer,
            )
        }
        StmtKind::ForIn { var, key, iter, body, of, else_body, is_async } => {
            let (body, mut has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::ForIn { var, key, iter, body, of, else_body: next_else, is_async }),
                has_defer,
            )
        }
        StmtKind::While { cond, body, else_body } => {
            let (body, mut has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::While { cond, body, else_body: next_else }),
                has_defer,
            )
        }
        StmtKind::DoWhile { body, cond, until } => {
            let (body, has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            (
                Statement::new(StmtKind::DoWhile { body, cond, until }),
                has_defer,
            )
        }
        StmtKind::Switch { expr, cases, default } => {
            let mut has_defer = false;
            let next_cases = cases
                .into_iter()
                .map(|case| {
                    let (body, nested_has_defer) = lower_go_defer_statements(case.body, env, signatures, state, stack_name);
                    has_defer |= nested_has_defer;
                    SwitchCase {
                        conditions: case.conditions,
                        body,
                    }
                })
                .collect();
            let next_default = if let Some(body) = default {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Switch { expr, cases: next_cases, default: next_default }),
                has_defer,
            )
        }
        StmtKind::Try { body, catches, else_body, finally } => {
            let (body, mut has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
            let next_catches = catches
                .into_iter()
                .map(|catch| {
                    let (body, nested_has_defer) = lower_go_defer_statements(catch.body, env, signatures, state, stack_name);
                    has_defer |= nested_has_defer;
                    CatchClause {
                        types: catch.types,
                        var_name: catch.var_name,
                        stack_var: catch.stack_var,
                        body,
                        when_clause: catch.when_clause,
                    }
                })
                .collect();
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            let next_finally = if let Some(body) = finally {
                let (body, nested_has_defer) = lower_go_defer_statements(body, env, signatures, state, stack_name);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Try {
                    body,
                    catches: next_catches,
                    else_body: next_else,
                    finally: next_finally,
                }),
                has_defer,
            )
        }
        _ => (stmt, false),
    }
}

fn go_extract_defer_expr(stmt: &Statement) -> Option<Expression> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_defer") || args.len() != 1 {
        return None;
    }
    Some(args[0].value.clone())
}

fn go_lower_defer_stmt(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
) -> Vec<Statement> {
    let mut stmts = Vec::new();
    let mut captures = Vec::new();

    let deferred_expr = match expr.kind {
        ExprKind::Call { callee, args, optional } => {
            let deferred_callee = match callee.as_ref() {
                Expression {
                    kind: ExprKind::Member { object, field, null_safe },
                    ..
                } => {
                    let receiver_type = go_expr_type_hint(object, env, signatures);
                    let deferred_object = if matches!(object.as_ref().kind, ExprKind::Ident(_)) && receiver_type.is_none() {
                        object.as_ref().clone()
                    } else {
                        let temp_name = fresh_go_temp(state, "__go_defer_recv");
                        stmts.push(go_defer_temp_decl(
                            temp_name.clone(),
                            receiver_type,
                            object.as_ref().clone(),
                        ));
                        captures.push(temp_name.clone());
                        Expression::ident(&temp_name)
                    };
                    Expression::new(ExprKind::Member {
                        object: Box::new(deferred_object),
                        field: field.clone(),
                        null_safe: *null_safe,
                    })
                }
                _ => {
                    let temp_name = fresh_go_temp(state, "__go_defer_fn");
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(callee.as_ref(), env, signatures),
                        callee.as_ref().clone(),
                    ));
                    captures.push(temp_name.clone());
                    Expression::ident(&temp_name)
                }
            };

            let deferred_args = args
                .into_iter()
                .map(|arg| {
                    let temp_name = fresh_go_temp(state, "__go_defer_arg");
                    let value = go_wrap_fixed_array_copy(arg.value, env, signatures);
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(&value, env, signatures),
                        value,
                    ));
                    captures.push(temp_name.clone());
                    Argument {
                        value: Expression::ident(&temp_name),
                        name: arg.name,
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect();

            Expression::new(ExprKind::Call {
                callee: Box::new(deferred_callee),
                args: deferred_args,
                optional,
            })
        }
        _ => expr,
    };

    let closure = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(deferred_expr))]),
        is_async: false,
        captures,
    });
    stmts.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(stack_name)],
        value: Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("fn"),
                value: closure,
            },
            ObjectProperty::KeyValue {
                key: Expression::string("next"),
                value: Expression::ident(stack_name),
            },
        ])),
    }));
    stmts
}

fn go_defer_temp_decl(name: String, type_hint: Option<String>, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn normalize_go_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    match &stmt.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles,
            is_async,
            is_generator,
            is_sub,
        } => {
            let mut fn_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                slice_caps: env.slice_caps.clone(),
                type_names: env.type_names.clone(),
                return_type: return_type.clone(),
            };
            for param in params {
                if let Some(type_hint) = param.type_hint.as_ref() {
                    fn_env.value_types.insert(param.name.clone(), type_hint.clone());
                }
                if let Some(type_hint) = param.type_hint.as_deref().filter(|hint| go_is_fixed_array_type(hint)) {
                    fn_env.fixed_arrays.insert(param.name.clone(), type_hint.to_string());
                }
            }

            vec![Statement::new(StmtKind::FunctionDecl {
                name: name.clone(),
                params: params.clone(),
                return_type: return_type.clone(),
                body: normalize_go_function_body(body, &mut fn_env, signatures, state),
                modifiers: modifiers.clone(),
                handles: handles.clone(),
                is_async: *is_async,
                is_generator: *is_generator,
                is_sub: *is_sub,
            })]
        }
        StmtKind::VarDecl { declarations, kind } => {
            let mut normalized = Vec::with_capacity(declarations.len());
            for decl in declarations {
                let mut next_decl = decl.clone();
                if let Some(pattern) = go_single_named_binding_pattern(&next_decl.pattern) {
                    next_decl.pattern = pattern;
                }
                next_decl.init = decl
                    .init
                    .as_ref()
                    .map(|expr| {
                        if go_is_two_value_binding_pattern(&decl.pattern) {
                            if let Some(tuple_expr) = go_normalize_map_lookup_tuple_expr(expr, env, signatures, state) {
                                return tuple_expr;
                            }
                        }
                        normalize_go_expr(expr, env, signatures, state)
                    });
                next_decl.array_bounds = decl
                    .array_bounds
                    .as_ref()
                    .map(|bounds| bounds.iter().map(|expr| normalize_go_expr(expr, env, signatures, state)).collect());

                if next_decl.init.is_none()
                    && next_decl
                        .type_hint
                        .as_deref()
                        .is_some_and(go_is_fixed_array_type)
                {
                    next_decl.init = next_decl
                        .type_hint
                        .as_deref()
                        .map(go_zero_value_expr);
                } else if let Some(init_expr) = next_decl.init.take() {
                    next_decl.init = Some(go_wrap_fixed_array_copy(init_expr, env, signatures));
                }

                if let Some((name, type_name)) = go_decl_fixed_array_binding(&next_decl, env, signatures) {
                    env.fixed_arrays.insert(name, type_name);
                }
                if let Some((name, type_name)) = go_decl_binding_type(&next_decl, env, signatures) {
                    env.value_types.insert(name, type_name);
                }
                if let Some((name, cap_expr)) = go_decl_slice_capacity_binding(decl, env, signatures, state) {
                    env.slice_caps.insert(name, cap_expr);
                }
                normalized.push(next_decl);
            }

            vec![Statement::new(StmtKind::VarDecl {
                declarations: normalized,
                kind: kind.clone(),
            })]
        }
        StmtKind::Expr(expr) => vec![Statement::new(StmtKind::Expr(normalize_go_expr(
            expr, env, signatures, state,
        )))],
        StmtKind::Assign { targets, value } => {
            let mut next_value = normalize_go_expr(value, env, signatures, state);
            next_value = go_wrap_fixed_array_copy(next_value, env, signatures);
            if let [target] = targets.as_slice() {
                if let ExprKind::Ident(name) = &target.kind {
                    if let Some(type_name) = go_expr_type_hint(&next_value, env, signatures) {
                        env.value_types.insert(name.clone(), type_name);
                    }
                    if let Some(cap_expr) = go_make_slice_capacity_expr(value, env, signatures, state) {
                        env.slice_caps.insert(name.clone(), cap_expr);
                    }
                }
            }
            vec![Statement::new(StmtKind::Assign {
                targets: targets
                    .iter()
                    .map(|target| normalize_go_lvalue_expr(target, env, signatures, state))
                    .collect(),
                value: next_value,
            })]
        }
        StmtKind::CompoundAssign { target, op, value } => vec![Statement::new(StmtKind::CompoundAssign {
            target: normalize_go_lvalue_expr(target, env, signatures, state),
            op: *op,
            value: normalize_go_expr(value, env, signatures, state),
        })],
        StmtKind::Return(expr) => {
            let next_expr = expr.as_ref().map(|value| {
                let normalized = normalize_go_expr(value, env, signatures, state);
                if env
                    .return_type
                    .as_deref()
                    .is_some_and(go_is_fixed_array_type)
                {
                    go_wrap_fixed_array_copy(normalized, env, signatures)
                } else {
                    normalized
                }
            });
            vec![Statement::new(StmtKind::Return(next_expr))]
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let next_elifs = elifs
                .iter()
                .map(|(elif_cond, elif_body)| {
                    (
                        normalize_go_expr(elif_cond, env, signatures, state),
                        normalize_go_block(elif_body, env, signatures, state),
                    )
                })
                .collect();
            let next_else = else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state));
            vec![Statement::new(StmtKind::If {
                cond: normalize_go_expr(cond, env, signatures, state),
                then_body: normalize_go_block(then_body, env, signatures, state),
                elifs: next_elifs,
                else_body: next_else,
            })]
        }
        StmtKind::Switch { expr, cases, default } => {
            let next_cases = cases
                .iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(value) => {
                                CaseCondition::Value(normalize_go_expr(value, env, signatures, state))
                            }
                            _ => condition.clone(),
                        })
                        .collect(),
                    body: normalize_go_block(&case.body, env, signatures, state),
                })
                .collect();
            vec![Statement::new(StmtKind::Switch {
                expr: normalize_go_expr(expr, env, signatures, state),
                cases: next_cases,
                default: default
                    .as_ref()
                    .map(|body| normalize_go_block(body, env, signatures, state)),
            })]
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_env = env.clone();
            let next_init = init.as_ref().map(|stmt| {
                Box::new(normalize_go_single_statement(stmt, &mut loop_env, signatures, state))
            });
            let next_cond = cond
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_update = update
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_body = normalize_go_block(body, &loop_env, signatures, state);
            vec![Statement::new(StmtKind::For {
                init: next_init,
                cond: next_cond,
                update: next_update,
                body: next_body,
            })]
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            let next_iter = normalize_go_expr(iter, env, signatures, state);
            if *of && go_expr_is_fixed_array(&next_iter, env, signatures) {
                lower_go_fixed_array_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else {
                vec![Statement::new(StmtKind::ForIn {
                    var: var.clone(),
                    key: key.clone(),
                    iter: next_iter,
                    body: normalize_go_block(body, env, signatures, state),
                    of: *of,
                    else_body: else_body
                        .as_ref()
                        .map(|body| normalize_go_block(body, env, signatures, state)),
                    is_async: *is_async,
                })]
            }
        }
        StmtKind::While { cond, body, else_body } => vec![Statement::new(StmtKind::While {
            cond: normalize_go_expr(cond, env, signatures, state),
            body: normalize_go_block(body, env, signatures, state),
            else_body: else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: normalize_go_block(body, env, signatures, state),
            cond: normalize_go_expr(cond, env, signatures, state),
            until: *until,
        })],
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(normalize_go_block(
            body, env, signatures, state,
        )))],
        StmtKind::Throw { expr, cause } => vec![Statement::new(StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
            cause: cause
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
        })],
        StmtKind::StructDecl {
            name,
            interfaces,
            members,
            visibility,
            decorators,
        } => {
            let normalized_members = members
                .iter()
                .map(|member| match member {
                    ClassMember::Method(stmt) => {
                        let normalized_method = normalize_go_single_statement(stmt, env, signatures, state);
                        ClassMember::Method(Box::new(normalized_method))
                    }
                    ClassMember::Field {
                        name,
                        type_hint,
                        init,
                        modifiers,
                        with_events,
                        array_bounds,
                    } => ClassMember::Field {
                        name: name.clone(),
                        type_hint: type_hint.clone(),
                        init: init.as_ref().map(|expr| normalize_go_expr(expr, env, signatures, state)),
                        modifiers: modifiers.clone(),
                        with_events: *with_events,
                        array_bounds: array_bounds.as_ref().map(|bounds| {
                            bounds
                                .iter()
                                .map(|expr| normalize_go_expr(expr, env, signatures, state))
                                .collect()
                        }),
                    },
                    _ => member.clone(),
                })
                .collect();
            vec![Statement::new(StmtKind::StructDecl {
                name: name.clone(),
                interfaces: interfaces.clone(),
                members: normalized_members,
                visibility: *visibility,
                decorators: decorators.clone(),
            })]
        }
        _ => vec![stmt.clone()],
    }
}

fn normalize_go_single_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Statement {
    let mut normalized = normalize_go_statement(stmt, env, signatures, state);
    if normalized.len() == 1 {
        normalized.pop().unwrap()
    } else {
        Statement::new(StmtKind::Block(normalized))
    }
}

fn normalize_go_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Binary { op, left, right } => {
            let next_left = normalize_go_expr(left, env, signatures, state);
            let next_right = normalize_go_expr(right, env, signatures, state);
            let normalized_op = if *op == BinOp::Div
                && go_expr_type_hint(&next_left, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
                && go_expr_type_hint(&next_right, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
            {
                BinOp::IDiv
            } else {
                *op
            };

            if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && go_expr_is_fixed_array(&next_left, env, signatures)
                && go_expr_is_fixed_array(&next_right, env, signatures)
            {
                let equal = go_builtin_call("__go_fixed_array_equal", vec![next_left, next_right]);
                if normalized_op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                }
            } else {
                Expression::new(ExprKind::Binary {
                    op: normalized_op,
                    left: Box::new(next_left),
                    right: Box::new(next_right),
                })
            }
        }
        ExprKind::Unary { op: UnaryOp::AddrOf, expr } => {
            let next_expr = normalize_go_expr(expr, env, signatures, state);
            if let Some(place) = go_expr_to_place(&next_expr) {
                Expression::new(ExprKind::RefOf(Box::new(place)))
            } else {
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(next_expr),
                })
            }
        }
        ExprKind::Unary { op: UnaryOp::Deref, expr } => {
            Expression::new(ExprKind::RefLoad(Box::new(normalize_go_expr(expr, env, signatures, state))))
        }
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(normalize_go_expr(expr, env, signatures, state)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(normalize_go_expr(cond, env, signatures, state)),
            then: Box::new(normalize_go_expr(then, env, signatures, state)),
            else_: Box::new(normalize_go_expr(else_, env, signatures, state)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(normalize_go_expr(object, env, signatures, state)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let next_object = normalize_go_expr(object, env, signatures, state);
            let next_index = normalize_go_expr(index, env, signatures, state);
            if go_expr_type_hint(&next_object, env, signatures).as_deref() == Some("string") {
                go_member_call(next_object, "charCodeAt", vec![next_index])
            } else if let Some(value_type) = go_expr_type_hint(&next_object, env, signatures)
                .and_then(|type_name| go_map_value_type(&type_name))
            {
                go_build_map_read_expr(next_object, next_index, &value_type)
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(next_object),
                    index: Box::new(next_index),
                    null_safe: *null_safe,
                })
            }
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            let next_callee = normalize_go_expr(callee, env, signatures, state);
            let signature = match &next_callee.kind {
                ExprKind::Ident(name) => signatures.get(name),
                _ => None,
            };
            let mut next_args = args
                .iter()
                .enumerate()
                .map(|(idx, arg)| {
                    let mut value = normalize_go_expr(&arg.value, env, signatures, state);
                    if signature
                        .and_then(|sig| sig.params.get(idx))
                        .and_then(|hint| hint.as_deref())
                        .is_some_and(go_is_fixed_array_type)
                    {
                        value = go_wrap_fixed_array_copy(value, env, signatures);
                    }
                    Argument {
                        value,
                        name: arg.name.clone(),
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect::<Vec<_>>();

            let call_name = go_expr_call_name(&next_callee);

            if call_name.as_deref() == Some("make") {
                if let Some(type_name) = next_args.first().and_then(|arg| go_type_name_from_expr(&arg.value)) {
                    if go_is_channel_type(&type_name) {
                        let capacity = next_args.get(1).map(|arg| arg.value.clone());
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(go_channel_object_expr(capacity)),
                            type_name,
                        });
                    }
                    if go_is_slice_type(&type_name) {
                        let len_expr = next_args
                            .get(1)
                            .map(|arg| arg.value.clone())
                            .unwrap_or_else(|| Expression::int(0));
                        let init_expr = go_array_element_type(&type_name)
                            .map(|elem| go_zero_value_expr(&elem))
                            .unwrap_or_else(Expression::null);
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(go_array_make_expr(len_expr, init_expr)),
                            type_name,
                        });
                    }
                    if go_is_map_type(&type_name) {
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(Expression::new(ExprKind::Object(Vec::new()))),
                            type_name,
                        });
                    }
                }
            }

            if call_name.as_deref() == Some("strconv.Atoi") && next_args.len() == 1 {
                return Expression::new(ExprKind::Tuple(vec![
                    go_builtin_call("__go_to_int", vec![next_args[0].value.clone()]),
                    Expression::null(),
                ]));
            }

            if next_args.len() == 1 {
                if let Some(type_name) = call_name
                    .as_deref()
                    .filter(|name| go_is_type_conversion_target(name, env, signatures))
                {
                    return go_normalize_type_conversion(
                        type_name,
                        next_args[0].value.clone(),
                        env,
                        signatures,
                    );
                }
            }

            if call_name.as_deref() == Some("copy") && next_args.len() >= 2 {
                let target = next_args[0].value.clone();
                let clone = go_member_call(next_args[1].value.clone(), "slice", Vec::new());
                let value = if let Some(type_name) = go_expr_type_hint(&target, env, signatures) {
                    Expression::new(ExprKind::Cast {
                        expr: Box::new(clone),
                        type_name,
                    })
                } else {
                    clone
                };
                return Expression::new(ExprKind::Assign {
                    target: Box::new(target),
                    value: Box::new(value),
                });
            }

            if call_name.as_deref() == Some("append") && !next_args.is_empty() {
                let mut result = next_args[0].value.clone();
                for arg in next_args.iter().skip(1) {
                    let rhs = if arg.spread {
                        arg.value.clone()
                    } else {
                        Expression::new(ExprKind::Array(vec![ArrayElement {
                            key: None,
                            value: arg.value.clone(),
                            spread: false,
                            by_ref: false,
                        }]))
                    };
                    result = go_member_call(result, "concat", vec![rhs]);
                }
                return result;
            }

            if call_name.as_deref() == Some("cap") && next_args.len() == 1 {
                if let Some(cap_expr) = go_expr_capacity_hint(&next_args[0].value, env) {
                    return cap_expr;
                }
            }

            if call_name.as_deref() == Some("strings.Replace")
                && next_args.len() == 4
                && go_is_neg_one_expr(&next_args[3].value)
            {
                next_args.pop();
            }

            if call_name.as_deref() == Some("strings.Fields") && next_args.len() == 1 {
                let trimmed = go_member_call(next_args[0].value.clone(), "trim", Vec::new());
                return go_builtin_call(
                    "__go_regex_split_pat_first",
                    vec![Expression::string("\\s+"), trimmed],
                );
            }

            if call_name.as_deref() == Some("close") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(next_args[0].value.clone()),
                            field: "closed".to_string(),
                            null_safe: false,
                        })),
                        value: Box::new(Expression::bool(true)),
                    });
                }
            }

            if call_name.as_deref() == Some("__go_type_assert") && next_args.len() == 2 {
                if let Some(type_name) = go_type_name_from_expr(&next_args[1].value) {
                    return go_type_assert_value_expr(next_args[0].value.clone(), &type_name);
                }
            }

            Expression::new(ExprKind::Call {
                callee: Box::new(next_callee),
                args: next_args,
                optional: *optional,
            })
        }
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|key| normalize_go_expr(key, env, signatures, state)),
                    value: normalize_go_expr(&element.value, env, signatures, state),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        )),
        ExprKind::Object(props) => Expression::new(ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    ObjectProperty::Spread(value) => {
                        ObjectProperty::Spread(normalize_go_expr(value, env, signatures, state))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    _ => prop.clone(),
                })
                .collect(),
        )),
        ExprKind::Cast { expr, type_name } => Expression::new(ExprKind::Cast {
            expr: Box::new(normalize_go_expr(expr, env, signatures, state)),
            type_name: type_name.clone(),
        }),
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .iter()
                .map(|value| normalize_go_expr(value, env, signatures, state))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            let mut lambda_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                slice_caps: env.slice_caps.clone(),
                type_names: env.type_names.clone(),
                return_type: None,
            };
            for param in params {
                if let Some(type_hint) = param.type_hint.as_ref() {
                    lambda_env.value_types.insert(param.name.clone(), type_hint.clone());
                }
                if let Some(type_hint) = param.type_hint.as_deref().filter(|hint| go_is_fixed_array_type(hint)) {
                    lambda_env.fixed_arrays.insert(param.name.clone(), type_hint.to_string());
                }
            }
            let next_body = match body {
                LambdaBody::Expr(expr) => {
                    LambdaBody::Expr(Box::new(normalize_go_expr(expr, &lambda_env, signatures, state)))
                }
                LambdaBody::Block(stmts) => {
                    LambdaBody::Block(normalize_go_function_body(stmts, &mut lambda_env, signatures, state))
                }
            };
            Expression::new(ExprKind::Lambda {
                params: params.clone(),
                body: next_body,
                is_async: *is_async,
                captures: captures.clone(),
            })
        }
        _ => expr.clone(),
    }
}

fn lower_go_fixed_array_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_type = go_expr_type_hint(&iter, env, signatures).unwrap_or_default();
    let iter_name = fresh_go_temp(state, "__go_range_iter");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: (!iter_type.is_empty()).then(|| iter_type.clone()),
            init: Some(go_wrap_fixed_array_copy(iter, env, signatures)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    if !iter_type.is_empty() {
        body_env.fixed_arrays.insert(iter_name.clone(), iter_type.clone());
    }

    let mut lowered_body = Vec::new();
    match key {
        Some(key_name) => {
            if key_name != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(key_name.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: go_array_element_type(&iter_type),
                            init: Some(Expression::new(ExprKind::Index {
                                object: Box::new(Expression::ident(&iter_name)),
                                index: Box::new(Expression::ident(&index_name)),
                                null_safe: false,
                            })),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
        None => {
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(stmt, &mut body_env, signatures, state));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(go_builtin_call("len", vec![Expression::ident(&iter_name)])),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![iter_decl, for_stmt]))]
}

fn fresh_go_temp(state: &mut GoNormalizeState, prefix: &str) -> String {
    let name = format!("{}{}", prefix, state.next_temp);
    state.next_temp += 1;
    name
}

fn go_wrap_fixed_array_copy(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_fixed_array(&expr, env, signatures) && go_requires_fixed_array_copy(&expr) {
        go_builtin_call("__go_fixed_array_clone", vec![expr])
    } else {
        expr
    }
}

fn go_requires_fixed_array_copy(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. })
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
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

fn go_expr_to_place(expr: &Expression) -> Option<PlaceExpr> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(PlaceExpr::Ident(name.clone())),
        ExprKind::Member { object, field, null_safe } => Some(PlaceExpr::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index { object, index, null_safe } => Some(PlaceExpr::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::RefLoad(expr) => Some(PlaceExpr::Deref(expr.clone())),
        ExprKind::Unary { op: UnaryOp::Deref, expr } => Some(PlaceExpr::Deref(expr.clone())),
        _ => None,
    }
}

fn normalize_go_lvalue_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Index { object, index, null_safe } => Expression::new(ExprKind::Index {
            object: Box::new(normalize_go_expr(object, env, signatures, state)),
            index: Box::new(normalize_go_expr(index, env, signatures, state)),
            null_safe: *null_safe,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        _ => normalize_go_expr(expr, env, signatures, state),
    }
}

fn go_is_two_value_binding_pattern(pattern: &BindingPattern) -> bool {
    matches!(pattern, BindingPattern::Array(elems) if elems.len() == 2)
}

fn go_normalize_map_lookup_tuple_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &expr.kind else {
        return None;
    };
    let value_type = go_map_index_value_type(expr, env, signatures)?;
    let next_object = normalize_go_expr(object, env, signatures, state);
    let next_index = normalize_go_expr(index, env, signatures, state);
    Some(Expression::new(ExprKind::Tuple(vec![
        go_build_map_read_expr(next_object.clone(), next_index.clone(), &value_type),
        go_map_has_expr(next_object, next_index),
    ])))
}

fn go_map_has_expr(object: Expression, index: Expression) -> Expression {
    go_builtin_call("__go_map_has", vec![object, index])
}

fn go_map_index_value_type(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let ExprKind::Index { object, .. } = &expr.kind else {
        return None;
    };
    go_expr_type_hint(object, env, signatures).and_then(|type_name| go_map_value_type(&type_name))
}

fn go_build_map_read_expr(object: Expression, index: Expression, value_type: &str) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_map_has_expr(object.clone(), index.clone())),
        then: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        })),
        else_: Box::new(go_zero_value_expr(value_type)),
    })
}

fn go_member_call(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: field.to_string(),
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

fn go_expr_is_fixed_array(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    go_expr_type_hint(expr, env, signatures)
        .as_deref()
        .is_some_and(go_is_fixed_array_type)
}

fn go_expr_type_hint(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => env
            .value_types
            .get(name)
            .cloned()
            .or_else(|| env.fixed_arrays.get(name).cloned()),
        ExprKind::Lit(Literal::Int(_)) => Some("int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("float64".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("bool".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
        ExprKind::RefOf(place) => {
            let pointee_type = match place.as_ref() {
                PlaceExpr::Ident(name) => env
                    .value_types
                    .get(name)
                    .cloned()
                    .or_else(|| env.fixed_arrays.get(name).cloned()),
                PlaceExpr::Member { object, field, null_safe } => go_expr_type_hint(
                    &Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Index { object, index, null_safe } => go_expr_type_hint(
                    &Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Deref(expr) => go_expr_type_hint(expr, env, signatures).map(|type_name| {
                    type_name
                        .trim()
                        .trim_start_matches('*')
                        .trim_start_matches('^')
                        .trim()
                        .to_string()
                }),
            }?;
            Some(format!("*{}", pointee_type.trim()))
        }
        ExprKind::Unary { op: UnaryOp::AddrOf, expr } => go_expr_type_hint(expr, env, signatures)
            .map(|type_name| format!("*{}", type_name.trim())),
        ExprKind::Unary { op: UnaryOp::Deref, expr } | ExprKind::RefLoad(expr) => go_expr_type_hint(expr, env, signatures)
            .map(|type_name| {
                type_name
                    .trim()
                    .trim_start_matches('*')
                    .trim_start_matches('^')
                    .trim()
                    .to_string()
            }),
        ExprKind::IsType { .. } => Some("bool".to_string()),
        ExprKind::Index { object, .. } => go_expr_type_hint(object, env, signatures)
            .and_then(|type_name| {
                if type_name == "string" {
                    Some("byte".to_string())
                } else {
                    go_array_element_type(&type_name).or_else(|| go_map_value_type(&type_name))
                }
            }),
        ExprKind::Assign { value, .. } => go_expr_type_hint(value, env, signatures),
        ExprKind::Ternary { then, else_, .. } => {
            let then_type = go_expr_type_hint(then, env, signatures);
            let else_type = go_expr_type_hint(else_, env, signatures);
            if then_type == else_type { then_type } else { then_type.or(else_type) }
        }
        ExprKind::Binary { op, left, right } => {
            let left_type = go_expr_type_hint(left, env, signatures);
            let right_type = go_expr_type_hint(right, env, signatures);
            match op {
                BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::IDiv | BinOp::Mod
                | BinOp::BitAnd | BinOp::BitOr | BinOp::BitXor | BinOp::Shl | BinOp::Shr => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        left_type.or(right_type)
                    }
                }
                BinOp::Div => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        Some("float64".to_string())
                    }
                }
                _ => None,
            }
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) if name == "__go_fixed_array_clone" => {
                args.first().and_then(|arg| go_expr_type_hint(&arg.value, env, signatures))
            }
            ExprKind::Ident(name) if name == "__go_fixed_array_equal" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_regex_split_pat_first" => Some("[]string".to_string()),
            ExprKind::Ident(name) if name == "__go_map_has" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_to_int" => Some("int".to_string()),
            ExprKind::Ident(name) if name == "__go_str_from_char_code" => Some("string".to_string()),
            ExprKind::Ident(name) if name == "__go_type_assert" => {
                args.get(1).and_then(|arg| go_type_name_from_expr(&arg.value))
            }
            ExprKind::Member { field, .. } if field == "charCodeAt" => Some("int".to_string()),
            ExprKind::Ident(name) => signatures.get(name).and_then(|sig| sig.return_type.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn go_expr_call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let object_name = go_expr_call_name(object)?;
            Some(format!("{}.{}", object_name, field))
        }
        _ => None,
    }
}

fn go_is_neg_one_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(-1)) => true,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => matches!(expr.kind, ExprKind::Lit(Literal::Int(1))),
        _ => false,
    }
}

fn go_decl_fixed_array_binding(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };
    let type_name = decl
        .type_hint
        .clone()
        .or_else(|| decl.init.as_ref().and_then(|expr| go_expr_type_hint(expr, env, signatures)))?;
    go_is_fixed_array_type(&type_name).then(|| (name.clone(), type_name))
}

fn go_decl_binding_type(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };

    decl.type_hint
        .clone()
        .or_else(|| decl.init.as_ref().and_then(|expr| go_expr_type_hint(expr, env, signatures)))
        .map(|type_name| (name.clone(), type_name))
}

fn go_single_named_binding_pattern(pattern: &BindingPattern) -> Option<BindingPattern> {
    let BindingPattern::Array(elements) = pattern else {
        return None;
    };

    if elements.len() != 1 {
        return None;
    }

    let mut bound_name = None;
    for element in elements {
        match element {
            ArrayPatternElem::Hole => return None,
            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) => {
                if bound_name.is_some() {
                    return None;
                }
                bound_name = Some(name.clone());
            }
            _ => return None,
        }
    }

    bound_name.map(BindingPattern::Ident)
}

fn go_is_fixed_array_type(type_name: &str) -> bool {
    go_array_head(type_name)
        .map(|(head, _)| !head.trim().is_empty())
        .unwrap_or(false)
}

fn go_is_integer_type(type_name: &str) -> bool {
    matches!(
        type_name.trim(),
        "int" | "int8" | "int16" | "int32" | "int64"
            | "uint" | "uint8" | "uint16" | "uint32" | "uint64" | "uintptr"
            | "byte" | "rune"
    )
}

fn go_is_float_type(type_name: &str) -> bool {
    matches!(type_name.trim(), "float32" | "float64")
}

fn go_is_builtin_conversion_type(type_name: &str) -> bool {
    go_is_integer_type(type_name)
        || go_is_float_type(type_name)
        || matches!(type_name.trim(), "string" | "bool")
}

fn go_is_type_conversion_target(
    type_name: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    (go_is_builtin_conversion_type(type_name) || env.type_names.contains(type_name))
        && !signatures.contains_key(type_name)
}

fn go_normalize_type_conversion(
    type_name: &str,
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_is_integer_type(type_name) {
        let int_expr = go_builtin_call("__go_to_int", vec![expr]);
        if type_name == "int" {
            return int_expr;
        }
        return Expression::new(ExprKind::Cast {
            expr: Box::new(int_expr),
            type_name: type_name.to_string(),
        });
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(go_is_integer_type)
    {
        return go_builtin_call("__go_str_from_char_code", vec![expr]);
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn walk_package_clause(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let mut path = String::new();
    let mut alias: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::import_spec => {
                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => {
                            alias = Some(spec_inner.as_str().to_string());
                        }
                        Rule::string_literal => {
                            path = unquote(spec_inner.as_str());
                        }
                        _ => {}
                    }
                }
            }
            Rule::string_literal => {
                path = unquote(inner.as_str());
            }
            _ => {}
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span: Span::default(),
    })
}

fn unquote(s: &str) -> String {
    if s.len() < 2 {
        return s.to_string();
    }

    if s.starts_with('`') && s.ends_with('`') {
        return s[1..s.len() - 1].to_string();
    }

    if !((s.starts_with('"') && s.ends_with('"')) || (s.starts_with('\'') && s.ends_with('\''))) {
        return s.to_string();
    }

    let mut out = String::new();
    let mut chars = s[1..s.len() - 1].chars();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }

        match chars.next() {
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some('t') => out.push('\t'),
            Some('\\') => out.push('\\'),
            Some('"') => out.push('"'),
            Some('\'') => out.push('\''),
            Some(other) => out.push(other),
            None => out.push('\\'),
        }
    }

    out
}

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::function_declaration => Ok(Some(walk_function_decl(pair)?)),
        Rule::method_declaration => Ok(Some(walk_method_decl(pair)?)),
        Rule::var_declaration => Ok(Some(walk_var_decl(pair)?)),
        Rule::const_declaration => Ok(Some(walk_const_decl(pair)?)),
        Rule::type_declaration => walk_type_decl(pair),
        Rule::declaration => {
            for inner in pair.into_inner() {
                return walk_top_level(inner);
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

// ── Function declarations ─────────────────────────────────────────────────────────────

fn walk_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    for param in named_results.iter().rev() {
        body_stmts.insert(0, go_named_result_marker_stmt(&param.name, param.type_hint.as_deref().unwrap_or("object")));
    }

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_method_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut receiver_name = String::new();
    let mut receiver_type = String::new();
    let mut receiver_owner = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::receiver => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::ident_name => receiver_name = r_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            receiver_type = walk_type(r_inner.clone());
                            receiver_owner = go_named_receiver_type(&receiver_type)
                                .unwrap_or_else(|| receiver_type.clone());
                        }
                        _ => {}
                    }
                }
            }
            Rule::ident_name => method_name = inner.as_str().to_string(),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    for param in named_results.iter().rev() {
        body_stmts.insert(0, go_named_result_marker_stmt(&param.name, param.type_hint.as_deref().unwrap_or("object")));
    }

    // Prepend receiver as first parameter
    params.insert(0, Param {
        name: if receiver_name.is_empty() { "self".to_string() } else { receiver_name },
        type_hint: Some(receiver_type.clone()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    });

    let method_stmt = Statement::new(StmtKind::FunctionDecl {
        name: method_name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    });

    Ok(Statement::new(StmtKind::StructDecl {
        name: receiver_owner,
        interfaces: Vec::new(),
        members: vec![ClassMember::Method(Box::new(method_stmt))],
        visibility: Visibility::Public,
        decorators: Vec::new(),
    }))
}

fn walk_signature(pair: Pair<Rule>) -> Result<GoSignatureInfo, String> {
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::parameter_list => {
                params = walk_parameter_list(inner)?;
            }
            Rule::result => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::type_annotation => return_type = Some(walk_type(r_inner)),
                        Rule::parameter_list => {
                            let p = walk_parameter_list(r_inner)?;
                            named_results = p
                                .iter()
                                .filter(|param| !param.name.starts_with("__go_param_"))
                                .cloned()
                                .collect();
                            return_type = if p.len() == 1 {
                                p[0].type_hint.clone()
                            } else {
                                Some(format!("[{}]", p.len()))
                            };
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_annotation => {
                return_type = Some(walk_type(inner));
            }
            _ => {}
        }
    }

    Ok(GoSignatureInfo {
        params,
        return_type,
        named_results,
    })
}

fn go_named_result_marker_stmt(name: &str, type_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(go_builtin_call(
        "__go_named_result",
        vec![Expression::string(name), go_type_arg_expr(type_name.to_string())],
    )))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter_decl {
            let mut names = Vec::new();
            let mut type_hint: Option<String> = None;
            let mut is_rest = false;

            for p_inner in inner.into_inner() {
                match p_inner.as_rule() {
                    Rule::ident_name => names.push(p_inner.as_str().to_string()),
                    Rule::ident_list => {
                        for id in p_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::type_annotation => type_hint = Some(walk_type(p_inner)),
                    Rule::variadic_parameter_type => {
                        is_rest = true;
                        for v_inner in p_inner.into_inner() {
                            if v_inner.as_rule() == Rule::type_annotation {
                                type_hint = Some(walk_type(v_inner));
                            }
                        }
                    }
                    _ => {}
                }
            }

            if names.is_empty() && type_hint.is_some() {
                names.push(format!("__go_param_{}", params.len()));
            }

            for name in names {
                params.push(Param {
                    name,
                    type_hint: type_hint.clone(),
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                });
            }
        }
    }
    Ok(params)
}

fn walk_type(pair: Pair<Rule>) -> String {
    pair.as_str().to_string()
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block_statement | Rule::function_body => {
                stmts.append(&mut walk_block(inner)?);
            }
            Rule::statement_list => {
                for s in inner.into_inner() {
                    if s.as_rule() == Rule::statement {
                        stmts.push(walk_statement(s)?);
                    }
                }
            }
            Rule::statement => {
                stmts.push(walk_statement(inner)?);
            }
            _ => {}
        }
    }
    Ok(stmts)
}

// ── Variable declarations ─────────────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::var_spec | Rule::const_spec => {
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Let)?;
                declarations.append(&mut decls);
            }
            Rule::var_group | Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::var_spec || spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Let)?;
                        declarations.append(&mut decls);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    }))
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::const_spec => {
                let (mut decls, _, _) = walk_const_spec(inner, 0, None, None)?;
                declarations.append(&mut decls);
            }
            Rule::const_group => {
                let mut prev_inits: Option<Vec<Expression>> = None;
                let mut prev_type_hint: Option<String> = None;
                let mut iota_index = 0i64;
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::const_spec {
                        let (mut decls, next_inits, next_type_hint) =
                            walk_const_spec(spec, iota_index, prev_inits.clone(), prev_type_hint.clone())?;
                        declarations.append(&mut decls);
                        prev_inits = Some(next_inits);
                        prev_type_hint = next_type_hint;
                        iota_index += 1;
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const,
    }))
}

fn walk_const_spec(
    pair: Pair<Rule>,
    iota_index: i64,
    prev_inits: Option<Vec<Expression>>,
    prev_type_hint: Option<String>,
) -> Result<(Vec<VarDeclarator>, Vec<Expression>, Option<String>), String> {
    let (names, type_hint, init_values) = parse_go_var_spec(pair)?;
    let effective_type_hint = type_hint.or(prev_type_hint);
    let raw_inits = if init_values.is_empty() {
        prev_inits.unwrap_or_default()
    } else {
        init_values
    };
    let next_inits: Vec<Expression> = raw_inits
        .iter()
        .map(|expr| go_rewrite_iota_expr(expr, iota_index))
        .collect();

    if names.len() > 1 && !next_inits.is_empty() {
        let pattern = BindingPattern::Array(
            names.into_iter().map(|name| {
                if name == "_" {
                    ArrayPatternElem::Hole
                } else {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                }
            }).collect(),
        );
        let init = if next_inits.len() == 1 {
            next_inits[0].clone()
        } else {
            Expression::new(ExprKind::Tuple(next_inits.clone()))
        };
        return Ok((vec![VarDeclarator {
            pattern,
            init: Some(init),
            type_hint: effective_type_hint.clone(),
            array_bounds: None,
            with_events: false,
        }], raw_inits, effective_type_hint));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: next_inits.first().cloned(),
            type_hint: effective_type_hint.clone(),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, raw_inits, effective_type_hint))
}

fn parse_go_var_spec(pair: Pair<Rule>) -> Result<(Vec<String>, Option<String>, Vec<Expression>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) || id.as_str() == "_" {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => init_values = walk_expression_list(inner)?,
            Rule::expression => init_values.push(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok((names, type_hint, init_values))
}

fn go_rewrite_iota_expr(expr: &Expression, iota_index: i64) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if name == "iota" => Expression::int(iota_index),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(go_rewrite_iota_expr(left, iota_index)),
            right: Box::new(go_rewrite_iota_expr(right, iota_index)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_iota_expr(cond, iota_index)),
            then: Box::new(go_rewrite_iota_expr(then, iota_index)),
            else_: Box::new(go_rewrite_iota_expr(else_, iota_index)),
        }),
        ExprKind::Call { callee, args, optional } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_iota_expr(callee, iota_index)),
            args: args.iter().map(|arg| Argument {
                value: go_rewrite_iota_expr(&arg.value, iota_index),
                name: arg.name.clone(),
                by_ref: arg.by_ref,
                spread: arg.spread,
            }).collect(),
            optional: *optional,
        }),
        ExprKind::Member { object, field, null_safe } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index { object, index, null_safe } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            index: Box::new(go_rewrite_iota_expr(index, iota_index)),
            null_safe: *null_safe,
        }),
        ExprKind::Cast { expr, type_name } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
            type_name: type_name.clone(),
        }),
        _ => expr.clone(),
    }
}

fn walk_var_spec(pair: Pair<Rule>, _kind: VarDeclKind) -> Result<(Vec<VarDeclarator>, Option<String>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) || id.as_str() == "_" {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => {
                init_values = walk_expression_list(inner)?;
            }
            Rule::expression => {
                init_values.push(walk_expression(inner)?);
            }
            _ => {}
        }
    }

    if names.len() > 1 && !init_values.is_empty() {
        let pattern = BindingPattern::Array(
            names.into_iter().map(|name| {
                if name == "_" {
                    ArrayPatternElem::Hole
                } else {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                }
            }).collect(),
        );
        let init = if init_values.len() == 1 {
            init_values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(init_values))
        };
        return Ok((vec![VarDeclarator {
            pattern,
            init: Some(init),
            type_hint,
            array_bounds: None,
            with_events: false,
        }], None));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: init_values.first().cloned(),
            type_hint: type_hint.clone(),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, type_hint))
}

// ── Type declarations (struct, interface, type alias) ─────────────────────────────────

fn walk_type_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_spec => {
                let mut name = String::new();
                let mut type_str = String::new();

                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => name = spec_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            if let Some(type_stmt) = walk_named_type_annotation(name.clone(), spec_inner.clone())? {
                                return Ok(Some(type_stmt));
                            }
                            type_str = walk_type(spec_inner);
                        }
                        Rule::struct_type => {
                            return Ok(Some(walk_struct_type(name, spec_inner)?));
                        }
                        Rule::interface_type => {
                            return Ok(Some(walk_interface_type(name, spec_inner)?));
                        }
                        _ => {}
                    }
                }

                // Type alias — just create a variable with the type name
                if !type_str.is_empty() && !name.is_empty() {
                    return Ok(Some(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(name),
                            init: Some(Expression::new(ExprKind::Lit(Literal::Str(type_str)))),
                            type_hint: None,
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    })));
                }
            }
            Rule::type_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::type_spec {
                        return walk_type_decl(spec.into());
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_struct_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::field_decl {
            let mut field_names = Vec::new();
            let mut field_type: Option<String> = None;

            for f_inner in inner.into_inner() {
                match f_inner.as_rule() {
                    Rule::ident_list => {
                        for id in f_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                field_names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::ident_name => field_names.push(f_inner.as_str().to_string()),
                    Rule::type_annotation => field_type = Some(walk_type(f_inner)),
                    _ => {}
                }
            }

            if field_names.is_empty() {
                if let Some(type_name) = field_type.as_deref().and_then(go_embedded_field_name) {
                    field_names.push(type_name);
                }
            }

            for fname in field_names {
                members.push(ClassMember::Field {
                    name: fname,
                    type_hint: field_type.clone(),
                    init: None,
                    modifiers: Modifiers::default(),
                    with_events: false,
                    array_bounds: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::StructDecl {
        name,
        interfaces: Vec::new(),
        members,
        visibility: Visibility::Public,
        decorators: Vec::new(),
    }))
}

fn walk_interface_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::interface_member {
            let mut method_name = String::new();
            let mut params = Vec::new();
            let mut return_type: Option<String> = None;

            for m_inner in inner.into_inner() {
                match m_inner.as_rule() {
                    Rule::ident_name => method_name = m_inner.as_str().to_string(),
                    Rule::signature => {
                        let sig = walk_signature(m_inner)?;
                        params = sig.params;
                        return_type = sig.return_type;
                    }
                    _ => {}
                }
            }

            if !method_name.is_empty() {
                members.push(InterfaceMember::Method {
                    name: method_name,
                    params,
                    return_type,
                    is_sub: false,
                    signature_source: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::InterfaceDecl {
        name,
        parents: Vec::new(),
        members,
        decorators: Vec::new(),
    }))
}

fn walk_named_type_annotation(name: String, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::struct_type => return Ok(Some(walk_struct_type(name, inner)?)),
            Rule::interface_type => return Ok(Some(walk_interface_type(name, inner)?)),
            _ => {}
        }
    }
    Ok(None)
}

fn go_embedded_field_name(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.rsplit('.').next().map(|name| name.to_string())
}

fn go_named_receiver_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    if trimmed.is_empty() {
        return None;
    }
    Some(trimmed.to_string())
}

// ── Statements ─────────────────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    if rule == Rule::statement {
        if let Some(inner) = pair.into_inner().next() {
            let mut s = walk_statement(inner)?;
            if s.span.start_line == 0 { s.span = span; }
            return Ok(s);
        }
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => StmtKind::Block(walk_block(pair)?),
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::assignment_statement => walk_assignment(pair)?,
        Rule::short_var_declaration => walk_short_var_decl(pair)?,
        Rule::inc_dec_statement => walk_inc_dec(pair)?,
        Rule::var_declaration => walk_var_decl(pair)?.kind,
        Rule::const_declaration => walk_const_decl(pair)?.kind,
        Rule::if_statement => walk_if(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::select_statement => walk_select(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::goto_statement => StmtKind::GoTo(walk_goto(pair)?),
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::defer_statement => walk_defer_stmt(pair)?,
        Rule::go_statement => walk_go_stmt(pair)?,
        Rule::send_statement => walk_send_stmt(pair)?,
        _ => StmtKind::Empty,
    };
    Ok(Statement::with_span(kind, span))
}

fn walk_defer_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    Ok(StmtKind::Expr(go_builtin_call(
        "__go_defer",
        vec![walk_expression(first_meaningful(pair)?)?],
    )))
}

fn walk_go_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|inner| inner.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?
        .unwrap_or_else(Expression::null);

    Ok(StmtKind::Expr(go_builtin_call(
        "__go_spawn",
        vec![go_wrap_spawn_expr(expr)],
    )))
}

fn walk_send_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression {
            exprs.push(walk_expression(inner)?);
        }
    }

    if exprs.len() == 2 {
        Ok(StmtKind::Expr(go_channel_send_expr(
            exprs.remove(0),
            exprs.remove(0),
        )))
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut op = "=";
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression_list => {
                if targets.is_empty() {
                    targets = walk_expression_list(inner)?;
                } else {
                    values = walk_expression_list(inner)?;
                }
            }
            Rule::assign_op => op = inner.as_str(),
            _ => {}
        }
    }

    if op != "=" {
        // Compound assignment
        if targets.len() == 1 && values.len() == 1 {
            let target = targets[0].clone();
            let value = values[0].clone();
            let compound_op = match op {
                "+=" => Some(CompoundOp::Add),
                "-=" => Some(CompoundOp::Sub),
                "*=" => Some(CompoundOp::Mul),
                "/=" => Some(CompoundOp::Div),
                "%=" => Some(CompoundOp::Mod),
                "&=" => Some(CompoundOp::BitAnd),
                "|=" => Some(CompoundOp::BitOr),
                "^=" => Some(CompoundOp::BitXor),
                "<<=" => Some(CompoundOp::Shl),
                ">>=" => Some(CompoundOp::Shr),
                _ => None,
            };
            if let Some(compound_op) = compound_op {
                return Ok(StmtKind::CompoundAssign {
                    target,
                    op: compound_op,
                    value,
                });
            }
            if op == "&^=" {
                let rhs = Expression::new(ExprKind::Unary {
                    op: UnaryOp::BitNot,
                    expr: Box::new(value),
                });
                return Ok(StmtKind::Assign {
                    targets: vec![target.clone()],
                    value: Expression::new(ExprKind::Binary {
                        op: BinOp::BitAnd,
                        left: Box::new(target),
                        right: Box::new(rhs),
                    }),
                });
            }
        }
    }

    if targets.len() > 1 {
        let patterns = targets
            .iter()
            .map(|target| match &target.kind {
                ExprKind::Ident(name) => {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                }
                _ => ArrayPatternElem::Hole,
            })
            .collect();
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        return Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Destructure(
                DestructurePattern::Array(patterns),
            ))],
            value,
        });
    }

    if values.len() == 1 {
        Ok(StmtKind::Assign {
            targets,
            value: values.into_iter().next().unwrap(),
        })
    } else if !values.is_empty() {
        Ok(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Tuple(values)),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_short_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::expression_list => {
                values = walk_expression_list(inner)?;
            }
            _ => {}
        }
    }

    let mut declarations = Vec::new();
    if names.len() == 2 && values.len() == 1 {
        if let Some((expr, type_name)) = go_extract_type_assert_expr(&values[0]) {
            let pattern = BindingPattern::Array(
                names.into_iter().map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                }).collect(),
            );
            declarations.push(VarDeclarator {
                pattern,
                init: Some(Expression::new(ExprKind::Tuple(vec![
                    go_type_assert_value_expr(expr.clone(), &type_name),
                    Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name,
                    }),
                ]))),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            });
            return Ok(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let,
            });
        }
    }

    if names.len() > 1 && !values.is_empty() {
        let pattern = BindingPattern::Array(
            names.into_iter().map(|name| {
                if name == "_" {
                    ArrayPatternElem::Hole
                } else {
                    ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                }
            }).collect(),
        );
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        declarations.push(VarDeclarator {
            pattern,
            init: Some(value),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        });
    } else {
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else if !values.is_empty() {
            Expression::new(ExprKind::Tuple(values))
        } else {
            Expression::new(ExprKind::Lit(Literal::Null))
        };

        for name in names {
            if name == "_" {
                continue;
            }
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            });
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_inc_dec(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let is_inc = !pair
        .as_str()
        .trim_end()
        .trim_end_matches(';')
        .trim_end()
        .ends_with("--");

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::primary => expr = Some(walk_primary(inner)?),
            _ => {}
        }
    }

    if let Some(target) = expr {
        Ok(StmtKind::CompoundAssign {
            target,
            op: if is_inc { CompoundOp::Add } else { CompoundOp::Sub },
            value: Expression::new(ExprKind::Lit(Literal::Int(1))),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(inner)?);
                }
            }
            Rule::block_statement => {
                if then_body.is_empty() {
                    then_body = walk_block(inner)?;
                }
            }
            Rule::else_clause => {
                for e_inner in inner.into_inner() {
                    match e_inner.as_rule() {
                        Rule::block_statement => else_body = Some(walk_block(e_inner)?),
                        Rule::if_statement => {
                            let elif = walk_if(e_inner)?;
                            match elif {
                                StmtKind::If {
                                    cond: c,
                                    then_body: t,
                                    elifs: nested_elifs,
                                    else_body: nested_else,
                                } => {
                                    elifs.push((c, t));
                                    elifs.extend(nested_elifs);
                                    else_body = nested_else;
                                }
                                StmtKind::Block(stmts) => else_body = Some(stmts),
                                _ => {}
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            _ => {}
        }
    }

    let if_stmt = Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body,
        elifs,
        else_body,
    });

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![*pre, if_stmt]))
    } else {
        Ok(if_stmt.kind)
    }
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    match pair.as_rule() {
        Rule::switch_statement => {
            if let Some(inner) = pair.into_inner().next() {
                return walk_switch(inner);
            }
            return Ok(StmtKind::Empty);
        }
        Rule::type_switch_stmt => return walk_type_switch(pair),
        Rule::expr_switch_stmt => {}
        _ => return Ok(StmtKind::Empty),
    }

    let mut expr = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::expr_case_clause => {
                let mut conditions: Vec<CaseCondition> = Vec::new();
                let mut body = Vec::new();

                for c_inner in inner.into_inner() {
                    match c_inner.as_rule() {
                        Rule::expr_switch_case => {
                            for sc_inner in c_inner.into_inner() {
                                if sc_inner.as_rule() == Rule::expression_list {
                                    for expr in walk_expression_list(sc_inner)? {
                                        conditions.push(CaseCondition::Value(expr));
                                    }
                                } else if sc_inner.as_rule() == Rule::kw_default {
                                    // default case
                                }
                            }
                        }
                        Rule::statement_list => {
                            body = walk_statement_list(c_inner)?;
                        }
                        _ => {}
                    }
                }

                if conditions.is_empty() {
                    default = Some(body);
                } else {
                    cases.push(SwitchCase { conditions, body });
                }
            }
            _ => {}
        }
    }

    if expr.is_none() {
        let mut first_case: Option<(Expression, Vec<Statement>)> = None;
        let mut elifs = Vec::new();
        for case in cases {
            let cond = case
                .conditions
                .into_iter()
                .filter_map(|condition| match condition {
                    CaseCondition::Value(expr) => Some(expr),
                    _ => None,
                })
                .reduce(|left, right| Expression::new(ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(left),
                    right: Box::new(right),
                }))
                .unwrap_or_else(|| Expression::bool(false));
            if first_case.is_none() {
                first_case = Some((cond, case.body));
            } else {
                elifs.push((cond, case.body));
            }
        }

        if let Some((cond, then_body)) = first_case {
            return Ok(StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body: default,
            });
        }

        return Ok(StmtKind::Block(default.unwrap_or_default()));
    }

    Ok(StmtKind::Switch {
        expr: expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        cases,
        default,
    })
}

fn walk_type_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut binding_name: Option<String> = None;
    let mut switch_expr: Option<Expression> = None;
    let mut first_case: Option<(Expression, Vec<Statement>)> = None;
    let mut elifs = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_switch_guard => {
                for guard_inner in inner.into_inner() {
                    match guard_inner.as_rule() {
                        Rule::ident_name => binding_name = Some(guard_inner.as_str().to_string()),
                        Rule::primary => switch_expr = Some(walk_primary(guard_inner)?),
                        _ => {}
                    }
                }
            }
            Rule::type_case_clause => {
                let mut case_types = Vec::new();
                let mut body = Vec::new();
                for case_inner in inner.into_inner() {
                    match case_inner.as_rule() {
                        Rule::type_switch_case => {
                            for switch_case_inner in case_inner.into_inner() {
                                match switch_case_inner.as_rule() {
                                    Rule::type_list => {
                                        for ty in switch_case_inner.into_inner() {
                                            if ty.as_rule() == Rule::type_annotation {
                                                case_types.push(walk_type(ty));
                                            }
                                        }
                                    }
                                    Rule::kw_default => {}
                                    _ => {}
                                }
                            }
                        }
                        Rule::statement_list => body = walk_statement_list(case_inner)?,
                        _ => {}
                    }
                }

                if case_types.is_empty() {
                    default_body = Some(body);
                } else {
                    let expr = switch_expr.clone().unwrap_or_else(Expression::null);
                    let cond = go_type_switch_case_cond(expr.clone(), &case_types);
                    let case_body = go_type_switch_case_body(
                        body,
                        binding_name.as_deref(),
                        expr,
                        &case_types[0],
                    );
                    if first_case.is_none() {
                        first_case = Some((cond, case_body));
                    } else {
                        elifs.push((cond, case_body));
                    }
                }
            }
            _ => {}
        }
    }

    if let Some((cond, then_body)) = first_case {
        Ok(StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body: default_body,
        })
    } else {
        Ok(StmtKind::Block(default_body.unwrap_or_default()))
    }
}

fn walk_select(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_case_clause => body.extend(walk_select_case_clause(inner)?),
            Rule::select_default_clause => body.extend(walk_select_default_clause(inner)?),
            _ => {}
        }
    }

    Ok(StmtKind::Block(body))
}

fn walk_select_case_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut prefix = Vec::new();
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_comm_clause => prefix.extend(walk_select_comm_clause(inner)?),
            Rule::statement_list => body.extend(walk_statement_list(inner)?),
            _ => {}
        }
    }

    prefix.extend(body);
    Ok(prefix)
}

fn walk_select_default_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement_list {
            return walk_statement_list(inner);
        }
    }
    Ok(Vec::new())
}

fn walk_select_comm_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_send_clause => {
                let mut exprs = Vec::new();
                for part in inner.into_inner() {
                    if part.as_rule() == Rule::expression {
                        exprs.push(walk_expression(part)?);
                    }
                }
                if exprs.len() == 2 {
                    stmts.push(Statement::new(StmtKind::Expr(go_channel_send_expr(
                        exprs.remove(0),
                        exprs.remove(0),
                    ))));
                }
            }
            Rule::select_receive_clause => {
                let mut names = Vec::new();
                let mut recv_expr = None;

                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::ident_list => {
                            for id in part.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    names.push(id.as_str().to_string());
                                }
                            }
                        }
                        Rule::expression => recv_expr = Some(walk_expression(part)?),
                        _ => {}
                    }
                }

                if let Some(expr) = recv_expr {
                    if names.is_empty() {
                        stmts.push(Statement::new(StmtKind::Expr(expr)));
                    } else {
                        stmts.push(go_short_var_decl_from_parts(names, expr));
                    }
                }
            }
            _ => {}
        }
    }

    Ok(stmts)
}

fn go_short_var_decl_from_parts(names: Vec<String>, value: Expression) -> Statement {
    let declarations = if names.len() > 1 {
        vec![VarDeclarator {
            pattern: BindingPattern::Array(
                names
                    .into_iter()
                    .map(|name| {
                        if name == "_" {
                            ArrayPatternElem::Hole
                        } else {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                        }
                    })
                    .collect(),
            ),
            init: Some(Expression::new(ExprKind::Tuple(vec![value, Expression::bool(true)]))),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        }]
    } else {
        names
            .into_iter()
            .filter(|name| name != "_")
            .map(|name| VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            })
            .collect()
    };

    Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement {
            stmts.push(walk_statement(inner)?);
        }
    }
    Ok(stmts)
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body = Vec::new();
    let mut is_range = false;
    let mut range_vars = Vec::new();
    let mut range_iter = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::for_clause => {
                for fc_inner in inner.into_inner() {
                    match fc_inner.as_rule() {
                        Rule::for_short_var_nosemi => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::short_var_declaration => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::expression_statement => {
                            let expr = walk_expression(first_meaningful(fc_inner)?)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            } else if update.is_none() {
                                update = Some(expr);
                            }
                        }
                        Rule::assignment_statement => {
                            let assign = walk_assignment(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(assign)));
                            } else if update.is_none() {
                                if let StmtKind::Assign { targets, value } = assign {
                                    if let Some(target) = targets.into_iter().next() {
                                        update = Some(Expression::new(ExprKind::Assign {
                                            target: Box::new(target),
                                            value: Box::new(value),
                                        }));
                                    }
                                }
                            }
                        }
                        Rule::inc_dec_statement => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(inc_dec)));
                            } else if update.is_none() {
                                if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                    let bin_op = match op {
                                        CompoundOp::Add => BinOp::Add,
                                        CompoundOp::Sub => BinOp::Sub,
                                        CompoundOp::Mul => BinOp::Mul,
                                        CompoundOp::Div => BinOp::Div,
                                        CompoundOp::Mod => BinOp::Mod,
                                        _ => BinOp::Add,
                                    };
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target.clone()),
                                        value: Box::new(Expression::new(ExprKind::Binary {
                                            op: bin_op,
                                            left: Box::new(target),
                                            right: Box::new(value),
                                        })),
                                    }));
                                }
                            }
                        }
                        Rule::for_inc_dec => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                let bin_op = match op {
                                    CompoundOp::Add => BinOp::Add,
                                    CompoundOp::Sub => BinOp::Sub,
                                    CompoundOp::Mul => BinOp::Mul,
                                    CompoundOp::Div => BinOp::Div,
                                    CompoundOp::Mod => BinOp::Mod,
                                    _ => BinOp::Add,
                                };
                                update = Some(Expression::new(ExprKind::Assign {
                                    target: Box::new(target.clone()),
                                    value: Box::new(Expression::new(ExprKind::Binary {
                                        op: bin_op,
                                        left: Box::new(target),
                                        right: Box::new(value),
                                    })),
                                }));
                            }
                        }
                        Rule::for_assign_nosemi => {
                            let assign = walk_assignment(fc_inner)?;
                            if let StmtKind::Assign { targets, value } = assign {
                                if let Some(target) = targets.into_iter().next() {
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target),
                                        value: Box::new(value),
                                    }));
                                }
                            }
                        }
                        Rule::expression => {
                            if cond.is_none() {
                                cond = Some(walk_expression(fc_inner)?);
                            } else if update.is_none() {
                                update = Some(walk_expression(fc_inner)?);
                            }
                        }
                        Rule::block_statement => {
                            body = walk_block(fc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::range_clause => {
                is_range = true;
                for rc_inner in inner.into_inner() {
                    match rc_inner.as_rule() {
                        Rule::expression_list => {
                            for expr in walk_expression_list(rc_inner)? {
                                let name = if let ExprKind::Ident(id) = &expr.kind {
                                    id.clone()
                                } else {
                                    "_".to_string()
                                };
                                range_vars.push(BindingPattern::Ident(name));
                            }
                        }
                        Rule::ident_list => {
                            for id in rc_inner.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    range_vars.push(BindingPattern::Ident(id.as_str().to_string()));
                                }
                            }
                        }
                        Rule::expression => {
                            range_iter = Some(walk_expression(rc_inner)?);
                        }
                        Rule::block_statement => {
                            body = walk_block(rc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::expression => {
                cond = Some(walk_expression(inner)?);
            }
            Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    if is_range {
        let var = if range_vars.len() > 1 {
            range_vars.get(1).cloned().unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        } else {
            range_vars.get(0).cloned().unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        };
        let var_name = match var {
            BindingPattern::Ident(name) => name,
            _ => "_".to_string(),
        };
        let key = if range_vars.len() > 1 {
            let key_pat = range_vars.get(0).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else {
            None
        };

        Ok(StmtKind::ForIn {
            var: var_name,
            key,
            iter: range_iter.unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new()))),
            body,
            of: true,
            else_body: None,
            is_async: false,
        })
    } else if init.is_none() && update.is_none() && cond.is_some() {
        Ok(StmtKind::While {
            cond: cond.unwrap(),
            body,
            else_body: None,
        })
    } else {
        Ok(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression_list {
            values = walk_expression_list(inner)?;
        } else if inner.as_rule() == Rule::expression {
            values.push(walk_expression(inner)?);
        }
    }

    if values.len() == 1 {
        Ok(StmtKind::Return(Some(values.into_iter().next().unwrap())))
    } else if values.len() > 1 {
        let arr_elems: Vec<ArrayElement> = values.into_iter().map(|v| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(StmtKind::Return(Some(Expression::new(ExprKind::Array(arr_elems)))))
    } else {
        Ok(StmtKind::Return(None))
    }
}

fn walk_goto(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut label = String::new();
    let mut stmt = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => label = inner.as_str().to_string(),
            Rule::statement => stmt = Some(walk_statement(inner)?),
            _ => {}
        }
    }

    if let Some(s) = stmt {
        Ok(StmtKind::Block(vec![
            Statement::new(StmtKind::Label(label)),
            s,
        ]))
    } else {
        Ok(StmtKind::Label(label))
    }
}

// ── Expressions ─────────────────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::expression {
        let mut operands = Vec::new();
        let mut operators: Vec<String> = Vec::new();

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::unary_expression => operands.push(walk_unary_expression(inner)?),
                Rule::binary_op => {
                    let op = inner.as_str().to_string();
                    while operators
                        .last()
                        .is_some_and(|top| go_binary_precedence(top) >= go_binary_precedence(&op))
                    {
                        go_reduce_binary_expr(&mut operands, &mut operators)?;
                    }
                    operators.push(op);
                }
                _ => {}
            }
        }

        while !operators.is_empty() {
            go_reduce_binary_expr(&mut operands, &mut operators)?;
        }

        if let Some(result) = operands.pop() {
            return Ok(result);
        }
    } else if pair.as_rule() == Rule::unary_expression {
        return walk_unary_expression(pair);
    } else if pair.as_rule() == Rule::primary {
        return walk_primary(pair);
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut op = None;
    let mut operand = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::unary_op => op = Some(inner.as_str().to_string()),
            Rule::unary_expression => operand = Some(walk_unary_expression(inner)?),
            Rule::primary => operand = Some(walk_primary(inner)?),
            _ => {}
        }
    }

    if let Some(uop) = op {
        let un_op = match uop.as_str() {
            "-" => UnaryOp::Neg,
            "!" => UnaryOp::Not,
            "+" => UnaryOp::Pos,
            "^" => UnaryOp::BitNot,
            "*" => UnaryOp::Deref,
            "&" => UnaryOp::AddrOf,
            "<-" => {
                return Ok(go_channel_receive_expr(
                    operand.unwrap_or_else(Expression::null),
                ));
            }
            _ => UnaryOp::Pos,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: un_op,
            expr: Box::new(operand.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        }))
    } else {
        operand.ok_or_else(|| "Empty unary expression".to_string())
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut base = None;
    let mut chain = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::operand => {
                base = Some(walk_operand(inner)?);
            }
            Rule::selector => {
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::ident_name {
                        chain.push(PrimaryChain::Member(s_inner.as_str().to_string()));
                    }
                }
            }
            Rule::index => {
                for i_inner in inner.into_inner() {
                    if i_inner.as_rule() == Rule::expression {
                        chain.push(PrimaryChain::Index(walk_expression(i_inner)?));
                    }
                }
            }
            Rule::two_index_slice | Rule::three_index_slice => {
                let slice_source = inner.as_str();
                let mut start = None;
                let mut end = None;
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::expression {
                        if start.is_none() && !slice_source.starts_with("[:") {
                            start = Some(walk_expression(s_inner)?);
                        } else if end.is_none() {
                            end = Some(walk_expression(s_inner)?);
                        }
                    }
                }
                chain.push(PrimaryChain::Slice { start, end });
            }
            Rule::call => {
                let mut args = Vec::new();
                for c_inner in inner.into_inner() {
                    if c_inner.as_rule() == Rule::argument_list {
                        for arg_inner in c_inner.into_inner() {
                            if arg_inner.as_rule() == Rule::argument {
                                let mut spread = false;
                                let mut val = None;
                                for expr_inner in arg_inner.into_inner() {
                                    if expr_inner.as_rule() == Rule::expression {
                                        val = Some(walk_expression(expr_inner)?);
                                    } else if expr_inner.as_rule() == Rule::type_annotation {
                                        val = Some(go_type_arg_expr(walk_type(expr_inner)));
                                    } else if expr_inner.as_str() == "..." {
                                        spread = true;
                                    }
                                }
                                if let Some(expr) = val {
                                    args.push(Argument {
                                        value: expr,
                                        name: None,
                                        by_ref: false,
                                        spread,
                                    });
                                }
                            }
                        }
                    }
                }
                chain.push(PrimaryChain::Call(args));
            }
            Rule::type_assertion => {
                for t_inner in inner.into_inner() {
                    if t_inner.as_rule() == Rule::type_annotation {
                        chain.push(PrimaryChain::TypeAssert(walk_type(t_inner)));
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(mut result) = base {
        for item in chain {
            result = match item {
                PrimaryChain::Member(name) => Expression::new(ExprKind::Member {
                    object: Box::new(result),
                    field: name,
                    null_safe: false,
                }),
                PrimaryChain::Index(idx) => Expression::new(ExprKind::Index {
                    object: Box::new(result),
                    index: Box::new(idx),
                    null_safe: false,
                }),
                PrimaryChain::Slice { start, end } => {
                    let start_expr = start.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
                    let mut args = vec![Argument {
                        value: start_expr,
                        name: None,
                        by_ref: false,
                        spread: false,
                    }];
                    if let Some(end_expr) = end {
                        args.push(Argument {
                            value: end_expr,
                            name: None,
                            by_ref: false,
                            spread: false,
                        });
                    }
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(result),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args,
                        optional: false,
                    })
                }
                PrimaryChain::Call(args) => Expression::new(ExprKind::Call {
                    callee: Box::new(result),
                    args,
                    optional: false,
                }),
                PrimaryChain::TypeAssert(type_name) => go_type_assert_expr(result, type_name),
            };
        }
        Ok(result)
    } else {
        Ok(Expression::new(ExprKind::Lit(Literal::Null)))
    }
}

#[derive(Clone)]
enum PrimaryChain {
    Member(String),
    Index(Expression),
    Slice { start: Option<Expression>, end: Option<Expression> },
    Call(Vec<Argument>),
    TypeAssert(String),
}

fn walk_operand(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal => return walk_literal(inner),
            Rule::slice_conversion => return walk_slice_conversion(inner),
            Rule::ident_name => {
                let name = inner.as_str();
                // Go builtins
                match name {
                    "nil" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))),
                    "true" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
                    "false" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                    _ => return Ok(Expression::new(ExprKind::Ident(name.to_string()))),
                }
            }
            Rule::expression => return walk_expression(inner),
            Rule::composite_literal => return walk_composite_literal(inner),
            Rule::function_literal => return walk_function_literal(inner),
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_slice_conversion(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::from("[]");
    let mut expr = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => type_name.push_str(inner.as_str()),
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        type_name,
    }))
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::numeric_literal => {
                let s = inner.as_str().replace('_', "");
                if s.starts_with("0x") || s.starts_with("0X") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 16) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 2) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 8) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.contains('.') || s.contains('e') || s.contains('E') || s.contains('p') || s.contains('P') {
                    if let Ok(f) = s.parse::<f64>() {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Float(f))));
                    }
                } else if let Ok(n) = s.parse::<i64>() {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                }
            }
            Rule::string_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(inner.as_str())))));
            }
            Rule::bool_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Bool(inner.as_str() == "true"))));
            }
            Rule::nil_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_composite_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::new();
    let mut elements = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal_type => {
                type_name = inner.as_str().to_string();
            }
            Rule::literal_value => {
                for lv_inner in inner.into_inner() {
                    if lv_inner.as_rule() == Rule::element_list {
                        elements = walk_element_list(lv_inner)?;
                    }
                }
            }
            _ => {}
        }
    }

    if type_name.starts_with("map[") {
        // Build a dict/object literal
        let mut props = Vec::new();
        for (key, val) in elements {
            props.push(ObjectProperty::KeyValue {
                key,
                value: val,
            });
        }
        Ok(go_typed_composite_expr(Expression::new(ExprKind::Object(props)), &type_name))
    } else if go_is_array_like_type(&type_name) {
        let mut values: Vec<Expression> = elements.into_iter().map(|(_, value)| value).collect();
        if let Some(target_len) = go_fixed_array_len(&type_name, values.len()) {
            if let Some(elem_type) = go_array_element_type(&type_name) {
                while values.len() < target_len {
                    values.push(go_zero_value_expr(&elem_type));
                }
            }
        }
        let arr_elems: Vec<ArrayElement> = values.into_iter().map(|value| ArrayElement {
            key: None,
            value,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(go_typed_composite_expr(Expression::new(ExprKind::Array(arr_elems)), &type_name))
    } else if !type_name.is_empty() && elements.iter().all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null))) {
        let mut props = Vec::new();
        for (key, val) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::new(ExprKind::Lit(Literal::Str(name))),
                ExprKind::Lit(Literal::Int(n)) => Expression::new(ExprKind::Lit(Literal::Str(n.to_string()))),
                _ => key,
            };
            props.push(ObjectProperty::KeyValue {
                key,
                value: val,
            });
        }
        Ok(go_typed_composite_expr(Expression::new(ExprKind::Object(props)), &type_name))
    } else {
        // Untyped composite literal fallback.
        let arr_elems: Vec<ArrayElement> = elements.into_iter().map(|(_, v)| ArrayElement {
            key: None,
            value: v,
            spread: false,
            by_ref: false,
        }).collect();
        Ok(Expression::new(ExprKind::Array(arr_elems)))
    }
}

fn go_typed_composite_expr(expr: Expression, type_name: &str) -> Expression {
    if type_name.is_empty() {
        expr
    } else {
        Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        })
    }
}

fn go_is_array_like_type(type_name: &str) -> bool {
    go_array_head(type_name).is_some()
}

fn go_array_head(type_name: &str) -> Option<(&str, &str)> {
    let trimmed = type_name.trim();
    if !trimmed.starts_with('[') {
        return None;
    }
    let close = trimmed.find(']')?;
    Some((&trimmed[1..close], trimmed[close + 1..].trim()))
}

fn go_array_element_type(type_name: &str) -> Option<String> {
    let (_, tail) = go_array_head(type_name)?;
    (!tail.is_empty()).then(|| tail.to_string())
}

fn go_fixed_array_len(type_name: &str, inferred_len: usize) -> Option<usize> {
    let (head, _) = go_array_head(type_name)?;
    let head = head.trim();
    if head.is_empty() {
        None
    } else if head == "..." {
        Some(inferred_len)
    } else {
        head.parse::<usize>().ok()
    }
}

fn go_zero_value_expr(type_name: &str) -> Expression {
    let trimmed = type_name.trim();
    let lower = trimmed.to_ascii_lowercase();

    if let Some(len) = go_fixed_array_len(trimmed, 0) {
        if let Some(elem_type) = go_array_element_type(trimmed) {
            let elements = (0..len).map(|_| ArrayElement {
                key: None,
                value: go_zero_value_expr(&elem_type),
                spread: false,
                by_ref: false,
            }).collect();
            return go_typed_composite_expr(Expression::new(ExprKind::Array(elements)), trimmed);
        }
    }

    if lower.starts_with("[]") || lower.starts_with("map[") || lower.starts_with("chan ") || lower.starts_with('*') {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    match lower.as_str() {
        "bool" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        "string" => Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        "float32" | "float64" => Expression::new(ExprKind::Lit(Literal::Float(0.0))),
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16"
        | "uint32" | "uint64" | "uintptr" | "byte" | "rune" => {
            Expression::new(ExprKind::Lit(Literal::Int(0)))
        }
        _ => go_typed_composite_expr(Expression::new(ExprKind::Object(Vec::new())), trimmed),
    }
}

fn go_map_value_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim();
    if !trimmed.starts_with("map[") {
        return None;
    }

    let mut depth = 0usize;
    for (idx, ch) in trimmed.char_indices() {
        match ch {
            '[' => depth += 1,
            ']' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    let tail = trimmed.get(idx + 1..)?.trim();
                    return (!tail.is_empty()).then(|| tail.to_string());
                }
            }
            _ => {}
        }
    }
    None
}

fn walk_element_list(pair: Pair<Rule>) -> Result<Vec<(Expression, Expression)>, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::keyed_element {
            let parts: Vec<_> = inner.into_inner().collect();
            if parts.len() >= 2 {
                let key = go_keyed_element_key(parts[0].clone())?;
                let value = go_keyed_element_value(parts[1].clone())?;
                elements.push((key, value));
            } else if let Some(value_pair) = parts.into_iter().next() {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    go_keyed_element_value(value_pair)?,
                ));
            } else {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    Expression::new(ExprKind::Lit(Literal::Null)),
                ));
            }
        }
    }
    Ok(elements)
}

fn go_keyed_element_key(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(pair.as_str()))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(pair.as_str() == "true")))),
        Rule::numeric_literal => {
            let literal = pair.as_str().replace('_', "");
            if let Ok(n) = literal.parse::<i64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            } else if let Ok(f) = literal.parse::<f64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Float(f))))
            } else {
                Ok(Expression::new(ExprKind::Lit(Literal::Null)))
            }
        }
        Rule::expression => walk_expression(pair),
        Rule::element => go_keyed_element_value(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn go_keyed_element_value(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::element => {
            let Some(inner) = pair.into_inner().next() else {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            };
            go_keyed_element_value(inner)
        }
        Rule::expression => walk_expression(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(pair.as_str()))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(pair.as_str() == "true")))),
        Rule::numeric_literal => go_keyed_element_key(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn walk_literal_value_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::element_list {
            elements = walk_element_list(inner)?;
        }
    }

    if elements.iter().all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null))) {
        let mut props = Vec::new();
        for (key, value) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::string(&name),
                ExprKind::Lit(Literal::Int(n)) => Expression::string(&n.to_string()),
                _ => key,
            };
            props.push(ObjectProperty::KeyValue { key, value });
        }
        Ok(Expression::new(ExprKind::Object(props)))
    } else {
        Ok(Expression::new(ExprKind::Array(
            elements
                .into_iter()
                .map(|(_, value)| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )))
    }
}

fn walk_function_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
            }
            Rule::function_body | Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    }))
}

fn walk_expression_list(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression {
            exprs.push(walk_expression(inner)?);
        }
    }
    Ok(exprs)
}

// ── Helpers ───────────────────────────────────────────────────────────────────────────────

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() != Rule::EOI {
            return Ok(inner);
        }
    }
    Err("No meaningful child".to_string())
}

fn parse_bin_op(op: &str) -> BinOp {
    match op {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "&^" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn build_go_binary_expr(op: &str, left: Expression, right: Expression) -> Expression {
    if op == "&^" {
        Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(left),
            right: Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(right),
            })),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op: parse_bin_op(op),
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}

fn go_reduce_binary_expr(
    operands: &mut Vec<Expression>,
    operators: &mut Vec<String>,
) -> Result<(), String> {
    let Some(op) = operators.pop() else {
        return Ok(());
    };
    let Some(right) = operands.pop() else {
        return Err(format!("missing right operand for Go binary operator {op}"));
    };
    let Some(left) = operands.pop() else {
        return Err(format!("missing left operand for Go binary operator {op}"));
    };
    operands.push(build_go_binary_expr(&op, left, right));
    Ok(())
}

fn go_binary_precedence(op: &str) -> u8 {
    match op {
        "||" => 1,
        "&&" => 2,
        "==" | "!=" | "<" | "<=" | ">" | ">=" => 3,
        "+" | "-" | "|" | "^" => 4,
        "*" | "/" | "%" | "<<" | ">>" | "&" | "&^" => 5,
        _ => 0,
    }
}

fn go_type_arg_expr(type_name: String) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::null()),
        type_name,
    })
}

fn go_type_assert_expr(expr: Expression, type_name: String) -> Expression {
    go_builtin_call("__go_type_assert", vec![expr, go_type_arg_expr(type_name)])
}

fn go_extract_type_assert_expr(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_type_assert") || args.len() != 2 {
        return None;
    }
    Some((args[0].value.clone(), go_type_name_from_expr(&args[1].value)?))
}

fn go_type_assert_value_expr(expr: Expression, type_name: &str) -> Expression {
    if !matches!(type_name.trim(),
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16"
        | "uint32" | "uint64" | "uintptr" | "byte" | "rune" | "float32" | "float64"
        | "string" | "bool"
    ) {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        });
    }

    let cond = go_build_is_type(expr.clone(), type_name);
    let then_expr = Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(go_zero_value_expr(type_name)),
    })
}

fn go_type_switch_case_cond(expr: Expression, case_types: &[String]) -> Expression {
    let mut iter = case_types.iter();
    let first = iter
        .next()
        .map(|type_name| go_build_is_type(expr.clone(), type_name))
        .unwrap_or_else(|| Expression::bool(false));
    iter.fold(first, |acc, type_name| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(acc),
            right: Box::new(go_build_is_type(expr.clone(), type_name)),
        })
    })
}

fn go_build_is_type(expr: Expression, type_name: &str) -> Expression {
    let typeof_tag = match type_name.trim() {
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16"
        | "uint32" | "uint64" | "uintptr" | "byte" | "rune" | "float32" | "float64" => {
            Some("number")
        }
        "string" => Some("string"),
        "bool" => Some("boolean"),
        _ => None,
    };

    if let Some(tag) = typeof_tag {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::string(tag)),
        });
    }

    Expression::new(ExprKind::IsType {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn go_type_switch_case_body(
    mut body: Vec<Statement>,
    binding_name: Option<&str>,
    expr: Expression,
    case_type: &str,
) -> Vec<Statement> {
    if let Some(name) = binding_name {
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.to_string()),
                    init: Some(Expression::new(ExprKind::Cast {
                        expr: Box::new(expr),
                        type_name: case_type.to_string(),
                    })),
                    type_hint: Some(case_type.to_string()),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
        );
    }
    body
}

fn go_wrap_spawn_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(expr))]),
        is_async: false,
        captures: Vec::new(),
    })
}

fn go_channel_receive_expr(channel: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(channel),
            field: "queue".to_string(),
            null_safe: false,
        })),
        index: Box::new(Expression::int(0)),
        null_safe: false,
    })
}

fn go_channel_send_expr(channel: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(channel),
                field: "queue".to_string(),
                null_safe: false,
            })),
            field: "push".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn go_channel_object_expr(capacity: Option<Expression>) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("queue"),
            value: Expression::new(ExprKind::Array(Vec::new())),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("closed"),
            value: Expression::bool(false),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("capacity"),
            value: capacity.unwrap_or_else(|| Expression::int(0)),
        },
    ]))
}

fn go_type_name_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Cast { expr, type_name } if matches!(expr.kind, ExprKind::Lit(Literal::Null)) => {
            Some(type_name.clone())
        }
        _ => None,
    }
}

fn go_is_slice_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("[]")
}

fn go_is_map_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("map[")
}

fn go_is_channel_type(type_name: &str) -> bool {
    let trimmed = type_name.trim_start();
    trimmed.starts_with("chan") || trimmed.starts_with("<-chan")
}

fn go_array_make_expr(len_expr: Expression, init_expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![Argument::positional(len_expr), Argument::positional(init_expr)],
        optional: false,
    })
}

fn go_make_slice_capacity_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("make") {
        return None;
    }
    let type_name = args.first().and_then(|arg| go_type_name_from_expr(&arg.value))?;
    if !go_is_slice_type(&type_name) {
        return None;
    }
    let cap_arg = args.get(2).or_else(|| args.get(1))?;
    Some(normalize_go_expr(&cap_arg.value, env, signatures, state))
}

fn go_decl_slice_capacity_binding(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<(String, Expression)> {
    let name = go_binding_name(&decl.pattern)?;
    let cap = go_make_slice_capacity_expr(decl.init.as_ref()?, env, signatures, state)?;
    Some((name, cap))
}

fn go_expr_capacity_hint(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Ident(name) => env.slice_caps.get(name).cloned(),
        ExprKind::Array(elements) => Some(Expression::int(elements.len() as i64)),
        _ => None,
    }
}

fn go_binding_name(pattern: &BindingPattern) -> Option<String> {
    match pattern {
        BindingPattern::Ident(name) => Some(name.clone()),
        _ => go_single_named_binding_pattern(pattern).and_then(|pattern| match pattern {
            BindingPattern::Ident(name) => Some(name),
            _ => None,
        }),
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span { start_line: sl as u32, start_col: sc as u32, end_line: el as u32, end_col: ec as u32 }
}
