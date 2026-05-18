//! Go walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//!
//! ## Go-specific normalisations
//!
//! - **Multiple return values**: Go functions can return multiple values.
//!   For simplicity we compile to returning a single array/tuple.
//! - **Short variable declaration** (`:=`): Maps to `VarDecl` with `Let`.
//! - **Methods**: Go methods on structs are compiled as regular functions
//!   with the receiver as the first parameter.
//! - **Structs**: Mapped to `ClassDecl` with fields.
//! - **Interfaces**: Mapped to `InterfaceDecl`.
//! - **`range`**: Mapped to `ForIn` with `of: true`.
//! - **`defer`**: Currently ignored (no-op) — Go's defer semantics require
//!   runtime support not yet available.
//! - **`go`**: Currently ignored (no-op) — goroutines require runtime support.
//! - **`fallthrough`**: Not yet supported in switch.
//! - **`select`**: Not yet supported.
//! - **`chan` / `<-`**: Not yet supported.
//! - **`nil`**: Mapped to `ExprKind::Lit(Literal::Null)`.
//! - **`make` / `new`**: `make` for slices/maps is rewritten to array/dict
//!   creation. `new(T)` becomes `&T{}` (pointer to zero value).
//! - **`append`**: Rewritten to array push.
//! - **`len` / `cap`**: Builtin functions mapped to host calls.
//! - **`panic` / `recover`**: Mapped to throw/try-catch.
//! - **`_` blank identifier**: Ignored in assignments.

use pest::Parser;
use pest::iterators::Pair;
use std::collections::HashMap;
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
    return_type: Option<String>,
}

#[derive(Default)]
struct GoNormalizeState {
    next_temp: usize,
}

fn normalize_go_module(mut module: Module) -> Module {
    let signatures = collect_go_function_signatures(&module.body);
    let globals = collect_go_global_fixed_arrays(&module.body, &signatures);
    let mut state = GoNormalizeState::default();
    let mut env = GoNormalizeEnv {
        value_types: HashMap::new(),
        fixed_arrays: globals.clone(),
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
    signatures
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
                body: normalize_go_block(body, &fn_env, signatures, state),
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
                    .map(|expr| normalize_go_expr(expr, env, signatures, state));
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
                }
            }
            vec![Statement::new(StmtKind::Assign {
                targets: targets
                    .iter()
                    .map(|target| normalize_go_expr(target, env, signatures, state))
                    .collect(),
                value: next_value,
            })]
        }
        StmtKind::CompoundAssign { target, op, value } => vec![Statement::new(StmtKind::CompoundAssign {
            target: normalize_go_expr(target, env, signatures, state),
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
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(next_object),
                    index: Box::new(next_index),
                    null_safe: *null_safe,
                })
            }
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_expr(target, env, signatures, state)),
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
                    LambdaBody::Block(normalize_go_block(stmts, &lambda_env, signatures, state))
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
        ExprKind::Index { object, .. } => go_expr_type_hint(object, env, signatures)
            .and_then(|type_name| go_array_element_type(&type_name)),
        ExprKind::Assign { value, .. } => go_expr_type_hint(value, env, signatures),
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

    let mut bound_name = None;
    for element in elements {
        match element {
            ArrayPatternElem::Hole => {}
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

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::signature => {
                let (p, rt) = walk_signature(inner)?;
                params = p;
                return_type = rt;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
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
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::receiver => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::ident_name => receiver_name = r_inner.as_str().to_string(),
                        Rule::type_annotation => receiver_type = walk_type(r_inner),
                        _ => {}
                    }
                }
            }
            Rule::ident_name => method_name = inner.as_str().to_string(),
            Rule::signature => {
                let (p, rt) = walk_signature(inner)?;
                params = p;
                return_type = rt;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    // Prepend receiver as first parameter
    params.insert(0, Param {
        name: if receiver_name.is_empty() { "self".to_string() } else { receiver_name },
        type_hint: Some(receiver_type),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    });

    Ok(Statement::new(StmtKind::FunctionDecl {
        name: method_name,
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

fn walk_signature(pair: Pair<Rule>) -> Result<(Vec<Param>, Option<String>), String> {
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;

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
                            // Multiple return values — represent as array
                            let p = walk_parameter_list(r_inner)?;
                            return_type = Some(format!("[{}]", p.len()));
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

    Ok((params, return_type))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter_decl {
            let mut names = Vec::new();
            let mut type_hint: Option<String> = None;

            for p_inner in inner.into_inner() {
                match p_inner.as_rule() {
                    Rule::ident_name => names.push(p_inner.as_str().to_string()),
                    Rule::ident_list => {
                        for id in p_inner.into_inner() {
                            if id.as_rule() == Rule::ident_name {
                                names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::type_annotation => type_hint = Some(walk_type(p_inner)),
                    _ => {}
                }
            }

            for name in names {
                params.push(Param {
                    name,
                    type_hint: type_hint.clone(),
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
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Const)?;
                declarations.append(&mut decls);
            }
            Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Const)?;
                        declarations.append(&mut decls);
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

fn walk_var_spec(pair: Pair<Rule>, _kind: VarDeclKind) -> Result<(Vec<VarDeclarator>, Option<String>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if id.as_rule() == Rule::ident_name {
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
                        Rule::type_annotation => type_str = walk_type(spec_inner),
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
                            if id.as_rule() == Rule::ident_name {
                                field_names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::ident_name => field_names.push(f_inner.as_str().to_string()),
                    Rule::type_annotation => field_type = Some(walk_type(f_inner)),
                    _ => {}
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

    Ok(Statement::new(StmtKind::ClassDecl {
        name,
        parents: Vec::new(),
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
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
                        let (p, rt) = walk_signature(m_inner)?;
                        params = p;
                        return_type = rt;
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

// ── Statements ─────────────────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let rule = pair.as_rule();
    if rule == Rule::statement {
        if let Some(inner) = pair.into_inner().next() {
            return walk_statement(inner);
        }
        return Ok(Statement::new(StmtKind::Empty));
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
        Rule::for_statement => walk_for(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::goto_statement => StmtKind::GoTo(walk_goto(pair)?),
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::defer_statement => StmtKind::Empty, // TODO: implement defer
        Rule::go_statement => StmtKind::Empty,    // TODO: implement goroutines
        Rule::send_statement => StmtKind::Empty,  // TODO: implement channels
        _ => StmtKind::Empty,
    };
    Ok(Statement::new(kind))
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
                    if id.as_rule() == Rule::ident_name {
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
                            if let StmtKind::If { cond: c, then_body: t, else_body: e, .. } = elif {
                                then_body.push(Statement::new(StmtKind::If {
                                    cond: c,
                                    then_body: t,
                                    elifs: Vec::new(),
                                    else_body: e,
                                }));
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

    let mut then = then_body;
    if let Some(pre) = pre_stmt {
        then.insert(0, *pre);
    }

    Ok(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body: then,
        elifs: Vec::new(),
        else_body,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
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

    Ok(StmtKind::Switch {
        expr: expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        cases,
        default,
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
                                if id.as_rule() == Rule::ident_name {
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
        let mut left = None;
        let mut ops = Vec::new();

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::unary_expression => {
                    if left.is_none() {
                        left = Some(walk_unary_expression(inner)?);
                    } else {
                        ops.push((None, walk_unary_expression(inner)?));
                    }
                }
                Rule::binary_op => {
                    ops.push((Some(inner.as_str().to_string()), Expression::new(ExprKind::Lit(Literal::Null))));
                }
                _ => {}
            }
        }

        if let Some(mut result) = left {
            let mut i = 0;
            while i < ops.len() {
                if let (Some(op), _) = &ops[i] {
                    if i + 1 < ops.len() {
                        let right = ops[i + 1].1.clone();
                        result = build_go_binary_expr(op, result, right);
                        i += 2;
                        continue;
                    }
                }
                i += 1;
            }
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
            "<-" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))), // channel receive — not supported
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
                // type assertions like .(Type) — ignore for now
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
                    let end_expr = end.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(result),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument { value: start_expr, name: None, by_ref: false, spread: false },
                            Argument { value: end_expr, name: None, by_ref: false, spread: false },
                        ],
                        optional: false,
                    })
                }
                PrimaryChain::Call(args) => Expression::new(ExprKind::Call {
                    callee: Box::new(result),
                    args,
                    optional: false,
                }),
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
            let key_str = match &key.kind {
                ExprKind::Lit(Literal::Str(s)) => s.clone(),
                ExprKind::Ident(s) => s.clone(),
                ExprKind::Lit(Literal::Int(n)) => n.to_string(),
                _ => format!("{:?}", key),
            };
            props.push(ObjectProperty::KeyValue {
                key: Expression::new(ExprKind::Lit(Literal::Str(key_str))),
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

fn walk_element_list(pair: Pair<Rule>) -> Result<Vec<(Expression, Expression)>, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::keyed_element {
            let mut key = None;
            let mut val = None;

            for ke_inner in inner.into_inner() {
                match ke_inner.as_rule() {
                    Rule::ident_name => {
                        if key.is_none() {
                            key = Some(Expression::new(ExprKind::Ident(ke_inner.as_str().to_string())));
                        } else if val.is_none() {
                            val = Some(Expression::new(ExprKind::Ident(ke_inner.as_str().to_string())));
                        }
                    }
                    Rule::expression => {
                        if key.is_none() {
                            key = Some(walk_expression(ke_inner)?);
                        } else if val.is_none() {
                            val = Some(walk_expression(ke_inner)?);
                        }
                    }
                    Rule::element => {
                        for e_inner in ke_inner.into_inner() {
                            if e_inner.as_rule() == Rule::expression {
                                if val.is_none() {
                                    val = Some(walk_expression(e_inner)?);
                                }
                            } else if e_inner.as_rule() == Rule::literal_value {
                                val = Some(walk_literal_value_expr(e_inner)?);
                            }
                        }
                    }
                    Rule::literal_value => {
                        val = Some(walk_literal_value_expr(ke_inner)?);
                    }
                    _ => {}
                }
            }

            let value = val.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
            if let Some(k) = key {
                elements.push((k, value));
            } else {
                elements.push((Expression::new(ExprKind::Lit(Literal::Null)), value));
            }
        }
    }
    Ok(elements)
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
                let (p, _) = walk_signature(inner)?;
                params = p;
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
