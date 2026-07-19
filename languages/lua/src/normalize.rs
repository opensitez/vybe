//! Lua → JS-shaped AST normalization.
//! Normalizes Lua-specific operations to adapter calls that handle metamethods at runtime.

use vybe_ast::*;

/// Normalize module: transform statements and expressions
pub fn normalize_module(module: &mut Module) {
    for stmt in &mut module.body {
        normalize_stmt(&mut stmt.kind);
    }
}

fn lua_float_repr(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_float_repr".to_string()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn expr_is_lua_float(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Unary {
            op: UnaryOp::Neg | UnaryOp::Pos,
            expr,
        } => expr_is_lua_float(expr),
        ExprKind::Binary { op, left, right } => match op {
            BinOp::Div | BinOp::Pow => true,
            BinOp::Add | BinOp::Sub | BinOp::Mul | BinOp::Mod | BinOp::FloorDiv => {
                expr_is_lua_float(left) || expr_is_lua_float(right)
            }
            _ => false,
        },
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name)
                if matches!(name.as_str(), "tonumber" | "__lua_tonumber")
                    && matches!(
                        args.first().map(|arg| &arg.value.kind),
                        Some(ExprKind::Lit(Literal::Str(s))) if s.contains('e') || s.contains('E')
                    ) =>
            {
                true
            }
            ExprKind::Ident(name) if name == "__lua_div" || name == "__lua_pow" => true,
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__lua_add" | "__lua_sub" | "__lua_mul" | "__lua_mod" | "__lua_idiv" | "__lua_unm"
                ) => args.iter().any(|arg| expr_is_lua_float(&arg.value)),
            _ => false,
        },
        _ => false,
    }
}

fn wrap_lua_float_display_arg(arg: &mut Argument) {
    if arg.name.is_none() && !arg.spread && expr_is_lua_float(&arg.value) {
        let value = std::mem::replace(&mut arg.value, Expression::new(ExprKind::Lit(Literal::Null)));
        arg.value = lua_float_repr(value);
    }
}

fn lua_call_name(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.as_str()),
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "string") =>
        {
            Some(match field.as_str() {
                "find" => "string.find",
                "gsub" => "string.gsub",
                "gmatch" => "string.gmatch",
                "match" => "string.match",
                "byte" => "string.byte",
                "unpack" => "string.unpack",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "table") =>
        {
            Some(match field.as_str() {
                "sort" => "table.sort",
                "unpack" => "table.unpack",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "math") =>
        {
            Some(match field.as_str() {
                "modf" => "math.modf",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "coroutine") =>
        {
            Some(match field.as_str() {
                "create" => "coroutine.create",
                "resume" => "coroutine.resume",
                "yield" => "coroutine.yield",
                "status" => "coroutine.status",
                "running" => "coroutine.running",
                "wrap" => "coroutine.wrap",
                "close" => "coroutine.close",
                "isyieldable" => "coroutine.isyieldable",
                _ => return None,
            })
        }
        _ => None,
    }
}

fn is_lua_global_env(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if name == "_G")
}

fn lua_static_key(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(key)) => Some(key.clone()),
        _ => None,
    }
}

fn lua_ident(name: impl Into<String>) -> Expression {
    Expression::new(ExprKind::Ident(name.into()))
}

fn is_lua_math_member(expr: &Expression, member: &str) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, field, .. }
            if field == member && matches!(&object.kind, ExprKind::Ident(name) if name == "math")
    )
}

fn lua_static_math_type_arg(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("integer"),
        ExprKind::Lit(Literal::Float(_)) => Some("float"),
        ExprKind::Call { callee, args, .. }
            if args.is_empty()
                && (is_lua_math_member(callee, "maxinteger")
                    || is_lua_math_member(callee, "mininteger")) =>
        {
            Some("integer")
        }
        ExprKind::Member { .. }
            if is_lua_math_member(expr, "maxinteger") || is_lua_math_member(expr, "mininteger") =>
        {
            Some("integer")
        }
        _ => None,
    }
}

/// Lua's `_G` is the language spelling for the VM global namespace. Keep that
/// alias at normalization time so the emitter continues to use GLOBAL_GET/SET
/// instead of growing a second Lua-only global table.
fn lua_global_alias_read(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if is_lua_global_env(object) => {
            Some(lua_ident(field.clone()))
        }
        ExprKind::Index { object, index, .. } if is_lua_global_env(object) => {
            lua_static_key(index).map(lua_ident)
        }
        ExprKind::Call { callee, args, .. }
            if lua_call_name(callee) == Some("rawget")
                && args.len() == 2
                && is_lua_global_env(&args[0].value) =>
        {
            lua_static_key(&args[1].value).map(lua_ident)
        }
        _ => None,
    }
}

fn lua_global_alias_write(target: &Expression, mut value: Expression) -> Option<Statement> {
    let name = match &target.kind {
        ExprKind::Member { object, field, .. } if is_lua_global_env(object) => Some(field.clone()),
        ExprKind::Index { object, index, .. } if is_lua_global_env(object) => lua_static_key(index),
        _ => None,
    }?;
    normalize_expr(&mut value);
    Some(Statement::new(StmtKind::Assign {
        targets: vec![lua_ident(name)],
        value,
    }))
}

fn lua_global_alias_rawset_stmt(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if lua_call_name(callee) != Some("rawset")
        || args.len() != 3
        || !is_lua_global_env(&args[0].value)
    {
        return None;
    }
    let name = lua_static_key(&args[1].value)?;
    let mut value = args[2].value.clone();
    normalize_expr(&mut value);
    Some(StmtKind::Assign {
        targets: vec![lua_ident(name)],
        value,
    })
}

fn is_lua_multi_return_call(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => match lua_call_name(callee) {
            Some("string.find" | "string.gsub" | "string.match") => true,
            Some("string.unpack" | "table.unpack" | "math.modf") => true,
            Some("coroutine.resume" | "coroutine.running" | "coroutine.yield" | "__lua_wrap_resume") => true,
            Some("next" | "pcall" | "xpcall") => true,
            Some("assert") => args.len() > 1,
            Some("string.byte") => args.len() >= 3,
            _ => false,
        },
        _ => false,
    }
}

fn is_lua_assert_call(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. } if lua_call_name(callee).as_deref() == Some("assert")
    )
}

fn is_lua_internal_multi_helper(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(ref name)
            if name == "__lua_first" || name == "__lua_index0" || name == "__lua_multi_row"
    )
}

fn lua_multi_row(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_multi_row".to_string()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_multi_row_from_values(values: Vec<Expression>) -> Expression {
    lua_multi_row(Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    )))
}

fn lua_is_multi_row_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. } if lua_call_name(callee).as_deref() == Some("__lua_multi_row")
    )
}

fn lua_mark_coroutine_returns(stmts: &mut [Statement]) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Return(Some(expr)) => {
                if !lua_is_multi_row_expr(expr) {
                    let value = std::mem::replace(
                        expr,
                        Expression::new(ExprKind::Lit(Literal::Null)),
                    );
                    *expr = lua_multi_row_from_values(vec![value]);
                }
            }
            StmtKind::Block(body) => lua_mark_coroutine_returns(body),
            StmtKind::For { body, .. } => lua_mark_coroutine_returns(body),
            StmtKind::ForIn {
                body, else_body, ..
            } => {
                lua_mark_coroutine_returns(body);
                if let Some(else_body) = else_body {
                    lua_mark_coroutine_returns(else_body);
                }
            }
            StmtKind::While { body, .. } | StmtKind::DoWhile { body, .. } => {
                lua_mark_coroutine_returns(body);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                lua_mark_coroutine_returns(then_body);
                for (_, body) in elifs {
                    lua_mark_coroutine_returns(body);
                }
                if let Some(else_body) = else_body {
                    lua_mark_coroutine_returns(else_body);
                }
            }
            _ => {}
        }
    }
}

fn lua_first(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_first".to_string()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_coroutine_generator_function(mut func: Expression) -> Expression {
    let ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    } = func.kind
    else {
        normalize_expr(&mut func);
        return func;
    };

    let mut generator_body = match body {
        LambdaBody::Block(stmts) => stmts,
        LambdaBody::Expr(expr) => vec![Statement::new(StmtKind::Return(Some(*expr)))],
    };
    for stmt in &mut generator_body {
        normalize_stmt(&mut stmt.kind);
    }
    lua_mark_coroutine_returns(&mut generator_body);

    let generator_fn = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body: generator_body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async,
            is_generator: true,
            is_sub: false,
        },
    ))));

    let gen_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__gen_fn".to_string()),
            type_hint: None,
            init: Some(generator_fn),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Const,
    });

    let call_gen = Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__gen_fn")),
        args: Vec::new(),
        optional: false,
    });

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(vec![gen_decl, Statement::new(StmtKind::Return(Some(call_gen)))]),
        is_async: false,
        captures,
    })
}

fn lua_table_from_pairs(rows: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_table_from_pairs".to_string()))),
        args: vec![Argument::positional(rows)],
        optional: false,
    })
}

fn lua_sort_compare_result(cond: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(lua_truthy(cond)),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
    })
}

fn normalize_lua_sort_comparator(arg: &mut Argument) {
    if arg.name.is_some() || arg.spread {
        return;
    }
    match &mut arg.value.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => {
                let cond = std::mem::replace(
                    expr,
                    Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                );
                *expr = Box::new(lua_sort_compare_result(*cond));
            }
            LambdaBody::Block(stmts) => {
                if stmts.len() == 1 {
                    if let StmtKind::Return(Some(ret)) = &mut stmts[0].kind {
                        let cond =
                            std::mem::replace(ret, Expression::new(ExprKind::Lit(Literal::Null)));
                        *ret = lua_sort_compare_result(cond);
                    }
                }
            }
        },
        _ => {}
    }
}

fn mark_last_lua_multi_return_arg_spread(args: &mut [Argument]) {
    let Some(last) = args.last_mut() else {
        return;
    };
    if last.name.is_some() || last.spread {
        return;
    }
    if is_lua_multi_return_call(&last.value) {
        if is_lua_assert_call(&last.value) {
            return;
        }
        last.spread = true;
        return;
    }
    let ExprKind::Call {
        callee,
        args: first_args,
        ..
    } = &mut last.value.kind
    else {
        return;
    };
    if !is_lua_internal_multi_helper(callee) || first_args.len() != 1 {
        return;
    }
    if is_lua_multi_return_call(&first_args[0].value) {
        if is_lua_assert_call(&first_args[0].value) {
            return;
        }
        last.value = first_args[0].value.clone();
        last.spread = true;
    }
}

fn normalize_lua_multi_return_source(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            let keep_profile_member = matches!(
                &callee.kind,
                ExprKind::Member { object, .. }
                    if matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if matches!(
                                name.as_str(),
                                "string" | "table" | "math" | "io" | "os" | "debug" | "coroutine"
                            )
                    )
            );
            if !keep_profile_member {
                normalize_expr(callee);
            }
            for arg in args.iter_mut() {
                normalize_expr(&mut arg.value);
            }
            mark_last_lua_multi_return_arg_spread(args);
        }
        _ => normalize_expr(expr),
    }
}

fn lua_multi_index(source: Expression, index: i64) -> Expression {
    lua_multi_index_expr(source, Expression::new(ExprKind::Lit(Literal::Int(index))))
}

fn lua_multi_index_expr(source: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_index0".to_string()))),
        args: vec![
            Argument::positional(source),
            Argument::positional(index),
        ],
        optional: false,
    })
}

fn normalize_expr(expr: &mut Expression) {
    if let Some(alias) = lua_global_alias_read(expr) {
        expr.kind = alias.kind;
        return;
    }

    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            // Recursively normalize operands first
            normalize_expr(left);
            normalize_expr(right);

            if *op == BinOp::And {
                let left_expr = left.as_ref().clone();
                expr.kind = ExprKind::Ternary {
                    cond: Box::new(lua_truthy(left_expr.clone())),
                    then: Box::new(right.as_ref().clone()),
                    else_: Box::new(left_expr),
                };
                return;
            }

            if *op == BinOp::Or {
                let left_expr = left.as_ref().clone();
                expr.kind = ExprKind::Ternary {
                    cond: Box::new(lua_truthy(left_expr.clone())),
                    then: Box::new(left_expr),
                    else_: Box::new(right.as_ref().clone()),
                };
                return;
            }

            // Wrap operations in adapters that check for metamethods at runtime
            let adapter_name = match op {
                BinOp::Add => "__lua_add",
                BinOp::Sub => "__lua_sub",
                BinOp::Mul => "__lua_mul",
                BinOp::Div => "__lua_div",
                BinOp::FloorDiv => "__lua_idiv",
                BinOp::Mod => "__lua_mod",
                BinOp::Pow => "__lua_pow",
                BinOp::Lt => "__lua_lt",
                BinOp::LtEq => "__lua_le",
                BinOp::Gt => "__lua_gt",
                BinOp::GtEq => "__lua_ge",
                BinOp::Eq => "__lua_eq",
                BinOp::NotEq => "__lua_ne",
                BinOp::Concat => "__lua_concat",
                BinOp::BitAnd => "__lua_band",
                BinOp::BitOr => "__lua_bor",
                BinOp::BitXor => "__lua_bxor",
                BinOp::Shl => "__lua_shl",
                BinOp::Shr => "__lua_shr",
                _ => return, // Other ops pass through
            };

            // Replace with call to adapter: __lua_add(left, right)
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(adapter_name.to_string()))),
                args: vec![
                    Argument::positional(left.as_ref().clone()),
                    Argument::positional(right.as_ref().clone()),
                ],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: inner,
        } => {
            normalize_expr(inner);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_unm".to_string()))),
                args: vec![Argument::positional(inner.as_ref().clone())],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: inner,
        } => {
            normalize_expr(inner);
            expr.kind = ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(lua_truthy(inner.as_ref().clone())),
            };
        }
        ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr: inner,
        } => {
            normalize_expr(inner);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_bnot".to_string()))),
                args: vec![Argument::positional(inner.as_ref().clone())],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Index { object, index, .. } => {
            normalize_expr(object);
            normalize_expr(index);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_index".to_string()))),
                args: vec![
                    Argument::positional(object.as_ref().clone()),
                    Argument::positional(index.as_ref().clone()),
                ],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Call { callee, args, .. } => {
            let call_name_before = lua_call_name(callee).map(str::to_string);
            if call_name_before.as_deref() == Some("coroutine.yield") {
                let values = std::mem::take(args)
                    .into_iter()
                    .map(|mut arg| {
                        normalize_expr(&mut arg.value);
                        arg.value
                    })
                    .collect::<Vec<_>>();
                expr.kind = ExprKind::Yield(Some(Box::new(lua_multi_row_from_values(values))));
                return;
            }
            if matches!(
                call_name_before.as_deref(),
                Some("coroutine.create" | "coroutine.wrap")
            ) {
                if let Some(first) = args.first_mut() {
                    if first.name.is_none() && !first.spread {
                        let value = std::mem::replace(
                            &mut first.value,
                            Expression::new(ExprKind::Lit(Literal::Null)),
                        );
                        first.value = lua_coroutine_generator_function(value);
                    }
                }
            }
            let internal_multi_helper = is_lua_internal_multi_helper(callee);
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if !matches!(&object.kind, ExprKind::Ident(name) if name == "string")
                    && matches!(
                        field.as_str(),
                        "byte"
                            | "char"
                            | "find"
                            | "format"
                            | "gmatch"
                            | "gsub"
                            | "len"
                            | "lower"
                            | "match"
                            | "rep"
                            | "reverse"
                            | "sub"
                            | "upper"
                    )
                {
                    let method = field.clone();
                    callee.kind = ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Ident("string".to_string()))),
                        field: method,
                        null_safe: false,
                    };
                }
            }
            let keep_profile_member = matches!(
                &callee.kind,
                ExprKind::Member { object, .. }
                    if matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if matches!(
                                name.as_str(),
                                "string" | "table" | "math" | "io" | "os" | "debug" | "coroutine"
                            )
                    )
            );
            if !keep_profile_member {
                normalize_expr(callee);
            }
            if !internal_multi_helper {
                for arg in args.iter_mut() {
                    normalize_expr(&mut arg.value);
                }
                mark_last_lua_multi_return_arg_spread(args);
            }
            if lua_call_name(callee).as_deref() == Some("table.sort") && args.len() >= 2 {
                normalize_lua_sort_comparator(&mut args[1]);
            }
            if is_lua_math_member(callee, "type") && args.len() == 1 && !args[0].spread {
                if let Some(kind) = lua_static_math_type_arg(&args[0].value) {
                    expr.kind = ExprKind::Lit(Literal::Str(kind.to_string()));
                    return;
                }
            }
            let callee_name = if let ExprKind::Ident(name) = &callee.kind {
                Some(name.as_str())
            } else {
                None
            };
            if matches!(callee_name, Some("print") | Some("__lua_print")) {
                if args.len() == 1 && args[0].spread {
                    let row = args[0].value.clone();
                    expr.kind = ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident(
                            "__lua_print_row".to_string(),
                        ))),
                        args: vec![Argument::positional(row)],
                        optional: false,
                    };
                    return;
                }
                for arg in args.iter_mut() {
                    wrap_lua_float_display_arg(arg);
                }
            } else if matches!(callee_name, Some("tostring") | Some("__lua_tostring")) {
                if let Some(arg) = args.first_mut() {
                    wrap_lua_float_display_arg(arg);
                }
            } else if args.len() == 1 && args[0].spread {
                let fn_expr = (**callee).clone();
                let row = args[0].value.clone();
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident(
                        "__lua_apply_row".to_string(),
                    ))),
                    args: vec![Argument::positional(fn_expr), Argument::positional(row)],
                    optional: false,
                };
            } else if is_lua_multi_return_call(expr) {
                let value = expr.clone();
                expr.kind = lua_first(value).kind;
            }
        }
        ExprKind::Array(elems) => {
            for elem in elems.iter_mut() {
                if elem.key.is_none() && !elem.spread {
                    match &elem.value.kind {
                        ExprKind::Spread(inner)
                            if matches!(inner.kind, ExprKind::Lit(Literal::Null)) =>
                        {
                            elem.value = Expression::new(ExprKind::Ident("...".to_string()));
                            elem.spread = true;
                        }
                        ExprKind::Spread(inner) => {
                            elem.value = inner.as_ref().clone();
                            normalize_expr(&mut elem.value);
                            elem.spread = true;
                        }
                        _ if is_lua_multi_return_call(&elem.value) => {
                            let mut value = elem.value.clone();
                            normalize_lua_multi_return_source(&mut value);
                            elem.value = lua_multi_row(value);
                            elem.spread = true;
                        }
                        _ => normalize_expr(&mut elem.value),
                    }
                } else {
                    normalize_expr(&mut elem.value);
                }
                if let Some(key) = &mut elem.key {
                    normalize_expr(key);
                }
            }
            let has_key = elems.iter().any(|elem| elem.key.is_some());
            let has_unkeyed = elems.iter().any(|elem| elem.key.is_none());
            let all_explicit_positive_int_keys = !elems.is_empty()
                && elems.iter().all(|elem| {
                    !elem.spread
                        && matches!(
                            elem.key.as_ref().map(|key| &key.kind),
                            Some(ExprKind::Lit(Literal::Int(n))) if *n > 0
                        )
                });
            let explicit_positive_max_key = if all_explicit_positive_int_keys {
                elems
                    .iter()
                    .filter_map(|elem| match elem.key.as_ref().map(|key| &key.kind) {
                        Some(ExprKind::Lit(Literal::Int(n))) => Some((*n - 1) as usize),
                        _ => None,
                    })
                    .max()
            } else {
                None
            };
            if all_explicit_positive_int_keys
                && explicit_positive_max_key.is_some_and(|max_key| max_key + 1 == elems.len())
            {
                let max_key = explicit_positive_max_key.unwrap_or(0);
                let mut dense = (0..=max_key)
                    .map(|_| ArrayElement {
                        key: None,
                        value: Expression::new(ExprKind::Lit(Literal::Undefined)),
                        spread: false,
                        by_ref: false,
                    })
                    .collect::<Vec<_>>();
                for elem in std::mem::take(elems) {
                    if let Some(ExprKind::Lit(Literal::Int(n))) =
                        elem.key.as_ref().map(|key| &key.kind)
                    {
                        dense[(*n - 1) as usize] = ArrayElement {
                            key: None,
                            value: elem.value,
                            spread: false,
                            by_ref: false,
                        };
                    }
                }
                *elems = dense;
            } else if has_key && has_unkeyed {
                let mut next_auto_key = 1_i64;
                for elem in elems.iter_mut() {
                    if elem.key.is_none() {
                        elem.key = Some(Expression::new(ExprKind::Lit(Literal::Int(next_auto_key))));
                        next_auto_key += 1;
                    }
                }
            }
            if elems.iter().any(|elem| elem.key.is_some()) {
                let rows = std::mem::take(elems)
                    .into_iter()
                    .map(|elem| {
                        let key = elem
                            .key
                            .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                        ArrayElement {
                            key: None,
                            value: Expression::new(ExprKind::Array(vec![
                                ArrayElement {
                                    key: None,
                                    value: key,
                                    spread: false,
                                    by_ref: false,
                                },
                                ArrayElement {
                                    key: None,
                                    value: elem.value,
                                    spread: false,
                                    by_ref: false,
                                },
                            ])),
                            spread: false,
                            by_ref: false,
                        }
                    })
                    .collect::<Vec<_>>();
                expr.kind = lua_table_from_pairs(Expression::new(ExprKind::Array(rows))).kind;
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_expr(expr),
            LambdaBody::Block(stmts) => {
                for stmt in stmts {
                    normalize_stmt(&mut stmt.kind);
                }
            }
        },
        ExprKind::Member { object, field, .. } => {
            if matches!(object.as_ref().kind, ExprKind::Ident(ref name) if name == "io" && field == "stdout") {
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__lua_stdout".to_string()))),
                    args: Vec::new(),
                    optional: false,
                };
                return;
            }
            let math_const = match object.as_ref().kind {
                ExprKind::Ident(ref name) if name == "math" && field == "maxinteger" => {
                    Some("math_maxinteger")
                }
                ExprKind::Ident(ref name) if name == "math" && field == "mininteger" => {
                    Some("math_mininteger")
                }
                ExprKind::Ident(ref name) if name == "math" && field == "huge" => {
                    expr.kind = ExprKind::Lit(Literal::Float(f64::INFINITY));
                    return;
                }
                _ => None,
            };
            if let Some(name) = math_const {
                let field = if name == "math_maxinteger" {
                    "maxinteger"
                } else {
                    "mininteger"
                };
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Ident("math".to_string()))),
                        field: field.to_string(),
                        null_safe: false,
                    })),
                    args: Vec::new(),
                    optional: false,
                };
                return;
            }
            normalize_expr(object);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_index".to_string()))),
                args: vec![
                    Argument::positional(object.as_ref().clone()),
                    Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(field.clone())))),
                ],
                optional: false,
            };
            expr.kind = call;
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_expr(cond);
            normalize_expr(then);
            normalize_expr(else_);
        }
        _ => {}
    }
}

fn lua_truthy(mut expr: Expression) -> Expression {
    normalize_expr(&mut expr);
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_truthy".to_string()))),
        args: vec![Argument::positional(expr)],
        optional: false,
    })
}

fn lua_write_stmt(mut target: Expression, mut value: Expression) -> Statement {
    if let Some(stmt) = lua_global_alias_write(&target, value.clone()) {
        return stmt;
    }

    match target.kind {
        ExprKind::Index {
            mut object,
            mut index,
            ..
        } => {
            normalize_expr(&mut object);
            normalize_expr(&mut index);
            normalize_expr(&mut value);
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_newindex".to_string()))),
                args: vec![
                    Argument::positional(*object),
                    Argument::positional(*index),
                    Argument::positional(value),
                ],
                optional: false,
            })))
        }
        ExprKind::Member {
            mut object, field, ..
        } => {
            normalize_expr(&mut object);
            normalize_expr(&mut value);
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_newindex".to_string()))),
                args: vec![
                    Argument::positional(*object),
                    Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(field)))),
                    Argument::positional(value),
                ],
                optional: false,
            })))
        }
        _ => {
            normalize_expr(&mut target);
            normalize_expr(&mut value);
            Statement::new(StmtKind::Assign {
                targets: vec![target],
                value,
            })
        }
    }
}

fn normalize_stmt(kind: &mut StmtKind) {
    match kind {
        StmtKind::Expr(expr) => {
            if let Some(alias_stmt) = lua_global_alias_rawset_stmt(expr) {
                *kind = alias_stmt;
                return;
            }
            normalize_expr(expr);
        }
        StmtKind::Assign { targets, value } => {
            if targets.len() > 1 {
                if let ExprKind::Array(elems) = &value.kind {
                    let mut temp_decls = Vec::new();
                    let mut assigns = Vec::new();
                    for (i, elem) in elems.iter().enumerate() {
                        let mut rhs = elem.value.clone();
                        normalize_expr(&mut rhs);
                        let temp_name = format!("__lua_assign_tmp_{i}");
                        temp_decls.push(VarDeclarator {
                            pattern: BindingPattern::Ident(temp_name.clone()),
                            type_hint: None,
                            init: Some(rhs),
                            array_bounds: None,
                            with_events: false,
                        });
                    }
                    assigns.push(Statement::new(StmtKind::VarDecl {
                        declarations: temp_decls,
                        kind: VarDeclKind::Let,
                    }));
                    for (i, target) in targets.iter().enumerate() {
                        let rhs = if i < elems.len() {
                            Expression::new(ExprKind::Ident(format!("__lua_assign_tmp_{i}")))
                        } else {
                            Expression::new(ExprKind::Lit(Literal::Null))
                        };
                        assigns.push(lua_write_stmt(target.clone(), rhs));
                    }
                    *kind = StmtKind::Block(assigns);
                    return;
                }
                if is_lua_multi_return_call(value) {
                    let mut call_value = value.clone();
                    normalize_expr(&mut call_value);
                    if let ExprKind::Call { args, .. } = &mut call_value.kind {
                        if args.len() == 1 && matches!(&args[0].value.kind, ExprKind::Call { .. }) {
                            call_value = args[0].value.clone();
                        }
                    }
                    let temp_name = "__lua_multi_tmp".to_string();
                    let mut assigns = vec![Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(temp_name.clone()),
                            type_hint: None,
                            init: Some(lua_multi_row(call_value)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    })];
                    for (i, target) in targets.iter().enumerate() {
                        assigns.push(lua_write_stmt(
                            target.clone(),
                            lua_multi_index(Expression::new(ExprKind::Ident(temp_name.clone())), i as i64),
                        ));
                    }
                    *kind = StmtKind::Block(assigns);
                    return;
                }
            }
            if targets.len() == 1 {
                let target = targets[0].clone();
                let rhs = value.clone();
                let stmt = lua_write_stmt(target, rhs);
                *kind = stmt.kind;
            } else {
                for t in targets {
                    normalize_expr(t);
                }
                normalize_expr(value);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            if declarations.len() == 1 {
                if let Some(decl) = declarations.first() {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(ExprKind::Call { callee, args, .. }) =
                            decl.init.as_ref().map(|init| &init.kind)
                        {
                            if lua_call_name(callee) == Some("coroutine.wrap") && args.len() == 1 {
                                let co_name = format!("__lua_wrap_co_{name}");
                                let create_arg = args[0].value.clone();
                                let co_decl = VarDeclarator {
                                        pattern: BindingPattern::Ident(co_name.clone()),
                                        type_hint: None,
                                        init: Some(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(lua_ident("coroutine")),
                                                field: "create".to_string(),
                                                null_safe: false,
                                            })),
                                            args: vec![Argument::positional(create_arg)],
                                            optional: false,
                                        })),
                                        array_bounds: None,
                                        with_events: false,
                                    };
                                let wrapper = Expression::new(ExprKind::Lambda {
                                    params: vec![Param {
                                        name: "...".to_string(),
                                        type_hint: None,
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: true,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    }],
                                    body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(lua_ident("__lua_wrap_resume")),
                                            args: vec![
                                                Argument::positional(lua_ident(co_name)),
                                                Argument {
                                                    name: None,
                                                    value: lua_ident("..."),
                                                    spread: true,
                                                    by_ref: false,
                                                },
                                            ],
                                            optional: false,
                                        }),
                                    )))]),
                                    is_async: false,
                                    captures: Vec::new(),
                                });
                                let wrapper_decl = VarDeclarator {
                                        pattern: BindingPattern::Ident(name.clone()),
                                        type_hint: decl.type_hint.clone(),
                                        init: Some(wrapper),
                                        array_bounds: decl.array_bounds.clone(),
                                        with_events: decl.with_events,
                                    };
                                *kind = StmtKind::VarDecl {
                                    declarations: vec![co_decl, wrapper_decl],
                                    kind: VarDeclKind::Let,
                                };
                                normalize_stmt(kind);
                                return;
                            }
                        }
                    }
                }
            }
            if declarations.len() > 1
                && let Some(first_init) = declarations.first().and_then(|decl| decl.init.as_ref())
                && is_lua_multi_return_call(first_init)
            {
                let mut call_value = first_init.clone();
                normalize_expr(&mut call_value);
                if let ExprKind::Call { args, .. } = &mut call_value.kind {
                    if args.len() == 1 && matches!(&args[0].value.kind, ExprKind::Call { .. }) {
                        call_value = args[0].value.clone();
                    }
                }
                let temp_name = "__lua_multi_tmp".to_string();
                let mut expanded = vec![VarDeclarator {
                    pattern: BindingPattern::Ident(temp_name.clone()),
                    type_hint: None,
                    init: Some(lua_multi_row(call_value)),
                    array_bounds: None,
                    with_events: false,
                }];
                for (i, decl) in declarations.iter().enumerate() {
                    let BindingPattern::Ident(name) = &decl.pattern else {
                        continue;
                    };
                    expanded.push(VarDeclarator {
                        pattern: BindingPattern::Ident(name.clone()),
                        type_hint: decl.type_hint.clone(),
                        init: Some(lua_multi_index(
                            Expression::new(ExprKind::Ident(temp_name.clone())),
                            i as i64,
                        )),
                        array_bounds: decl.array_bounds.clone(),
                        with_events: decl.with_events,
                    });
                }
                *kind = StmtKind::VarDecl {
                    declarations: expanded,
                    kind: VarDeclKind::Let,
                };
                return;
            }
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_expr(init);
                }
            }
        }
        StmtKind::Block(body) => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init_stmt) = init {
                normalize_stmt(&mut init_stmt.kind);
            }
            if let Some(c) = cond {
                normalize_expr(c);
            }
            if let Some(u) = update {
                normalize_expr(u);
            }
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_expr(iter);
            let is_gmatch = matches!(
                &iter.kind,
                ExprKind::Call { callee, .. }
                    if lua_call_name(callee).as_deref() == Some("string.gmatch")
            );
            let is_row_iterator = matches!(
                &iter.kind,
                ExprKind::Call { callee, .. }
                    if matches!(
                        lua_call_name(callee).as_deref(),
                        Some("string.gmatch" | "pairs" | "ipairs")
                    )
            );
            if is_gmatch || is_row_iterator {
                let first = var.clone();
                let second = key.take();
                let rows = format!("__lua_rows_{}", first);
                let idx = format!("__lua_idx_{}", first);
                let item = format!("__lua_item_{}", first);
                let pairs_table = if let ExprKind::Call { callee, args, .. } = &iter.kind {
                    if lua_call_name(callee).as_deref() == Some("pairs") {
                        args.first().map(|arg| (format!("__lua_pairs_table_{}", first), arg.value.clone()))
                    } else {
                        None
                    }
                } else {
                    None
                };
                let mut loop_body = vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(item.clone()),
                            type_hint: None,
                            init: Some(lua_multi_index_expr(
                                Expression::new(ExprKind::Ident(rows.clone())),
                                Expression::new(ExprKind::Ident(idx.clone())),
                            )),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(first),
                            type_hint: None,
                            init: Some(lua_multi_index(
                                Expression::new(ExprKind::Ident(item.clone())),
                                0,
                            )),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                ];
                if let Some(second) = second {
                    loop_body.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(second),
                            type_hint: None,
                            init: Some(lua_multi_index(Expression::new(ExprKind::Ident(item)), 1)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }));
                }
                loop_body.extend(std::mem::take(body));
                let rows_expr = if let Some((table_var, _)) = &pairs_table {
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("pairs".to_string()))),
                        args: vec![Argument::positional(Expression::new(ExprKind::Ident(table_var.clone())))],
                        optional: false,
                    })
                } else {
                    iter.clone()
                };
                let idx_expr = Expression::new(ExprKind::Ident(idx.clone()));
                let len_call = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".to_string()))),
                    args: vec![Argument::positional(Expression::new(ExprKind::Ident(rows.clone())))],
                    optional: false,
                });
                let mut block = Vec::new();
                if let Some((table_var, table_expr)) = &pairs_table {
                    block.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(table_var.clone()),
                            type_hint: None,
                            init: Some(table_expr.clone()),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }));
                }
                block.push(
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(rows),
                            type_hint: None,
                            init: Some(rows_expr),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    })
                );
                block.push(Statement::new(StmtKind::For {
                        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(idx.clone()),
                                type_hint: None,
                                init: Some(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }))),
                        cond: Some(Expression::new(ExprKind::Binary {
                            op: BinOp::Lt,
                            left: Box::new(idx_expr.clone()),
                            right: Box::new(len_call),
                        })),
                        update: Some(Expression::new(ExprKind::Assign {
                            target: Box::new(idx_expr.clone()),
                            value: Box::new(Expression::new(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(idx_expr),
                                right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
                            })),
                        })),
                        body: vec![Statement::new(StmtKind::Block(loop_body))],
                    }));
                if let Some((table_var, _)) = &pairs_table {
                    block.push(Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("__lua_iter_end".to_string()))),
                        args: vec![Argument::positional(Expression::new(ExprKind::Ident(table_var.clone())))],
                        optional: false,
                    }))));
                }
                *kind = StmtKind::Block(block);
                normalize_stmt(kind);
                return;
            }
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            if let Some(eb) = else_body {
                for s in eb.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
        }
        StmtKind::While { cond, body, .. } => {
            let wrapped = lua_truthy(cond.clone());
            *cond = wrapped;
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            let wrapped = lua_truthy(cond.clone());
            *cond = wrapped;
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let wrapped = lua_truthy(cond.clone());
            *cond = wrapped;
            for s in then_body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
            for (c, body) in elifs.iter_mut() {
                let wrapped = lua_truthy(c.clone());
                *c = wrapped;
                for s in body.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
            if let Some(body) = else_body {
                for s in body.iter_mut() {
                    normalize_stmt(&mut s.kind);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body.iter_mut() {
                normalize_stmt(&mut s.kind);
            }
        }
        StmtKind::Return(Some(expr)) => {
            if is_lua_multi_return_call(expr) {
                let mut value = expr.clone();
                normalize_expr(&mut value);
                if let ExprKind::Call { args, .. } = &mut value.kind {
                    if args.len() == 1 && matches!(&args[0].value.kind, ExprKind::Call { .. }) {
                        value = args[0].value.clone();
                    }
                }
                *expr = lua_multi_row(value);
            } else {
                normalize_expr(expr);
            }
        }
        _ => {}
    }
}

/// Lua numeric for -> canonical C-style for loop.
pub(crate) fn build_numeric_for(
    index_var: String,
    start: Expression,
    limit: Expression,
    step: Expression,
    body: Vec<Statement>,
) -> StmtKind {
    let ctrl_var = format!("__lua_for_{}", index_var);
    let limit_var = format!("__lua_limit_{}", index_var);
    let step_var = format!("__lua_step_{}", index_var);

    let ctrl_expr = Expression::new(ExprKind::Ident(ctrl_var.clone()));
    let limit_expr = Expression::new(ExprKind::Ident(limit_var.clone()));
    let step_expr_cond = Expression::new(ExprKind::Ident(step_var.clone()));

    // Lua numeric for. A zero step must not hang the VM: run only when the
    // initial value equals the limit, then force the update past the limit.
    let step_zero = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(step_expr_cond.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
    });

    let step_pos = Expression::new(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(step_expr_cond.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
    });

    let ctrl_lte_limit = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr.clone()),
    });

    let ctrl_gte_limit = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr.clone()),
    });

    let ctrl_eq_limit = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(ctrl_expr.clone()),
        right: Box::new(limit_expr.clone()),
    });

    // Lua numeric for loop decision without `and`/`or` value semantics:
    // step == 0 ? ctrl == limit : (step > 0 ? ctrl <= limit : ctrl >= limit)
    let cond = Expression::new(ExprKind::Ternary {
        cond: Box::new(step_zero.clone()),
        then: Box::new(ctrl_eq_limit),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(step_pos),
            then: Box::new(ctrl_lte_limit),
            else_: Box::new(ctrl_gte_limit),
        })),
    });

    // Bind user-visible loop variable each iteration
    let bind_loop_var = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(index_var.clone()),
            type_hint: None,
            init: Some(ctrl_expr.clone()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut scoped_body = vec![bind_loop_var];
    scoped_body.extend(body);

    let increment = Expression::new(ExprKind::Assign {
        target: Box::new(ctrl_expr.clone()),
        value: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(step_zero),
            then: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(limit_expr),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
            })),
            else_: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(ctrl_expr),
                right: Box::new(Expression::new(ExprKind::Ident(step_var.clone()))),
            })),
        })),
    });

    StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![
                VarDeclarator {
                    pattern: BindingPattern::Ident(ctrl_var.clone()),
                    type_hint: None,
                    init: Some(start),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(limit_var),
                    type_hint: None,
                    init: Some(limit),
                    array_bounds: None,
                    with_events: false,
                },
                VarDeclarator {
                    pattern: BindingPattern::Ident(step_var.clone()),
                    type_hint: None,
                    init: Some(step),
                    array_bounds: None,
                    with_events: false,
                },
            ],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(cond),
        update: Some(increment),
        body: vec![Statement::new(StmtKind::Block(scoped_body))],
    }
}
