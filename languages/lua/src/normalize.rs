//! Lua → JS-shaped AST normalization.
//! Normalizes Lua-specific operations to adapter calls that handle metamethods at runtime.

use vybe_ast::*;

const LUA_VARARGS: &str = "_lua_varargs";

fn lua_varargs() -> Expression {
    lua_ident(LUA_VARARGS)
}

fn lua_for_key_vars(key: Option<String>) -> Vec<String> {
    key.map(|value| {
        value
            .split("__lua_extra__")
            .filter(|part| !part.is_empty())
            .map(str::to_string)
            .collect()
    })
    .unwrap_or_default()
}

/// Normalize module: transform statements and expressions
pub fn normalize_module(module: &mut Module) {
    normalize_lua_stmt_sequence(&mut module.body);
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
                "format" => "string.format",
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
                "pack" => "table.pack",
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
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "debug") =>
        {
            Some(match field.as_str() {
                "gethook" => "debug.gethook",
                "getinfo" => "debug.getinfo",
                "getlocal" => "debug.getlocal",
                "getupvalue" => "debug.getupvalue",
                "sethook" => "debug.sethook",
                "setlocal" => "debug.setlocal",
                "setupvalue" => "debug.setupvalue",
                "traceback" => "debug.traceback",
                "type" => "debug.type",
                "upvalueid" => "debug.upvalueid",
                "upvaluejoin" => "debug.upvaluejoin",
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

fn lua_destructure_target(names: &[String]) -> Expression {
    Expression::new(ExprKind::Destructure(DestructurePattern::Array(
        names
            .iter()
            .map(|name| ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None))
            .collect(),
    )))
}

fn lua_decl_ident_names(declarations: &[VarDeclarator]) -> Option<Vec<String>> {
    declarations
        .iter()
        .map(|decl| match &decl.pattern {
            BindingPattern::Ident(name) => Some(name.clone()),
            _ => None,
        })
        .collect()
}

fn lua_target_ident_names(targets: &[Expression]) -> Option<Vec<String>> {
    targets
        .iter()
        .map(|target| match &target.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        })
        .collect()
}

fn lua_decl_init_is_empty(decl: &VarDeclarator) -> bool {
    decl.init
        .as_ref()
        .is_none_or(|expr| matches!(expr.kind, ExprKind::Lit(Literal::Null)))
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
        ExprKind::Ident(name) if name == LUA_VARARGS => true,
        ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)) => true,
        ExprKind::Call { callee, args, .. } => match lua_call_name(callee) {
            Some("string.find" | "string.gsub" | "string.match") => true,
            Some(
                "string.unpack"
                | "table.unpack"
                | "math.modf"
                | "debug.gethook"
                | "debug.getlocal"
                | "debug.getupvalue",
            ) => true,
            Some("coroutine.resume" | "coroutine.running" | "coroutine.yield" | "__lua_wrap_resume") => true,
            Some("next" | "pcall" | "xpcall") => true,
            Some("select") => !matches!(
                args.first().map(|arg| &arg.value.kind),
                Some(ExprKind::Lit(Literal::Str(value))) if value == "#"
            ),
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
            if name == "__lua_first"
                || name == "__lua_index0"
                || name == "__lua_multi_get0"
                || name == "__lua_multi_row"
                || name == "__lua_as_multi_row"
    )
}

fn is_lua_direct_identifier_call(name: &str) -> bool {
    name.starts_with("__lua_")
        || matches!(
            name,
            "print"
                | "tonumber"
                | "tostring"
                | "type"
                | "rawlen"
                | "rawget"
                | "rawset"
                | "assert"
                | "setmetatable"
                | "getmetatable"
                | "pairs"
                | "ipairs"
                | "next"
                | "pcall"
                | "xpcall"
                | "error"
                | "select"
        )
}

fn lua_multi_row(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_multi_row".to_string()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_as_multi_row(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_as_multi_row".to_string()))),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_mark_rest(value: Expression, fixed_count: usize) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_mark_rest")),
        args: vec![
            Argument::positional(value),
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(
                fixed_count as i64,
            )))),
        ],
        optional: false,
    })
}

fn lua_lower_rest_param(params: &mut [Param]) -> Option<usize> {
    let fixed_count = params.len().checked_sub(1)?;
    let last = params.last_mut()?;
    if !last.is_rest {
        return None;
    }
    last.is_rest = false;
    Some(fixed_count)
}

fn lua_multi_row_from_values(values: Vec<Expression>) -> Expression {
    lua_multi_row(lua_array_from_values(values))
}

fn lua_array_from_values(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
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
        ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)) => {
            expr.kind = lua_varargs().kind;
        }
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
        callee: Box::new(Expression::new(ExprKind::Ident("__lua_multi_get0".to_string()))),
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
            let callee_is_lua_internal = matches!(
                &callee.kind,
                ExprKind::Ident(name) if name.starts_with("__lua_") && name != "__lua_print"
            );
            if !internal_multi_helper {
                let last_arg_index = args.len().saturating_sub(1);
                for (arg_index, arg) in args.iter_mut().enumerate() {
                    match &arg.value.kind {
                        ExprKind::Spread(inner)
                            if matches!(inner.kind, ExprKind::Lit(Literal::Null))
                                && arg_index == last_arg_index =>
                        {
                            arg.value = lua_varargs();
                            arg.spread = true;
                            continue;
                        }
                        ExprKind::Spread(inner)
                            if matches!(inner.kind, ExprKind::Lit(Literal::Null)) =>
                        {
                            arg.value = lua_first(lua_multi_row(lua_varargs()));
                        }
                        _ => {}
                    }
                    normalize_expr(&mut arg.value);
                }
                if call_name_before.as_deref() == Some("error") {
                    if let Some(last) = args.last_mut() {
                        if last.name.is_none()
                            && !last.spread
                            && is_lua_multi_return_call(&last.value)
                        {
                            let value = last.value.clone();
                            last.value = lua_multi_index(value, 0);
                        }
                    }
                } else if !callee_is_lua_internal {
                    mark_last_lua_multi_return_arg_spread(args);
                }
            }
            if internal_multi_helper {
                let helper_name = lua_call_name(callee);
                for (arg_index, arg) in args.iter_mut().enumerate() {
                    if arg_index == 0
                        && matches!(
                            helper_name.as_deref(),
                            Some(
                                "__lua_first"
                                    | "__lua_multi_get0"
                                    | "__lua_multi_row"
                                    | "__lua_as_multi_row"
                            )
                        )
                    {
                        normalize_lua_multi_return_source(&mut arg.value);
                    } else {
                        normalize_expr(&mut arg.value);
                    }
                }
            }
            if lua_call_name(callee).as_deref() == Some("table.sort") && args.len() >= 2 {
                normalize_lua_sort_comparator(&mut args[1]);
            }
            if lua_call_name(callee).as_deref() == Some("select") && args.len() == 2 && args[1].spread {
                args[1].spread = false;
                args[1].value = lua_multi_row(args[1].value.clone());
            }
            if is_lua_math_member(callee, "type") && args.len() == 1 && !args[0].spread {
                if let Some(kind) = lua_static_math_type_arg(&args[0].value) {
                    expr.kind = ExprKind::Lit(Literal::Str(kind.to_string()));
                    return;
                }
            }
            if args.last().is_some_and(|arg| arg.name.is_none() && arg.spread) {
                let row = args.last().map(|arg| arg.value.clone()).unwrap();
                if lua_call_name(callee).as_deref() == Some("table.pack") && args.len() == 1 {
                    expr.kind = ExprKind::Call {
                        callee: Box::new(lua_ident("__lua_table_pack_row")),
                        args: vec![Argument::positional(row)],
                        optional: false,
                    };
                    return;
                }
                if lua_call_name(callee).as_deref() == Some("string.format") && args.len() >= 2 {
                    let prefix = args[..args.len() - 1]
                        .iter()
                        .map(|arg| arg.value.clone())
                        .collect::<Vec<_>>();
                    expr.kind = ExprKind::Call {
                        callee: Box::new(lua_ident("__lua_string_format_row")),
                        args: vec![
                            Argument::positional(lua_array_from_values(prefix)),
                            Argument::positional(row),
                        ],
                        optional: false,
                    };
                    return;
                }
            }
            let callee_name = if let ExprKind::Ident(name) = &callee.kind {
                Some(name.as_str())
            } else {
                None
            };
            let direct_callee_name = callee_name.or(call_name_before.as_deref());
            if matches!(direct_callee_name, Some("print") | Some("__lua_print")) {
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
            } else if matches!(direct_callee_name, Some("tostring") | Some("__lua_tostring")) {
                if let Some(arg) = args.first_mut() {
                    wrap_lua_float_display_arg(arg);
                }
            } else if args.last().is_some_and(|arg| arg.name.is_none() && arg.spread) {
                let fn_expr = (**callee).clone();
                let row = args.last().map(|arg| arg.value.clone()).unwrap();
                if args.len() > 1 {
                    let prefix = args[..args.len() - 1]
                        .iter()
                        .map(|arg| arg.value.clone())
                        .collect::<Vec<_>>();
                    expr.kind = ExprKind::Call {
                        callee: Box::new(lua_ident("__lua_apply_row_prefix")),
                        args: vec![
                            Argument::positional(fn_expr),
                            Argument::positional(lua_array_from_values(prefix)),
                            Argument::positional(row),
                        ],
                        optional: false,
                    };
                    return;
                }
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident(
                        "__lua_apply_row".to_string(),
                    ))),
                    args: vec![Argument::positional(fn_expr), Argument::positional(row)],
                    optional: false,
                };
            } else if let ExprKind::Ident(name) = &callee.kind {
                if !is_lua_direct_identifier_call(name) {
                    let fn_expr = (**callee).clone();
                    let mut call_args = Vec::with_capacity(args.len() + 1);
                    call_args.push(Argument::positional(fn_expr));
                    call_args.extend(std::mem::take(args));
                    expr.kind = ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Ident("__lua_call".to_string()))),
                        args: call_args,
                        optional: false,
                    };
                }
            } else if !keep_profile_member {
                let fn_expr = (**callee).clone();
                let mut call_args = Vec::with_capacity(args.len() + 1);
                call_args.push(Argument::positional(fn_expr));
                call_args.extend(std::mem::take(args));
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__lua_call".to_string()))),
                    args: call_args,
                    optional: false,
                };
            } else if is_lua_multi_return_call(expr) {
                let value = expr.clone();
                expr.kind = lua_multi_index(value, 0).kind;
            }
        }
        ExprKind::Array(elems) => {
            let last_elem_index = elems.len().saturating_sub(1);
            for (elem_index, elem) in elems.iter_mut().enumerate() {
                if elem.key.is_none() && !elem.spread {
                    match &elem.value.kind {
                        ExprKind::Spread(inner)
                            if matches!(inner.kind, ExprKind::Lit(Literal::Null))
                                && elem_index == last_elem_index =>
                        {
                            elem.value = lua_varargs();
                            elem.spread = true;
                        }
                        ExprKind::Spread(inner)
                            if matches!(inner.kind, ExprKind::Lit(Literal::Null)) =>
                        {
                            elem.value = lua_first(lua_multi_row(lua_varargs()));
                            normalize_expr(&mut elem.value);
                        }
                        ExprKind::Spread(inner) if elem_index == last_elem_index => {
                            elem.value = inner.as_ref().clone();
                            normalize_expr(&mut elem.value);
                            elem.spread = true;
                        }
                        ExprKind::Spread(inner) => {
                            elem.value = lua_multi_index(inner.as_ref().clone(), 0);
                            normalize_expr(&mut elem.value);
                        }
                        _ if is_lua_multi_return_call(&elem.value) && elem_index == last_elem_index => {
                            let mut value = elem.value.clone();
                            normalize_lua_multi_return_source(&mut value);
                            elem.value = lua_multi_row(value);
                            elem.spread = true;
                        }
                        _ if is_lua_multi_return_call(&elem.value) => {
                            let value = elem.value.clone();
                            elem.value = lua_multi_index(value, 0);
                            normalize_expr(&mut elem.value);
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
        ExprKind::Lambda { params, body, .. } => {
            match body {
                LambdaBody::Expr(expr) => {
                    rewrite_lua_current_getinfo_expr(expr, params);
                    normalize_expr(expr);
                }
                LambdaBody::Block(stmts) => {
                    rewrite_lua_current_getinfo_stmts(stmts, params);
                    normalize_lua_stmt_sequence(stmts);
                }
            }
            if let Some(fixed_count) = lua_lower_rest_param(params) {
                let lambda = expr.clone();
                expr.kind = lua_mark_rest(lambda, fixed_count).kind;
            }
        }
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

fn lua_int_literal(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) if value.fract() == 0.0 => Some(*value as i64),
        _ => None,
    }
}

fn collect_lua_local_decls(kind: &StmtKind, locals: &mut Vec<String>) {
    if let StmtKind::VarDecl { declarations, .. } = kind {
        for decl in declarations {
            if let BindingPattern::Ident(name) = &decl.pattern {
                locals.push(name.clone());
            }
        }
    }
}

fn lua_debug_setlocal_assignment(expr: &Expression, locals: &[String]) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if lua_call_name(callee) != Some("debug.setlocal") || args.len() < 3 {
        return None;
    }
    if lua_int_literal(&args[0].value)? != 1 {
        return None;
    }
    let index = lua_int_literal(&args[1].value)?;
    if index <= 0 {
        return None;
    }
    let name = locals.get(index as usize - 1)?.clone();
    Some(StmtKind::Assign {
        targets: vec![lua_ident(name)],
        value: args[2].value.clone(),
    })
}

fn lua_debug_setlocal_decl_block(kind: &StmtKind, locals: &[String]) -> Option<StmtKind> {
    let StmtKind::VarDecl { declarations, kind } = kind else {
        return None;
    };
    if declarations.len() != 1 {
        return None;
    }
    let decl = &declarations[0];
    let Some(init) = &decl.init else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &init.kind else {
        return None;
    };
    if lua_call_name(callee) != Some("debug.setlocal") || args.len() < 3 {
        return None;
    }
    if lua_int_literal(&args[0].value)? != 1 {
        return None;
    }
    let index = lua_int_literal(&args[1].value)?;
    if index <= 0 {
        return None;
    }
    let name = locals.get(index as usize - 1)?.clone();
    Some(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: decl.pattern.clone(),
            type_hint: decl.type_hint.clone(),
            init: Some(Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Assign {
                    target: Box::new(lua_ident(name.clone())),
                    value: Box::new(args[2].value.clone()),
                }),
                Expression::new(ExprKind::Lit(Literal::Str(name))),
            ]))),
            array_bounds: decl.array_bounds.clone(),
            with_events: decl.with_events,
        }],
        kind: kind.clone(),
    })
}

fn lua_debug_setupvalue_decl_block(kind: &StmtKind, locals: &[String]) -> Option<StmtKind> {
    let StmtKind::VarDecl { declarations, kind } = kind else {
        return None;
    };
    if declarations.len() != 1 {
        return None;
    }
    let decl = &declarations[0];
    let Some(init) = &decl.init else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &init.kind else {
        return None;
    };
    if lua_call_name(callee) != Some("debug.setupvalue") || args.len() < 3 {
        return None;
    }
    if lua_int_literal(&args[1].value)? != 1 {
        return None;
    }
    let name = locals.first()?.clone();
    Some(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: decl.pattern.clone(),
            type_hint: decl.type_hint.clone(),
            init: Some(Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Assign {
                    target: Box::new(lua_ident(name.clone())),
                    value: Box::new(args[2].value.clone()),
                }),
                Expression::new(ExprKind::Lit(Literal::Str(name))),
            ]))),
            array_bounds: decl.array_bounds.clone(),
            with_events: decl.with_events,
        }],
        kind: kind.clone(),
    })
}

fn lua_debug_getinfo_static(params: &[Param]) -> Expression {
    let nparams = params.iter().filter(|param| !param.is_rest).count() as i64;
    let isvararg = params.iter().any(|param| param.is_rest);
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_debug_getinfo_static")),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(nparams)))),
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Bool(isvararg)))),
        ],
        optional: false,
    })
}

fn rewrite_lua_current_getinfo_expr(expr: &mut Expression, params: &[Param]) {
    if let ExprKind::Call { callee, args, .. } = &mut expr.kind
        && lua_call_name(callee) == Some("debug.getinfo")
        && matches!(
            args.first().and_then(|arg| lua_int_literal(&arg.value)),
            Some(1)
        )
    {
        *expr = lua_debug_getinfo_static(params);
        return;
    }

    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_lua_current_getinfo_expr(left, params);
            rewrite_lua_current_getinfo_expr(right, params);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => rewrite_lua_current_getinfo_expr(expr, params),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_lua_current_getinfo_expr(cond, params);
            rewrite_lua_current_getinfo_expr(then, params);
            rewrite_lua_current_getinfo_expr(else_, params);
        }
        ExprKind::Member { object, .. } => rewrite_lua_current_getinfo_expr(object, params),
        ExprKind::Index { object, index, .. } => {
            rewrite_lua_current_getinfo_expr(object, params);
            rewrite_lua_current_getinfo_expr(index, params);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_lua_current_getinfo_expr(callee, params);
            for arg in args {
                rewrite_lua_current_getinfo_expr(&mut arg.value, params);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_lua_current_getinfo_expr(class, params);
            for arg in args {
                rewrite_lua_current_getinfo_expr(&mut arg.value, params);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_lua_current_getinfo_expr(target, params);
            rewrite_lua_current_getinfo_expr(value, params);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    rewrite_lua_current_getinfo_expr(key, params);
                }
                rewrite_lua_current_getinfo_expr(&mut elem.value, params);
            }
        }
        ExprKind::Tuple(values) | ExprKind::Set(values) | ExprKind::Sequence(values) => {
            for value in values {
                rewrite_lua_current_getinfo_expr(value, params);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                rewrite_lua_current_getinfo_expr(value, params);
            }
        }
        ExprKind::Yield(Some(value)) => rewrite_lua_current_getinfo_expr(value, params),
        ExprKind::Lambda { params: inner_params, body, .. } => match body {
            LambdaBody::Expr(value) => rewrite_lua_current_getinfo_expr(value, inner_params),
            LambdaBody::Block(stmts) => rewrite_lua_current_getinfo_stmts(stmts, inner_params),
        },
        _ => {}
    }
}

fn rewrite_lua_current_getinfo_stmts(stmts: &mut [Statement], params: &[Param]) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                rewrite_lua_current_getinfo_expr(expr, params);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        rewrite_lua_current_getinfo_expr(init, params);
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    rewrite_lua_current_getinfo_expr(target, params);
                }
                rewrite_lua_current_getinfo_expr(value, params);
            }
            StmtKind::Block(body)
            | StmtKind::For { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. } => rewrite_lua_current_getinfo_stmts(body, params),
            StmtKind::ForIn { body, else_body, .. } => {
                rewrite_lua_current_getinfo_stmts(body, params);
                if let Some(else_body) = else_body {
                    rewrite_lua_current_getinfo_stmts(else_body, params);
                }
            }
            StmtKind::If { cond, then_body, elifs, else_body } => {
                rewrite_lua_current_getinfo_expr(cond, params);
                rewrite_lua_current_getinfo_stmts(then_body, params);
                for (cond, body) in elifs {
                    rewrite_lua_current_getinfo_expr(cond, params);
                    rewrite_lua_current_getinfo_stmts(body, params);
                }
                if let Some(else_body) = else_body {
                    rewrite_lua_current_getinfo_stmts(else_body, params);
                }
            }
            _ => {}
        }
    }
}

fn is_lua_synthetic_multi_decl_block(kind: &StmtKind) -> bool {
    let StmtKind::Block(stmts) = kind else {
        return false;
    };
    if let Some(Statement {
        kind:
            StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let,
            },
        ..
    }) = stmts.first()
    {
        if declarations.iter().any(|decl| {
            matches!(&decl.pattern, BindingPattern::Ident(name) if name.starts_with("__lua_"))
        }) {
            return true;
        }
    }
    let [
        Statement {
            kind:
                StmtKind::VarDecl {
                    kind: VarDeclKind::Let,
                    ..
                },
            ..
        },
        Statement {
            kind: StmtKind::Assign { targets, .. },
            ..
        },
    ] = stmts.as_slice()
    else {
        return false;
    };
    targets.len() == 1
        && matches!(
            targets[0].kind,
            ExprKind::Destructure(DestructurePattern::Array(_))
        )
}

fn take_lua_synthetic_multi_decl_block(kind: &mut StmtKind) -> Option<Vec<Statement>> {
    if let StmtKind::Block(stmts) = kind {
        if stmts.len() >= 2
            && matches!(
                stmts.first().map(|stmt| &stmt.kind),
                Some(StmtKind::VarDecl {
                    kind: VarDeclKind::Let,
                    ..
                })
            )
            && is_lua_synthetic_multi_decl_block(&stmts[1].kind)
        {
            let mut stmts = match std::mem::replace(kind, StmtKind::Block(Vec::new())) {
                StmtKind::Block(stmts) => stmts,
                other => {
                    *kind = other;
                    return None;
                }
            };
            let mut flattened = Vec::new();
            flattened.push(stmts.remove(0));
            if let StmtKind::Block(inner) = stmts.remove(0).kind {
                flattened.extend(inner);
            }
            flattened.extend(stmts);
            return Some(flattened);
        }
    }
    if !is_lua_synthetic_multi_decl_block(kind) {
        return None;
    }
    match std::mem::replace(kind, StmtKind::Block(Vec::new())) {
        StmtKind::Block(stmts) => Some(stmts),
        other => {
            *kind = other;
            None
        }
    }
}

fn normalize_lua_stmt_sequence(body: &mut Vec<Statement>) {
    let mut locals = Vec::new();
    let mut i = 0;
    while i < body.len() {
        {
            let stmt = &mut body[i];
            if let StmtKind::Expr(expr) = &stmt.kind
                && let Some(assign) = lua_debug_setlocal_assignment(expr, &locals)
            {
                stmt.kind = assign;
            }
            if let Some(block) = lua_debug_setlocal_decl_block(&stmt.kind, &locals) {
                stmt.kind = block;
            } else if let Some(block) = lua_debug_setupvalue_decl_block(&stmt.kind, &locals) {
                stmt.kind = block;
            }
        }
        if let Some(stmts) = take_lua_synthetic_multi_decl_block(&mut body[i].kind) {
            body.splice(i..=i, stmts);
            continue;
        }
        {
            let stmt = &mut body[i];
            normalize_stmt(&mut stmt.kind);
        }
        if let Some(stmts) = take_lua_synthetic_multi_decl_block(&mut body[i].kind) {
            body.splice(i..=i, stmts);
            continue;
        }
        collect_lua_local_decls(&body[i].kind, &mut locals);
        i += 1;
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
            if targets.len() == 1
                && let ExprKind::Destructure(DestructurePattern::Array(patterns)) = &targets[0].kind
            {
                if is_lua_multi_return_call(value) {
                    let mut call_value = value.clone();
                    normalize_lua_multi_return_source(&mut call_value);
                    let temp_name = "__lua_multi_tmp".to_string();
                    let mut assigns = vec![Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(temp_name.clone()),
                            type_hint: None,
                            init: Some(lua_as_multi_row(call_value)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    })];
                    for (i, pattern) in patterns.iter().enumerate() {
                        if let ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) = pattern {
                            assigns.push(lua_write_stmt(
                                lua_ident(name.clone()),
                                lua_multi_index(
                                    Expression::new(ExprKind::Ident(temp_name.clone())),
                                    i as i64,
                                ),
                            ));
                        }
                    }
                    *kind = StmtKind::Block(assigns);
                    return;
                }
                if let ExprKind::Call { args, .. } = &mut value.kind {
                    for arg in args.iter_mut() {
                        normalize_expr(&mut arg.value);
                    }
                    return;
                }
            }
            if targets.len() > 1 {
                if let ExprKind::Array(elems) = &value.kind {
                    let mut temp_decls = Vec::new();
                    let mut assigns = Vec::new();
                    let last_elem_index = elems.len().saturating_sub(1);
                    let mut temp_is_row = Vec::with_capacity(elems.len());
                    for (i, elem) in elems.iter().enumerate() {
                        let mut rhs = elem.value.clone();
                        let may_return_multi = is_lua_multi_return_call(&rhs)
                            || matches!(rhs.kind, ExprKind::Call { .. });
                        let is_last_row = i == last_elem_index && may_return_multi;
                        if is_last_row {
                            if is_lua_multi_return_call(&rhs) {
                                normalize_lua_multi_return_source(&mut rhs);
                                rhs = lua_as_multi_row(rhs);
                            } else {
                                normalize_expr(&mut rhs);
                                rhs = lua_as_multi_row(rhs);
                            }
                        } else if may_return_multi {
                            if is_lua_multi_return_call(&rhs) {
                                normalize_lua_multi_return_source(&mut rhs);
                                rhs = lua_first(lua_as_multi_row(rhs));
                            } else {
                                normalize_expr(&mut rhs);
                                rhs = lua_first(lua_as_multi_row(rhs));
                            }
                        } else {
                            normalize_expr(&mut rhs);
                        }
                        let temp_name = format!("__lua_assign_tmp_{i}");
                        temp_decls.push(VarDeclarator {
                            pattern: BindingPattern::Ident(temp_name.clone()),
                            type_hint: None,
                            init: Some(rhs),
                            array_bounds: None,
                            with_events: false,
                        });
                        temp_is_row.push(is_last_row);
                    }
                    assigns.push(Statement::new(StmtKind::VarDecl {
                        declarations: temp_decls,
                        kind: VarDeclKind::Let,
                    }));
                    for (i, target) in targets.iter().enumerate() {
                        let rhs = if i < elems.len() {
                            let temp = Expression::new(ExprKind::Ident(format!("__lua_assign_tmp_{i}")));
                            if temp_is_row.get(i).copied().unwrap_or(false) {
                                lua_multi_index(temp, 0)
                            } else {
                                temp
                            }
                        } else if temp_is_row.last().copied().unwrap_or(false) {
                            lua_multi_index(
                                Expression::new(ExprKind::Ident(format!(
                                    "__lua_assign_tmp_{}",
                                    elems.len().saturating_sub(1)
                                ))),
                                (i - elems.len() + 1) as i64,
                            )
                        } else {
                            Expression::new(ExprKind::Lit(Literal::Null))
                        };
                        assigns.push(lua_write_stmt(target.clone(), rhs));
                    }
                    *kind = StmtKind::Block(assigns);
                    return;
                }
                if matches!(value.kind, ExprKind::Call { .. })
                    && !is_lua_multi_return_call(value)
                    && let Some(names) = lua_target_ident_names(targets)
                {
                    *kind = StmtKind::Assign {
                        targets: vec![lua_destructure_target(&names)],
                        value: value.clone(),
                    };
                    return;
                }
                if is_lua_multi_return_call(value) {
                    let mut call_value = value.clone();
                    normalize_lua_multi_return_source(&mut call_value);
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
                                        name: LUA_VARARGS.to_string(),
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
                                                    value: lua_varargs(),
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
            if declarations.len() == 1 {
                if let Some(decl) = declarations.first_mut() {
                    if matches!(&decl.pattern, BindingPattern::Ident(name) if name.starts_with("__lua_")) {
                        if let Some(init) = &mut decl.init {
                            normalize_expr(init);
                        }
                        return;
                    }
                    if let Some(init) = &mut decl.init {
                        if is_lua_multi_return_call(init) {
                            let mut call_value = init.clone();
                            normalize_lua_multi_return_source(&mut call_value);
                            decl.init = Some(lua_first(lua_as_multi_row(call_value)));
                            return;
                        }
                        if matches!(init.kind, ExprKind::Call { .. }) {
                            let mut call_value = init.clone();
                            normalize_expr(&mut call_value);
                            decl.init = Some(lua_first(lua_as_multi_row(call_value)));
                            return;
                        }
                    }
                }
            }
            if declarations.len() > 1
                && let Some(first_init) = declarations.first().and_then(|decl| decl.init.as_ref())
                && matches!(first_init.kind, ExprKind::Call { .. })
                && !is_lua_multi_return_call(first_init)
                && declarations.iter().skip(1).all(lua_decl_init_is_empty)
                && let Some(names) = lua_decl_ident_names(declarations)
            {
                let decls = declarations
                    .iter()
                    .map(|decl| VarDeclarator {
                        pattern: decl.pattern.clone(),
                        type_hint: decl.type_hint.clone(),
                        init: None,
                        array_bounds: decl.array_bounds.clone(),
                        with_events: decl.with_events,
                    })
                    .collect();
                *kind = StmtKind::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: decls,
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![lua_destructure_target(&names)],
                        value: first_init.clone(),
                    }),
                ]);
                return;
            }
            if declarations.len() > 1
                && declarations.iter().any(|decl| {
                    decl.init.as_ref().is_some_and(|init| {
                        is_lua_multi_return_call(init)
                            || matches!(init.kind, ExprKind::Call { .. })
                    })
                })
            {
                let last_init_index = declarations
                    .iter()
                    .rposition(|decl| decl.init.is_some())
                    .unwrap_or(0);
                let last_init_may_return_multi = declarations[last_init_index]
                    .init
                    .as_ref()
                    .is_some_and(|init| {
                        is_lua_multi_return_call(init)
                            || matches!(init.kind, ExprKind::Call { .. })
                    });
                let row_name = "__lua_local_multi_tmp".to_string();
                let mut expanded = Vec::new();
                if last_init_may_return_multi && last_init_index + 1 < declarations.len() {
                    let mut last_init = declarations[last_init_index].init.clone().unwrap();
                    if is_lua_multi_return_call(&last_init) {
                        normalize_lua_multi_return_source(&mut last_init);
                        last_init = lua_as_multi_row(last_init);
                    } else {
                        normalize_expr(&mut last_init);
                        last_init = lua_as_multi_row(last_init);
                    }
                    expanded.push(VarDeclarator {
                        pattern: BindingPattern::Ident(row_name.clone()),
                        type_hint: None,
                        init: Some(last_init),
                        array_bounds: None,
                        with_events: false,
                    });
                }

                for (i, decl) in declarations.iter().enumerate() {
                    let mut init = if last_init_may_return_multi && i >= last_init_index {
                        Some(lua_multi_index(
                            Expression::new(ExprKind::Ident(row_name.clone())),
                            (i - last_init_index) as i64,
                        ))
                    } else if let Some(mut value) = decl.init.clone() {
                        let may_return_multi = is_lua_multi_return_call(&value)
                            || matches!(value.kind, ExprKind::Call { .. });
                        if may_return_multi {
                            if is_lua_multi_return_call(&value) {
                                normalize_lua_multi_return_source(&mut value);
                                Some(lua_first(lua_as_multi_row(value)))
                            } else {
                                normalize_expr(&mut value);
                                Some(lua_first(lua_as_multi_row(value)))
                            }
                        } else {
                            normalize_expr(&mut value);
                            Some(value)
                        }
                    } else {
                        None
                    };
                    if i > last_init_index && !last_init_may_return_multi {
                        init = None;
                    }
                    expanded.push(VarDeclarator {
                        pattern: decl.pattern.clone(),
                        type_hint: decl.type_hint.clone(),
                        init,
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
            if declarations.len() > 1
                && let Some(first_init) = declarations.first().and_then(|decl| decl.init.as_ref())
                && is_lua_multi_return_call(first_init)
                && declarations.iter().skip(1).all(|decl| decl.init.is_none())
            {
                let mut call_value = first_init.clone();
                normalize_lua_multi_return_source(&mut call_value);
                if let ExprKind::Call { args, .. } = &mut call_value.kind {
                    if args.len() == 1 && matches!(&args[0].value.kind, ExprKind::Call { .. }) {
                        call_value = args[0].value.clone();
                    }
                }
                let temp_name = "__lua_multi_tmp".to_string();
                let mut expanded = vec![VarDeclarator {
                    pattern: BindingPattern::Ident(temp_name.clone()),
                    type_hint: None,
                    init: Some(lua_as_multi_row(call_value)),
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
            normalize_lua_stmt_sequence(body);
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
            if let ExprKind::Sequence(parts) = &iter.kind {
                if !parts.is_empty() {
                    let first = var.clone();
                    let extra_vars = lua_for_key_vars(key.take());
                    let iter_name = format!("__lua_iter_{}", first);
                    let state_name = format!("__lua_state_{}", first);
                    let ctrl_name = format!("__lua_ctrl_{}", first);
                    let row_name = format!("__lua_row_{}", first);

                    let iter_expr = parts[0].clone();
                    let state_expr = parts
                        .get(1)
                        .cloned()
                        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));
                    let ctrl_expr = parts
                        .get(2)
                        .cloned()
                        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)));

                    let iter_callee = if matches!(&iter_expr.kind, ExprKind::Ident(name) if name == "next") {
                        iter_expr.clone()
                    } else {
                        Expression::new(ExprKind::Ident(iter_name.clone()))
                    };
                    let iter_call = Expression::new(ExprKind::Call {
                        callee: Box::new(iter_callee),
                        args: vec![
                            Argument::positional(Expression::new(ExprKind::Ident(state_name.clone()))),
                            Argument::positional(Expression::new(ExprKind::Ident(ctrl_name.clone()))),
                        ],
                        optional: false,
                    });

                    let mut loop_body = vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(row_name.clone()),
                                type_hint: None,
                                init: Some(iter_call),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(first.clone()),
                                type_hint: None,
                                init: Some(lua_first(Expression::new(ExprKind::Ident(row_name.clone())))),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::If {
                            cond: Expression::new(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(Expression::new(ExprKind::Ident(first.clone()))),
                                right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                            }),
                            then_body: vec![Statement::new(StmtKind::Break(BreakTarget::Implicit))],
                            elifs: Vec::new(),
                            else_body: None,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![Expression::new(ExprKind::Ident(ctrl_name.clone()))],
                            value: Expression::new(ExprKind::Ident(first.clone())),
                        }),
                    ];

                    for (extra_index, extra_name) in extra_vars.into_iter().enumerate() {
                        loop_body.push(Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(extra_name),
                                type_hint: None,
                                init: Some(lua_multi_index(
                                    Expression::new(ExprKind::Ident(row_name.clone())),
                                    (extra_index + 1) as i64,
                                )),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }));
                    }

                    loop_body.extend(std::mem::take(body));
                    let mut declarations = Vec::new();
                    if !matches!(&parts[0].kind, ExprKind::Ident(name) if name == "next") {
                        declarations.push(VarDeclarator {
                            pattern: BindingPattern::Ident(iter_name),
                            type_hint: None,
                            init: Some(iter_expr),
                            array_bounds: None,
                            with_events: false,
                        });
                    }
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(state_name),
                        type_hint: None,
                        init: Some(state_expr),
                        array_bounds: None,
                        with_events: false,
                    });
                    declarations.push(VarDeclarator {
                        pattern: BindingPattern::Ident(ctrl_name),
                        type_hint: None,
                        init: Some(ctrl_expr),
                        array_bounds: None,
                        with_events: false,
                    });
                    let mut block = vec![Statement::new(StmtKind::VarDecl {
                        declarations,
                        kind: VarDeclKind::Let,
                    })];
                    block.push(Statement::new(StmtKind::While {
                        cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
                        body: vec![Statement::new(StmtKind::Block(loop_body))],
                        else_body: None,
                    }));
                    *kind = StmtKind::Block(block);
                    normalize_stmt(kind);
                    return;
                }
            }
            let is_builtin_row_iterator = matches!(
                &iter.kind,
                ExprKind::Call { callee, .. }
                    if matches!(
                        lua_call_name(callee).as_deref(),
                        Some("string.gmatch" | "pairs" | "ipairs")
                    )
            );
            if !is_builtin_row_iterator {
                let first = var.clone();
                let extra_vars = lua_for_key_vars(key.take());
                let source_name = format!("__lua_iter_source_{}", first);
                let iter_name = format!("__lua_iter_{}", first);
                let state_name = format!("__lua_state_{}", first);
                let ctrl_name = format!("__lua_ctrl_{}", first);
                let row_name = format!("__lua_row_{}", first);

                let iter_call = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident(iter_name.clone()))),
                    args: vec![
                        Argument::positional(Expression::new(ExprKind::Ident(state_name.clone()))),
                        Argument::positional(Expression::new(ExprKind::Ident(ctrl_name.clone()))),
                    ],
                    optional: false,
                });

                let mut loop_body = vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(row_name.clone()),
                            type_hint: None,
                            init: Some(iter_call),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(first.clone()),
                            type_hint: None,
                            init: Some(lua_first(Expression::new(ExprKind::Ident(row_name.clone())))),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::If {
                        cond: Expression::new(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(Expression::new(ExprKind::Ident(first.clone()))),
                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
                        }),
                        then_body: vec![Statement::new(StmtKind::Break(BreakTarget::Implicit))],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::new(ExprKind::Ident(ctrl_name.clone()))],
                        value: Expression::new(ExprKind::Ident(first.clone())),
                    }),
                ];

                for (extra_index, extra_name) in extra_vars.into_iter().enumerate() {
                    loop_body.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(extra_name),
                            type_hint: None,
                            init: Some(lua_multi_index(
                                Expression::new(ExprKind::Ident(row_name.clone())),
                                (extra_index + 1) as i64,
                            )),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }));
                }

                loop_body.extend(std::mem::take(body));

                let source_expr = Expression::new(ExprKind::Ident(source_name.clone()));
                let block = vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(source_name.clone()),
                            type_hint: None,
                            init: Some(iter.clone()),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![
                            VarDeclarator {
                                pattern: BindingPattern::Ident(iter_name),
                                type_hint: None,
                                init: Some(lua_first(source_expr.clone())),
                                array_bounds: None,
                                with_events: false,
                            },
                            VarDeclarator {
                                pattern: BindingPattern::Ident(state_name),
                                type_hint: None,
                                init: Some(lua_multi_index(source_expr.clone(), 1)),
                                array_bounds: None,
                                with_events: false,
                            },
                            VarDeclarator {
                                pattern: BindingPattern::Ident(ctrl_name),
                                type_hint: None,
                                init: Some(lua_multi_index(source_expr, 2)),
                                array_bounds: None,
                                with_events: false,
                            },
                        ],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::While {
                        cond: Expression::new(ExprKind::Lit(Literal::Bool(true))),
                        body: vec![Statement::new(StmtKind::Block(loop_body))],
                        else_body: None,
                    }),
                ];
                *kind = StmtKind::Block(block);
                normalize_stmt(kind);
                return;
            }
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
                let extra_vars = lua_for_key_vars(key.take());
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
                for (extra_index, extra_name) in extra_vars.into_iter().enumerate() {
                    loop_body.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(extra_name),
                            type_hint: None,
                            init: Some(lua_multi_index(
                                Expression::new(ExprKind::Ident(item.clone())),
                                (extra_index + 1) as i64,
                            )),
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
        StmtKind::FunctionDecl {
            name,
            params,
            body,
            is_async,
            ..
        } => {
            rewrite_lua_current_getinfo_stmts(body, params);
            normalize_lua_stmt_sequence(body);
            if let Some(fixed_count) = lua_lower_rest_param(params) {
                let lambda = Expression::new(ExprKind::Lambda {
                    params: params.clone(),
                    body: LambdaBody::Block(std::mem::take(body)),
                    is_async: *is_async,
                    captures: Vec::new(),
                });
                *kind = StmtKind::Assign {
                    targets: vec![lua_ident(name.clone())],
                    value: lua_mark_rest(lambda, fixed_count),
                };
                normalize_stmt(kind);
            }
        }
        StmtKind::Return(Some(expr)) => {
            if matches!(&expr.kind, ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null))) {
                *expr = lua_multi_row(lua_varargs());
            } else if is_lua_multi_return_call(expr) {
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
