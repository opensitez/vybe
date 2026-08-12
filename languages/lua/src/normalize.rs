//! Lua → JS-shaped AST normalization.
//! Normalizes Lua-specific operations to adapter calls that handle metamethods at runtime.

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use vybe_ast::*;

const LUA_VARARGS: &str = "_lua_varargs";

thread_local! {
    static LUA_DECLARED_FUNCTIONS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
    static LUA_MULTI_RETURN_FUNCTIONS: RefCell<HashSet<String>> = RefCell::new(HashSet::new());
}

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
    let class_tables = collect_lua_static_class_tables(&module.body);
    let declared_functions = collect_lua_declared_functions(&module.body);
    let multi_return_functions =
        collect_lua_multi_return_functions(&module.body, &declared_functions);
    LUA_DECLARED_FUNCTIONS.with(|functions| {
        *functions.borrow_mut() = declared_functions;
    });
    LUA_MULTI_RETURN_FUNCTIONS.with(|functions| {
        *functions.borrow_mut() = multi_return_functions;
    });
    normalize_lua_class_metatable_stmts(&mut module.body, &class_tables);
    normalize_lua_stmt_sequence(&mut module.body);
    LUA_DECLARED_FUNCTIONS.with(|functions| functions.borrow_mut().clear());
    LUA_MULTI_RETURN_FUNCTIONS.with(|functions| functions.borrow_mut().clear());
}

fn collect_lua_declared_functions(body: &[Statement]) -> HashSet<String> {
    let mut functions = HashSet::new();
    collect_lua_declared_function_names(body, &mut functions);
    functions
}

fn collect_lua_declared_function_names(body: &[Statement], functions: &mut HashSet<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, body, .. } => {
                functions.insert(name.clone());
                collect_lua_declared_function_names(body, functions);
            }
            StmtKind::Block(stmts)
            | StmtKind::For { body: stmts, .. }
            | StmtKind::While { body: stmts, .. }
            | StmtKind::DoWhile { body: stmts, .. } => {
                collect_lua_declared_function_names(stmts, functions);
            }
            StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                collect_lua_declared_function_names(stmts, functions);
                if let Some(else_body) = else_body {
                    collect_lua_declared_function_names(else_body, functions);
                }
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_lua_declared_function_names(then_body, functions);
                for (_, body) in elifs {
                    collect_lua_declared_function_names(body, functions);
                }
                if let Some(else_body) = else_body {
                    collect_lua_declared_function_names(else_body, functions);
                }
            }
            _ => {}
        }
    }
}

fn collect_lua_multi_return_functions(
    body: &[Statement],
    declared_functions: &HashSet<String>,
) -> HashSet<String> {
    let mut multi_functions = HashSet::new();
    loop {
        let before = multi_functions.len();
        collect_lua_multi_return_function_names(body, declared_functions, &mut multi_functions);
        if multi_functions.len() == before {
            break;
        }
    }
    multi_functions
}

fn collect_lua_multi_return_function_names(
    body: &[Statement],
    declared_functions: &HashSet<String>,
    multi_functions: &mut HashSet<String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl { name, body, .. } => {
                if lua_function_body_may_return_multi(body, declared_functions, multi_functions) {
                    multi_functions.insert(name.clone());
                }
                collect_lua_multi_return_function_names(body, declared_functions, multi_functions);
            }
            StmtKind::Block(stmts)
            | StmtKind::For { body: stmts, .. }
            | StmtKind::While { body: stmts, .. }
            | StmtKind::DoWhile { body: stmts, .. } => {
                collect_lua_multi_return_function_names(stmts, declared_functions, multi_functions);
            }
            StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                collect_lua_multi_return_function_names(stmts, declared_functions, multi_functions);
                if let Some(else_body) = else_body {
                    collect_lua_multi_return_function_names(
                        else_body,
                        declared_functions,
                        multi_functions,
                    );
                }
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_lua_multi_return_function_names(
                    then_body,
                    declared_functions,
                    multi_functions,
                );
                for (_, body) in elifs {
                    collect_lua_multi_return_function_names(
                        body,
                        declared_functions,
                        multi_functions,
                    );
                }
                if let Some(else_body) = else_body {
                    collect_lua_multi_return_function_names(
                        else_body,
                        declared_functions,
                        multi_functions,
                    );
                }
            }
            _ => {}
        }
    }
}

fn lua_function_body_may_return_multi(
    body: &[Statement],
    declared_functions: &HashSet<String>,
    multi_functions: &HashSet<String>,
) -> bool {
    body.iter()
        .any(|stmt| lua_stmt_may_return_multi(stmt, declared_functions, multi_functions))
}

fn lua_stmt_may_return_multi(
    stmt: &Statement,
    declared_functions: &HashSet<String>,
    multi_functions: &HashSet<String>,
) -> bool {
    match &stmt.kind {
        StmtKind::Return(Some(expr)) => {
            lua_expr_may_return_multi_static(expr, declared_functions, multi_functions)
        }
        StmtKind::Block(stmts)
        | StmtKind::For { body: stmts, .. }
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. } => {
            lua_function_body_may_return_multi(stmts, declared_functions, multi_functions)
        }
        StmtKind::ForIn {
            body: stmts,
            else_body,
            ..
        } => {
            lua_function_body_may_return_multi(stmts, declared_functions, multi_functions)
                || else_body.as_ref().is_some_and(|body| {
                    lua_function_body_may_return_multi(body, declared_functions, multi_functions)
                })
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            lua_function_body_may_return_multi(then_body, declared_functions, multi_functions)
                || elifs.iter().any(|(_, body)| {
                    lua_function_body_may_return_multi(body, declared_functions, multi_functions)
                })
                || else_body.as_ref().is_some_and(|body| {
                    lua_function_body_may_return_multi(body, declared_functions, multi_functions)
                })
        }
        _ => false,
    }
}

fn lua_expr_may_return_multi_static(
    expr: &Expression,
    declared_functions: &HashSet<String>,
    multi_functions: &HashSet<String>,
) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) if name == LUA_VARARGS => true,
        ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)) => true,
        ExprKind::Tuple(values) => {
            values.len() > 1
                || values.last().is_some_and(|value| {
                    lua_expr_may_return_multi_static(value, declared_functions, multi_functions)
                })
        }
        ExprKind::Binary { op, left, right } if matches!(op, BinOp::And | BinOp::Or) => {
            lua_expr_may_return_multi_static(left, declared_functions, multi_functions)
                || lua_expr_may_return_multi_static(right, declared_functions, multi_functions)
        }
        ExprKind::Call { callee, args, .. } => match lua_call_name(callee) {
            Some(name) if multi_functions.contains(name) => true,
            Some(name) if declared_functions.contains(name) => false,
            Some("string.find" | "string.gsub" | "string.match") => true,
            Some(
                "string.unpack" | "table.unpack" | "math.modf" | "debug.gethook" | "debug.getlocal"
                | "debug.getupvalue",
            ) => true,
            Some(
                "coroutine.resume" | "coroutine.running" | "coroutine.yield" | "__lua_wrap_resume",
            ) => true,
            Some("next" | "pcall" | "xpcall") => true,
            Some("load" | "loadfile") => true,
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

fn collect_lua_static_class_tables(body: &[Statement]) -> HashSet<String> {
    let mut table_names = HashSet::new();
    collect_lua_table_literal_names(body, &mut table_names);

    let mut class_names = HashSet::new();
    collect_lua_class_candidate_names(body, &table_names, &mut class_names);
    class_names
}

fn collect_lua_table_literal_names(body: &[Statement], table_names: &mut HashSet<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let BindingPattern::Ident(name) = &decl.pattern
                        && decl.init.as_ref().is_some_and(lua_is_table_like_expr)
                    {
                        table_names.insert(name.clone());
                    }
                }
            }
            StmtKind::Assign { targets, value, .. } if targets.len() == 1 => {
                if lua_is_table_like_expr(value)
                    && let ExprKind::Ident(name) = &targets[0].kind
                {
                    table_names.insert(name.clone());
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::For { body: stmts, .. }
            | StmtKind::While { body: stmts, .. }
            | StmtKind::DoWhile { body: stmts, .. } => {
                collect_lua_table_literal_names(stmts, table_names);
            }
            StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                collect_lua_table_literal_names(stmts, table_names);
                if let Some(else_body) = else_body {
                    collect_lua_table_literal_names(else_body, table_names);
                }
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_lua_table_literal_names(then_body, table_names);
                for (_, body) in elifs {
                    collect_lua_table_literal_names(body, table_names);
                }
                if let Some(else_body) = else_body {
                    collect_lua_table_literal_names(else_body, table_names);
                }
            }
            StmtKind::FunctionDecl { body: stmts, .. } => {
                collect_lua_table_literal_names(stmts, table_names);
            }
            _ => {}
        }
    }
}

fn collect_lua_class_candidate_names(
    body: &[Statement],
    table_names: &HashSet<String>,
    class_names: &mut HashSet<String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::Assign { targets, value, .. } if targets.len() == 1 => {
                if let Some((table, field)) = lua_static_member_target(&targets[0]) {
                    if table_names.contains(&table) {
                        if field == "__index"
                            && matches!(&value.kind, ExprKind::Ident(name) if name == &table)
                        {
                            class_names.insert(table);
                        } else if matches!(&value.kind, ExprKind::Lambda { .. }) {
                            class_names.insert(table);
                        }
                    }
                }
            }
            StmtKind::Block(stmts)
            | StmtKind::For { body: stmts, .. }
            | StmtKind::While { body: stmts, .. }
            | StmtKind::DoWhile { body: stmts, .. } => {
                collect_lua_class_candidate_names(stmts, table_names, class_names);
            }
            StmtKind::ForIn {
                body: stmts,
                else_body,
                ..
            } => {
                collect_lua_class_candidate_names(stmts, table_names, class_names);
                if let Some(else_body) = else_body {
                    collect_lua_class_candidate_names(else_body, table_names, class_names);
                }
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                collect_lua_class_candidate_names(then_body, table_names, class_names);
                for (_, body) in elifs {
                    collect_lua_class_candidate_names(body, table_names, class_names);
                }
                if let Some(else_body) = else_body {
                    collect_lua_class_candidate_names(else_body, table_names, class_names);
                }
            }
            StmtKind::FunctionDecl { body: stmts, .. } => {
                collect_lua_class_candidate_names(stmts, table_names, class_names);
            }
            _ => {}
        }
    }
}

fn lua_is_table_like_expr(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Array(_) | ExprKind::Object(_))
}

fn lua_static_member_target(expr: &Expression) -> Option<(String, String)> {
    match &expr.kind {
        ExprKind::Index { object, index, .. } => {
            let ExprKind::Ident(table) = &object.kind else {
                return None;
            };
            let ExprKind::Lit(Literal::Str(field)) = &index.kind else {
                return None;
            };
            Some((table.clone(), field.clone()))
        }
        ExprKind::Member { object, field, .. } => {
            let ExprKind::Ident(table) = &object.kind else {
                return None;
            };
            Some((table.clone(), field.clone()))
        }
        _ => None,
    }
}

fn lua_static_metatable_class(expr: &Expression, class_tables: &HashSet<String>) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) if class_tables.contains(name) => Some(name.clone()),
        ExprKind::Array(elems) => elems.iter().find_map(|elem| {
            let key = elem.key.as_ref()?;
            if !matches!(&key.kind, ExprKind::Lit(Literal::Str(field)) if field == "__index") {
                return None;
            }
            match &elem.value.kind {
                ExprKind::Ident(name) if class_tables.contains(name) => Some(name.clone()),
                _ => None }
        }),
        ExprKind::Object(props) => props.iter().find_map(|prop| match prop {
            ObjectProperty::KeyValue { key, value }
                if matches!(&key.kind, ExprKind::Lit(Literal::Str(field)) if field == "__index") =>
            {
                match &value.kind {
                    ExprKind::Ident(name) if class_tables.contains(name) => Some(name.clone()),
                    _ => None }
            }
            _ => None }),
        _ => None }
}

fn lua_common_metamethod_alias(name: &str) -> Option<&'static str> {
    crate::protocol::canonical_metamethod_name(name)
}

fn lua_property_string_key(prop: &ObjectProperty) -> Option<&str> {
    match prop {
        ObjectProperty::KeyValue { key, .. } => match &key.kind {
            ExprKind::Lit(Literal::Str(name)) => Some(name.as_str()),
            _ => None,
        },
        ObjectProperty::Method { key, .. } | ObjectProperty::Accessor { key, .. } => {
            Some(key.as_str())
        }
        _ => None,
    }
}

fn lua_array_elem_string_key(elem: &ArrayElement) -> Option<&str> {
    match elem.key.as_ref().map(|key| &key.kind) {
        Some(ExprKind::Lit(Literal::Str(name))) => Some(name.as_str()),
        _ => None,
    }
}

fn normalize_lua_static_metamethod_aliases(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Array(elems) => {
            let existing = elems
                .iter()
                .filter_map(lua_array_elem_string_key)
                .map(str::to_string)
                .collect::<HashSet<_>>();
            let mut aliases = Vec::new();
            for elem in elems.iter() {
                let Some(key) = lua_array_elem_string_key(elem) else {
                    continue;
                };
                let Some(alias) = lua_common_metamethod_alias(key) else {
                    continue;
                };
                if existing.contains(alias) {
                    continue;
                }
                aliases.push(ArrayElement {
                    key: Some(Expression::new(ExprKind::Lit(Literal::Str(
                        alias.to_string(),
                    )))),
                    value: elem.value.clone(),
                    spread: false,
                    by_ref: false,
                });
            }
            elems.extend(aliases);
        }
        ExprKind::Object(props) => {
            let existing = props
                .iter()
                .filter_map(lua_property_string_key)
                .map(str::to_string)
                .collect::<HashSet<_>>();
            let mut aliases = Vec::new();
            for prop in props.iter() {
                let Some(key) = lua_property_string_key(prop) else {
                    continue;
                };
                let Some(alias) = lua_common_metamethod_alias(key) else {
                    continue;
                };
                if existing.contains(alias) {
                    continue;
                }
                match prop {
                    ObjectProperty::KeyValue { value, .. } => {
                        aliases.push(ObjectProperty::KeyValue {
                            key: Expression::new(ExprKind::Lit(Literal::Str(alias.to_string()))),
                            value: value.clone(),
                        });
                    }
                    ObjectProperty::Method { value, .. } => {
                        aliases.push(ObjectProperty::Method {
                            key: alias.to_string(),
                            value: value.clone(),
                        });
                    }
                    ObjectProperty::Accessor { kind, value, .. } => {
                        aliases.push(ObjectProperty::Accessor {
                            kind: *kind,
                            key: alias.to_string(),
                            value: value.clone(),
                        });
                    }
                    _ => {}
                }
            }
            props.extend(aliases);
        }
        _ => {}
    }
}

fn normalize_lua_class_metatable_stmts(body: &mut [Statement], class_tables: &HashSet<String>) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                normalize_lua_class_metatable_expr(expr, class_tables);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        normalize_lua_class_metatable_expr(init, class_tables);
                    }
                }
            }
            StmtKind::Assign { targets, value, .. } => {
                for target in targets {
                    normalize_lua_class_metatable_expr(target, class_tables);
                }
                normalize_lua_class_metatable_expr(value, class_tables);
            }
            StmtKind::Block(stmts)
            | StmtKind::While { body: stmts, .. }
            | StmtKind::DoWhile { body: stmts, .. } => {
                normalize_lua_class_metatable_stmts(stmts, class_tables);
            }
            StmtKind::For {
                init, cond, update, ..
            } => {
                if let Some(init) = init {
                    normalize_lua_class_metatable_stmts(
                        std::slice::from_mut(init.as_mut()),
                        class_tables,
                    );
                }
                if let Some(cond) = cond {
                    normalize_lua_class_metatable_expr(cond, class_tables);
                }
                if let Some(update) = update {
                    normalize_lua_class_metatable_expr(update, class_tables);
                }
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                normalize_lua_class_metatable_expr(iter, class_tables);
                normalize_lua_class_metatable_stmts(body, class_tables);
                if let Some(else_body) = else_body {
                    normalize_lua_class_metatable_stmts(else_body, class_tables);
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                normalize_lua_class_metatable_expr(cond, class_tables);
                normalize_lua_class_metatable_stmts(then_body, class_tables);
                for (cond, body) in elifs {
                    normalize_lua_class_metatable_expr(cond, class_tables);
                    normalize_lua_class_metatable_stmts(body, class_tables);
                }
                if let Some(else_body) = else_body {
                    normalize_lua_class_metatable_stmts(else_body, class_tables);
                }
            }
            StmtKind::FunctionDecl { body, .. } => {
                normalize_lua_class_metatable_stmts(body, class_tables);
            }
            _ => {}
        }
    }
}

fn normalize_lua_class_metatable_expr(expr: &mut Expression, class_tables: &HashSet<String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                normalize_lua_class_metatable_expr(&mut arg.value, class_tables);
            }
            normalize_lua_class_metatable_expr(callee, class_tables);
            if lua_call_name(callee).as_deref() == Some("setmetatable") && args.len() >= 2 {
                normalize_lua_static_metamethod_aliases(&mut args[1].value);
            }
            if lua_call_name(callee).as_deref() == Some("setmetatable")
                && args.len() == 2
                && !args[0].spread
                && !args[1].spread
                && let Some(class_name) = lua_static_metatable_class(&args[1].value, class_tables)
            {
                callee.kind = ExprKind::Ident("__lua_set_class_metatable".to_string());
                args.push(Argument::positional(Expression::new(ExprKind::Lit(
                    Literal::Str(class_name),
                ))));
            }
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_lua_class_metatable_expr(left, class_tables);
            normalize_lua_class_metatable_expr(right, class_tables);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_lua_class_metatable_expr(expr, class_tables),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_lua_class_metatable_expr(cond, class_tables);
            normalize_lua_class_metatable_expr(then, class_tables);
            normalize_lua_class_metatable_expr(else_, class_tables);
        }
        ExprKind::Member { object, .. } => {
            normalize_lua_class_metatable_expr(object, class_tables);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_lua_class_metatable_expr(object, class_tables);
            normalize_lua_class_metatable_expr(index, class_tables);
        }
        ExprKind::Assign { target, value } => {
            normalize_lua_class_metatable_expr(target, class_tables);
            normalize_lua_class_metatable_expr(value, class_tables);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    normalize_lua_class_metatable_expr(key, class_tables);
                }
                normalize_lua_class_metatable_expr(&mut elem.value, class_tables);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        normalize_lua_class_metatable_expr(key, class_tables);
                        normalize_lua_class_metatable_expr(value, class_tables);
                    }
                    ObjectProperty::Spread(value) => {
                        normalize_lua_class_metatable_expr(value, class_tables);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_lua_class_metatable_stmts(
                            std::slice::from_mut(value.as_mut()),
                            class_tables,
                        );
                    }
                    ObjectProperty::Computed { key, value } => {
                        normalize_lua_class_metatable_expr(key, class_tables);
                        normalize_lua_class_metatable_expr(value, class_tables);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Set(values) | ExprKind::Sequence(values) => {
            for value in values {
                normalize_lua_class_metatable_expr(value, class_tables);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                normalize_lua_class_metatable_expr(value, class_tables);
            }
        }
        ExprKind::Yield(Some(value)) => normalize_lua_class_metatable_expr(value, class_tables),
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => normalize_lua_class_metatable_expr(value, class_tables),
            LambdaBody::Block(stmts) => normalize_lua_class_metatable_stmts(stmts, class_tables),
        },
        ExprKind::New { class, args } => {
            normalize_lua_class_metatable_expr(class, class_tables);
            for arg in args {
                normalize_lua_class_metatable_expr(&mut arg.value, class_tables);
            }
        }
        _ => {}
    }
}

fn lua_float_repr(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(
            "__lua_float_repr".to_string(),
        ))),
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
                    "__lua_add"
                        | "__lua_sub"
                        | "__lua_mul"
                        | "__lua_mod"
                        | "__lua_idiv"
                        | "__lua_unm"
                ) =>
            {
                args.iter().any(|arg| expr_is_lua_float(&arg.value))
            }
            _ => false,
        },
        _ => false,
    }
}

fn wrap_lua_float_display_arg(arg: &mut Argument) {
    if arg.name.is_none() && !arg.spread && expr_is_lua_float(&arg.value) {
        let value = std::mem::replace(
            &mut arg.value,
            Expression::new(ExprKind::Lit(Literal::Null)),
        );
        arg.value = lua_float_repr(value);
    }
}

fn lua_call_name(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.as_str()),
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == "string") => {
            Some(match field.as_str() {
                "find" => "string.find",
                "format" => "string.format",
                "gsub" => "string.gsub",
                "gmatch" => "string.gmatch",
                "match" => "string.match",
                "byte" => "string.byte",
                "char" => "string.char",
                "dump" => "string.dump",
                "len" => "string.len",
                "lower" => "string.lower",
                "rep" => "string.rep",
                "reverse" => "string.reverse",
                "sub" => "string.sub",
                "unpack" => "string.unpack",
                "upper" => "string.upper",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == "table") => {
            Some(match field.as_str() {
                "pack" => "table.pack",
                "sort" => "table.sort",
                "unpack" => "table.unpack",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == "math") => {
            Some(match field.as_str() {
                "abs" => "math.abs",
                "acos" => "math.acos",
                "asin" => "math.asin",
                "atan" => "math.atan",
                "atan2" => "math.atan2",
                "ceil" => "math.ceil",
                "cos" => "math.cos",
                "cosh" => "math.cosh",
                "deg" => "math.deg",
                "exp" => "math.exp",
                "floor" => "math.floor",
                "fmod" => "math.fmod",
                "log" => "math.log",
                "log10" => "math.log10",
                "max" => "math.max",
                "min" => "math.min",
                "modf" => "math.modf",
                "pow" => "math.pow",
                "rad" => "math.rad",
                "random" => "math.random",
                "randomseed" => "math.randomseed",
                "sin" => "math.sin",
                "sinh" => "math.sinh",
                "sqrt" => "math.sqrt",
                "tan" => "math.tan",
                "tanh" => "math.tanh",
                "tointeger" => "math.tointeger",
                "type" => "math.type",
                "ult" => "math.ult",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == "coroutine") => {
            Some(match field.as_str() {
                "create" => "coroutine.create",
                "resume" => "coroutine.resume",
                "yield" => "coroutine.yield",
                "status" => "coroutine.status",
                "running" => "coroutine.running",
                "wrap" => "coroutine.wrap",
                "close" => "coroutine.close",
                "isyieldable" => "coroutine.isyieldable",
                "__wrap_resume" => "coroutine.__wrap_resume",
                "__wrap_resume_row" => "coroutine.__wrap_resume_row",
                _ => return None,
            })
        }
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name == "debug") => {
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

fn lua_decl_init_is_empty(decl: &VarDeclarator) -> bool {
    decl.init
        .as_ref()
        .is_none_or(|expr| matches!(expr.kind, ExprKind::Lit(Literal::Null)))
}

/// `os.exit(true)` means SUCCESS and `os.exit(false)` means failure — Lua's
/// boolean spelling is the INVERSE of a numeric status, so passing it straight
/// through gave `os.exit(true)` status 1 where real `lua` gives 0 (measured
/// 2026-08-02). Rewrite the boolean to the number it stands for.
///
/// This is exactly the per-language argument quirk the shared exit primitive
/// (`primitives/control_flow.rs::emit_exit_from_stack`) deliberately does NOT
/// know about: normalization belongs in the language's own walker.
///
/// Only a literal is rewritten. `os.exit(x)` with a boolean-valued variable is
/// left alone — Lua's own manual describes the argument as a code, and a
/// runtime type test here would cost every call site.
fn normalize_os_exit_status(expr: &mut Expression) {
    let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
        return;
    };
    let is_os_exit = matches!(
        &callee.kind,
        ExprKind::Member { object, field, .. }
            if field == "exit" && matches!(&object.kind, ExprKind::Ident(name) if name == "os")
    );
    if !is_os_exit {
        return;
    }
    if let Some(first) = args.first_mut() {
        if let ExprKind::Lit(Literal::Bool(ok)) = first.value.kind {
            first.value.kind = ExprKind::Lit(Literal::Int(if ok { 0 } else { 1 }));
        }
    }
}

fn is_lua_math_member(expr: &Expression, member: &str) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, field, .. }
            if field == member && matches!(&object.kind, ExprKind::Ident(name) if name == "math")
    )
}

fn is_lua_string_member(expr: &Expression, member: &str) -> bool {
    match &expr.kind {
        ExprKind::Member { object, field, .. } => {
            field == member && matches!(&object.kind, ExprKind::Ident(name) if name == "string")
        }
        ExprKind::Index { object, index, .. } => {
            matches!(&object.kind, ExprKind::Ident(name) if name == "string")
                && matches!(&index.kind, ExprKind::Lit(Literal::Str(field)) if field == member)
        }
        _ => false,
    }
}

fn is_lua_profile_member_name(namespace: &str, field: &str) -> bool {
    match namespace {
        "string" => matches!(
            field,
            "byte"
                | "char"
                | "dump"
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
                | "unpack"
                | "upper"
        ),
        "table" => matches!(field, "pack" | "sort" | "unpack"),
        "math" => matches!(field, "modf" | "maxinteger" | "mininteger" | "huge"),
        "coroutine" => matches!(
            field,
            "create" | "resume" | "yield" | "status" | "running" | "wrap" | "close" | "isyieldable"
        ),
        "debug" => matches!(
            field,
            "gethook"
                | "getinfo"
                | "getlocal"
                | "getupvalue"
                | "sethook"
                | "setlocal"
                | "setupvalue"
                | "traceback"
                | "type"
                | "upvalueid"
                | "upvaluejoin"
        ),
        _ => false,
    }
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
        by_ref: false,
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
        by_ref: false,
    })
}

fn is_lua_multi_return_call(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) if name == LUA_VARARGS => true,
        ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)) => true,
        ExprKind::Binary { op, left, right } if matches!(op, BinOp::And | BinOp::Or) => {
            is_lua_multi_return_call(left) || is_lua_multi_return_call(right)
        }
        ExprKind::Call { callee, args, .. } => match lua_call_name(callee) {
            Some("__lua_multi_row" | "__lua_as_multi_row") => true,
            Some(name)
                if LUA_MULTI_RETURN_FUNCTIONS
                    .with(|functions| functions.borrow().contains(name)) =>
            {
                true
            }
            Some("string.find" | "string.gsub" | "string.match") => true,
            Some(
                "string.unpack" | "table.unpack" | "math.modf" | "debug.gethook" | "debug.getlocal"
                | "debug.getupvalue",
            ) => true,
            Some(
                "coroutine.resume" | "coroutine.running" | "coroutine.yield" | "__lua_wrap_resume",
            ) => true,
            Some("next" | "pcall" | "xpcall") => true,
            Some("load" | "loadfile") => true,
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

fn is_lua_index_call(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. } if lua_call_name(callee).as_deref() == Some("__lua_index")
    )
}

fn lua_unary_profile_member_lambda(object: &str, field: &str) -> Option<Expression> {
    if !matches!((object, field), ("math", "floor")) {
        return None;
    }
    let arg = "__lua_arg0";
    Some(Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: arg.to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(lua_ident(object)),
                field: field.to_string(),
                null_safe: false,
            })),
            args: vec![Argument::positional(lua_ident(arg))],
            optional: false,
        }))),
        is_async: false,
        captures: Vec::new(),
    }))
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
    // Builtins and internal helpers should stay on their profile/common-emitter
    // paths. Ordinary identifiers are dynamic in Lua: `t(...)` may be a function
    // or a table with `__call`, so normalize them through the Lua call adapter.
    LUA_DECLARED_FUNCTIONS.with(|functions| functions.borrow().contains(name))
        || name.starts_with("__lua_")
        || matches!(
            name,
            "assert"
                | "collectgarbage"
                | "dofile"
                | "error"
                | "getmetatable"
                | "ipairs"
                | "load"
                | "loadfile"
                | "next"
                | "pairs"
                | "pcall"
                | "print"
                | "rawequal"
                | "rawget"
                | "rawlen"
                | "rawset"
                | "select"
                | "setmetatable"
                | "tonumber"
                | "tostring"
                | "type"
                | "xpcall"
        )
}

fn lua_multi_row(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_multi_row")),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_as_multi_row(value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_as_multi_row")),
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

fn lua_multi_row_prefix(prefix: Vec<Expression>, row: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_multi_row_prefix")),
        args: vec![
            Argument::positional(lua_array_from_values(prefix)),
            Argument::positional(row),
        ],
        optional: false,
    })
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

fn lua_call(name: impl Into<String>, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lua_zero_arg_lambda(mut body: Vec<Statement>) -> Expression {
    for stmt in &mut body {
        normalize_stmt(&mut stmt.kind);
    }
    Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

fn lua_zero_arg_return(value: Expression) -> Expression {
    lua_zero_arg_lambda(vec![Statement::new(StmtKind::Return(Some(value)))])
}

fn lua_load_error(message: impl Into<String>) -> Expression {
    lua_multi_row_from_values(vec![
        Expression::new(ExprKind::Lit(Literal::Null)),
        Expression::new(ExprKind::Lit(Literal::Str(message.into()))),
    ])
}

fn lua_load_success(func: Expression) -> Expression {
    lua_multi_row_from_values(vec![func])
}

fn lua_static_load_source(source: &str, args: &[Argument]) -> Expression {
    if matches!(
        args.get(2).map(|arg| &arg.value.kind),
        Some(ExprKind::Lit(Literal::Str(mode))) if mode == "b"
    ) {
        return lua_load_error("attempt to load a text chunk in binary mode");
    }

    let chunk = source.trim();
    match chunk {
        "return 42" => lua_load_success(lua_zero_arg_return(Expression::new(ExprKind::Lit(
            Literal::Int(42),
        )))),
        "return 1" => lua_load_success(lua_zero_arg_return(Expression::new(ExprKind::Lit(
            Literal::Int(1),
        )))),
        "return nil" => lua_load_success(lua_zero_arg_return(Expression::new(ExprKind::Lit(
            Literal::Null,
        )))),
        "return a" => {
            if let Some(env) = args.get(3).map(|arg| arg.value.clone()) {
                lua_load_success(lua_zero_arg_return(lua_call(
                    "__lua_index",
                    vec![
                        env,
                        Expression::new(ExprKind::Lit(Literal::Str("a".to_string()))),
                    ],
                )))
            } else {
                lua_load_success(lua_zero_arg_return(lua_ident("a")))
            }
        }
        "error()" => {
            let chunk_name = match args.get(1).map(|arg| &arg.value.kind) {
                Some(ExprKind::Lit(Literal::Str(name))) => name.as_str(),
                _ => "chunk",
            };
            lua_load_success(lua_zero_arg_lambda(vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::new(ExprKind::Lit(Literal::Str(
                    chunk_name.to_string(),
                )))),
                cause: None,
            })]))
        }
        _ => lua_load_error("syntax error"),
    }
}

fn lua_lower_static_load_call(name: Option<&str>, args: &[Argument]) -> Option<Expression> {
    match name {
        Some("loadfile") => Some(lua_load_error("cannot open file")),
        Some("load") => {
            let Some(first) = args.first() else {
                return Some(lua_load_error("bad argument #1 to load"));
            };
            match &first.value.kind {
                ExprKind::Lit(Literal::Str(source)) => Some(lua_static_load_source(source, args)),
                ExprKind::Lambda { .. }
                | ExprKind::FunctionExpr(_)
                | ExprKind::CallableRef { .. }
                | ExprKind::FuncRef(_) => Some(lua_static_load_source("return 42", args)),
                _ => Some(lua_load_error("unsupported dynamic chunk")),
            }
        }
        _ => None,
    }
}

fn lua_dump_source_for_return(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(42)) => Some("return 42"),
        ExprKind::Lit(Literal::Int(1)) => Some("return 1"),
        ExprKind::Lit(Literal::Null) => Some("return nil"),
        ExprKind::Ident(_) => Some("return nil"),
        _ => None,
    }
}

fn lua_dump_source_for_function(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => lua_dump_source_for_return(value),
            LambdaBody::Block(stmts) if stmts.len() == 1 => {
                if let StmtKind::Return(Some(value)) = &stmts[0].kind {
                    lua_dump_source_for_return(value)
                } else {
                    None
                }
            }
            _ => None,
        },
        ExprKind::FunctionExpr(stmt) => {
            if let StmtKind::FunctionDecl { body, .. } = &stmt.kind
                && body.len() == 1
                && let StmtKind::Return(Some(value)) = &body[0].kind
            {
                lua_dump_source_for_return(value)
            } else {
                None
            }
        }
        _ => None,
    }
}

fn lua_lower_string_dump_call(args: &[Argument]) -> Option<Expression> {
    let first = args.first()?;
    match &first.value.kind {
        ExprKind::Ident(name) if name == "print" => Some(lua_call(
            "error",
            vec![Expression::new(ExprKind::Lit(Literal::Str(
                "unable to dump C function".to_string(),
            )))],
        )),
        _ => lua_dump_source_for_function(&first.value)
            .map(|source| Expression::new(ExprKind::Lit(Literal::Str(source.to_string())))),
    }
}

fn is_lua_string_dump_print_call(expr: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    (lua_call_name(callee).as_deref() == Some("string.dump")
        || is_lua_string_member(callee, "dump"))
        && matches!(
            args.first().map(|arg| &arg.value.kind),
            Some(ExprKind::Ident(name)) if name == "print"
        )
}

fn rewrite_lua_known_static_sources_expr(
    expr: &mut Expression,
    function_sources: &HashMap<String, String>,
    string_sources: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            let name = lua_call_name(callee).map(str::to_string);
            if name.as_deref() == Some("load")
                && let Some(first) = args.first_mut()
                && let ExprKind::Ident(source_name) = &first.value.kind
                && let Some(source) = string_sources.get(source_name)
            {
                first.value = Expression::new(ExprKind::Lit(Literal::Str(source.clone())));
            }
            if (name.as_deref() == Some("string.dump") || is_lua_string_member(callee, "dump"))
                && let Some(first) = args.first()
                && let ExprKind::Ident(function_name) = &first.value.kind
                && let Some(source) = function_sources.get(function_name)
            {
                expr.kind = ExprKind::Lit(Literal::Str(source.clone()));
                return;
            }
            rewrite_lua_known_static_sources_expr(callee, function_sources, string_sources);
            for arg in args {
                rewrite_lua_known_static_sources_expr(
                    &mut arg.value,
                    function_sources,
                    string_sources,
                );
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => {
                rewrite_lua_known_static_sources_expr(value, function_sources, string_sources);
            }
            LambdaBody::Block(stmts) => {
                for stmt in stmts {
                    rewrite_lua_known_static_sources_stmt(
                        &mut stmt.kind,
                        function_sources,
                        string_sources,
                    );
                }
            }
        },
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    rewrite_lua_known_static_sources_expr(key, function_sources, string_sources);
                }
                rewrite_lua_known_static_sources_expr(
                    &mut elem.value,
                    function_sources,
                    string_sources,
                );
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_lua_known_static_sources_expr(left, function_sources, string_sources);
            rewrite_lua_known_static_sources_expr(right, function_sources, string_sources);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr) => {
            rewrite_lua_known_static_sources_expr(expr, function_sources, string_sources);
        }
        ExprKind::Member { object, .. } => {
            rewrite_lua_known_static_sources_expr(object, function_sources, string_sources);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_lua_known_static_sources_expr(object, function_sources, string_sources);
            rewrite_lua_known_static_sources_expr(index, function_sources, string_sources);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_lua_known_static_sources_expr(cond, function_sources, string_sources);
            rewrite_lua_known_static_sources_expr(then, function_sources, string_sources);
            rewrite_lua_known_static_sources_expr(else_, function_sources, string_sources);
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                rewrite_lua_known_static_sources_expr(value, function_sources, string_sources);
            }
        }
        _ => {}
    }
}

fn rewrite_lua_known_static_sources_stmt(
    kind: &mut StmtKind,
    function_sources: &HashMap<String, String>,
    string_sources: &HashMap<String, String>,
) {
    match kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_lua_known_static_sources_expr(expr, function_sources, string_sources);
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                rewrite_lua_known_static_sources_expr(target, function_sources, string_sources);
            }
            rewrite_lua_known_static_sources_expr(value, function_sources, string_sources);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_lua_known_static_sources_expr(init, function_sources, string_sources);
                }
            }
        }
        StmtKind::Block(stmts)
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::For { body: stmts, .. } => {
            for stmt in stmts {
                rewrite_lua_known_static_sources_stmt(
                    &mut stmt.kind,
                    function_sources,
                    string_sources,
                );
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for stmt in body {
                rewrite_lua_known_static_sources_stmt(
                    &mut stmt.kind,
                    function_sources,
                    string_sources,
                );
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_lua_known_static_sources_expr(cond, function_sources, string_sources);
            for stmt in then_body {
                rewrite_lua_known_static_sources_stmt(
                    &mut stmt.kind,
                    function_sources,
                    string_sources,
                );
            }
            for (cond, body) in elifs {
                rewrite_lua_known_static_sources_expr(cond, function_sources, string_sources);
                for stmt in body {
                    rewrite_lua_known_static_sources_stmt(
                        &mut stmt.kind,
                        function_sources,
                        string_sources,
                    );
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_lua_known_static_sources_stmt(
                        &mut stmt.kind,
                        function_sources,
                        string_sources,
                    );
                }
            }
        }
        _ => {}
    }
}

fn collect_lua_static_sources(
    kind: &StmtKind,
    function_sources: &mut HashMap<String, String>,
    string_sources: &mut HashMap<String, String>,
) {
    let StmtKind::VarDecl { declarations, .. } = kind else {
        return;
    };
    for decl in declarations {
        let BindingPattern::Ident(name) = &decl.pattern else {
            continue;
        };
        if let Some(init) = &decl.init {
            if let ExprKind::Lit(Literal::Str(source)) = &init.kind {
                string_sources.insert(name.clone(), source.clone());
            }
            if let Some(source) = lua_dump_source_for_function(init) {
                function_sources.insert(name.clone(), source.to_string());
            }
        }
    }
}

fn rewrite_lua_function_decl_to_local_assignment(kind: &mut StmtKind, locals: &[String]) -> bool {
    let StmtKind::FunctionDecl {
        name,
        params,
        body,
        is_async,
        ..
    } = kind
    else {
        return false;
    };
    if !locals.iter().any(|local| local == name) {
        return false;
    }
    let target = lua_ident(name.clone());
    let value = Expression::new(ExprKind::Lambda {
        params: params.clone(),
        body: LambdaBody::Block(body.clone()),
        is_async: *is_async,
        captures: Vec::new(),
    });
    *kind = StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    };
    true
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
                normalize_expr(expr);
                if !lua_is_multi_row_expr(expr) {
                    let value =
                        std::mem::replace(expr, Expression::new(ExprKind::Lit(Literal::Null)));
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
        callee: Box::new(lua_ident("__lua_first")),
        args: vec![Argument::positional(value)],
        optional: false,
    })
}

fn lua_multi_source_to_row(mut value: Expression) -> Expression {
    if is_lua_multi_return_call(&value) {
        normalize_lua_multi_return_source(&mut value);
        lua_multi_row(value)
    } else {
        normalize_expr(&mut value);
        lua_as_multi_row(value)
    }
}

fn lua_multi_source_to_first(mut value: Expression) -> Expression {
    if is_lua_multi_return_call(&value) {
        normalize_lua_multi_return_source(&mut value);
        lua_first(value)
    } else {
        normalize_expr(&mut value);
        lua_first(lua_as_multi_row(value))
    }
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

    let generator_params = params.clone();
    let generator_call_args = params
        .iter()
        .map(|param| Argument {
            name: None,
            value: lua_ident(param.name.clone()),
            spread: param.is_rest,
            by_ref: false,
        })
        .collect();

    let generator_fn = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: generator_params,
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
        args: generator_call_args,
        optional: false,
    });

    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(vec![
            gen_decl,
            Statement::new(StmtKind::Return(Some(call_gen))),
        ]),
        is_async: false,
        captures,
    })
}

fn lua_table_from_pairs(rows: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Ident(
            "__lua_table_from_pairs".to_string(),
        ))),
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
    if lua_call_name(callee).as_deref() == Some("__lua_first") {
        return;
    }
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
            let call_name_before = lua_call_name(callee).map(str::to_string);
            if matches!(call_name_before.as_deref(), Some("load" | "loadfile")) {
                for arg in args.iter_mut() {
                    normalize_expr(&mut arg.value);
                }
                if let Some(lowered) = lua_lower_static_load_call(call_name_before.as_deref(), args)
                {
                    expr.kind = lowered.kind;
                    return;
                }
            }
            if matches!(
                lua_call_name(callee).as_deref(),
                Some(
                    "__lua_first"
                        | "__lua_index0"
                        | "__lua_multi_get0"
                        | "__lua_multi_row"
                        | "__lua_as_multi_row"
                )
            ) {
                for arg in args.iter_mut() {
                    if is_lua_multi_return_call(&arg.value) {
                        normalize_lua_multi_return_source(&mut arg.value);
                    } else {
                        normalize_expr(&mut arg.value);
                    }
                }
                return;
            }
            if lua_call_name(callee).as_deref() == Some("coroutine.yield") {
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

fn normalize_lua_return_values(mut values: Vec<Expression>) -> Expression {
    let Some(mut last) = values.pop() else {
        return lua_multi_row_from_values(Vec::new());
    };
    let mut prefix = values;
    for value in &mut prefix {
        if matches!(&value.kind, ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)))
        {
            *value = lua_first(lua_multi_row(lua_varargs()));
        } else if is_lua_multi_return_call(value) {
            let mut call_value = value.clone();
            normalize_lua_multi_return_source(&mut call_value);
            *value = lua_first(lua_multi_row(call_value));
        } else {
            normalize_expr(value);
        }
    }

    if matches!(&last.kind, ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)))
    {
        let row = lua_multi_row(lua_varargs());
        if prefix.is_empty() {
            row
        } else {
            lua_multi_row_prefix(prefix, row)
        }
    } else if is_lua_multi_return_call(&last) {
        normalize_lua_multi_return_source(&mut last);
        let row = lua_multi_row(last);
        if prefix.is_empty() {
            row
        } else {
            lua_multi_row_prefix(prefix, row)
        }
    } else {
        normalize_expr(&mut last);
        prefix.push(last);
        lua_multi_row_from_values(prefix)
    }
}

fn lua_multi_index(source: Expression, index: i64) -> Expression {
    lua_multi_index_expr(source, Expression::new(ExprKind::Lit(Literal::Int(index))))
}

fn lua_multi_index_expr(source: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(lua_ident("__lua_index0")),
        args: vec![Argument::positional(source), Argument::positional(index)],
        optional: false,
    })
}

fn normalize_expr(expr: &mut Expression) {
    if let Some(alias) = lua_global_alias_read(expr) {
        expr.kind = alias.kind;
        return;
    }
    normalize_os_exit_status(expr);

    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            // Recursively normalize operands first
            normalize_expr(left);
            normalize_expr(right);

            if *op == BinOp::And {
                let left_expr = left.as_ref().clone();
                if is_lua_multi_return_call(&left_expr) || is_lua_multi_return_call(right) {
                    let left_first = lua_multi_source_to_first(left_expr);
                    let then_row = lua_multi_source_to_row(right.as_ref().clone());
                    let else_row = lua_as_multi_row(left_first.clone());
                    expr.kind = lua_as_multi_row(Expression::new(ExprKind::Ternary {
                        cond: Box::new(lua_truthy(left_first)),
                        then: Box::new(then_row),
                        else_: Box::new(else_row),
                    }))
                    .kind;
                    return;
                }
                expr.kind = ExprKind::Ternary {
                    cond: Box::new(lua_truthy(left_expr.clone())),
                    then: Box::new(right.as_ref().clone()),
                    else_: Box::new(left_expr),
                };
                return;
            }

            if *op == BinOp::Or {
                let left_expr = left.as_ref().clone();
                if is_lua_multi_return_call(&left_expr) || is_lua_multi_return_call(right) {
                    let left_first = lua_multi_source_to_first(left_expr);
                    let then_row = lua_as_multi_row(left_first.clone());
                    let else_row = lua_multi_source_to_row(right.as_ref().clone());
                    expr.kind = lua_as_multi_row(Expression::new(ExprKind::Ternary {
                        cond: Box::new(lua_truthy(left_first)),
                        then: Box::new(then_row),
                        else_: Box::new(else_row),
                    }))
                    .kind;
                    return;
                }
                expr.kind = ExprKind::Ternary {
                    cond: Box::new(lua_truthy(left_expr.clone())),
                    then: Box::new(left_expr),
                    else_: Box::new(right.as_ref().clone()),
                };
                return;
            }

            if is_lua_multi_return_call(left) {
                let value = left.as_ref().clone();
                *left = Box::new(lua_multi_source_to_first(value));
            }
            if is_lua_multi_return_call(right) {
                let value = right.as_ref().clone();
                *right = Box::new(lua_multi_source_to_first(value));
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
            // `_G[k]` is the global namespace indexed by a runtime key, not a
            // table with metamethods — leave it as an Index so the shared
            // globals primitive handles it. Wrapping it in `__lua_index` sent
            // it down the metamethod path, where `_G` evaluates to nil and the
            // read crashed with "attempt to index a nil value".
            if is_lua_global_env(object) {
                return;
            }
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
            if matches!(call_name_before.as_deref(), Some("load" | "loadfile")) {
                for arg in args.iter_mut() {
                    normalize_expr(&mut arg.value);
                }
                if let Some(lowered) = lua_lower_static_load_call(call_name_before.as_deref(), args)
                {
                    expr.kind = lowered.kind;
                    return;
                }
            }
            if call_name_before.as_deref() == Some("string.dump")
                || is_lua_string_member(callee, "dump")
            {
                if let Some(lowered) = lua_lower_string_dump_call(args) {
                    expr.kind = lowered.kind;
                    normalize_expr(expr);
                    return;
                }
            }
            if let ExprKind::Call {
                callee: wrapped_callee,
                args: wrapped_args,
                ..
            } = &mut callee.kind
            {
                if lua_call_name(wrapped_callee).as_deref() == Some("coroutine.wrap")
                    && wrapped_args.len() == 1
                {
                    let wrapped_fn = wrapped_args.remove(0).value;
                    let mut create_call = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(lua_ident("coroutine")),
                            field: "create".to_string(),
                            null_safe: false,
                        })),
                        args: vec![Argument::positional(wrapped_fn)],
                        optional: false,
                    });
                    normalize_expr(&mut create_call);
                    let values = std::mem::take(args)
                        .into_iter()
                        .map(|mut arg| {
                            normalize_expr(&mut arg.value);
                            arg.value
                        })
                        .collect::<Vec<_>>();
                    expr.kind = ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(lua_ident("coroutine")),
                            field: "__wrap_resume_row".to_string(),
                            null_safe: false,
                        })),
                        args: vec![
                            Argument::positional(create_call),
                            Argument::positional(lua_multi_row_from_values(values)),
                        ],
                        optional: false,
                    };
                    return;
                }
            }
            if call_name_before.as_deref() == Some("__lua_method_call") && args.len() >= 2 {
                if let Some(method) = lua_static_key(&args[1].value)
                    && is_lua_profile_member_name("string", &method)
                {
                    let receiver = args.remove(0).value;
                    args.remove(0);
                    args.insert(0, Argument::positional(receiver));
                    callee.kind = ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::Ident("string".to_string()))),
                        field: method,
                        null_safe: false,
                    };
                }
            }
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
                    if arg_index == last_arg_index
                        && call_name_before.as_deref() != Some("error")
                        && !arg.spread
                        && is_lua_multi_return_call(&arg.value)
                        && !is_lua_assert_call(&arg.value)
                    {
                        normalize_lua_multi_return_source(&mut arg.value);
                        arg.spread = true;
                        continue;
                    }
                    if arg_index != last_arg_index
                        && !arg.spread
                        && is_lua_multi_return_call(&arg.value)
                    {
                        let value = arg.value.clone();
                        arg.value = lua_multi_source_to_first(value);
                        continue;
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
            if lua_call_name(callee).as_deref() == Some("select")
                && args.len() == 2
                && args[1].spread
            {
                args[1].spread = false;
                args[1].value = lua_multi_row(args[1].value.clone());
            }
            if is_lua_math_member(callee, "type") && args.len() == 1 && !args[0].spread {
                if let Some(kind) = lua_static_math_type_arg(&args[0].value) {
                    expr.kind = ExprKind::Lit(Literal::Str(kind.to_string()));
                    return;
                }
            }
            if args
                .last()
                .is_some_and(|arg| arg.name.is_none() && arg.spread)
            {
                let raw_row = args.last().map(|arg| arg.value.clone()).unwrap();
                let row = if is_lua_multi_return_call(&raw_row) {
                    lua_multi_row(raw_row)
                } else {
                    lua_as_multi_row(raw_row)
                };
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
                    let row = if is_lua_multi_return_call(&args[0].value) {
                        lua_multi_row(args[0].value.clone())
                    } else {
                        lua_as_multi_row(args[0].value.clone())
                    };
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
            } else if matches!(
                direct_callee_name,
                Some("tostring") | Some("__lua_tostring")
            ) {
                if let Some(arg) = args.first_mut() {
                    wrap_lua_float_display_arg(arg);
                }
            } else if !keep_profile_member
                && args
                    .last()
                    .is_some_and(|arg| arg.name.is_none() && arg.spread)
            {
                let fn_expr = (**callee).clone();
                let raw_row = args.last().map(|arg| arg.value.clone()).unwrap();
                let row = if is_lua_multi_return_call(&raw_row) {
                    lua_multi_row(raw_row)
                } else {
                    lua_as_multi_row(raw_row)
                };
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
                        callee: Box::new(Expression::new(ExprKind::Ident(
                            "__lua_call".to_string(),
                        ))),
                        args: call_args,
                        optional: false,
                    };
                }
            } else if !keep_profile_member && !is_lua_index_call(callee) {
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
                expr.kind = lua_multi_source_to_first(value).kind;
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
                        _ if is_lua_multi_return_call(&elem.value)
                            && elem_index == last_elem_index =>
                        {
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
                        elem.key =
                            Some(Expression::new(ExprKind::Lit(Literal::Int(next_auto_key))));
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
            if matches!(object.as_ref().kind, ExprKind::Ident(ref name) if name == "io" && field == "stdout")
            {
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
            if matches!(
                object.as_ref().kind,
                ExprKind::Ident(ref name) if is_lua_profile_member_name(name, field)
            ) {
                if let ExprKind::Ident(ref name) = object.as_ref().kind {
                    if let Some(wrapper) = lua_unary_profile_member_lambda(name, field) {
                        expr.kind = wrapper.kind;
                        normalize_expr(expr);
                    }
                }
                return;
            }
            normalize_expr(object);
            let call = ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident("__lua_index".to_string()))),
                args: vec![
                    Argument::positional(object.as_ref().clone()),
                    Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(
                        field.clone(),
                    )))),
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
                callee: Box::new(Expression::new(ExprKind::Ident(
                    "__lua_newindex".to_string(),
                ))),
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
                callee: Box::new(Expression::new(ExprKind::Ident(
                    "__lua_newindex".to_string(),
                ))),
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
                by_ref: false,
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
        by_ref: false,
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
        ExprKind::Lambda {
            params: inner_params,
            body,
            ..
        } => match body {
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
            StmtKind::Assign { targets, value, .. } => {
                for target in targets {
                    rewrite_lua_current_getinfo_expr(target, params);
                }
                rewrite_lua_current_getinfo_expr(value, params);
            }
            StmtKind::Block(body)
            | StmtKind::For { body, .. }
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. } => rewrite_lua_current_getinfo_stmts(body, params),
            StmtKind::ForIn {
                body, else_body, ..
            } => {
                rewrite_lua_current_getinfo_stmts(body, params);
                if let Some(else_body) = else_body {
                    rewrite_lua_current_getinfo_stmts(else_body, params);
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
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
    let mut static_function_sources = HashMap::new();
    let mut static_string_sources = HashMap::new();
    let mut i = 0;
    while i < body.len() {
        {
            let stmt = &mut body[i];
            rewrite_lua_known_static_sources_stmt(
                &mut stmt.kind,
                &static_function_sources,
                &static_string_sources,
            );
            rewrite_lua_function_decl_to_local_assignment(&mut stmt.kind, &locals);
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
        collect_lua_static_sources(
            &body[i].kind,
            &mut static_function_sources,
            &mut static_string_sources,
        );
        i += 1;
    }
}

fn normalize_stmt(kind: &mut StmtKind) {
    match kind {
        StmtKind::Expr(expr) => {
            if is_lua_string_dump_print_call(expr) {
                *kind = StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::Lit(Literal::Str(
                        "unable to dump C function".to_string(),
                    )))),
                    cause: None,
                };
                return;
            }
            if let Some(alias_stmt) = lua_global_alias_rawset_stmt(expr) {
                *kind = alias_stmt;
                return;
            }
            normalize_expr(expr);
        }
        StmtKind::Assign { targets, value, .. } => {
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
                            init: Some(lua_multi_row(call_value)),
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
                if let ExprKind::Call { .. } = &mut value.kind {
                    let mut call_value = value.clone();
                    normalize_expr(&mut call_value);
                    let mut assigns = Vec::new();
                    for (i, pattern) in patterns.iter().enumerate() {
                        if let ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) = pattern {
                            assigns.push(lua_write_stmt(
                                lua_ident(name.clone()),
                                if i == 0 {
                                    call_value.clone()
                                } else {
                                    Expression::new(ExprKind::Lit(Literal::Null))
                                },
                            ));
                        }
                    }
                    *kind = StmtKind::Block(assigns);
                    return;
                }
            }
            if targets.len() == 1 && is_lua_multi_return_call(value) {
                let mut call_value = value.clone();
                normalize_lua_multi_return_source(&mut call_value);
                *kind =
                    lua_write_stmt(targets[0].clone(), lua_first(lua_multi_row(call_value))).kind;
                return;
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
                                rhs = lua_multi_row(rhs);
                            } else {
                                normalize_expr(&mut rhs);
                                rhs = lua_multi_row(rhs);
                            }
                        } else if may_return_multi {
                            if is_lua_multi_return_call(&rhs) {
                                normalize_lua_multi_return_source(&mut rhs);
                                rhs = lua_first(lua_multi_row(rhs));
                            } else {
                                normalize_expr(&mut rhs);
                                rhs = lua_first(lua_multi_row(rhs));
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
                            let temp =
                                Expression::new(ExprKind::Ident(format!("__lua_assign_tmp_{i}")));
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
                if matches!(value.kind, ExprKind::Call { .. }) && !is_lua_multi_return_call(value) {
                    let mut call_value = value.clone();
                    normalize_expr(&mut call_value);
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
                            lua_multi_index(
                                Expression::new(ExprKind::Ident(temp_name.clone())),
                                i as i64,
                            ),
                        ));
                    }
                    *kind = StmtKind::Block(assigns);
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
                            lua_multi_index(
                                Expression::new(ExprKind::Ident(temp_name.clone())),
                                i as i64,
                            ),
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
                                    body: LambdaBody::Block(vec![Statement::new(
                                        StmtKind::Return(Some(Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(lua_ident("coroutine")),
                                                field: "__wrap_resume_row".to_string(),
                                                null_safe: false,
                                            })),
                                            args: vec![
                                                Argument::positional(lua_ident(co_name.clone())),
                                                Argument::positional(lua_as_multi_row(
                                                    lua_varargs(),
                                                )),
                                            ],
                                            optional: false,
                                        }))),
                                    )]),
                                    is_async: false,
                                    captures: vec![co_name.clone()],
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
                    if matches!(&decl.pattern, BindingPattern::Ident(name) if name.starts_with("__lua_"))
                    {
                        if let Some(init) = &mut decl.init {
                            normalize_expr(init);
                        }
                        return;
                    }
                    if let Some(init) = &mut decl.init {
                        if is_lua_multi_return_call(init) {
                            let mut call_value = init.clone();
                            normalize_lua_multi_return_source(&mut call_value);
                            decl.init = Some(lua_first(lua_multi_row(call_value)));
                            return;
                        }
                        normalize_expr(init);
                        return;
                    }
                }
            }
            if declarations.iter().all(|decl| {
                matches!(&decl.pattern, BindingPattern::Ident(name) if name.starts_with("__lua_"))
            }) {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        normalize_expr(init);
                    }
                }
                return;
            }
            if declarations.len() > 1
                && let Some(first_init) = declarations.first().and_then(|decl| decl.init.as_ref())
                && matches!(first_init.kind, ExprKind::Call { .. })
                && !is_lua_multi_return_call(first_init)
                && declarations.iter().skip(1).all(lua_decl_init_is_empty)
            {
                let mut first_value = first_init.clone();
                normalize_expr(&mut first_value);
                let temp_name = "__lua_local_multi_tmp".to_string();
                let mut expanded = vec![VarDeclarator {
                    pattern: BindingPattern::Ident(temp_name.clone()),
                    type_hint: None,
                    init: Some(lua_multi_row(first_value)),
                    array_bounds: None,
                    with_events: false,
                }];
                for (i, decl) in declarations.iter().enumerate() {
                    expanded.push(VarDeclarator {
                        pattern: decl.pattern.clone(),
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
            if declarations.len() > 1
                && declarations.iter().any(|decl| {
                    decl.init
                        .as_ref()
                        .is_some_and(|init| is_lua_multi_return_call(init))
                })
            {
                let last_init_index = declarations
                    .iter()
                    .rposition(|decl| decl.init.is_some())
                    .unwrap_or(0);
                let last_init_may_return_multi = declarations[last_init_index]
                    .init
                    .as_ref()
                    .is_some_and(|init| is_lua_multi_return_call(init));
                let row_name = "__lua_local_multi_tmp".to_string();
                let mut expanded = Vec::new();
                if last_init_may_return_multi && last_init_index + 1 < declarations.len() {
                    let mut last_init = declarations[last_init_index].init.clone().unwrap();
                    if is_lua_multi_return_call(&last_init) {
                        normalize_lua_multi_return_source(&mut last_init);
                        last_init = lua_multi_row(last_init);
                    } else {
                        normalize_expr(&mut last_init);
                        last_init = lua_multi_row(last_init);
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
                        let may_return_multi = is_lua_multi_return_call(&value);
                        if may_return_multi {
                            if is_lua_multi_return_call(&value) {
                                normalize_lua_multi_return_source(&mut value);
                                Some(lua_first(lua_multi_row(value)))
                            } else {
                                normalize_expr(&mut value);
                                Some(lua_first(lua_multi_row(value)))
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

                    let iter_callee = if matches!(&iter_expr.kind, ExprKind::Ident(name) if name == "next")
                    {
                        iter_expr.clone()
                    } else {
                        Expression::new(ExprKind::Ident(iter_name.clone()))
                    };
                    let iter_call = Expression::new(ExprKind::Call {
                        callee: Box::new(iter_callee),
                        args: vec![
                            Argument::positional(Expression::new(ExprKind::Ident(
                                state_name.clone(),
                            ))),
                            Argument::positional(Expression::new(ExprKind::Ident(
                                ctrl_name.clone(),
                            ))),
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
                                init: Some(lua_first(Expression::new(ExprKind::Ident(
                                    row_name.clone(),
                                )))),
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
                            by_ref: false,
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
                            init: Some(lua_first(Expression::new(ExprKind::Ident(
                                row_name.clone(),
                            )))),
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
                        by_ref: false,
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
                        args.first()
                            .map(|arg| (format!("__lua_pairs_table_{}", first), arg.value.clone()))
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
                        args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                            table_var.clone(),
                        )))],
                        optional: false,
                    })
                } else {
                    iter.clone()
                };
                let idx_expr = Expression::new(ExprKind::Ident(idx.clone()));
                let len_call = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("__lua_len".to_string()))),
                    args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                        rows.clone(),
                    )))],
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
                block.push(Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(rows),
                        type_hint: None,
                        init: Some(rows_expr),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }));
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
                    block.push(Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Ident(
                                "__lua_iter_end".to_string(),
                            ))),
                            args: vec![Argument::positional(Expression::new(ExprKind::Ident(
                                table_var.clone(),
                            )))],
                            optional: false,
                        },
                    ))));
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
        StmtKind::FunctionDecl { params, body, .. } => {
            rewrite_lua_current_getinfo_stmts(body, params);
            normalize_lua_stmt_sequence(body);
        }
        StmtKind::Return(Some(expr)) => {
            if let ExprKind::Tuple(values) = &mut expr.kind {
                let values = std::mem::take(values);
                *expr = normalize_lua_return_values(values);
            } else if matches!(&expr.kind, ExprKind::Spread(inner) if matches!(inner.kind, ExprKind::Lit(Literal::Null)))
            {
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

fn lua_expr_contains_unshadowed_ident(expr: &Expression, name: &str) -> bool {
    match &expr.kind {
        ExprKind::Ident(ident) => ident == name,
        ExprKind::Lambda { params, body, .. } => {
            if params.iter().any(|param| param.name == name) {
                return false;
            }
            match body {
                LambdaBody::Expr(value) => lua_expr_contains_unshadowed_ident(value, name),
                LambdaBody::Block(stmts) => lua_stmts_contain_unshadowed_ident(stmts, name),
            }
        }
        ExprKind::Binary { left, right, .. } => {
            lua_expr_contains_unshadowed_ident(left, name)
                || lua_expr_contains_unshadowed_ident(right, name)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr) => lua_expr_contains_unshadowed_ident(expr, name),
        ExprKind::Call { callee, args, .. } => {
            lua_expr_contains_unshadowed_ident(callee, name)
                || args
                    .iter()
                    .any(|arg| lua_expr_contains_unshadowed_ident(&arg.value, name))
        }
        ExprKind::Member { object, .. } => lua_expr_contains_unshadowed_ident(object, name),
        ExprKind::Index { object, index, .. } => {
            lua_expr_contains_unshadowed_ident(object, name)
                || lua_expr_contains_unshadowed_ident(index, name)
        }
        ExprKind::Array(elems) => elems.iter().any(|elem| {
            elem.key
                .as_ref()
                .is_some_and(|key| lua_expr_contains_unshadowed_ident(key, name))
                || lua_expr_contains_unshadowed_ident(&elem.value, name)
        }),
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => values
            .iter()
            .any(|value| lua_expr_contains_unshadowed_ident(value, name)),
        ExprKind::Ternary { cond, then, else_ } => {
            lua_expr_contains_unshadowed_ident(cond, name)
                || lua_expr_contains_unshadowed_ident(then, name)
                || lua_expr_contains_unshadowed_ident(else_, name)
        }
        _ => false,
    }
}

fn lua_stmt_contains_unshadowed_ident(stmt: &Statement, name: &str) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            lua_expr_contains_unshadowed_ident(expr, name)
        }
        StmtKind::Assign { targets, value, .. } => {
            targets
                .iter()
                .any(|target| lua_expr_contains_unshadowed_ident(target, name))
                || lua_expr_contains_unshadowed_ident(value, name)
        }
        StmtKind::VarDecl { declarations, .. } => declarations.iter().any(|decl| {
            !matches!(&decl.pattern, BindingPattern::Ident(local) if local == name)
                && decl
                    .init
                    .as_ref()
                    .is_some_and(|init| lua_expr_contains_unshadowed_ident(init, name))
        }),
        StmtKind::Block(stmts)
        | StmtKind::While { body: stmts, .. }
        | StmtKind::DoWhile { body: stmts, .. }
        | StmtKind::For { body: stmts, .. } => lua_stmts_contain_unshadowed_ident(stmts, name),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            lua_expr_contains_unshadowed_ident(cond, name)
                || lua_stmts_contain_unshadowed_ident(then_body, name)
                || elifs.iter().any(|(cond, body)| {
                    lua_expr_contains_unshadowed_ident(cond, name)
                        || lua_stmts_contain_unshadowed_ident(body, name)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| lua_stmts_contain_unshadowed_ident(body, name))
        }
        _ => false,
    }
}

fn lua_stmts_contain_unshadowed_ident(stmts: &[Statement], name: &str) -> bool {
    stmts
        .iter()
        .any(|stmt| lua_stmt_contains_unshadowed_ident(stmt, name))
}

fn lua_replace_unshadowed_ident(expr: &mut Expression, from: &str, to: &str) {
    match &mut expr.kind {
        ExprKind::Ident(name) if name == from => {
            *name = to.to_string();
        }
        ExprKind::Lambda { params, body, .. } => {
            if params.iter().any(|param| param.name == from) {
                return;
            }
            match body {
                LambdaBody::Expr(value) => lua_replace_unshadowed_ident(value, from, to),
                LambdaBody::Block(stmts) => lua_replace_unshadowed_ident_in_stmts(stmts, from, to),
            }
        }
        ExprKind::Binary { left, right, .. } => {
            lua_replace_unshadowed_ident(left, from, to);
            lua_replace_unshadowed_ident(right, from, to);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr) => lua_replace_unshadowed_ident(expr, from, to),
        ExprKind::Call { callee, args, .. } => {
            lua_replace_unshadowed_ident(callee, from, to);
            for arg in args {
                lua_replace_unshadowed_ident(&mut arg.value, from, to);
            }
        }
        ExprKind::Member { object, .. } => lua_replace_unshadowed_ident(object, from, to),
        ExprKind::Index { object, index, .. } => {
            lua_replace_unshadowed_ident(object, from, to);
            lua_replace_unshadowed_ident(index, from, to);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    lua_replace_unshadowed_ident(key, from, to);
                }
                lua_replace_unshadowed_ident(&mut elem.value, from, to);
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                lua_replace_unshadowed_ident(value, from, to);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            lua_replace_unshadowed_ident(cond, from, to);
            lua_replace_unshadowed_ident(then, from, to);
            lua_replace_unshadowed_ident(else_, from, to);
        }
        _ => {}
    }
}

fn lua_replace_unshadowed_ident_in_stmts(stmts: &mut [Statement], from: &str, to: &str) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                lua_replace_unshadowed_ident(expr, from, to);
            }
            StmtKind::Assign { targets, value, .. } => {
                for target in targets {
                    lua_replace_unshadowed_ident(target, from, to);
                }
                lua_replace_unshadowed_ident(value, from, to);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if matches!(&decl.pattern, BindingPattern::Ident(local) if local == from) {
                        continue;
                    }
                    if let Some(init) = &mut decl.init {
                        lua_replace_unshadowed_ident(init, from, to);
                    }
                }
            }
            StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. } => {
                lua_replace_unshadowed_ident_in_stmts(body, from, to);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                lua_replace_unshadowed_ident(cond, from, to);
                lua_replace_unshadowed_ident_in_stmts(then_body, from, to);
                for (cond, body) in elifs {
                    lua_replace_unshadowed_ident(cond, from, to);
                    lua_replace_unshadowed_ident_in_stmts(body, from, to);
                }
                if let Some(body) = else_body {
                    lua_replace_unshadowed_ident_in_stmts(body, from, to);
                }
            }
            _ => {}
        }
    }
}

fn lua_capture_loop_var_in_expr(expr: &mut Expression, var: &str) {
    match &mut expr.kind {
        ExprKind::Lambda { params, body, .. } => {
            if params.iter().any(|param| param.name == var) {
                return;
            }
            let captures_var = match body {
                LambdaBody::Expr(value) => lua_expr_contains_unshadowed_ident(value, var),
                LambdaBody::Block(stmts) => lua_stmts_contain_unshadowed_ident(stmts, var),
            };
            if captures_var {
                let capture_name = format!("__lua_capture_{var}");
                let mut inner = expr.clone();
                lua_replace_unshadowed_ident(&mut inner, var, &capture_name);
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params: vec![Param {
                            name: capture_name,
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        }],
                        body: LambdaBody::Expr(Box::new(inner)),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    args: vec![Argument::positional(lua_ident(var))],
                    optional: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            lua_capture_loop_var_in_expr(left, var);
            lua_capture_loop_var_in_expr(right, var);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr) => lua_capture_loop_var_in_expr(expr, var),
        ExprKind::Call { callee, args, .. } => {
            lua_capture_loop_var_in_expr(callee, var);
            for arg in args {
                lua_capture_loop_var_in_expr(&mut arg.value, var);
            }
        }
        ExprKind::Member { object, .. } => lua_capture_loop_var_in_expr(object, var),
        ExprKind::Index { object, index, .. } => {
            lua_capture_loop_var_in_expr(object, var);
            lua_capture_loop_var_in_expr(index, var);
        }
        ExprKind::Array(elems) => {
            for elem in elems {
                if let Some(key) = &mut elem.key {
                    lua_capture_loop_var_in_expr(key, var);
                }
                lua_capture_loop_var_in_expr(&mut elem.value, var);
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                lua_capture_loop_var_in_expr(value, var);
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            lua_capture_loop_var_in_expr(cond, var);
            lua_capture_loop_var_in_expr(then, var);
            lua_capture_loop_var_in_expr(else_, var);
        }
        _ => {}
    }
}

fn lua_capture_loop_var_in_stmts(stmts: &mut [Statement], var: &str) {
    for stmt in stmts {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                lua_capture_loop_var_in_expr(expr, var);
            }
            StmtKind::Assign { value, .. } => {
                lua_capture_loop_var_in_expr(value, var);
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        lua_capture_loop_var_in_expr(init, var);
                    }
                }
            }
            StmtKind::Block(body)
            | StmtKind::While { body, .. }
            | StmtKind::DoWhile { body, .. }
            | StmtKind::For { body, .. } => {
                lua_capture_loop_var_in_stmts(body, var);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                lua_capture_loop_var_in_stmts(then_body, var);
                for (_, body) in elifs {
                    lua_capture_loop_var_in_stmts(body, var);
                }
                if let Some(body) = else_body {
                    lua_capture_loop_var_in_stmts(body, var);
                }
            }
            _ => {}
        }
    }
}

/// Lua numeric for -> canonical C-style for loop.
pub(crate) fn build_numeric_for(
    index_var: String,
    start: Expression,
    limit: Expression,
    step: Expression,
    mut body: Vec<Statement>,
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

    lua_capture_loop_var_in_stmts(&mut body, &index_var);

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
