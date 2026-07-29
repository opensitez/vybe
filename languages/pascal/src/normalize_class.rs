//! Pascal `ClassDecl` → `NormalClass` walker pass.
//!
//! Pascal / Delphi / Free Pascal class specifics:
//!   - `constructor Create;` / `constructor Init;` → NormalConstructor.
//!     Pascal's convention is `Create`; Free Pascal also allows `Init`.
//!   - `destructor Destroy;` → destructor.
//!   - `property Foo read GetFoo write SetFoo` → NormalProperty. Walker
//!     already links property accessors to their accessor methods.
//!   - `class operator Add(...)` / `class operator Equal(...)` →
//!     SpecialMethodKind::Add / Eq. The walker marks these as
//!     `operator_Add` / `operator_Equal` so role binding stays syntax-driven;
//!     a plain `procedure Add` is just a user method.
//!   - `override` / `virtual` / `reintroduce` → flag carries through.
//!   - Case-insensitive: Pascal method names lowercase to canonical.

use std::collections::{HashMap, HashSet};
use vybe_ast::{
    Argument, CaseCondition, ClassMember, ClassModifiers, ExprKind, Expression, Literal, Modifiers,
    PropertySetter, Span, Statement, StmtKind,
};
use vybe_bytecode::class_normalize::{
    NormalMembers, build_normal_method, from_method_stmt, types::*,
};

const PASCAL_NO_BASE_CTOR_MARKER: &str = "__pascal_no_base_ctor__";

fn property_field_name(body: &[Statement], field_names: &HashSet<String>) -> Option<String> {
    let [stmt] = body else {
        return None;
    };
    match &stmt.kind {
        StmtKind::Return(Some(expr)) => match &expr.kind {
            ExprKind::Member { object, field, .. } if matches!(object.kind, ExprKind::This) => {
                field_names
                    .contains(&field.to_ascii_lowercase())
                    .then(|| field.clone())
            }
            ExprKind::Ident(field) => field_names
                .contains(&field.to_ascii_lowercase())
                .then(|| field.clone()),
            ExprKind::Call { callee, args, .. } if args.is_empty() => match &callee.kind {
                ExprKind::Member { object, field, .. } if matches!(object.kind, ExprKind::This) => {
                    field_names
                        .contains(&field.to_ascii_lowercase())
                        .then(|| field.clone())
                }
                _ => None,
            },
            _ => None,
        },
        StmtKind::Expr(expr) => match &expr.kind {
            ExprKind::Assign { target, value } if matches!(value.kind, ExprKind::Ident(ref name) if name.eq_ignore_ascii_case("value")) => {
                match &target.kind {
                    ExprKind::Member { object, field, .. }
                        if matches!(object.kind, ExprKind::This) =>
                    {
                        field_names
                            .contains(&field.to_ascii_lowercase())
                            .then(|| field.clone())
                    }
                    ExprKind::Ident(field) => field_names
                        .contains(&field.to_ascii_lowercase())
                        .then(|| field.clone()),
                    _ => None,
                }
            }
            ExprKind::Call { callee, args, .. } if args.len() == 1 => match &callee.kind {
                ExprKind::Member { object, field, .. }
                    if matches!(object.kind, ExprKind::This)
                        && matches!(args[0].value.kind, ExprKind::Ident(ref name) if name.eq_ignore_ascii_case("value")) =>
                {
                    field_names
                        .contains(&field.to_ascii_lowercase())
                        .then(|| field.clone())
                }
                _ => None,
            },
            _ => None,
        },
        _ => None,
    }
}

fn rewrite_property_getter_body(
    body: &[Statement],
    field_names: &HashSet<String>,
) -> Vec<Statement> {
    let Some(field_name) = property_field_name(body, field_names) else {
        return body.to_vec();
    };
    vec![Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: field_name,
            null_safe: false,
        },
    ))))]
}

fn rewrite_property_setter_body(
    body: &[Statement],
    field_names: &HashSet<String>,
) -> Vec<Statement> {
    let Some(field_name) = property_field_name(body, field_names) else {
        return body.to_vec();
    };
    vec![Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: field_name,
            null_safe: false,
        })],
        value: Expression::ident("value"),
    })]
}

fn self_member_expr(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: name.to_string(),
        null_safe: false,
    })
}

fn rewrite_implicit_self_members_in_methods(
    methods: &mut [NormalMethod],
    member_names: &HashSet<String>,
) {
    for method in methods {
        let mut shadowed: HashSet<String> = method
            .params
            .iter()
            .map(|param| param.name.to_ascii_lowercase())
            .collect();
        rewrite_implicit_self_members_in_body(&mut method.body, member_names, &mut shadowed);
    }
}

fn extend_gcl_member_names(member_names: &mut HashSet<String>, parents: &[String]) {
    let classes = vybe_platform_plib::emitter::gcl::gcl_classes();
    let mut pending: Vec<String> = parents.to_vec();
    while let Some(class_name) = pending.pop() {
        let Some(class) = classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(&class_name))
        else {
            continue;
        };
        for property in class.properties {
            member_names.insert(property.to_ascii_lowercase());
        }
        if let Some(parent) = class.parent {
            pending.push(parent.to_string());
        }
    }
}

fn rewrite_implicit_self_members_in_constructors(
    constructors: &mut [NormalConstructor],
    member_names: &HashSet<String>,
) {
    for constructor in constructors {
        let mut shadowed: HashSet<String> = constructor
            .params
            .iter()
            .map(|param| param.name.to_ascii_lowercase())
            .collect();
        rewrite_implicit_self_members_in_body(&mut constructor.body, member_names, &mut shadowed);
    }
}

fn static_access_expr(class_name: &str, member_name: &str) -> Expression {
    Expression::new(ExprKind::StaticAccess {
        class: Box::new(Expression::ident(class_name)),
        member: Box::new(Expression::ident(member_name)),
    })
}

fn rewrite_static_value_members_in_methods(
    methods: &mut [NormalMethod],
    class_name: &str,
    member_names: &HashSet<String>,
) {
    for method in methods {
        let mut shadowed: HashSet<String> = method
            .params
            .iter()
            .map(|param| param.name.to_ascii_lowercase())
            .collect();
        rewrite_static_value_members_in_body(
            &mut method.body,
            class_name,
            member_names,
            &mut shadowed,
        );
    }
}

fn rewrite_static_value_members_in_constructors(
    constructors: &mut [NormalConstructor],
    class_name: &str,
    member_names: &HashSet<String>,
) {
    for constructor in constructors {
        let mut shadowed: HashSet<String> = constructor
            .params
            .iter()
            .map(|param| param.name.to_ascii_lowercase())
            .collect();
        rewrite_static_value_members_in_body(
            &mut constructor.body,
            class_name,
            member_names,
            &mut shadowed,
        );
    }
}

fn rewrite_static_value_members_in_body(
    body: &mut [Statement],
    class_name: &str,
    member_names: &HashSet<String>,
    shadowed: &mut HashSet<String>,
) {
    for stmt in body {
        rewrite_static_value_members_stmt(stmt, class_name, member_names, shadowed);
    }
}

fn rewrite_static_value_members_stmt(
    stmt: &mut Statement,
    class_name: &str,
    member_names: &HashSet<String>,
    shadowed: &mut HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_static_value_members_expr(expr, class_name, member_names, shadowed);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    rewrite_static_value_members_expr(init, class_name, member_names, shadowed);
                }
            }
            for decl in declarations {
                if let vybe_ast::BindingPattern::Ident(name) = &decl.pattern {
                    shadowed.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_static_value_members_expr(target, class_name, member_names, shadowed);
            }
            rewrite_static_value_members_expr(value, class_name, member_names, shadowed);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_static_value_members_expr(target, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(value, class_name, member_names, shadowed);
        }
        StmtKind::Block(body) => {
            let mut scoped = shadowed.clone();
            rewrite_static_value_members_in_body(body, class_name, member_names, &mut scoped);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_static_value_members_expr(cond, class_name, member_names, shadowed);
            rewrite_static_value_members_in_body(
                then_body,
                class_name,
                member_names,
                &mut shadowed.clone(),
            );
            for (cond, body) in elifs {
                rewrite_static_value_members_expr(cond, class_name, member_names, shadowed);
                rewrite_static_value_members_in_body(
                    body,
                    class_name,
                    member_names,
                    &mut shadowed.clone(),
                );
            }
            if let Some(body) = else_body {
                rewrite_static_value_members_in_body(
                    body,
                    class_name,
                    member_names,
                    &mut shadowed.clone(),
                );
            }
        }
        StmtKind::While { cond, body, .. } => {
            rewrite_static_value_members_expr(cond, class_name, member_names, shadowed);
            rewrite_static_value_members_in_body(
                body,
                class_name,
                member_names,
                &mut shadowed.clone(),
            );
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = shadowed.clone();
            if let Some(init) = init {
                rewrite_static_value_members_stmt(init, class_name, member_names, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_static_value_members_expr(cond, class_name, member_names, &mut scoped);
            }
            if let Some(update) = update {
                rewrite_static_value_members_expr(update, class_name, member_names, &mut scoped);
            }
            rewrite_static_value_members_in_body(body, class_name, member_names, &mut scoped);
        }
        _ => {}
    }
}

fn rewrite_static_value_members_expr(
    expr: &mut Expression,
    class_name: &str,
    member_names: &HashSet<String>,
    shadowed: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name)
            if member_names.contains(&name.to_ascii_lowercase())
                && !shadowed.contains(&name.to_ascii_lowercase()) =>
        {
            *expr = static_access_expr(class_name, name);
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(name) = &callee.kind {
                if member_names.contains(&name.to_ascii_lowercase())
                    && !shadowed.contains(&name.to_ascii_lowercase())
                {
                    *callee = Box::new(static_access_expr(class_name, name));
                }
            } else {
                rewrite_static_value_members_expr(callee, class_name, member_names, shadowed);
            }
            for arg in args {
                rewrite_static_value_members_expr(
                    &mut arg.value,
                    class_name,
                    member_names,
                    shadowed,
                );
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_static_value_members_expr(object, class_name, member_names, shadowed);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_static_value_members_expr(object, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(index, class_name, member_names, shadowed);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_static_value_members_expr(left, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(right, class_name, member_names, shadowed);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_static_value_members_expr(expr, class_name, member_names, shadowed);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_static_value_members_expr(cond, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(then, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(else_, class_name, member_names, shadowed);
        }
        ExprKind::Assign { target, value } => {
            rewrite_static_value_members_expr(target, class_name, member_names, shadowed);
            rewrite_static_value_members_expr(value, class_name, member_names, shadowed);
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_static_value_members_expr(
                    &mut arg.value,
                    class_name,
                    member_names,
                    shadowed,
                );
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_static_value_members_expr(key, class_name, member_names, shadowed);
                }
                rewrite_static_value_members_expr(
                    &mut element.value,
                    class_name,
                    member_names,
                    shadowed,
                );
            }
        }
        _ => {}
    }
}

fn normalize_destructor_inherited_calls(body: &mut [Statement], has_parent: bool) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::Expr(expr) => normalize_destructor_inherited_expr(expr, has_parent),
            StmtKind::Block(body) => normalize_destructor_inherited_calls(body, has_parent),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                normalize_destructor_inherited_calls(then_body, has_parent);
                for (_, body) in elifs {
                    normalize_destructor_inherited_calls(body, has_parent);
                }
                if let Some(body) = else_body {
                    normalize_destructor_inherited_calls(body, has_parent);
                }
            }
            StmtKind::While { body, .. }
            | StmtKind::For { body, .. }
            | StmtKind::ForIn { body, .. }
            | StmtKind::DoWhile { body, .. } => {
                normalize_destructor_inherited_calls(body, has_parent);
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                normalize_destructor_inherited_calls(body, has_parent);
                for catch in catches {
                    normalize_destructor_inherited_calls(&mut catch.body, has_parent);
                }
                if let Some(body) = else_body {
                    normalize_destructor_inherited_calls(body, has_parent);
                }
                if let Some(body) = finally {
                    normalize_destructor_inherited_calls(body, has_parent);
                }
            }
            _ => {}
        }
    }
}

fn normalize_destructor_inherited_expr(expr: &mut Expression, has_parent: bool) {
    match &mut expr.kind {
        ExprKind::SuperCall { method, .. } if method.is_none() => {
            if has_parent {
                *method = Some("Destroy".to_string());
            } else {
                *expr = Expression::null();
            }
        }
        ExprKind::SuperCall { method, .. }
            if !has_parent
                && method
                    .as_deref()
                    .is_some_and(|name| name.eq_ignore_ascii_case("Destroy")) =>
        {
            *expr = Expression::null();
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_destructor_inherited_expr(callee, has_parent);
            for arg in args {
                normalize_destructor_inherited_expr(&mut arg.value, has_parent);
            }
        }
        ExprKind::Member { object, .. } => normalize_destructor_inherited_expr(object, has_parent),
        ExprKind::Index { object, index, .. } => {
            normalize_destructor_inherited_expr(object, has_parent);
            normalize_destructor_inherited_expr(index, has_parent);
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_destructor_inherited_expr(left, has_parent);
            normalize_destructor_inherited_expr(right, has_parent);
        }
        ExprKind::Unary { expr, .. } => normalize_destructor_inherited_expr(expr, has_parent),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_destructor_inherited_expr(cond, has_parent);
            normalize_destructor_inherited_expr(then, has_parent);
            normalize_destructor_inherited_expr(else_, has_parent);
        }
        ExprKind::Assign { target, value } => {
            normalize_destructor_inherited_expr(target, has_parent);
            normalize_destructor_inherited_expr(value, has_parent);
        }
        _ => {}
    }
}

fn rewrite_implicit_self_members_in_body(
    body: &mut [Statement],
    member_names: &HashSet<String>,
    shadowed: &mut HashSet<String>,
) {
    for stmt in body {
        rewrite_implicit_self_members_stmt(stmt, member_names, shadowed);
    }
}

fn rewrite_implicit_self_members_stmt(
    stmt: &mut Statement,
    member_names: &HashSet<String>,
    shadowed: &mut HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_implicit_self_members_expr(expr, member_names, shadowed, false);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if let Some(init) = &mut decl.init {
                    rewrite_implicit_self_members_expr(init, member_names, shadowed, false);
                }
            }
            for decl in declarations {
                if let vybe_ast::BindingPattern::Ident(name) = &decl.pattern {
                    shadowed.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_implicit_self_members_expr(target, member_names, shadowed, true);
            }
            rewrite_implicit_self_members_expr(value, member_names, shadowed, false);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_implicit_self_members_expr(target, member_names, shadowed, true);
            rewrite_implicit_self_members_expr(value, member_names, shadowed, false);
        }
        StmtKind::Block(body) => {
            let mut scoped = shadowed.clone();
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_implicit_self_members_expr(cond, member_names, shadowed, false);
            let mut scoped = shadowed.clone();
            rewrite_implicit_self_members_in_body(then_body, member_names, &mut scoped);
            for (cond, body) in elifs {
                rewrite_implicit_self_members_expr(cond, member_names, shadowed, false);
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
            if let Some(body) = else_body {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = shadowed.clone();
            if let Some(init) = init {
                rewrite_implicit_self_members_stmt(init, member_names, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_implicit_self_members_expr(cond, member_names, &mut scoped, false);
            }
            if let Some(update) = update {
                rewrite_implicit_self_members_expr(update, member_names, &mut scoped, false);
            }
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_implicit_self_members_expr(iter, member_names, shadowed, false);
            let mut scoped = shadowed.clone();
            scoped.insert(var.to_ascii_lowercase());
            if let Some(key) = key {
                scoped.insert(key.to_ascii_lowercase());
            }
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            if let Some(body) = else_body {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_implicit_self_members_expr(cond, member_names, shadowed, false);
            let mut scoped = shadowed.clone();
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            if let Some(body) = else_body {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut scoped = shadowed.clone();
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            rewrite_implicit_self_members_expr(cond, member_names, shadowed, false);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_implicit_self_members_expr(expr, member_names, shadowed, false);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        vybe_ast::CaseCondition::Value(expr)
                        | vybe_ast::CaseCondition::Comparison { expr, .. } => {
                            rewrite_implicit_self_members_expr(expr, member_names, shadowed, false);
                        }
                        vybe_ast::CaseCondition::Range { from, to } => {
                            rewrite_implicit_self_members_expr(from, member_names, shadowed, false);
                            rewrite_implicit_self_members_expr(to, member_names, shadowed, false);
                        }
                    }
                }
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(&mut case.body, member_names, &mut scoped);
            }
            if let Some(body) = default {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let mut scoped = shadowed.clone();
            rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            for catch in catches {
                let mut scoped = shadowed.clone();
                if let Some(var) = &catch.var_name {
                    scoped.insert(var.to_ascii_lowercase());
                }
                rewrite_implicit_self_members_in_body(&mut catch.body, member_names, &mut scoped);
            }
            if let Some(body) = else_body {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
            if let Some(body) = finally {
                let mut scoped = shadowed.clone();
                rewrite_implicit_self_members_in_body(body, member_names, &mut scoped);
            }
        }
        StmtKind::With { items, .. } => {
            for item in items {
                rewrite_implicit_self_members_expr(&mut item.expr, member_names, shadowed, false);
            }
            // Bare names inside `with` belong to the with-target, not Self.
        }
        _ => {}
    }
}

fn rewrite_implicit_self_members_expr(
    expr: &mut Expression,
    member_names: &HashSet<String>,
    shadowed: &HashSet<String>,
    assignment_target: bool,
) {
    match &mut expr.kind {
        ExprKind::Ident(name)
            if member_names.contains(&name.to_ascii_lowercase())
                && (assignment_target || !shadowed.contains(&name.to_ascii_lowercase())) =>
        {
            *expr = self_member_expr(name);
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Ident(name) = &callee.kind {
                let key = name.to_ascii_lowercase();
                if member_names.contains(&key) && !shadowed.contains(&key) {
                    *callee = Box::new(self_member_expr(name));
                }
            } else {
                rewrite_implicit_self_members_expr(callee, member_names, shadowed, false);
            }
            for arg in args {
                rewrite_implicit_self_members_expr(&mut arg.value, member_names, shadowed, false);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_implicit_self_members_expr(object, member_names, shadowed, false);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_implicit_self_members_expr(object, member_names, shadowed, assignment_target);
            rewrite_implicit_self_members_expr(index, member_names, shadowed, false);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_implicit_self_members_expr(left, member_names, shadowed, false);
            rewrite_implicit_self_members_expr(right, member_names, shadowed, false);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_implicit_self_members_expr(expr, member_names, shadowed, false);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_implicit_self_members_expr(cond, member_names, shadowed, false);
            rewrite_implicit_self_members_expr(then, member_names, shadowed, false);
            rewrite_implicit_self_members_expr(else_, member_names, shadowed, false);
        }
        ExprKind::Assign { target, value } => {
            rewrite_implicit_self_members_expr(target, member_names, shadowed, true);
            rewrite_implicit_self_members_expr(value, member_names, shadowed, false);
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_implicit_self_members_expr(&mut arg.value, member_names, shadowed, false);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_implicit_self_members_expr(key, member_names, shadowed, false);
                }
                rewrite_implicit_self_members_expr(
                    &mut element.value,
                    member_names,
                    shadowed,
                    false,
                );
            }
        }
        _ => {}
    }
}

fn gcl_accessor_property_names(parents: &[String]) -> HashSet<String> {
    let classes = vybe_platform_plib::emitter::gcl::gcl_classes();
    let mut names = HashSet::new();
    let mut pending: Vec<String> = parents.to_vec();
    while let Some(class_name) = pending.pop() {
        let Some(class) = classes
            .iter()
            .find(|class| class.name.eq_ignore_ascii_case(&class_name))
        else {
            continue;
        };
        for property in class.properties {
            let lower = property.to_ascii_lowercase();
            if !matches!(lower.as_str(), "controls" | "items" | "components") {
                names.insert(lower);
            }
        }
        if let Some(parent) = class.parent {
            pending.push(parent.to_string());
        }
    }
    names
}

fn gcl_accessor_call(
    object: Expression,
    prefix: &str,
    field: &str,
    args: Vec<Argument>,
) -> Expression {
    let key = format!("{}_{}", prefix, field.to_ascii_lowercase());
    // `this` is NOT passed explicitly: plib's bind_ref stamps
    // `__vybe_method_receiver` on every accessor ref, and the call path
    // prepends that receiver. An explicit object arg here doubled the
    // receiver — setters got (this, this, value) and the value fell off
    // the arity-2 chunk.
    let explicit_args = args;
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(Expression::new(ExprKind::Lit(Literal::Str(key)))),
            null_safe: false,
        })),
        args: explicit_args,
        optional: false,
    })
}

fn rewrite_gcl_property_accessors_in_methods(
    methods: &mut [NormalMethod],
    property_names: &HashSet<String>,
) {
    for method in methods {
        rewrite_gcl_property_accessors_in_body(&mut method.body, property_names);
    }
}

fn rewrite_gcl_property_accessors_in_constructors(
    constructors: &mut [NormalConstructor],
    property_names: &HashSet<String>,
) {
    for constructor in constructors {
        rewrite_gcl_property_accessors_in_body(&mut constructor.body, property_names);
    }
}

fn rewrite_gcl_property_accessors_in_body(
    body: &mut [Statement],
    property_names: &HashSet<String>,
) {
    for stmt in body {
        rewrite_gcl_property_accessors_stmt(stmt, property_names);
    }
}

fn rewrite_gcl_property_accessors_stmt(stmt: &mut Statement, property_names: &HashSet<String>) {
    let setter_rewrite = match &mut stmt.kind {
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            rewrite_gcl_property_accessors_expr(value, property_names);
            rewrite_gcl_property_setter_target(&mut targets[0], value.clone(), property_names)
        }
        StmtKind::Expr(expr) => {
            if let ExprKind::Assign { target, value } = &mut expr.kind {
                rewrite_gcl_property_accessors_expr(value, property_names);
                rewrite_gcl_property_setter_target(target, (**value).clone(), property_names)
            } else {
                rewrite_gcl_property_accessors_expr(expr, property_names);
                None
            }
        }
        _ => None,
    };
    if let Some(rewritten) = setter_rewrite {
        *stmt = Statement::new(StmtKind::Expr(rewritten));
        return;
    }

    match &mut stmt.kind {
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_gcl_property_accessors_target_object(target, property_names);
            }
            rewrite_gcl_property_accessors_expr(value, property_names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_gcl_property_accessors_target_object(target, property_names);
            rewrite_gcl_property_accessors_expr(value, property_names);
        }
        StmtKind::Block(stmts) => rewrite_gcl_property_accessors_in_body(stmts, property_names),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_gcl_property_accessors_expr(init, property_names);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_gcl_property_accessors_expr(bound, property_names);
                    }
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_gcl_property_accessors_expr(cond, property_names);
            rewrite_gcl_property_accessors_in_body(then_body, property_names);
            for (cond, body) in elifs {
                rewrite_gcl_property_accessors_expr(cond, property_names);
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
            if let Some(body) = else_body {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
            ..
        } => {
            rewrite_gcl_property_accessors_expr(cond, property_names);
            rewrite_gcl_property_accessors_in_body(body, property_names);
            if let Some(body) = else_body {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
            ..
        } => {
            if let Some(init) = init {
                rewrite_gcl_property_accessors_stmt(init, property_names);
            }
            if let Some(cond) = cond {
                rewrite_gcl_property_accessors_expr(cond, property_names);
            }
            if let Some(update) = update {
                rewrite_gcl_property_accessors_expr(update, property_names);
            }
            rewrite_gcl_property_accessors_in_body(body, property_names);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_gcl_property_accessors_expr(iter, property_names);
            rewrite_gcl_property_accessors_in_body(body, property_names);
            if let Some(body) = else_body {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_gcl_property_accessors_in_body(body, property_names);
            rewrite_gcl_property_accessors_expr(cond, property_names);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_gcl_property_accessors_expr(expr, property_names);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) => {
                            rewrite_gcl_property_accessors_expr(expr, property_names);
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_gcl_property_accessors_expr(from, property_names);
                            rewrite_gcl_property_accessors_expr(to, property_names);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_gcl_property_accessors_expr(expr, property_names);
                        }
                    }
                }
                rewrite_gcl_property_accessors_in_body(&mut case.body, property_names);
            }
            if let Some(body) = default {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        StmtKind::Return(Some(expr)) => {
            rewrite_gcl_property_accessors_expr(expr, property_names);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_gcl_property_accessors_expr(expr, property_names);
            }
            if let Some(cause) = cause {
                rewrite_gcl_property_accessors_expr(cause, property_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_gcl_property_accessors_in_body(body, property_names);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_gcl_property_accessors_expr(when_clause, property_names);
                }
                rewrite_gcl_property_accessors_in_body(&mut catch.body, property_names);
            }
            if let Some(body) = else_body {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
            if let Some(body) = finally {
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        StmtKind::Using { resource, body, .. } => {
            rewrite_gcl_property_accessors_expr(resource, property_names);
            rewrite_gcl_property_accessors_in_body(body, property_names);
        }
        StmtKind::With { items, body, .. } => {
            for item in &mut *items {
                rewrite_gcl_property_accessors_expr(&mut item.expr, property_names);
            }
            if let Some(item) = items.first_mut() {
                let receiver = item
                    .var
                    .get_or_insert_with(|| "__gcl_with_target".to_string())
                    .clone();
                rewrite_gcl_with_receiver_body(body, &receiver, property_names);
                rewrite_gcl_property_accessors_in_body(body, property_names);
            }
        }
        _ => {}
    }
}

fn gcl_with_receiver_member(receiver: &str, name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(receiver)),
        field: name.to_string(),
        null_safe: false,
    })
}

fn rewrite_gcl_with_receiver_body(
    body: &mut [Statement],
    receiver: &str,
    property_names: &HashSet<String>,
) {
    for stmt in body {
        rewrite_gcl_with_receiver_stmt(stmt, receiver, property_names);
    }
}

fn rewrite_gcl_with_receiver_stmt(
    stmt: &mut Statement,
    receiver: &str,
    property_names: &HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false)
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_gcl_with_receiver_expr(target, receiver, property_names, true);
            }
            rewrite_gcl_with_receiver_expr(value, receiver, property_names, false);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_gcl_with_receiver_expr(target, receiver, property_names, true);
            rewrite_gcl_with_receiver_expr(value, receiver, property_names, false);
        }
        StmtKind::Block(stmts) => rewrite_gcl_with_receiver_body(stmts, receiver, property_names),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_gcl_with_receiver_expr(init, receiver, property_names, false);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_gcl_with_receiver_expr(bound, receiver, property_names, false);
                    }
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
            rewrite_gcl_with_receiver_body(then_body, receiver, property_names);
            for (cond, body) in elifs {
                rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
            if let Some(body) = else_body {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
            if let Some(body) = else_body {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_gcl_with_receiver_stmt(init, receiver, property_names);
            }
            if let Some(cond) = cond {
                rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
            }
            if let Some(update) = update {
                rewrite_gcl_with_receiver_expr(update, receiver, property_names, false);
            }
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_gcl_with_receiver_expr(iter, receiver, property_names, false);
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
            if let Some(body) = else_body {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
            rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false);
            for case in cases {
                for cond in &mut case.conditions {
                    match cond {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false);
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_gcl_with_receiver_expr(from, receiver, property_names, false);
                            rewrite_gcl_with_receiver_expr(to, receiver, property_names, false);
                        }
                    }
                }
                rewrite_gcl_with_receiver_body(&mut case.body, receiver, property_names);
            }
            if let Some(body) = default {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
        }
        StmtKind::Return(Some(expr)) => {
            rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false);
            }
            if let Some(cause) = cause {
                rewrite_gcl_with_receiver_expr(cause, receiver, property_names, false);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_gcl_with_receiver_expr(when_clause, receiver, property_names, false);
                }
                rewrite_gcl_with_receiver_body(&mut catch.body, receiver, property_names);
            }
            if let Some(body) = else_body {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
            if let Some(body) = finally {
                rewrite_gcl_with_receiver_body(body, receiver, property_names);
            }
        }
        StmtKind::Using { resource, body, .. } => {
            rewrite_gcl_with_receiver_expr(resource, receiver, property_names, false);
            rewrite_gcl_with_receiver_body(body, receiver, property_names);
        }
        StmtKind::With { items, .. } => {
            for item in items {
                rewrite_gcl_with_receiver_expr(&mut item.expr, receiver, property_names, false);
            }
        }
        _ => {}
    }
}

fn rewrite_gcl_with_receiver_expr(
    expr: &mut Expression,
    receiver: &str,
    property_names: &HashSet<String>,
    assignment_target: bool,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) if property_names.contains(&name.to_ascii_lowercase()) => {
            *expr = gcl_with_receiver_member(receiver, name);
        }
        ExprKind::Call { callee, args, .. } => {
            if !matches!(&callee.kind, ExprKind::Ident(_)) {
                rewrite_gcl_with_receiver_expr(callee, receiver, property_names, false);
            }
            for arg in args {
                rewrite_gcl_with_receiver_expr(&mut arg.value, receiver, property_names, false);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_gcl_with_receiver_expr(object, receiver, property_names, false);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_gcl_with_receiver_expr(object, receiver, property_names, assignment_target);
            rewrite_gcl_with_receiver_expr(index, receiver, property_names, false);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_gcl_with_receiver_expr(left, receiver, property_names, false);
            rewrite_gcl_with_receiver_expr(right, receiver, property_names, false);
        }
        ExprKind::Unary { expr, .. } => {
            rewrite_gcl_with_receiver_expr(expr, receiver, property_names, false);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_gcl_with_receiver_expr(cond, receiver, property_names, false);
            rewrite_gcl_with_receiver_expr(then, receiver, property_names, false);
            rewrite_gcl_with_receiver_expr(else_, receiver, property_names, false);
        }
        ExprKind::Assign { target, value } => {
            rewrite_gcl_with_receiver_expr(target, receiver, property_names, true);
            rewrite_gcl_with_receiver_expr(value, receiver, property_names, false);
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_gcl_with_receiver_expr(&mut arg.value, receiver, property_names, false);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_gcl_with_receiver_expr(key, receiver, property_names, false);
                }
                rewrite_gcl_with_receiver_expr(&mut element.value, receiver, property_names, false);
            }
        }
        _ => {}
    }
}

fn rewrite_gcl_property_setter_target(
    target: &mut Expression,
    value: Expression,
    property_names: &HashSet<String>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &mut target.kind else {
        rewrite_gcl_property_accessors_target_object(target, property_names);
        return None;
    };
    rewrite_gcl_property_accessors_expr(object, property_names);
    if !property_names.contains(&field.to_ascii_lowercase()) {
        return None;
    }
    Some(gcl_accessor_call(
        (**object).clone(),
        "__set",
        field,
        vec![Argument::positional(value)],
    ))
}

fn rewrite_gcl_property_accessors_target_object(
    target: &mut Expression,
    property_names: &HashSet<String>,
) {
    match &mut target.kind {
        ExprKind::Member { object, .. } => {
            rewrite_gcl_property_accessors_expr(object, property_names)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_gcl_property_accessors_expr(object, property_names);
            rewrite_gcl_property_accessors_expr(index, property_names);
        }
        _ => rewrite_gcl_property_accessors_expr(target, property_names),
    }
}

fn rewrite_gcl_property_accessors_expr(expr: &mut Expression, property_names: &HashSet<String>) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            rewrite_gcl_property_accessors_expr(object, property_names);
            if property_names.contains(&field.to_ascii_lowercase()) {
                let object_expr = (**object).clone();
                let field_name = field.clone();
                *expr = gcl_accessor_call(object_expr, "__get", &field_name, Vec::new());
            }
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_gcl_property_accessors_expr(callee, property_names);
            for arg in args {
                rewrite_gcl_property_accessors_expr(&mut arg.value, property_names);
            }
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_gcl_property_accessors_expr(object, property_names);
            rewrite_gcl_property_accessors_expr(index, property_names);
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_gcl_property_accessors_expr(left, property_names);
            rewrite_gcl_property_accessors_expr(right, property_names);
        }
        ExprKind::Unary { expr, .. } => rewrite_gcl_property_accessors_expr(expr, property_names),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_gcl_property_accessors_expr(cond, property_names);
            rewrite_gcl_property_accessors_expr(then, property_names);
            rewrite_gcl_property_accessors_expr(else_, property_names);
        }
        ExprKind::Assign { target, value } => {
            rewrite_gcl_property_accessors_expr(value, property_names);
            if let Some(rewritten) =
                rewrite_gcl_property_setter_target(target, (**value).clone(), property_names)
            {
                *expr = rewritten;
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_gcl_property_accessors_expr(&mut arg.value, property_names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_gcl_property_accessors_expr(key, property_names);
                }
                rewrite_gcl_property_accessors_expr(&mut element.value, property_names);
            }
        }
        _ => {}
    }
}

const GCL_FORM_CREATE_AUTOINIT: &str = "__gcl_form_create_autoinit";

fn call_self_form_create_body() -> Vec<Statement> {
    vec![Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Call {
            callee: Box::new(Expression::ident("FormCreate")),
            args: vec![vybe_ast::Argument::positional(Expression::new(
                ExprKind::This,
            ))],
            optional: false,
        },
    )))]
}

fn stmt_calls_method(stmt: &Statement, method_name: &str) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => expr_calls_method(expr, method_name),
        StmtKind::Block(body)
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => {
            body.iter().any(|stmt| stmt_calls_method(stmt, method_name))
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            expr_calls_method(cond, method_name)
                || then_body
                    .iter()
                    .any(|stmt| stmt_calls_method(stmt, method_name))
                || elifs.iter().any(|(cond, body)| {
                    expr_calls_method(cond, method_name)
                        || body.iter().any(|stmt| stmt_calls_method(stmt, method_name))
                })
                || else_body.as_ref().is_some_and(|body| {
                    body.iter().any(|stmt| stmt_calls_method(stmt, method_name))
                })
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            init.as_ref()
                .is_some_and(|stmt| stmt_calls_method(stmt, method_name))
                || cond
                    .as_ref()
                    .is_some_and(|expr| expr_calls_method(expr, method_name))
                || update
                    .as_ref()
                    .is_some_and(|expr| expr_calls_method(expr, method_name))
                || body.iter().any(|stmt| stmt_calls_method(stmt, method_name))
        }
        StmtKind::ForIn { iter, body, .. } => {
            expr_calls_method(iter, method_name)
                || body.iter().any(|stmt| stmt_calls_method(stmt, method_name))
        }
        StmtKind::Assign { targets, value } => {
            targets
                .iter()
                .any(|expr| expr_calls_method(expr, method_name))
                || expr_calls_method(value, method_name)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_calls_method(target, method_name) || expr_calls_method(value, method_name)
        }
        _ => false,
    }
}

fn expr_calls_method(expr: &Expression, method_name: &str) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            match &callee.kind {
                ExprKind::Ident(name) | ExprKind::Member { field: name, .. }
                    if name.eq_ignore_ascii_case(method_name) =>
                {
                    return true;
                }
                _ => {}
            }
            expr_calls_method(callee, method_name)
                || args
                    .iter()
                    .any(|arg| expr_calls_method(&arg.value, method_name))
        }
        ExprKind::Member { object, .. } => expr_calls_method(object, method_name),
        ExprKind::Index { object, index, .. } => {
            expr_calls_method(object, method_name) || expr_calls_method(index, method_name)
        }
        ExprKind::Binary { left, right, .. } => {
            expr_calls_method(left, method_name) || expr_calls_method(right, method_name)
        }
        ExprKind::Unary { expr, .. } => expr_calls_method(expr, method_name),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_calls_method(cond, method_name)
                || expr_calls_method(then, method_name)
                || expr_calls_method(else_, method_name)
        }
        ExprKind::Assign { target, value } => {
            expr_calls_method(target, method_name) || expr_calls_method(value, method_name)
        }
        ExprKind::New { class, args } => {
            expr_calls_method(class, method_name)
                || args
                    .iter()
                    .any(|arg| expr_calls_method(&arg.value, method_name))
        }
        _ => false,
    }
}

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    _modifiers: &ClassModifiers,
) -> NormalClass {
    let mut out = NormalMembers::default();
    let instance_field_names: HashSet<String> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name, modifiers, ..
            } if !modifiers.is_static => Some(name.to_ascii_lowercase()),
            _ => None,
        })
        .collect();
    let static_value_member_names: HashSet<String> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name, modifiers, ..
            } if modifiers.is_static => Some(name.to_ascii_lowercase()),
            ClassMember::Const { name, .. } => Some(name.to_ascii_lowercase()),
            ClassMember::Property {
                name, modifiers, ..
            } if modifiers.is_static => Some(name.to_ascii_lowercase()),
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl {
                    name, modifiers, ..
                } if modifiers.is_static => Some(name.to_ascii_lowercase()),
                _ => None,
            },
            _ => None,
        })
        .collect();
    let mut implicit_self_member_names = instance_field_names.clone();

    for member in members {
        match member {
            ClassMember::Field {
                name: fname,
                type_hint,
                init,
                modifiers: m,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::from(m.visibility.clone()),
                    readonly: m.is_readonly,
                };
                out.push_field(m.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: src_name,
                    modifiers: m,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };

                // Pascal protocol binding is syntax-driven for operators:
                // `class operator Add` arrives from the walker as
                // `operator_Add`, while a plain `procedure Add` is just a
                // user method named Add and must remain callable as `add`
                // across languages.
                let operator_name = src_name.strip_prefix("operator_");
                let protocol_raw_name = operator_name.unwrap_or(src_name);
                let protocol_source_name = protocol_raw_name
                    .split_once("__pascal_overload_")
                    .map_or(protocol_raw_name, |(base, _)| base);
                let (canonical, mut special_kind) =
                    crate::protocol::canonical_method(protocol_source_name);
                if operator_name.is_none()
                    && !matches!(
                        special_kind,
                        Some(SpecialMethodKind::Destructor | SpecialMethodKind::ToString)
                    )
                {
                    special_kind = None;
                }
                // `destructor Destroy` — a lifecycle member, not a method. The
                // spelling is declared in the shared canonical table; the
                // `inherited` rewrite below is genuine Pascal semantics and
                // stays, which is why this arm builds its own statement rather
                // than routing on the kind after the fact.
                if special_kind == Some(SpecialMethodKind::Destructor) {
                    let mut stmt = (**stmt).clone();
                    if let StmtKind::FunctionDecl { body, .. } = &mut stmt.kind {
                        normalize_destructor_inherited_calls(body, !parents.is_empty());
                    }
                    if let Some(d) = from_method_stmt(
                        span.clone(),
                        &stmt,
                        &canonical,
                        Access::from(m.visibility.clone()),
                    ) {
                        out.destructor = Some(d);
                    }
                    continue;
                }

                let mut callable_stmt;
                let method_stmt =
                    if operator_name.is_some() && !src_name.contains("__pascal_overload_") {
                        callable_stmt = (**stmt).clone();
                        if let StmtKind::FunctionDecl { name, .. } = &mut callable_stmt.kind {
                            *name = canonical.clone();
                        }
                        &callable_stmt
                    } else {
                        stmt
                    };

                let Some(method) = from_method_stmt(
                    span.clone(),
                    method_stmt,
                    &canonical,
                    Access::from(m.visibility.clone()),
                ) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                // An instance method's name is what a bare identifier inside
                // another method can resolve to as `Self.<name>` — a Pascal
                // rule, so it is stated here rather than in the router.
                if !m.is_static {
                    implicit_self_member_names.insert(src_name.to_ascii_lowercase());
                }
                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                name: _constructor_name,
                params,
                body,
                base_args,
                initializer_target,
                visibility: _constructor_visibility,
                ..
            } => {
                let mut body = body.clone();
                let suppress_base_call = body.first().is_some_and(is_pascal_no_base_ctor_marker);
                if suppress_base_call {
                    body.remove(0);
                }
                out.push_constructor(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body,
                    base_call: match base_args {
                        Some(args) => {
                            let args = args
                                .iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect();
                            match initializer_target {
                                vybe_ast::ConstructorInitializerTarget::This => {
                                    BaseCall::This(args)
                                }
                                vybe_ast::ConstructorInitializerTarget::Base => {
                                    BaseCall::Explicit(args)
                                }
                            }
                        }
                        // Pascal: `inherited;` or `inherited Create;` is
                        // the explicit call. Walker today emits base_args
                        // = None when absent — mirror with Auto if there's
                        // a parent, None otherwise.
                        None => {
                            if parents.is_empty() || suppress_base_call {
                                BaseCall::None
                            } else {
                                BaseCall::Auto
                            }
                        }
                    },
                    // Pascal's constructor spelling (`Create`, `CreateFoo`,
                    // `Init`) remains source syntax for construction calls.
                    // The current shared constructor emitter treats named
                    // constructors as a different callable surface, so keep
                    // Pascal variants primary until constructor roles are
                    // represented as protocol slots.
                    named_name: None,
                });
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                implicit_self_member_names.insert(pname.to_ascii_lowercase());
                let (canonical, _) = crate::protocol::canonical_method(pname);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![],
                        None,
                        rewrite_property_getter_body(body, &instance_field_names),
                        Access::from(m.visibility.clone()),
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![s.param.clone()],
                        None,
                        rewrite_property_setter_body(&s.body, &instance_field_names),
                        Access::from(m.visibility.clone()),
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            // Object Pascal has single inheritance plus interfaces; class
            // helpers extend a type from outside and are the prototype-fallback
            // mechanism (§4d), not an augmentation of the declaration.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    extend_gcl_member_names(&mut implicit_self_member_names, parents);
    if !static_value_member_names.is_empty() {
        rewrite_static_value_members_in_methods(
            &mut out.instance_methods,
            name,
            &static_value_member_names,
        );
        rewrite_static_value_members_in_methods(
            &mut out.static_methods,
            name,
            &static_value_member_names,
        );
        rewrite_static_value_members_in_constructors(
            &mut out.constructors,
            name,
            &static_value_member_names,
        );
        if let Some(destructor) = out.destructor.as_mut() {
            rewrite_static_value_members_in_methods(
                std::slice::from_mut(destructor),
                name,
                &static_value_member_names,
            );
        }
    }
    rewrite_implicit_self_members_in_methods(
        &mut out.instance_methods,
        &implicit_self_member_names,
    );
    rewrite_implicit_self_members_in_constructors(
        &mut out.constructors,
        &implicit_self_member_names,
    );
    let gcl_accessor_property_names = gcl_accessor_property_names(parents);
    if !gcl_accessor_property_names.is_empty() {
        rewrite_gcl_property_accessors_in_methods(
            &mut out.instance_methods,
            &gcl_accessor_property_names,
        );
        rewrite_gcl_property_accessors_in_constructors(
            &mut out.constructors,
            &gcl_accessor_property_names,
        );
    }

    out.instance_methods = lower_pascal_method_overloads(out.instance_methods, &span);
    out.static_methods = lower_pascal_method_overloads(out.static_methods, &span);
    let constructor_calls_form_create = out.constructors.iter().any(|ctor| {
        ctor.body
            .iter()
            .any(|stmt| stmt_calls_method(stmt, "FormCreate"))
    });
    // The rewrites above (static-value members, implicit `Self.`, GCL
    // accessors) mutate `out.constructors` IN PLACE, so the view `push_
    // constructor` cloned at the walk site is the un-rewritten original.
    // Re-derive it under the same primary rule now that the list is final.
    out.resync_constructor_view();

    let is_gcl_form = parents
        .iter()
        .any(|parent| parent.eq_ignore_ascii_case("TForm"));
    let has_form_create = out.instance_methods.iter().any(|method| {
        method.source_name.eq_ignore_ascii_case("FormCreate")
            || method.canonical_name.eq_ignore_ascii_case("formcreate")
    });
    if is_gcl_form && has_form_create && !constructor_calls_form_create {
        let mut auto_init = build_normal_method(
            span.clone(),
            GCL_FORM_CREATE_AUTOINIT,
            GCL_FORM_CREATE_AUTOINIT,
            Vec::new(),
            None,
            call_self_form_create_body(),
            Access::Public,
            false,
            false,
            true,
            Modifiers::default(),
        );
        rewrite_implicit_self_members_in_body(
            &mut auto_init.body,
            &implicit_self_member_names,
            &mut HashSet::new(),
        );
        out.instance_methods.push(auto_init);
        out.auto_init_methods
            .push(GCL_FORM_CREATE_AUTOINIT.to_string());
    }

    // Pascal's implicit root: every class descends from TObject, so `is
    // TObject` must answer true. This is an ADDITION to the declared list —
    // the declared interfaces are filled centrally (see
    // `normalize_class_from_ast`), so this only states the Pascal rule.
    let mut normalized_interfaces = Vec::new();
    if !name.eq_ignore_ascii_case("TObject")
        && !interfaces
            .iter()
            .any(|iface| iface.eq_ignore_ascii_case("TObject"))
    {
        normalized_interfaces.push("TObject".to_string());
    }

    NormalClass {
        interfaces: normalized_interfaces,
        explicit_self_param: false, // Pascal: Self is implicit
        implicit_self_fields: true, // Pascal: bare field names resolve to Self.field inside methods
        ..Default::default()
    }
    .with_members(out)
}

#[derive(Clone)]
struct PascalOverloadCase {
    arity: usize,
    hidden_name: String,
}

fn lower_pascal_method_overloads(methods: Vec<NormalMethod>, span: &Span) -> Vec<NormalMethod> {
    let mut groups: HashMap<String, Vec<NormalMethod>> = HashMap::new();
    let mut order = Vec::new();

    for method in methods {
        if !groups.contains_key(&method.canonical_name) {
            order.push(method.canonical_name.clone());
        }
        groups
            .entry(method.canonical_name.clone())
            .or_default()
            .push(method);
    }

    let mut lowered = Vec::new();
    for key in order {
        let Some(group) = groups.remove(&key) else {
            continue;
        };
        if group.len() <= 1 || has_duplicate_arities(group.iter().map(|m| m.params.len())) {
            lowered.extend(group);
            continue;
        }

        let mut hidden_methods = Vec::new();
        let mut cases = Vec::new();
        let mut sorted = group;
        sorted.sort_by_key(|method| method.params.len());
        let wrapper_template = sorted.last().cloned().unwrap();

        for method in sorted {
            let hidden_name = format!(
                "__vybe_overload_{}_{}",
                method.canonical_name,
                method.params.len()
            );
            cases.push(PascalOverloadCase {
                arity: method.params.len(),
                hidden_name: hidden_name.clone(),
            });
            hidden_methods.push(build_normal_method(
                method.span.clone(),
                &hidden_name,
                &hidden_name,
                method.params.clone(),
                method.return_type.clone(),
                method.body.clone(),
                method.access,
                method.is_async,
                method.is_generator,
                method.is_sub,
                method.raw_modifiers.clone(),
            ));
        }

        lowered.extend(hidden_methods);
        lowered.push(build_normal_method(
            span.clone(),
            &wrapper_template.canonical_name,
            &wrapper_template.source_name,
            wrapper_template.params.clone(),
            wrapper_template.return_type.clone(),
            build_pascal_overload_dispatch(
                &cases,
                &wrapper_template.params,
                wrapper_template.return_type.is_none() && wrapper_template.is_sub,
            ),
            wrapper_template.access,
            false,
            false,
            wrapper_template.is_sub,
            Modifiers::default(),
        ));
    }

    lowered
}

fn build_pascal_overload_dispatch(
    cases: &[PascalOverloadCase],
    wrapper_params: &[vybe_ast::Param],
    is_sub: bool,
) -> Vec<Statement> {
    if cases.is_empty() {
        return Vec::new();
    }

    let first = &cases[0];
    let call_args: Vec<vybe_ast::Argument> = wrapper_params
        .iter()
        .take(first.arity)
        .map(|param| vybe_ast::Argument::positional(Expression::ident(&param.name)))
        .collect();
    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&first.hidden_name)),
        args: call_args,
        optional: false,
    });
    let invoke_stmt = if is_sub {
        Statement::new(StmtKind::Expr(call_expr))
    } else {
        Statement::new(StmtKind::Return(Some(call_expr)))
    };

    if cases.len() == 1 {
        return vec![invoke_stmt];
    }

    let gate_param = &wrapper_params[first.arity].name;
    let is_null = Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Eq,
        left: Box::new(Expression::ident(gate_param)),
        right: Box::new(Expression::null()),
    });
    let is_undefined = Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Eq,
        left: Box::new(Expression::ident(gate_param)),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
    });
    let cond = Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Or,
        left: Box::new(is_null),
        right: Box::new(is_undefined),
    });

    vec![Statement::new(StmtKind::If {
        cond,
        then_body: vec![invoke_stmt],
        elifs: Vec::new(),
        else_body: Some(build_pascal_overload_dispatch(
            &cases[1..],
            wrapper_params,
            is_sub,
        )),
    })]
}

fn has_duplicate_arities<I>(arities: I) -> bool
where
    I: IntoIterator<Item = usize>,
{
    let mut seen = std::collections::HashSet::new();
    for arity in arities {
        if !seen.insert(arity) {
            return true;
        }
    }
    false
}

fn is_pascal_no_base_ctor_marker(stmt: &Statement) -> bool {
    matches!(
        &stmt.kind,
        StmtKind::Expr(Expression {
            kind: ExprKind::Ident(name),
            ..
        }) if name == PASCAL_NO_BASE_CTOR_MARKER
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Modifiers;
    use vybe_ast::{ConstructorInitializerTarget, Visibility};

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    }

    fn make_method_with_visibility(src_name: &str, visibility: Visibility) -> ClassMember {
        let mut modifiers = Modifiers::default();
        modifiers.visibility = visibility;
        ClassMember::Method(Box::new(vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    }

    fn make_field_with_visibility(src_name: &str, visibility: Visibility) -> ClassMember {
        let mut modifiers = Modifiers::default();
        modifiers.visibility = visibility;
        ClassMember::Field {
            name: src_name.into(),
            type_hint: None,
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
        }
    }

    fn make_property_with_visibility(src_name: &str, visibility: Visibility) -> ClassMember {
        let mut modifiers = Modifiers::default();
        modifiers.visibility = visibility;
        ClassMember::Property {
            name: src_name.into(),
            type_hint: None,
            getter: Some(vec![Statement::new(StmtKind::Return(Some(
                Expression::ident("FValue"),
            )))]),
            setter: None,
            is_auto: false,
            modifiers,
        }
    }

    #[test]
    fn destroy_goes_to_destructor_case_insensitive() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("Destroy")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());

        // case variant
        let nc2 = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("destroy")],
            &ClassModifiers::default(),
        );
        assert!(nc2.destructor.is_some());
    }

    #[test]
    fn plain_add_is_not_protocol_add() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("Add")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert!(nc.special_methods.is_empty());
    }

    #[test]
    fn operator_add_maps_to_protocol_add() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("operator_Add")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.instance_methods[0].source_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
        assert_eq!(nc.special_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].source_name, "operator_Add");
    }

    #[test]
    fn hidden_operator_overload_keeps_callable_binding_and_role_slot() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("operator_Add__pascal_overload_0")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(
            nc.instance_methods[0].source_name,
            "operator_Add__pascal_overload_0"
        );
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
        assert_eq!(nc.special_methods[0].canonical_name, "add");
        assert_eq!(
            nc.special_methods[0].source_name,
            "operator_Add__pascal_overload_0"
        );
    }

    #[test]
    fn constructors_normalize_as_constructors_not_methods() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[ClassMember::Constructor {
                name: Some("Create".into()),
                params: vec![],
                body: vec![],
                base_args: None,
                initializer_target: ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            }],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.constructors.len(), 1);
        assert!(nc.instance_methods.is_empty());
        assert!(nc.special_methods.is_empty());
        assert!(matches!(nc.constructors[0].base_call, BaseCall::None));
    }

    #[test]
    fn member_visibility_normalizes_to_shared_access() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[
                make_field_with_visibility("Secret", Visibility::Private),
                make_method_with_visibility("Touch", Visibility::Protected),
                make_property_with_visibility("Value", Visibility::Public),
            ],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_fields[0].access, Access::Private);
        assert_eq!(nc.instance_methods[0].access, Access::Protected);
        assert_eq!(
            nc.properties[0].getter.as_ref().unwrap().access,
            Access::Public
        );
    }

    #[test]
    fn constructor_this_initializer_normalizes_to_this_base_call() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[ClassMember::Constructor {
                name: Some("Create".into()),
                params: vec![],
                body: vec![],
                base_args: Some(vec![Expression::new(ExprKind::Lit(Literal::Int(1)))]),
                initializer_target: ConstructorInitializerTarget::This,
                visibility: Visibility::Public,
            }],
            &ClassModifiers::default(),
        );
        assert!(matches!(nc.constructors[0].base_call, BaseCall::This(_)));
    }
}
