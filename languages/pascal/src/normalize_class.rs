//! Pascal `ClassDecl` → `NormalClass` walker pass.
//!
//! Pascal / Delphi / Free Pascal class specifics:
//!   - `constructor Create;` / `constructor Init;` → NormalConstructor.
//!     Pascal's convention is `Create`; Free Pascal also allows `Init`.
//!   - `destructor Destroy;` → destructor.
//!   - `property Foo read GetFoo write SetFoo` → NormalProperty. Walker
//!     already links property accessors to their accessor methods.
//!   - `class operator Add(...)` / `class operator Equal(...)` →
//!     SpecialMethodKind::Add / Eq. Pascal operator overloads arrive
//!     with names like "Add" / "Subtract" / "Multiply" / "Divide" /
//!     "Equal" per Delphi convention.
//!   - `override` / `virtual` / `reintroduce` → flag carries through.
//!   - Case-insensitive: Pascal method names lowercase to canonical.

use std::collections::{HashMap, HashSet};
use vybe_ast::{
    Argument, CaseCondition, ClassMember, ClassModifiers, ExprKind, Expression, Literal, Modifiers,
    PropertySetter, Span, Statement, StmtKind,
};
use vybe_plugin::class_normalize::{
    build_normal_method,
    canonical::{ClassLang, canonicalize_method},
    from_method_stmt,
    types::*,
};

fn property_field_name(body: &[Statement], field_names: &HashSet<String>) -> Option<String> {
    let [stmt] = body else {
        return None;
    };
    match &stmt.kind {
        StmtKind::Return(Some(expr)) => match &expr.kind {
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
            if !matches!(&callee.kind, ExprKind::Ident(_)) {
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
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut raw_extra_members: Vec<ClassMember> = Vec::new();
    let mut instance_fields: Vec<NormalField> = Vec::new();
    let mut static_fields: Vec<NormalField> = Vec::new();
    let mut instance_methods: Vec<NormalMethod> = Vec::new();
    let mut static_methods: Vec<NormalMethod> = Vec::new();
    let mut properties: Vec<NormalProperty> = Vec::new();
    let mut constructors: Vec<NormalConstructor> = Vec::new();
    let mut destructor: Option<NormalMethod> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();
    let field_names: HashSet<String> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field { name, .. } => Some(name.to_ascii_lowercase()),
            _ => None,
        })
        .collect();
    let mut implicit_self_member_names = field_names.clone();

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
                    access: Access::Public,
                    readonly: m.is_readonly,
                };
                if m.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
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

                // Pascal destructor: `destructor Destroy;`. Case-insensitive.
                if src_name.eq_ignore_ascii_case("Destroy") {
                    if let Some(d) = from_method_stmt(span.clone(), stmt, "destroy", Access::Public)
                    {
                        destructor = Some(d);
                    }
                    continue;
                }

                let (canonical, special_kind) = canonicalize_method(ClassLang::Pascal, src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public)
                else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                if m.is_static {
                    static_methods.push(method);
                } else {
                    implicit_self_member_names.insert(src_name.to_ascii_lowercase());
                    instance_methods.push(method);
                }
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                ..
            } => {
                constructors.push(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect(),
                        ),
                        // Pascal: `inherited;` or `inherited Create;` is
                        // the explicit call. Walker today emits base_args
                        // = None when absent — mirror with Auto if there's
                        // a parent, None otherwise.
                        None => {
                            if parents.is_empty() {
                                BaseCall::None
                            } else {
                                BaseCall::Auto
                            }
                        }
                    },
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
                let (canonical, _) = canonicalize_method(ClassLang::Pascal, pname);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        Vec::new(),
                        vec![],
                        None,
                        rewrite_property_getter_body(body, &field_names),
                        Access::Public,
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
                        Vec::new(),
                        vec![s.param.clone()],
                        None,
                        rewrite_property_setter_body(&s.body, &field_names),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                raw_extra_members.push(other.clone());
            }
        }
    }

    extend_gcl_member_names(&mut implicit_self_member_names, parents);
    rewrite_implicit_self_members_in_methods(&mut instance_methods, &implicit_self_member_names);
    rewrite_implicit_self_members_in_constructors(&mut constructors, &implicit_self_member_names);
    let gcl_accessor_property_names = gcl_accessor_property_names(parents);
    if !gcl_accessor_property_names.is_empty() {
        rewrite_gcl_property_accessors_in_methods(
            &mut instance_methods,
            &gcl_accessor_property_names,
        );
        rewrite_gcl_property_accessors_in_constructors(
            &mut constructors,
            &gcl_accessor_property_names,
        );
    }

    instance_methods = lower_pascal_method_overloads(instance_methods, &span);
    static_methods = lower_pascal_method_overloads(static_methods, &span);
    let constructor_calls_form_create = constructors.iter().any(|ctor| {
        ctor.body
            .iter()
            .any(|stmt| stmt_calls_method(stmt, "FormCreate"))
    });
    let (ctor_helper_methods, constructor) =
        lower_pascal_constructor_overloads(constructors, &span);
    instance_methods.extend(ctor_helper_methods);
    let mut auto_init_methods = Vec::new();

    let is_gcl_form = parents
        .iter()
        .any(|parent| parent.eq_ignore_ascii_case("TForm"));
    let has_form_create = instance_methods.iter().any(|method| {
        method.source_name.eq_ignore_ascii_case("FormCreate")
            || method.canonical_name.eq_ignore_ascii_case("formcreate")
    });
    if is_gcl_form && has_form_create && !constructor_calls_form_create {
        let mut auto_init = build_normal_method(
            span.clone(),
            GCL_FORM_CREATE_AUTOINIT,
            GCL_FORM_CREATE_AUTOINIT,
            Vec::new(),
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
        instance_methods.push(auto_init);
        auto_init_methods.push(GCL_FORM_CREATE_AUTOINIT.to_string());
    }

    if let Some(destructor_method) = destructor.clone() {
        instance_methods.push(destructor_method);

        let has_free = instance_methods.iter().any(|method| {
            method.source_name.eq_ignore_ascii_case("Free")
                || method.canonical_name.eq_ignore_ascii_case("free")
        });
        if !has_free {
            let free_body = vec![Statement::new(StmtKind::Expr(Expression::new(
                ExprKind::Call {
                    callee: Box::new(Expression::ident("Destroy")),
                    args: Vec::new(),
                    optional: false,
                },
            )))];
            instance_methods.push(build_normal_method(
                span.clone(),
                "free",
                "Free",
                Vec::new(),
                Vec::new(),
                None,
                free_body,
                Access::Public,
                false,
                false,
                true,
                Modifiers::default(),
            ));
        }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        bases: Vec::new(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: false,
        is_value_type: false,
        explicit_self_param: false, // Pascal: Self is implicit
        implicit_self_fields: true, // Pascal: bare field names resolve to Self.field inside methods
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors: Vec::new(),
        constructor,
        destructor,
        auto_init_methods,
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
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
                Vec::new(),
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
            wrapper_template.aliases.clone(),
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

fn lower_pascal_constructor_overloads(
    constructors: Vec<NormalConstructor>,
    span: &Span,
) -> (Vec<NormalMethod>, Option<NormalConstructor>) {
    if constructors.is_empty() {
        return (Vec::new(), None);
    }
    if constructors.len() == 1
        || has_duplicate_arities(constructors.iter().map(|ctor| ctor.params.len()))
    {
        return (Vec::new(), constructors.into_iter().last());
    }

    let mut sorted = constructors;
    sorted.sort_by_key(|ctor| ctor.params.len());
    let wrapper_template = sorted.last().cloned().unwrap();
    let mut helper_methods = Vec::new();
    let mut cases = Vec::new();

    for ctor in sorted {
        let hidden_name = format!("__vybe_ctor_create_{}", ctor.params.len());
        cases.push(PascalOverloadCase {
            arity: ctor.params.len(),
            hidden_name: hidden_name.clone(),
        });
        helper_methods.push(build_normal_method(
            ctor.span.clone(),
            &hidden_name,
            &hidden_name,
            Vec::new(),
            ctor.params.clone(),
            None,
            ctor.body.clone(),
            Access::Public,
            false,
            false,
            true,
            Modifiers::default(),
        ));
    }

    let wrapper = NormalConstructor {
        span: span.clone(),
        params: wrapper_template.params.clone(),
        body: build_pascal_overload_dispatch(&cases, &wrapper_template.params, true),
        base_call: wrapper_template.base_call,
        named_name: wrapper_template.named_name,
    };

    (helper_methods, Some(wrapper))
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
    let cond = Expression::new(ExprKind::Binary {
        op: vybe_ast::BinOp::Eq,
        left: Box::new(Expression::ident(gate_param)),
        right: Box::new(Expression::null()),
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

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Modifiers;

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
    fn add_operator_maps_to_canonical_add() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("Add")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }
}
