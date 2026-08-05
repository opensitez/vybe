use super::{JsParser, Rule};
use pest::Parser;
use pest::iterators::Pair;
use std::collections::HashMap;
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_ast::*;

// Monotonically increasing counter — unique template object slot per call site.
static TEMPLATE_COUNTER: AtomicUsize = AtomicUsize::new(0);

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs =
        JsParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    // pest wraps everything in the `program` rule — unwrap it
    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                body.push(walk_statement(top)?);
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE => continue,
                Rule::import_statement => imports.push(walk_import(pair)?),
                _ => body.push(walk_statement(pair)?) }
        }
    }
    // TC39 explicit-resource-management classes: when the module references
    // DisposableStack / AsyncDisposableStack, prepend the canonical JS
    // implementations so they compile through the common class pipeline —
    // same pattern as Pascal's synthetic Exception class.
    if source.contains("DisposableStack") {
        let mut injected = parse_runtime_class_snippet(DISPOSABLE_STACK_JS);
        if source.contains("AsyncDisposableStack") {
            injected.extend(parse_runtime_class_snippet(ASYNC_DISPOSABLE_STACK_JS));
        }
        injected.extend(body);
        body = injected;
    }

    // JS function hoisting: function declarations are visible before their
    // textual position. Reorder so they come first in the body. This mirrors
    // what the JS engine does at parse time — function decls are hoisted to
    // the top of their enclosing scope.
    let mut hoisted = Vec::new();
    let mut rest = Vec::new();
    for stmt in body {
        if matches!(stmt.kind, StmtKind::FunctionDecl { .. }) {
            hoisted.push(stmt);
        } else {
            rest.push(stmt);
        }
    }
    hoisted.append(&mut rest);
    let mut body = hoisted;

    // Const-folding pass for computed method/property names that
    // reference a top-level string constant: `const X = "greet"` makes
    // `class C { [X]() {…} }` and `{ [X]() {…} }` resolvable to method
    // name "greet" at compile time. Without this fold the method ends
    // up bound under the literal text "X" and `obj.greet()` misses.
    //
    // Pure walker work — no compiler state, no AST extension. The fold
    // only fires when the computed key is a single identifier whose
    // value is a string literal in scope; anything more complex falls
    // through to the existing literal-text path (still incorrect for
    // those cases, but those tests already need runtime install).
    fold_const_computed_names(&mut body);

    // ES2026 explicit resource management — desugar `using` / `await using`
    // markers into the spec's try/finally shape before any other pass.
    lower_using_declarations(&mut body);

    // Static TDZ pass — ECMA-262 §14.3.1 / §14.6: a `let` / `const` / `class`
    // binding is in a temporal dead zone until its declaration executes.
    apply_static_tdz(&mut body);
    validate_private_class_syntax(&body)?;

    for stmt in &mut body {
        stmt.walk_exprs_mut(&mut normalize_promise_statics);
    }

    Ok(Module {
        name: "main".into(),
        language: Lang::JavaScript,
        body,
        imports })
}

fn validate_private_class_syntax(body: &[Statement]) -> Result<(), String> {
    for stmt in body {
        validate_private_stmt(stmt)?;
    }
    Ok(())
}

fn validate_private_stmt(stmt: &Statement) -> Result<(), String> {
    match &stmt.kind {
        StmtKind::Expr(expr) => validate_private_expr(expr),
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. } => validate_private_class_syntax(body),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => validate_private_class_members(members),
        StmtKind::EnumDecl { body_members, .. } => validate_private_class_members(body_members),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    validate_private_expr(init)?;
                }
            }
            Ok(())
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body } => {
            validate_private_expr(cond)?;
            validate_private_class_syntax(then_body)?;
            for (cond, body) in elifs {
                validate_private_expr(cond)?;
                validate_private_class_syntax(body)?;
            }
            if let Some(body) = else_body {
                validate_private_class_syntax(body)?;
            }
            Ok(())
        }
        StmtKind::For {
            init,
            cond,
            update,
            body } => {
            if let Some(init) = init {
                validate_private_stmt(init)?;
            }
            if let Some(cond) = cond {
                validate_private_expr(cond)?;
            }
            if let Some(update) = update {
                validate_private_expr(update)?;
            }
            validate_private_class_syntax(body)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            validate_private_expr(iter)?;
            validate_private_class_syntax(body)?;
            if let Some(body) = else_body {
                validate_private_class_syntax(body)?;
            }
            Ok(())
        }
        StmtKind::While {
            cond,
            body,
            else_body } => {
            validate_private_expr(cond)?;
            validate_private_class_syntax(body)?;
            if let Some(body) = else_body {
                validate_private_class_syntax(body)?;
            }
            Ok(())
        }
        StmtKind::DoWhile { body, cond, .. } => {
            validate_private_class_syntax(body)?;
            validate_private_expr(cond)
        }
        StmtKind::Switch {
            expr,
            cases,
            default } => {
            validate_private_expr(expr)?;
            for case in cases {
                for cond in &case.conditions {
                    validate_private_case_condition(cond)?;
                }
                validate_private_class_syntax(&case.body)?;
            }
            if let Some(body) = default {
                validate_private_class_syntax(body)?;
            }
            Ok(())
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally } => {
            validate_private_class_syntax(body)?;
            for catch in catches {
                validate_private_class_syntax(&catch.body)?;
            }
            if let Some(body) = else_body {
                validate_private_class_syntax(body)?;
            }
            if let Some(body) = finally {
                validate_private_class_syntax(body)?;
            }
            Ok(())
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                validate_private_expr(&item.expr)?;
            }
            validate_private_class_syntax(body)
        }
        StmtKind::Using { resource, body, .. } => {
            validate_private_expr(resource)?;
            validate_private_class_syntax(body)
        }
        StmtKind::Lock { expr, body } => {
            validate_private_expr(expr)?;
            validate_private_class_syntax(body)
        }
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                validate_private_expr(expr)?;
            }
            Ok(())
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                validate_private_expr(expr)?;
            }
            if let Some(cause) = cause {
                validate_private_expr(cause)?;
            }
            Ok(())
        }
        StmtKind::Assign { targets, value, .. } => {
            for target in targets {
                validate_private_expr(target)?;
            }
            validate_private_expr(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            validate_private_expr(target)?;
            validate_private_expr(value)
        }
        StmtKind::ReDim { bounds, .. } => {
            for bound in bounds {
                validate_private_expr(bound)?;
            }
            Ok(())
        }
        StmtKind::AddHandler {
            control, handler, ..
        } => {
            validate_private_expr(control)?;
            validate_private_expr(handler)
        }
        _ => Ok(()) }
}

fn validate_private_class_members(members: &[ClassMember]) -> Result<(), String> {
    let mut instance_private = HashMap::new();
    let mut static_private = HashMap::new();
    for member in members {
        let (name, is_static, private_kind) = match member {
            ClassMember::Field {
                name,
                modifiers,
                init,
                ..
            } => {
                if let Some(init) = init {
                    validate_private_expr(init)?;
                }
                (name.as_str(), modifiers.is_static, 0b001)
            }
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    body,
                    modifiers,
                    ..
                } = &stmt.kind
                {
                    validate_private_class_syntax(body)?;
                    (name.as_str(), modifiers.is_static, 0b001)
                } else {
                    continue;
                }
            }
            ClassMember::Property {
                name,
                getter,
                setter,
                modifiers,
                ..
            } => {
                if let Some(body) = getter {
                    validate_private_class_syntax(body)?;
                }
                if let Some(setter) = setter {
                    validate_private_class_syntax(&setter.body)?;
                }
                let private_kind = match (getter.is_some(), setter.is_some()) {
                    (true, false) => 0b010,
                    (false, true) => 0b100,
                    (true, true) => 0b110,
                    (false, false) => 0 };
                (name.as_str(), modifiers.is_static, private_kind)
            }
            ClassMember::Constructor { body, .. } => {
                validate_private_class_syntax(body)?;
                continue;
            }
            ClassMember::Const { value, .. } => {
                validate_private_expr(value)?;
                continue;
            }
            ClassMember::NestedType(stmt) => {
                validate_private_stmt(stmt)?;
                continue;
            }
            ClassMember::Event { .. } | ClassMember::Augment(_) => continue };
        if name.starts_with('#') {
            let seen: &mut HashMap<String, u8> = if is_static {
                &mut static_private
            } else {
                &mut instance_private
            };
            let existing = seen.get(name).copied().unwrap_or(0);
            if existing & private_kind != 0
                || existing & 0b001 != 0
                || private_kind & 0b001 != 0 && existing != 0
            {
                return Err(format!("Duplicate private member {}", name));
            }
            seen.insert(name.to_string(), existing | private_kind);
        }
    }
    Ok(())
}

fn validate_private_expr(expr: &Expression) -> Result<(), String> {
    match &expr.kind {
        ExprKind::Async(op) => {
            for child in op.children() {
                validate_private_expr(child)?;
            }
            Ok(())
        }
        ExprKind::Chan(op) => {
            for child in op.children() {
                validate_private_expr(child)?;
            }
            Ok(())
        }
        ExprKind::Delete(inner) => {
            if matches!(
                &inner.kind,
                ExprKind::Member { field, .. } if field.starts_with('#')
            ) {
                return Err("Cannot delete private member".to_string());
            }
            validate_private_expr(inner)
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Range {
            start: left,
            end: right,
            ..
        } => {
            validate_private_expr(left)?;
            validate_private_expr(right)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. } => validate_private_expr(expr),
        ExprKind::RefOf(place) => validate_private_place(place),
        ExprKind::Ternary { cond, then, else_ } => {
            validate_private_expr(cond)?;
            validate_private_expr(then)?;
            validate_private_expr(else_)
        }
        ExprKind::Member { object, .. } => validate_private_expr(object),
        ExprKind::Index { object, index, .. } => {
            validate_private_expr(object)?;
            validate_private_expr(index)
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args } => {
            validate_private_expr(callee)?;
            for arg in args {
                validate_private_expr(&arg.value)?;
            }
            Ok(())
        }
        ExprKind::Assign { target, value } => {
            validate_private_expr(target)?;
            validate_private_expr(value)
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => validate_private_expr(expr),
            LambdaBody::Block(body) => validate_private_class_syntax(body) },
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &item.key {
                    validate_private_expr(key)?;
                }
                validate_private_expr(&item.value)?;
            }
            Ok(())
        }
        ExprKind::Tuple(items)
        | ExprKind::Set(items)
        | ExprKind::Sequence(items)
        | ExprKind::Zip {
            iterables: items, ..
        } => {
            for item in items {
                validate_private_expr(item)?;
            }
            Ok(())
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, expr) in fields {
                validate_private_expr(expr)?;
            }
            Ok(())
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        validate_private_expr(key)?;
                        validate_private_expr(value)?;
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => validate_private_stmt(value)?,
                    ObjectProperty::Spread(expr) => validate_private_expr(expr)?,
                    ObjectProperty::Shorthand(_) => {}
                }
            }
            Ok(())
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) = part {
                    validate_private_expr(expr)?;
                }
            }
            Ok(())
        }
        ExprKind::Yield(Some(expr)) => validate_private_expr(expr),
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            validate_private_expr(element)?;
            for generator in generators {
                validate_private_expr(&generator.target)?;
                validate_private_expr(&generator.iter)?;
                for cond in &generator.conditions {
                    validate_private_expr(cond)?;
                }
            }
            Ok(())
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(expr) = lower {
                validate_private_expr(expr)?;
            }
            if let Some(expr) = upper {
                validate_private_expr(expr)?;
            }
            if let Some(expr) = step {
                validate_private_expr(expr)?;
            }
            Ok(())
        }
        ExprKind::Walrus { target, value } => {
            validate_private_expr(target)?;
            validate_private_expr(value)
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                validate_private_expr(parent)?;
            }
            validate_private_class_members(members)
        }
        ExprKind::FunctionExpr(stmt) => validate_private_stmt(stmt),
        ExprKind::StaticAccess { class, member } => {
            validate_private_expr(class)?;
            validate_private_expr(member)
        }
        ExprKind::Match { subject, arms } => {
            validate_private_expr(subject)?;
            for arm in arms {
                if let Some(conditions) = &arm.conditions {
                    for cond in conditions {
                        validate_private_expr(cond)?;
                    }
                }
                validate_private_expr(&arm.body)?;
            }
            Ok(())
        }
        ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::DefaultOf(_)
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::Yield(None)
        | ExprKind::Destructure(_) => Ok(()) }
}

fn validate_private_case_condition(cond: &CaseCondition) -> Result<(), String> {
    match cond {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            validate_private_expr(expr)
        }
        CaseCondition::Range { from, to } => {
            validate_private_expr(from)?;
            validate_private_expr(to)
        }
    }
}

fn stmt_contains_await(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Expr(expr) => expr_contains_await(expr),
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. } => body.iter().any(stmt_contains_await),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => members.iter().any(class_member_contains_await),
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|decl| decl.init.as_ref().is_some_and(expr_contains_await)),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body } => {
            expr_contains_await(cond)
                || then_body.iter().any(stmt_contains_await)
                || elifs.iter().any(|(cond, body)| {
                    expr_contains_await(cond) || body.iter().any(stmt_contains_await)
                })
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
        }
        StmtKind::For {
            init,
            cond,
            update,
            body } => {
            init.as_ref().is_some_and(|stmt| stmt_contains_await(stmt))
                || cond.as_ref().is_some_and(expr_contains_await)
                || update.as_ref().is_some_and(expr_contains_await)
                || body.iter().any(stmt_contains_await)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            expr_contains_await(iter)
                || body.iter().any(stmt_contains_await)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
        }
        StmtKind::While {
            cond,
            body,
            else_body } => {
            expr_contains_await(cond)
                || body.iter().any(stmt_contains_await)
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
        }
        StmtKind::DoWhile { body, cond, .. } => {
            body.iter().any(stmt_contains_await) || expr_contains_await(cond)
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally } => {
            body.iter().any(stmt_contains_await)
                || catches
                    .iter()
                    .any(|catch| catch.body.iter().any(stmt_contains_await))
                || else_body
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
                || finally
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
        }
        StmtKind::Return(expr) => expr.as_ref().is_some_and(expr_contains_await),
        StmtKind::Throw { expr, cause } => {
            expr.as_ref().is_some_and(expr_contains_await)
                || cause.as_ref().is_some_and(expr_contains_await)
        }
        StmtKind::Assign { targets, value, .. } => {
            targets.iter().any(expr_contains_await) || expr_contains_await(value)
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            expr_contains_await(target) || expr_contains_await(value)
        }
        StmtKind::Switch {
            expr,
            cases,
            default } => {
            expr_contains_await(expr)
                || cases.iter().any(|case| {
                    case.conditions.iter().any(case_condition_contains_await)
                        || case.body.iter().any(stmt_contains_await)
                })
                || default
                    .as_ref()
                    .is_some_and(|body| body.iter().any(stmt_contains_await))
        }
        StmtKind::With { items, body, .. } => {
            items.iter().any(|item| expr_contains_await(&item.expr))
                || body.iter().any(stmt_contains_await)
        }
        StmtKind::Using { resource, body, .. } => {
            expr_contains_await(resource) || body.iter().any(stmt_contains_await)
        }
        StmtKind::Lock { expr, body } => {
            expr_contains_await(expr) || body.iter().any(stmt_contains_await)
        }
        _ => false }
}

fn class_member_contains_await(member: &ClassMember) -> bool {
    match member {
        ClassMember::Field { init, .. } => init.as_ref().is_some_and(expr_contains_await),
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => stmt_contains_await(stmt),
        ClassMember::Constructor { body, .. } => body.iter().any(stmt_contains_await),
        ClassMember::Property { getter, setter, .. } => {
            getter
                .as_ref()
                .is_some_and(|body| body.iter().any(stmt_contains_await))
                || setter
                    .as_ref()
                    .is_some_and(|setter| setter.body.iter().any(stmt_contains_await))
        }
        ClassMember::Const { value, .. } => expr_contains_await(value),
        ClassMember::Event { .. } | ClassMember::Augment(_) => false }
}

fn expr_contains_await(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Async(op) => op.children().into_iter().any(expr_contains_await),
        ExprKind::Chan(op) => op.children().into_iter().any(expr_contains_await),
        ExprKind::Await(_) => true,
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Range {
            start: left,
            end: right,
            ..
        } => expr_contains_await(left) || expr_contains_await(right),
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. } => expr_contains_await(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_await(cond) || expr_contains_await(then) || expr_contains_await(else_)
        }
        ExprKind::Member { object, .. } => expr_contains_await(object),
        ExprKind::Index { object, index, .. } => {
            expr_contains_await(object) || expr_contains_await(index)
        }
        ExprKind::Call { callee, args, .. }
        | ExprKind::New {
            class: callee,
            args } => expr_contains_await(callee) || args.iter().any(|arg| expr_contains_await(&arg.value)),
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            expr_contains_await(target) || expr_contains_await(value)
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => expr_contains_await(expr),
            LambdaBody::Block(body) => body.iter().any(stmt_contains_await) },
        ExprKind::Array(items) => items.iter().any(|item| {
            item.key.as_ref().is_some_and(expr_contains_await) || expr_contains_await(&item.value)
        }),
        ExprKind::Tuple(items)
        | ExprKind::Set(items)
        | ExprKind::Sequence(items)
        | ExprKind::Zip {
            iterables: items, ..
        } => {
            items.iter().any(expr_contains_await)
        }
        ExprKind::NamedTuple { fields, .. } => {
            fields.iter().any(|(_, expr)| expr_contains_await(expr))
        }
        ExprKind::Object(props) => props.iter().any(|prop| match prop {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                expr_contains_await(key) || expr_contains_await(value)
            }
            ObjectProperty::Spread(expr) => expr_contains_await(expr),
            ObjectProperty::Method { value, .. } | ObjectProperty::Accessor { value, .. } => {
                stmt_contains_await(value)
            }
            ObjectProperty::Shorthand(_) => false }),
        ExprKind::Interpolation(parts) => parts.iter().any(|part| match part {
            InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                expr_contains_await(expr)
            }
            InterpolPart::Text(_) => false }),
        ExprKind::Yield(Some(expr)) => expr_contains_await(expr),
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            expr_contains_await(element)
                || generators.iter().any(|generator| {
                    expr_contains_await(&generator.target)
                        || expr_contains_await(&generator.iter)
                        || generator.conditions.iter().any(expr_contains_await)
                })
        }
        ExprKind::Slice { lower, upper, step } => {
            lower.as_ref().is_some_and(|expr| expr_contains_await(expr))
                || upper.as_ref().is_some_and(|expr| expr_contains_await(expr))
                || step.as_ref().is_some_and(|expr| expr_contains_await(expr))
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            parent
                .as_ref()
                .is_some_and(|expr| expr_contains_await(expr))
                || members.iter().any(class_member_contains_await)
        }
        ExprKind::FunctionExpr(stmt) => stmt_contains_await(stmt),
        ExprKind::StaticAccess { class, member } => {
            expr_contains_await(class) || expr_contains_await(member)
        }
        ExprKind::Match { subject, arms } => {
            expr_contains_await(subject)
                || arms.iter().any(|arm| {
                    arm.conditions
                        .as_ref()
                        .is_some_and(|conditions| conditions.iter().any(expr_contains_await))
                        || expr_contains_await(&arm.body)
                })
        }
        ExprKind::RefOf(place) => place_contains_await(place),
        ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::DefaultOf(_)
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::Yield(None)
        | ExprKind::Destructure(_) => false }
}

fn case_condition_contains_await(cond: &CaseCondition) -> bool {
    match cond {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            expr_contains_await(expr)
        }
        CaseCondition::Range { from, to } => expr_contains_await(from) || expr_contains_await(to) }
}

fn place_contains_await(place: &PlaceExpr) -> bool {
    match place {
        PlaceExpr::Member { object, .. } => expr_contains_await(object),
        PlaceExpr::Index { object, index, .. } => {
            expr_contains_await(object) || expr_contains_await(index)
        }
        PlaceExpr::Deref(expr) => expr_contains_await(expr),
        PlaceExpr::Ident(_) => false }
}

fn validate_private_place(place: &PlaceExpr) -> Result<(), String> {
    match place {
        PlaceExpr::Member { object, .. } => validate_private_expr(object),
        PlaceExpr::Index { object, index, .. } => {
            validate_private_expr(object)?;
            validate_private_expr(index)
        }
        PlaceExpr::Deref(expr) => validate_private_expr(expr),
        PlaceExpr::Ident(_) => Ok(()) }
}

// ── Static TDZ pass — ECMA-262 §14.3.1 / §14.6 ──────────────────────────────
//
// A reference at a textually earlier statement in the SAME statement list is
// guaranteed to evaluate inside the binding's temporal dead zone, so the
// walker rewrites the reference itself into JS-shape AST that throws at
// evaluation time:
//
//   (() => { throw new ReferenceError("Cannot access 'x' before initialization") })()
//
// Replacing the read (not the statement) keeps evaluation order faithful:
// short-circuited or never-reached reads still don't throw. Deferred-execution
// bodies (Lambda / FunctionExpr / FunctionDecl / class bodies) are skipped — a
// closure may legally run after initialization. Nested lists that re-declare
// the name shadow it and stop the descent.

fn apply_static_tdz(stmts: &mut Vec<Statement>) {
    let mut decls: Vec<(usize, String)> = Vec::new();
    for (i, s) in stmts.iter().enumerate() {
        match &s.kind {
            StmtKind::VarDecl { declarations, kind }
                if matches!(kind, VarDeclKind::Let | VarDeclKind::Const) =>
            {
                for d in declarations {
                    if let BindingPattern::Ident(n) = &d.pattern {
                        decls.push((i, n.clone()));
                    }
                }
            }
            StmtKind::ClassDecl { name, .. } => decls.push((i, name.clone())),
            _ => {}
        }
    }
    for (decl_idx, name) in &decls {
        for s in stmts[..*decl_idx].iter_mut() {
            tdz_rewrite_stmt(s, name);
        }
    }
    // Each nested statement list is its own (block) scope with its own TDZ.
    for s in stmts.iter_mut() {
        visit_nested_stmt_lists(&mut s.kind, apply_static_tdz);
    }
}

/// Apply `f` to every statement list directly nested in this statement —
/// shared traversal for the walker's list-level passes (TDZ, `using`).
fn visit_nested_stmt_lists(kind: &mut StmtKind, f: fn(&mut Vec<Statement>)) {
    match kind {
        StmtKind::Block(b) | StmtKind::FunctionDecl { body: b, .. } => f(b),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            f(then_body);
            for (_, b) in elifs {
                f(b);
            }
            if let Some(b) = else_body {
                f(b);
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(i) = init {
                visit_nested_stmt_lists(&mut i.kind, f);
            }
            f(body);
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            f(body);
            if let Some(b) = else_body {
                f(b);
            }
        }
        StmtKind::DoWhile { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => f(body),
        StmtKind::Switch { cases, default, .. } => {
            for c in cases {
                f(&mut c.body);
            }
            if let Some(d) = default {
                f(d);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally } => {
            f(body);
            for c in catches {
                f(&mut c.body);
            }
            if let Some(b) = else_body {
                f(b);
            }
            if let Some(b) = finally {
                f(b);
            }
        }
        StmtKind::ClassDecl { members, .. } => {
            for m in members {
                match m {
                    ClassMember::Method(s) => visit_nested_stmt_lists(&mut s.kind, f),
                    ClassMember::Constructor { body, .. } => f(body),
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

/// Does this statement list declare `name` (any binding form)? Used as the
/// shadow check when the TDZ rewrite descends into a nested block scope.
fn tdz_list_declares(stmts: &[Statement], name: &str) -> bool {
    stmts.iter().any(|s| match &s.kind {
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .any(|d| matches!(&d.pattern, BindingPattern::Ident(n) if n == name)),
        StmtKind::FunctionDecl { name: n, .. } | StmtKind::ClassDecl { name: n, .. } => n == name,
        _ => false })
}

fn tdz_rewrite_list(stmts: &mut [Statement], name: &str) {
    if tdz_list_declares(stmts, name) {
        return;
    }
    for s in stmts {
        tdz_rewrite_stmt(s, name);
    }
}

fn tdz_rewrite_stmt(stmt: &mut Statement, name: &str) {
    match &mut stmt.kind {
        StmtKind::Expr(e) => tdz_rewrite_expr(e, name),
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(init) = &mut d.init {
                    tdz_rewrite_expr(init, name);
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            tdz_rewrite_expr(value, name);
            for t in targets {
                tdz_rewrite_place(t, name);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            tdz_rewrite_expr(value, name);
            tdz_rewrite_place(target, name);
        }
        StmtKind::Return(Some(e)) => tdz_rewrite_expr(e, name),
        StmtKind::Throw { expr, cause } => {
            if let Some(e) = expr {
                tdz_rewrite_expr(e, name);
            }
            if let Some(e) = cause {
                tdz_rewrite_expr(e, name);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body } => {
            tdz_rewrite_expr(cond, name);
            tdz_rewrite_list(then_body, name);
            for (c, b) in elifs {
                tdz_rewrite_expr(c, name);
                tdz_rewrite_list(b, name);
            }
            if let Some(b) = else_body {
                tdz_rewrite_list(b, name);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body } => {
            tdz_rewrite_expr(cond, name);
            tdz_rewrite_list(body, name);
            if let Some(b) = else_body {
                tdz_rewrite_list(b, name);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            tdz_rewrite_list(body, name);
            tdz_rewrite_expr(cond, name);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body } => {
            // `for (let name = …)` shadows the outer binding for the loop.
            let shadowed = init.as_ref().is_some_and(|i| {
                matches!(&i.kind, StmtKind::VarDecl { declarations, .. }
                    if declarations.iter().any(|d| matches!(&d.pattern, BindingPattern::Ident(n) if n == name)))
            });
            if let Some(i) = init {
                tdz_rewrite_stmt(i, name);
            }
            if !shadowed {
                if let Some(c) = cond {
                    tdz_rewrite_expr(c, name);
                }
                if let Some(u) = update {
                    tdz_rewrite_expr(u, name);
                }
                tdz_rewrite_list(body, name);
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
            tdz_rewrite_expr(iter, name);
            let shadowed = var == name || key.as_deref() == Some(name);
            if !shadowed {
                tdz_rewrite_list(body, name);
                if let Some(b) = else_body {
                    tdz_rewrite_list(b, name);
                }
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default } => {
            tdz_rewrite_expr(expr, name);
            // All switch arms share one block scope (§14.12).
            let shadowed = cases.iter().any(|c| tdz_list_declares(&c.body, name))
                || default.as_ref().is_some_and(|d| tdz_list_declares(d, name));
            if !shadowed {
                for c in cases {
                    for cond in &mut c.conditions {
                        match cond {
                            CaseCondition::Value(e) => tdz_rewrite_expr(e, name),
                            CaseCondition::Range { from, to } => {
                                tdz_rewrite_expr(from, name);
                                tdz_rewrite_expr(to, name);
                            }
                            CaseCondition::Comparison { expr, .. } => tdz_rewrite_expr(expr, name) }
                    }
                    for s in &mut c.body {
                        tdz_rewrite_stmt(s, name);
                    }
                }
                if let Some(d) = default {
                    for s in d {
                        tdz_rewrite_stmt(s, name);
                    }
                }
            }
        }
        StmtKind::Block(b) => tdz_rewrite_list(b, name),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally } => {
            tdz_rewrite_list(body, name);
            for c in catches {
                if c.var_name.as_deref() != Some(name) && c.stack_var.as_deref() != Some(name) {
                    if let Some(w) = &mut c.when_clause {
                        tdz_rewrite_expr(w, name);
                    }
                    tdz_rewrite_list(&mut c.body, name);
                }
            }
            if let Some(b) = else_body {
                tdz_rewrite_list(b, name);
            }
            if let Some(b) = finally {
                tdz_rewrite_list(b, name);
            }
        }
        StmtKind::With { items, body, .. } => {
            for it in items.iter_mut() {
                tdz_rewrite_expr(&mut it.expr, name);
            }
            let shadowed = items.iter().any(|it| it.var.as_deref() == Some(name));
            if !shadowed {
                tdz_rewrite_list(body, name);
            }
        }
        StmtKind::Using {
            var,
            resource,
            body } => {
            tdz_rewrite_expr(resource, name);
            if var != name {
                tdz_rewrite_list(body, name);
            }
        }
        StmtKind::Lock { expr, body } => {
            tdz_rewrite_expr(expr, name);
            tdz_rewrite_list(body, name);
        }
        // Deferred execution / own scope — a body that runs later may
        // legally observe the initialized binding.
        StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {}
        _ => {}
    }
}

/// Assignment-target position: a bare `Ident` is a write (left untouched by
/// this narrow pass); Member / Index targets still evaluate their object and
/// index subexpressions as reads.
fn tdz_rewrite_place(target: &mut Expression, name: &str) {
    match &mut target.kind {
        ExprKind::Member { object, .. } => tdz_rewrite_expr(object, name),
        ExprKind::Index { object, index, .. } => {
            tdz_rewrite_expr(object, name);
            tdz_rewrite_expr(index, name);
        }
        _ => {}
    }
}

fn tdz_rewrite_expr(e: &mut Expression, name: &str) {
    if let ExprKind::Ident(n) = &e.kind {
        if n == name {
            *e = tdz_throw_expr(name);
        }
        return;
    }
    match &mut e.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            tdz_rewrite_expr(left, name);
            tdz_rewrite_expr(right, name);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::RefLoad(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. } => tdz_rewrite_expr(expr, name),
        ExprKind::Yield(Some(expr)) => tdz_rewrite_expr(expr, name),
        ExprKind::Ternary { cond, then, else_ } => {
            tdz_rewrite_expr(cond, name);
            tdz_rewrite_expr(then, name);
            tdz_rewrite_expr(else_, name);
        }
        ExprKind::Member { object, .. } => tdz_rewrite_expr(object, name),
        ExprKind::Index { object, index, .. } => {
            tdz_rewrite_expr(object, name);
            tdz_rewrite_expr(index, name);
        }
        ExprKind::Call { callee, args, .. } => {
            tdz_rewrite_expr(callee, name);
            for a in args {
                tdz_rewrite_expr(&mut a.value, name);
            }
        }
        ExprKind::New { class, args } => {
            tdz_rewrite_expr(class, name);
            for a in args {
                tdz_rewrite_expr(&mut a.value, name);
            }
        }
        ExprKind::SuperCall { args, .. } => {
            for a in args {
                tdz_rewrite_expr(&mut a.value, name);
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            tdz_rewrite_place(target, name);
            tdz_rewrite_expr(value, name);
        }
        ExprKind::Array(elems) => {
            for el in elems {
                if let Some(k) = &mut el.key {
                    tdz_rewrite_expr(k, name);
                }
                tdz_rewrite_expr(&mut el.value, name);
            }
        }
        ExprKind::Tuple(v) | ExprKind::Set(v) | ExprKind::Sequence(v) => {
            for x in v {
                tdz_rewrite_expr(x, name);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                match p {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        tdz_rewrite_expr(key, name);
                        tdz_rewrite_expr(value, name);
                    }
                    ObjectProperty::Spread(e) => tdz_rewrite_expr(e, name),
                    // Shorthand is a read, but rewriting it needs a
                    // KeyValue expansion — out of scope for this pass.
                    // Method / Accessor bodies are deferred execution.
                    _ => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for p in parts {
                match p {
                    InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                        tdz_rewrite_expr(e, name)
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            for o in [lower, upper, step].into_iter().flatten() {
                tdz_rewrite_expr(o, name);
            }
        }
        ExprKind::Range { start, end, .. } => {
            tdz_rewrite_expr(start, name);
            tdz_rewrite_expr(end, name);
        }
        ExprKind::StaticAccess { class, member } => {
            tdz_rewrite_expr(class, name);
            tdz_rewrite_expr(member, name);
        }
        // Deferred execution / own scope.
        ExprKind::Lambda { .. }
        | ExprKind::FunctionExpr(_)
        | ExprKind::ClassExpr { .. }
        | ExprKind::Comprehension { .. } => {}
        _ => {}
    }
}

/// `(() => { throw new ReferenceError("Cannot access '<name>' before
/// initialization") })()` — pure JS-shape AST; throws exactly when (and only
/// when) the dead-zone reference would have been evaluated.
fn tdz_throw_expr(name: &str) -> Expression {
    let message = format!("Cannot access '{name}' before initialization");
    let throw_stmt = Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(Expression::ident("ReferenceError")),
            args: vec![Argument::positional(Expression::string(&message))] })),
        cause: None });
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![throw_stmt]),
            is_async: false,
            captures: Vec::new() })),
        args: Vec::new(),
        optional: false })
}

fn fold_const_computed_names(body: &mut [Statement]) {
    use std::collections::HashMap;
    let mut consts: HashMap<String, String> = HashMap::new();
    for stmt in body.iter() {
        if let StmtKind::VarDecl { declarations, kind } = &stmt.kind {
            if matches!(
                kind,
                VarDeclKind::Const | VarDeclKind::Let | VarDeclKind::Var
            ) {
                for d in declarations {
                    if let (BindingPattern::Ident(name), Some(init)) = (&d.pattern, &d.init) {
                        if let ExprKind::Lit(Literal::Str(s)) = &init.kind {
                            consts.insert(name.clone(), s.clone());
                        }
                    }
                }
            }
        }
    }
    if consts.is_empty() {
        return;
    }
    for stmt in body.iter_mut() {
        rewrite_class_method_names(stmt, &consts);
    }
}

// ── TC39 explicit-resource-management classes (§12 DisposableStack) ─────────
//
// Canonical JS sources compiled through the normal class pipeline. The
// `Symbol.dispose` / `Symbol.asyncDispose` hooks are constructor-assigned
// instance properties so the `using` lowering's `x[Symbol.dispose]` lookup
// resolves through the same runtime-key path object literals use.

const DISPOSABLE_STACK_JS: &str = r#"
class DisposableStack {
    constructor() {
        this.__entries = [];
        this.disposed = false;
        this[Symbol.dispose] = () => { this.dispose(); };
    }
    use(value) {
        if (value !== null && value !== undefined) {
            const d = value[Symbol.dispose];
            if (typeof d !== "function") { throw new TypeError("Object is not disposable"); }
            this.__entries.push(() => d.call(value));
        }
        return value;
    }
    adopt(value, onDispose) {
        this.__entries.push(() => onDispose(value));
        return value;
    }
    defer(onDispose) {
        this.__entries.push(onDispose);
    }
    move() {
        const next = new DisposableStack();
        next.__entries = this.__entries;
        this.__entries = [];
        this.disposed = true;
        return next;
    }
    dispose() {
        if (this.disposed) { return; }
        this.disposed = true;
        const entries = this.__entries;
        this.__entries = [];
        for (let i = entries.length - 1; i >= 0; i--) { entries[i](); }
    }
}
"#;

const ASYNC_DISPOSABLE_STACK_JS: &str = r#"
class AsyncDisposableStack {
    constructor() {
        this.__entries = [];
        this.disposed = false;
        this[Symbol.asyncDispose] = () => this.disposeAsync();
    }
    use(value) {
        if (value !== null && value !== undefined) {
            const d = value[Symbol.asyncDispose] ?? value[Symbol.dispose];
            if (typeof d !== "function") { throw new TypeError("Object is not disposable"); }
            this.__entries.push(() => d.call(value));
        }
        return value;
    }
    adopt(value, onDispose) {
        this.__entries.push(() => onDispose(value));
        return value;
    }
    defer(onDispose) {
        this.__entries.push(onDispose);
    }
    move() {
        const next = new AsyncDisposableStack();
        next.__entries = this.__entries;
        this.__entries = [];
        this.disposed = true;
        return next;
    }
    async disposeAsync() {
        if (this.disposed) { return; }
        this.disposed = true;
        const entries = this.__entries;
        this.__entries = [];
        for (let i = entries.length - 1; i >= 0; i--) { await entries[i](); }
    }
}
"#;

/// Parse a trusted runtime-class snippet into statements (no recursive
/// injection passes — the snippet goes through the main module's passes
/// after splicing).
fn parse_runtime_class_snippet(src: &str) -> Vec<Statement> {
    let Ok(pairs) = JsParser::parse(Rule::program, src) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for top in pairs {
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            Rule::EOI => continue,
            _ => {
                if let Ok(s) = walk_statement(top) {
                    out.push(s);
                }
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI | Rule::NEWLINE | Rule::import_statement => continue,
                _ => {
                    if let Ok(s) = walk_statement(pair) {
                        out.push(s);
                    }
                }
            }
        }
    }
    out
}

fn js_array_elision_marker() -> ArrayElement {
    ArrayElement {
        key: Some(Expression::int(-1)),
        value: Expression::new(ExprKind::Lit(Literal::Undefined)),
        spread: false,
        by_ref: false }
}

fn rewrite_class_method_names(
    stmt: &mut Statement,
    consts: &std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_expression_keys(expr, consts),
        StmtKind::ClassDecl { members, .. } => {
            rewrite_class_members(members, consts);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations.iter_mut() {
                rewrite_pattern_keys(&mut d.pattern, consts);
                if let Some(init) = d.init.as_mut() {
                    rewrite_expression_keys(init, consts);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for stmt in body.iter_mut() {
                rewrite_class_method_names(stmt, consts);
            }
        }
        StmtKind::Block(stmts) => {
            for s in stmts.iter_mut() {
                rewrite_class_method_names(s, consts);
            }
        }
        _ => {}
    }
}

fn rewrite_class_members(
    members: &mut [ClassMember],
    consts: &std::collections::HashMap<String, String>,
) {
    for member in members.iter_mut() {
        match member {
            ClassMember::Field { init, .. } => {
                if let Some(init) = init.as_mut() {
                    rewrite_expression_keys(init, consts);
                }
            }
            ClassMember::Method(box_stmt) => {
                if let StmtKind::FunctionDecl { name, .. } = &mut box_stmt.kind {
                    if let Some(resolved) = resolve_const_key(name, consts) {
                        *name = resolved;
                    }
                    if let Some(alias) = js_well_known_symbol_alias_from_raw(name) {
                        *name = alias.to_string();
                    }
                }
                rewrite_class_method_names(box_stmt, consts);
            }
            ClassMember::Constructor {
                body, base_args, ..
            } => {
                if let Some(args) = base_args.as_mut() {
                    for arg in args.iter_mut() {
                        rewrite_expression_keys(arg, consts);
                    }
                }
                for stmt in body.iter_mut() {
                    rewrite_class_method_names(stmt, consts);
                }
            }
            ClassMember::Property {
                name,
                getter,
                setter,
                ..
            } => {
                if let Some(resolved) = resolve_const_key(name, consts) {
                    *name = resolved;
                }
                if let Some(alias) = js_well_known_symbol_alias_from_raw(name) {
                    *name = alias.to_string();
                }
                if let Some(getter) = getter.as_mut() {
                    for stmt in getter.iter_mut() {
                        rewrite_class_method_names(stmt, consts);
                    }
                }
                if let Some(setter) = setter.as_mut() {
                    for stmt in setter.body.iter_mut() {
                        rewrite_class_method_names(stmt, consts);
                    }
                }
            }
            ClassMember::Const { value, .. } => rewrite_expression_keys(value, consts),
            ClassMember::NestedType(stmt) => rewrite_class_method_names(stmt, consts),
            ClassMember::Event { .. } | ClassMember::Augment(_) => {}
        }
    }
}

fn rewrite_expression_keys(
    expr: &mut Expression,
    consts: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Async(op) => {
            for child in op.children_mut() {
                rewrite_expression_keys(child, consts);
            }
        }
        ExprKind::Chan(op) => {
            for child in op.children_mut() {
                rewrite_expression_keys(child, consts);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right }
        | ExprKind::Walrus {
            target: left,
            value: right }
        | ExprKind::Range {
            start: left,
            end: right,
            ..
        } => {
            rewrite_expression_keys(left, consts);
            rewrite_expression_keys(right, consts);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr) => rewrite_expression_keys(expr, consts),
        ExprKind::RefOf(place) => rewrite_place_expression(place, consts),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_expression_keys(cond, consts);
            rewrite_expression_keys(then, consts);
            rewrite_expression_keys(else_, consts);
        }
        ExprKind::Member { object, .. } => rewrite_expression_keys(object, consts),
        ExprKind::Index { object, index, .. } => {
            rewrite_expression_keys(object, consts);
            rewrite_expression_keys(index, consts);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_expression_keys(callee, consts);
            for arg in args.iter_mut() {
                rewrite_expression_keys(&mut arg.value, consts);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_expression_keys(class, consts);
            for arg in args.iter_mut() {
                rewrite_expression_keys(&mut arg.value, consts);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_expression_keys(expr, consts),
            LambdaBody::Block(stmts) => {
                for stmt in stmts.iter_mut() {
                    rewrite_class_method_names(stmt, consts);
                }
            }
        },
        ExprKind::Array(elements) => {
            for element in elements.iter_mut() {
                if let Some(key) = element.key.as_mut() {
                    rewrite_expression_keys(key, consts);
                }
                rewrite_expression_keys(&mut element.value, consts);
            }
        }
        ExprKind::Tuple(items)
        | ExprKind::Set(items)
        | ExprKind::Sequence(items)
        | ExprKind::Zip {
            iterables: items, ..
        } => {
            for item in items.iter_mut() {
                rewrite_expression_keys(item, consts);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields.iter_mut() {
                rewrite_expression_keys(value, consts);
            }
        }
        ExprKind::Object(props) => {
            for prop in props.iter_mut() {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_expression_keys(key, consts);
                        if let ExprKind::Lit(Literal::Str(s)) = &mut key.kind {
                            if let Some(name) = s.strip_prefix("__get_") {
                                if let Some(resolved) = consts.get(name) {
                                    *s = format!("__get_{}", resolved);
                                }
                            } else if let Some(name) = s.strip_prefix("__set_") {
                                if let Some(resolved) = consts.get(name) {
                                    *s = format!("__set_{}", resolved);
                                }
                            }
                        }
                        rewrite_expression_keys(value, consts);
                    }
                    ObjectProperty::Spread(expr) => rewrite_expression_keys(expr, consts),
                    ObjectProperty::Method { key, value }
                    | ObjectProperty::Accessor { key, value, .. } => {
                        if let Some(resolved) = resolve_const_key(key, consts) {
                            *key = resolved;
                        }
                        rewrite_class_method_names(value, consts);
                    }
                    ObjectProperty::Computed { key, value } => {
                        rewrite_expression_keys(key, consts);
                        rewrite_expression_keys(value, consts);
                        if let ExprKind::Ident(name) = &key.kind {
                            if let Some(resolved) = consts.get(name.as_str()) {
                                *prop = ObjectProperty::KeyValue {
                                    key: Expression::string(resolved),
                                    value: value.clone() };
                            }
                        }
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts.iter_mut() {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_expression_keys(expr, consts);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Yield(Some(expr)) => rewrite_expression_keys(expr, consts),
        ExprKind::SuperCall { args, .. } => {
            for arg in args.iter_mut() {
                rewrite_expression_keys(&mut arg.value, consts);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent.as_mut() {
                rewrite_expression_keys(parent, consts);
            }
            rewrite_class_members(members, consts);
        }
        ExprKind::FunctionExpr(stmt) => rewrite_class_method_names(stmt, consts),
        ExprKind::StaticAccess { class, member } => {
            rewrite_expression_keys(class, consts);
            rewrite_expression_keys(member, consts);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_expression_keys(subject, consts);
            for arm in arms.iter_mut() {
                if let Some(conditions) = arm.conditions.as_mut() {
                    for condition in conditions.iter_mut() {
                        rewrite_expression_keys(condition, consts);
                    }
                }
                rewrite_expression_keys(&mut arm.body, consts);
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower.as_mut() {
                rewrite_expression_keys(lower, consts);
            }
            if let Some(upper) = upper.as_mut() {
                rewrite_expression_keys(upper, consts);
            }
            if let Some(step) = step.as_mut() {
                rewrite_expression_keys(step, consts);
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_expression_keys(element, consts);
            for generator in generators.iter_mut() {
                rewrite_expression_keys(&mut generator.iter, consts);
                for cond in generator.conditions.iter_mut() {
                    rewrite_expression_keys(cond, consts);
                }
            }
        }
        ExprKind::IsType { expr, .. } | ExprKind::Cast { expr, .. } => {
            rewrite_expression_keys(expr, consts);
        }
        ExprKind::Yield(None)
        | ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::DefaultOf(_)
        | ExprKind::AddressOf(_)
        | ExprKind::Destructure(_) => {}
    }
}

fn rewrite_place_expression(
    place: &mut PlaceExpr,
    consts: &std::collections::HashMap<String, String>,
) {
    match place {
        PlaceExpr::Member { object, .. } => rewrite_expression_keys(object, consts),
        PlaceExpr::Index { object, index, .. } => {
            rewrite_expression_keys(object, consts);
            rewrite_expression_keys(index, consts);
        }
        PlaceExpr::Deref(expr) => rewrite_expression_keys(expr, consts),
        PlaceExpr::Ident(_) => {}
    }
}

fn resolve_const_key(
    key: &str,
    consts: &std::collections::HashMap<String, String>,
) -> Option<String> {
    let trimmed = key.trim_start_matches('[').trim_end_matches(']').trim();
    consts.get(trimmed).cloned()
}

fn rewrite_pattern_keys(
    pat: &mut BindingPattern,
    consts: &std::collections::HashMap<String, String>,
) {
    if let BindingPattern::Object(props) = pat {
        for p in props.iter_mut() {
            // `[ident]: val` lands either as the bare ident text
            // (walker dropped the brackets) or as `[ident]`. Probe
            // both forms and resolve via the const map.
            let trimmed = p
                .key
                .trim_start_matches('[')
                .trim_end_matches(']')
                .trim()
                .to_string();
            if let Some(resolved) = consts.get(trimmed.as_str()) {
                p.key = resolved.clone();
            }
            if let Some(ref mut nested) = p.value {
                rewrite_pattern_keys(nested, consts);
            }
        }
    }
}

// ── Statements ──────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => {
            let stmts = pair
                .into_inner()
                .filter(|p| p.as_rule() != Rule::NEWLINE)
                .map(walk_statement)
                .collect::<Result<Vec<_>, _>>()?;
            StmtKind::Block(stmts)
        }
        Rule::variable_declaration => walk_var_decl(pair)?,
        Rule::function_declaration | Rule::async_function_declaration => walk_func_decl(pair)?,
        Rule::class_declaration => walk_class_decl(pair)?,
        Rule::if_statement => walk_if(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::while_statement => walk_while(pair)?,
        Rule::do_while_statement => walk_do_while(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => walk_break(pair)?,
        Rule::continue_statement => walk_continue(pair)?,
        Rule::throw_statement => walk_throw(pair)?,
        Rule::try_statement => walk_try(pair)?,
        Rule::export_statement => walk_export(pair)?,
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::debugger_statement => StmtKind::Empty,
        Rule::using_declaration => walk_using_decl(pair, false)?,
        Rule::await_using_declaration => walk_using_decl(pair, true)?,
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::NEWLINE => return Ok(Statement::new(StmtKind::Empty)),
        other => return Err(format!("Unexpected statement rule: {:?}", other)) };
    Ok(Statement::with_span(kind, span))
}

// ── Variable declaration ────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let kind_pair = next_rule(&mut inner, Rule::var_kind)?;
    let var_kind = match kind_pair.as_str() {
        "var" => VarDeclKind::Var,
        "let" => VarDeclKind::Let,
        "const" => VarDeclKind::Const,
        _ => VarDeclKind::Let };
    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            declarations.push(walk_var_declarator(p)?);
        }
    }
    // ECMA-262 §13.3.1: a `const` LexicalBinding requires an Initializer
    // (`const x;` is a SyntaxError). `let`/`var` may omit it. `for (const x of
    // …)` / `for (const x in …)` bindings don't reach here (they use
    // `extract_for_target`), so this only rejects genuinely-uninitialized
    // standalone `const` declarations.
    if matches!(var_kind, VarDeclKind::Const) {
        if declarations.iter().any(|d| d.init.is_none()) {
            return Err("Missing initializer in const declaration".to_string());
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind })
}

fn walk_var_declarator(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut inner = pair.into_inner();
    let pattern = walk_binding_pattern(inner.next().ok_or("Expected binding pattern")?)?;
    let init = inner.next().map(walk_expression).transpose()?;
    Ok(VarDeclarator {
        pattern,
        type_hint: None,
        init,
        array_bounds: None,
        with_events: false })
}

// ES2026 `using x = expr` / `await using x = expr` — emit a `StmtKind::Using`
// marker (empty body; an `__vybe_await_using` sentinel statement marks the
// await form). `lower_using_declarations` folds the rest of the enclosing
// statement list into the spec's try/finally desugaring. Multi-declarator or
// non-identifier forms fall back to plain `const` (spec only allows binding
// identifiers anyway).
fn walk_using_decl(pair: Pair<Rule>, is_await: bool) -> Result<StmtKind, String> {
    let mut declarations = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::using_declarator {
            declarations.push(walk_var_declarator(p)?);
        }
    }
    if declarations.len() == 1 {
        let d = &declarations[0];
        if let (BindingPattern::Ident(name), Some(init)) = (&d.pattern, &d.init) {
            let body = if is_await {
                vec![Statement::new(StmtKind::Expr(Expression::ident(
                    "__vybe_await_using",
                )))]
            } else {
                Vec::new()
            };
            return Ok(StmtKind::Using {
                var: name.clone(),
                resource: init.clone(),
                body });
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const })
}

// ── `using` lowering — TC39 explicit-resource-management ────────────────────
//
// `using x = expr; rest…` desugars in place (one statement list at a time) to
//
//   const x = expr;
//   const __vybe_using_dispose_N =
//       (x === null || x === undefined) ? undefined : x[Symbol.dispose];
//   if (x !== null && x !== undefined &&
//       typeof __vybe_using_dispose_N !== "function") {
//       throw new TypeError("Object is not disposable");   // §9.3.3
//   }
//   try { rest… } finally {
//       if (__vybe_using_dispose_N !== undefined) {
//           __vybe_using_dispose_N.call(x);                 // §9.3.5
//       }
//   }
//
// `await using` looks up `Symbol.asyncDispose ?? Symbol.dispose` (§9.3.4) and
// awaits the disposal call. LIFO order for multiple `using`s falls out of the
// recursive nesting. Pure JS-shape AST — no compiler or VM involvement.

static USING_COUNTER: AtomicUsize = AtomicUsize::new(0);

fn lower_using_declarations(stmts: &mut Vec<Statement>) {
    for (i, s) in stmts.iter().enumerate() {
        if matches!(&s.kind, StmtKind::Using { .. }) {
            let mut tail: Vec<Statement> = stmts.drain(i + 1..).collect();
            let marker = stmts.pop().expect("using marker");
            let StmtKind::Using {
                var,
                resource,
                body } = marker.kind
            else {
                unreachable!()
            };
            let is_await = !body.is_empty();
            lower_using_declarations(&mut tail);
            stmts.extend(lower_one_using(&var, resource, is_await, tail));
            break;
        }
    }
    for s in stmts.iter_mut() {
        visit_nested_stmt_lists(&mut s.kind, lower_using_declarations);
    }
}

fn lower_one_using(
    var: &str,
    resource: Expression,
    is_await: bool,
    tail: Vec<Statement>,
) -> Vec<Statement> {
    let decl_x = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(var.to_string()),
            type_hint: None,
            init: Some(resource),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Const });
    let mut out = vec![decl_x];
    out.extend(using_disposal_wrap(var, is_await, tail));
    out
}

/// The disposal half of the `using` desugaring — `var` is already bound
/// (by a const declaration or a for-of loop binding).
fn using_disposal_wrap(var: &str, is_await: bool, tail: Vec<Statement>) -> Vec<Statement> {
    let n = USING_COUNTER.fetch_add(1, Ordering::Relaxed);
    let disp = format!("__vybe_using_dispose_{n}");
    let bin = |op, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r) })
    };
    let sym_lookup = |field: &str| {
        Expression::new(ExprKind::Index {
            object: Box::new(Expression::ident(var)),
            index: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Symbol")),
                field: field.to_string(),
                null_safe: false })),
            null_safe: false })
    };
    let undefined_lit = || Expression::new(ExprKind::Lit(Literal::Undefined));
    let null_lit = || Expression::new(ExprKind::Lit(Literal::Null));
    let const_decl = |name: &str, init: Expression| {
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(name.to_string()),
                type_hint: None,
                init: Some(init),
                array_bounds: None,
                with_events: false }],
            kind: VarDeclKind::Const })
    };

    // §9.3.4 GetDisposeMethod — resolved once, at binding time.
    let method_lookup = if is_await {
        Expression::new(ExprKind::NullCoalesce {
            left: Box::new(sym_lookup("asyncDispose")),
            right: Box::new(sym_lookup("dispose")) })
    } else {
        sym_lookup("dispose")
    };
    let is_nullish = bin(
        BinOp::Or,
        bin(BinOp::StrictEq, Expression::ident(var), null_lit()),
        bin(BinOp::StrictEq, Expression::ident(var), undefined_lit()),
    );
    let decl_disp = const_decl(
        &disp,
        Expression::new(ExprKind::Ternary {
            cond: Box::new(is_nullish),
            then: Box::new(undefined_lit()),
            else_: Box::new(method_lookup) }),
    );

    // §9.3.3 AddDisposableResource — non-nullish resource without a callable
    // dispose method is a TypeError at binding time.
    let guard_cond = bin(
        BinOp::And,
        bin(
            BinOp::And,
            bin(BinOp::StrictNotEq, Expression::ident(var), null_lit()),
            bin(BinOp::StrictNotEq, Expression::ident(var), undefined_lit()),
        ),
        bin(
            BinOp::StrictNotEq,
            Expression::new(ExprKind::TypeOf(Box::new(Expression::ident(&disp)))),
            Expression::string("function"),
        ),
    );
    let guard = Statement::new(StmtKind::If {
        cond: guard_cond,
        then_body: vec![Statement::new(StmtKind::Throw {
            expr: Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident("TypeError")),
                args: vec![Argument::positional(Expression::string(
                    "Object is not disposable",
                ))] })),
            cause: None })],
        elifs: Vec::new(),
        else_body: None });

    // finally { if (disp !== undefined) [await] disp.call(x); }
    let mut call_disp = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&disp)),
            field: "call".to_string(),
            null_safe: false })),
        args: vec![Argument::positional(Expression::ident(var))],
        optional: false });
    if is_await {
        call_disp = Expression::new(ExprKind::Await(Box::new(call_disp)));
    }
    let finally_body = vec![Statement::new(StmtKind::If {
        cond: bin(
            BinOp::StrictNotEq,
            Expression::ident(&disp),
            undefined_lit(),
        ),
        then_body: vec![Statement::new(StmtKind::Expr(call_disp))],
        elifs: Vec::new(),
        else_body: None })];

    let try_stmt = Statement::new(StmtKind::Try {
        body: tail,
        catches: Vec::new(),
        else_body: None,
        finally: Some(finally_body) });

    vec![decl_disp, guard, try_stmt]
}

fn walk_binding_pattern(pair: Pair<Rule>) -> Result<BindingPattern, String> {
    match pair.as_rule() {
        Rule::ident_name => Ok(BindingPattern::Ident(pair.as_str().to_string())),
        Rule::binding_pattern => {
            walk_binding_pattern(pair.into_inner().next().ok_or("Empty binding")?)
        }
        Rule::object_pattern => {
            let props = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::object_pattern_prop)
                .map(walk_object_pattern_prop)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BindingPattern::Object(props))
        }
        Rule::array_pattern => {
            let elems = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::array_pattern_elem)
                .map(walk_array_pattern_elem)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(BindingPattern::Array(elems))
        }
        other => Err(format!("Unexpected binding pattern: {:?}", other)) }
}

fn walk_object_pattern_prop(pair: Pair<Rule>) -> Result<ObjectPatternProp, String> {
    let is_rest = pair.as_str().starts_with("...");
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty object pattern prop")?;
    let key = first.as_str().to_string();
    if is_rest {
        return Ok(ObjectPatternProp {
            key,
            value: None,
            default: None,
            is_rest: true });
    }
    let mut value = None;
    let mut default = None;
    for p in inner {
        match p.as_rule() {
            Rule::binding_pattern => value = Some(walk_binding_pattern(p)?),
            _ => default = Some(walk_expression(p)?) }
    }
    Ok(ObjectPatternProp {
        key,
        value,
        default,
        is_rest: false })
}

fn walk_array_pattern_elem(pair: Pair<Rule>) -> Result<ArrayPatternElem, String> {
    let src = pair.as_str().to_string();
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty array pattern elem")?;
    match first.as_rule() {
        Rule::array_hole => Ok(ArrayPatternElem::Hole),
        Rule::ident_name => {
            // Could be rest (...name) or simple binding
            let name = first.as_str().to_string();
            let default = inner.next().map(walk_expression).transpose()?;
            // If parent started with "..." it's rest — check source text
            if src.starts_with("...") {
                Ok(ArrayPatternElem::Rest(name))
            } else {
                Ok(ArrayPatternElem::Pattern(
                    BindingPattern::Ident(name),
                    default,
                ))
            }
        }
        Rule::binding_pattern => {
            let pat = walk_binding_pattern(first)?;
            let default = inner.next().map(walk_expression).transpose()?;
            Ok(ArrayPatternElem::Pattern(pat, default))
        }
        other => Err(format!("Unexpected array pattern elem: {:?}", other)) }
}

// ── Function declaration ────────────────────────────────────────────────────

/// Recursively scan a function body for `yield` / `yield from` expressions.
/// Does NOT descend into nested function/closure/class bodies — those are
/// their own generator scope.
fn body_contains_yield(stmts: &[Statement]) -> bool {
    vybe_ast::statements_contain_yield_outside_nested_scopes(stmts)
}

fn walk_func_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let is_async = matches!(
        pair.as_rule(),
        Rule::async_function_declaration | Rule::async_function_expression
    );
    let inner = pair.into_inner();
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut param_prologue = Vec::new();
    let mut has_generator_marker = false;

    for p in inner {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::generator_marker => has_generator_marker = true,
            Rule::param_list => {
                let (parsed_params, prologue) = walk_params_with_prologue(p)?;
                params = parsed_params;
                param_prologue = prologue;
            }
            Rule::function_body => body = walk_body(p)?,
            Rule::async_kw => {}
            _ => {}
        }
    }

    if !param_prologue.is_empty() {
        let mut full_body = param_prologue;
        full_body.extend(body);
        body = full_body;
    }

    let is_generator = has_generator_marker || body_contains_yield(&body);
    Ok(wrap_generator_if_needed(
        name,
        params,
        body,
        is_async,
        is_generator,
    ))
}

fn walk_params_with_prologue(pair: Pair<Rule>) -> Result<(Vec<Param>, Vec<Statement>), String> {
    let mut params = Vec::new();
    let mut prologue = Vec::new();
    let mut destructure_idx = 0usize;

    for p in pair.into_inner().filter(|p| p.as_rule() == Rule::param) {
        let (param, init_stmt) = walk_param_with_prologue(p, destructure_idx)?;
        destructure_idx += 1;
        params.push(param);
        if let Some(stmt) = init_stmt {
            prologue.push(stmt);
        }
    }

    Ok((params, prologue))
}

fn walk_param_with_prologue(
    pair: Pair<Rule>,
    destructure_idx: usize,
) -> Result<(Param, Option<Statement>), String> {
    let src = pair.as_str();
    let is_rest = src.starts_with("...");
    let mut binding = None;
    let mut default = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => binding = Some(BindingPattern::Ident(p.as_str().to_string())),
            Rule::binding_pattern => binding = Some(walk_binding_pattern(p)?),
            _ => default = Some(walk_expression(p)?) }
    }
    let binding = binding.ok_or("Expected parameter binding")?;

    match binding {
        BindingPattern::Ident(name) => Ok((
            Param {
                name,
                type_hint: None,
                default,
                pass_by: PassBy::Value,
                is_rest,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false },
            None,
        )),
        pattern => {
            let temp_name = format!("__param_destruct_{}", destructure_idx);
            let stmt = Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern,
                    type_hint: None,
                    init: Some(Expression::ident(&temp_name)),
                    array_bounds: None,
                    with_events: false }],
                kind: VarDeclKind::Let });
            Ok((
                Param {
                    name: temp_name,
                    type_hint: None,
                    default,
                    pass_by: PassBy::Value,
                    is_rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false },
                Some(stmt),
            ))
        }
    }
}

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .map(walk_statement)
        .collect()
}

// ── Class declaration ───────────────────────────────────────────────────────

fn walk_class_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();
    // `class X extends null {}` — legal definition, throwing construction.
    let mut extends_null = false;
    // Pre-class statements emitted to bind synthetic names for non-trivial
    // `extends <expression>` heads (e.g. `class X extends getBase()`).
    let mut pre_class_stmts: Vec<Statement> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::assignment_expression
            | Rule::conditional_expression
            | Rule::logical_expr
            | Rule::comparison => {
                // `extends Expr` — if Expr is a bare identifier we use
                // it as the parent name directly (back-compat). Otherwise
                // we lower to `var __extends_<class>_<n> = Expr;` before
                // the class and use that synthetic name as the parent.
                // Lets `class X extends getBase()` /
                // `class X extends Mixin(Base)` work without changing
                // the AST shape (parent stays a single ident name).
                let init = walk_expression(p.clone())?;
                let raw = extract_ident_name(&p);
                let is_ident_parent = matches!(&init.kind, ExprKind::Ident(_));
                // §15.7.5 ClassDefinitionEvaluation: `extends undefined`
                // throws a TypeError at DEFINITION time; `extends null`
                // is legal at definition — the class is heritage-less and
                // constructing it throws instead (node-verified).
                if raw == "undefined" {
                    pre_class_stmts.push(Statement::new(StmtKind::Throw {
                        expr: Some(Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident("TypeError")),
                            args: vec![Argument::positional(Expression::string(
                                "Class extends value undefined is not a constructor or null",
                            ))] })),
                        cause: None }));
                } else if raw == "null" {
                    extends_null = true;
                } else if matches!(
                    &init.kind,
                    ExprKind::Lit(Literal::Int(_))
                        | ExprKind::Lit(Literal::Float(_))
                        | ExprKind::Lit(Literal::BigInt(_))
                        | ExprKind::Lit(Literal::Str(_))
                        | ExprKind::Lit(Literal::Bool(_))
                        | ExprKind::Lit(Literal::Char(_))
                ) {
                    pre_class_stmts.push(Statement::new(StmtKind::Throw {
                        expr: Some(Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident("TypeError")),
                            args: vec![Argument::positional(Expression::string(
                                "Class extends value is not a constructor or null",
                            ))] })),
                        cause: None }));
                } else if is_ident_parent {
                    if let ExprKind::Ident(parent_name) = init.kind {
                        parents.push(parent_name);
                    }
                } else {
                    let synth = format!("__extends_{}_{}", name, parents.len());
                    pre_class_stmts.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(synth.clone()),
                            type_hint: None,
                            init: Some(init),
                            array_bounds: None,
                            with_events: false }],
                        kind: VarDeclKind::Var }));
                    parents.push(synth);
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        members.push(walk_class_member(m)?);
                    }
                }
            }
            _ => {}
        }
    }

    // `extends null` (§15.7.14): the implicit derived constructor would
    // call super — with a null heritage that's "Super constructor null …
    // is not a constructor". Synthesize a constructor that throws the
    // TypeError at construction time (node-verified). A user-written
    // constructor keeps its own body.
    if extends_null
        && !members
            .iter()
            .any(|m| matches!(m, ClassMember::Constructor { .. }))
    {
        members.push(ClassMember::Constructor {
            name: None,
            params: vec![],
            body: vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::new(ExprKind::New {
                    class: Box::new(Expression::ident("TypeError")),
                    args: vec![Argument::positional(Expression::string(&format!(
                        "Super constructor null of {} is not a constructor",
                        name
                    )))] })),
                cause: None })],
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public });
    }

    let class_stmt = StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![] };

    if pre_class_stmts.is_empty() {
        Ok(class_stmt)
    } else {
        // Wrap: pre-class extends bindings, then class declaration.
        let mut block = pre_class_stmts;
        block.push(Statement::new(class_stmt));
        Ok(StmtKind::Block(block))
    }
}

fn walk_class_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut is_static = false;
    let mut inner_pairs: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Skip TC39 decorator pairs — parsed but not yet executed
    while inner_pairs
        .first()
        .map_or(false, |p| p.as_rule() == Rule::decorator)
    {
        inner_pairs.remove(0);
    }

    // Check for static keyword
    if inner_pairs
        .first()
        .map_or(false, |p| p.as_rule() == Rule::static_kw)
    {
        is_static = true;
        inner_pairs.remove(0);
    }

    let member_pair = inner_pairs.into_iter().next().ok_or("Empty class member")?;

    // ES2022 static block — convert to a synthetic static method __static_init
    if member_pair.as_rule() == Rule::static_block {
        let stmts: Vec<Statement> = member_pair
            .into_inner()
            .filter(|p| !matches!(p.as_rule(), Rule::NEWLINE | Rule::static_kw))
            .map(walk_statement)
            .collect::<Result<_, _>>()?;
        if stmts.iter().any(stmt_contains_await) {
            return Err("Await is not allowed in class static initialization blocks".to_string());
        }
        let func = Statement::new(StmtKind::FunctionDecl {
            name: "__static_init".to_string(),
            params: vec![],
            return_type: None,
            body: stmts,
            modifiers: Modifiers {
                is_static: true,
                ..Default::default()
            },
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: true });
        return Ok(ClassMember::Method(Box::new(func)));
    }

    match member_pair.as_rule() {
        Rule::getter_method => {
            let mut name = String::new();
            let mut body = Vec::new();
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => {
                        name = extract_property_name(&p)
                    }
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            Ok(ClassMember::Property {
                name,
                type_hint: None,
                getter: Some(body),
                setter: None,
                is_auto: false,
                modifiers: Modifiers {
                    is_static,
                    ..Default::default()
                } })
        }
        Rule::setter_method => {
            let mut name = String::new();
            let mut param = Param {
                name: "value".into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false };
            let mut body = Vec::new();
            let mut param_prologue = Vec::new();
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => {
                        name = extract_property_name(&p)
                    }
                    Rule::param => {
                        let (parsed_param, init_stmt) = walk_param_with_prologue(p, 0)?;
                        param = parsed_param;
                        if let Some(stmt) = init_stmt {
                            param_prologue.push(stmt);
                        }
                    }
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            if !param_prologue.is_empty() {
                let mut full_body = param_prologue;
                full_body.extend(body);
                body = full_body;
            }
            Ok(ClassMember::Property {
                name,
                type_hint: None,
                getter: None,
                setter: Some(PropertySetter { param, body }),
                is_auto: false,
                modifiers: Modifiers {
                    is_static,
                    ..Default::default()
                } })
        }
        Rule::class_method => {
            let mut name = String::new();
            let mut params = Vec::new();
            let mut body = Vec::new();
            let mut is_async = false;
            let mut is_generator = false;
            let mut param_prologue = Vec::new();
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::generator_marker => is_generator = true,
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => {
                        name = extract_property_name(&p)
                    }
                    Rule::param_list => {
                        let (parsed_params, prologue) = walk_params_with_prologue(p)?;
                        params = parsed_params;
                        param_prologue = prologue;
                    }
                    Rule::function_body => body = walk_body(p)?,
                    _ => {}
                }
            }
            if !param_prologue.is_empty() {
                let mut full_body = param_prologue;
                full_body.extend(body);
                body = full_body;
            }
            if name == "constructor" {
                Ok(ClassMember::Constructor {
                    name: None,
                    params,
                    body,
                    base_args: None,
                    initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public })
            } else {
                if !is_generator {
                    is_generator = body_contains_yield(&body);
                }
                let wrapped = wrap_generator_if_needed(name, params, body, is_async, is_generator);
                if let StmtKind::FunctionDecl {
                    name,
                    params,
                    body,
                    is_async,
                    is_generator,
                    is_sub,
                    ..
                } = wrapped
                {
                    Ok(ClassMember::Method(Box::new(Statement::new(
                        StmtKind::FunctionDecl {
                            name,
                            params,
                            return_type: None,
                            body,
                            modifiers: Modifiers {
                                is_static,
                                ..Default::default()
                            },
                            handles: Vec::new(),
                            is_async,
                            is_generator,
                            is_sub },
                    ))))
                } else {
                    unreachable!()
                }
            }
        }
        Rule::class_property => {
            let mut name = String::new();
            let mut init = None;
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => {
                        name = extract_property_name(&p)
                    }
                    _ => init = Some(walk_expression(p)?) }
            }
            Ok(ClassMember::Field {
                name,
                type_hint: None,
                init,
                modifiers: Modifiers {
                    is_static,
                    ..Default::default()
                },
                with_events: false,
                array_bounds: None })
        }
        Rule::accessor_property => {
            // TC39 accessor auto-field: treat as a regular class field for now
            let mut name = String::new();
            let mut init = None;
            for p in member_pair.into_inner() {
                match p.as_rule() {
                    Rule::accessor_kw => {}
                    Rule::property_name | Rule::ident_name | Rule::ident_or_keyword => {
                        name = extract_property_name(&p)
                    }
                    _ => init = Some(walk_expression(p)?) }
            }
            Ok(ClassMember::Field {
                name,
                type_hint: None,
                init,
                modifiers: Modifiers {
                    is_static,
                    ..Default::default()
                },
                with_events: false,
                array_bounds: None })
        }
        other => Err(format!("Unexpected class member: {:?}", other)) }
}

// ── Control flow ────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let then_stmt = walk_statement(next_meaningful(&mut inner)?)?;
    // Skip NEWLINEs to find the optional else clause. The grammar's
    // eat_terminators between then and else may leave visible NEWLINE
    // tokens as siblings.
    let else_body = match next_meaningful(&mut inner) {
        Ok(p) => Some(vec![walk_statement(p)?]),
        Err(_) => None };
    Ok(StmtKind::If {
        cond,
        then_body: vec![then_stmt],
        elifs: Vec::new(),
        else_body })
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    // `for await (...)` — optional async marker between `for` and the
    // header. Captured by the grammar as a distinct `for_await_marker`
    // pair so we can route through `is_async = true` and emit `await`
    // before each body iteration.
    let mut is_for_await = false;
    let mut peek = inner.peek();
    if peek
        .as_ref()
        .map_or(false, |p| p.as_rule() == Rule::for_await_marker)
    {
        is_for_await = true;
        inner.next();
        peek = inner.peek();
        let _ = peek;
    }
    let header =
        next_rule(&mut inner, Rule::for_header).or_else(|_| next_meaningful(&mut inner))?;
    let header_inner = header.into_inner().next().ok_or("Empty for header")?;
    let body_pair = next_meaningful(&mut inner)?;
    let body = vec![walk_statement(body_pair)?];

    match header_inner.as_rule() {
        Rule::for_in_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let (var, prefix) = extract_for_target(&parts)?;
            let iter = walk_expression(
                parts
                    .into_iter()
                    .find(|p| {
                        !matches!(
                            p.as_rule(),
                            Rule::var_kind
                                | Rule::ident_name
                                | Rule::binding_pattern
                                | Rule::for_lhs_expr
                        )
                    })
                    .ok_or("missing iter expr")?,
            )?;
            let mut full_body = prefix;
            full_body.extend(body);
            Ok(StmtKind::ForIn {
                var,
                key: None,
                iter,
                body: full_body,
                of: false,
                else_body: None,
                is_async: is_for_await })
        }
        Rule::for_of_header => {
            // `for (using r of …)` — the "using" literal is anonymous in the
            // grammar, so detect it from the header text (no var_kind pair).
            let is_using_form = header_inner.as_str().trim_start().starts_with("using");
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let is_using_form =
                is_using_form && parts.iter().all(|p| p.as_rule() != Rule::var_kind);
            let is_let_const = parts
                .iter()
                .find(|p| p.as_rule() == Rule::var_kind)
                .map_or(false, |p| matches!(p.as_str(), "let" | "const"));
            let (var, prefix) = extract_for_target(&parts)?;
            let iter = walk_expression(
                parts
                    .into_iter()
                    .find(|p| {
                        !matches!(
                            p.as_rule(),
                            Rule::var_kind
                                | Rule::ident_name
                                | Rule::binding_pattern
                                | Rule::for_lhs_expr
                        )
                    })
                    .ok_or("missing iter expr")?,
            )?;
            let mut full_body = prefix;
            full_body.extend(body);
            // `for (using r of …)` disposes r at the end of every iteration
            // (TC39 explicit-resource-management §14.7.5).
            if is_using_form {
                full_body = using_disposal_wrap(&var, false, full_body);
            }
            // Per-iteration binding: wrap body in IIFE when const/let + closures present.
            let body_final = if is_let_const && body_contains_closure(&full_body, &[var.clone()]) {
                let params = vec![Param {
                    name: var.clone(),
                    type_hint: None,
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest: false,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false }];
                let args = vec![Argument::positional(Expression::ident(&var))];
                let iife = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(full_body),
                        is_async: false,
                        captures: Vec::new() })),
                    args,
                    optional: false });
                vec![Statement::new(StmtKind::Expr(iife))]
            } else {
                full_body
            };
            Ok(StmtKind::ForIn {
                var,
                key: None,
                iter,
                body: body_final,
                of: true,
                else_body: None,
                is_async: is_for_await })
        }
        Rule::for_c_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
            let mut init = None;
            let mut cond = None;
            let mut update = None;
            let mut let_vars: Vec<String> = Vec::new(); // track `let` loop vars

            for p in parts {
                match p.as_rule() {
                    Rule::for_c_init => {
                        let inner = p.into_inner().next().ok_or("Empty for init")?;
                        match inner.as_rule() {
                            Rule::variable_declaration_no_semi => {
                                let mut vi = inner.into_inner();
                                let kind_pair = next_rule(&mut vi, Rule::var_kind)?;
                                let var_kind = match kind_pair.as_str() {
                                    "var" => VarDeclKind::Var,
                                    "let" => VarDeclKind::Let,
                                    "const" => VarDeclKind::Const,
                                    _ => VarDeclKind::Let };
                                let mut decls = Vec::new();
                                for d in vi {
                                    if d.as_rule() == Rule::var_declarator {
                                        let decl = walk_var_declarator(d)?;
                                        if var_kind == VarDeclKind::Let
                                            || var_kind == VarDeclKind::Const
                                        {
                                            if let BindingPattern::Ident(ref name) = decl.pattern {
                                                let_vars.push(name.clone());
                                            }
                                        }
                                        decls.push(decl);
                                    }
                                }
                                init = Some(Box::new(Statement::new(StmtKind::VarDecl {
                                    declarations: decls,
                                    kind: var_kind })));
                            }
                            _ => {
                                let expr = walk_expression(inner)?;
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            }
                        }
                    }
                    Rule::expression => {
                        let expr = walk_expression(p)?;
                        // First expression seen is always the condition;
                        // second is always the update. The init is handled
                        // separately via Rule::for_c_init above.
                        if cond.is_none() {
                            cond = Some(expr);
                        } else {
                            update = Some(expr);
                        }
                    }
                    _ => {
                        // Try as expression
                        if let Ok(expr) = walk_expression(p) {
                            if cond.is_none() {
                                cond = Some(expr);
                            } else {
                                update = Some(expr);
                            }
                        }
                    }
                }
            }

            // Per-iteration `let` binding: wrap body in IIFE so closures
            // capture a fresh copy each iteration. Only apply when the body
            // contains function expressions/arrows that could close over the
            // loop variable — otherwise IIFE breaks break/continue.
            let body = if !let_vars.is_empty() && body_contains_closure(&body, &let_vars) {
                let params: Vec<Param> = let_vars
                    .iter()
                    .map(|v| Param {
                        name: v.clone(),
                        type_hint: None,
                        default: None,
                        pass_by: PassBy::Value,
                        is_rest: false,
                        is_kwargs: false,
                        is_optional: false,
                        is_nullable: false })
                    .collect();
                let args: Vec<Argument> = let_vars
                    .iter()
                    .map(|v| Argument::positional(Expression::ident(v)))
                    .collect();
                let iife = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(body),
                        is_async: false,
                        captures: Vec::new() })),
                    args,
                    optional: false });
                vec![Statement::new(StmtKind::Expr(iife))]
            } else {
                body
            };

            Ok(StmtKind::For {
                init,
                cond,
                update,
                body })
        }
        other => Err(format!("Unexpected for header: {:?}", other)) }
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::DoWhile {
        body,
        cond,
        until: false })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(next_meaningful(&mut inner)?)?;
    let mut cases = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::switch_case {
            // Grammar: switch_case = { ("case" expression | "default") ~ ":" ~ statements }
            // Detect default by looking at the source slice.
            // Default is emitted as a SwitchCase with empty conditions,
            // preserving its position among the other cases for fallthrough.
            let is_default = p.as_str().trim_start().starts_with("default");
            let mut case_inner = p.into_inner();
            if is_default {
                let stmts: Vec<Statement> = case_inner
                    .filter(|p| p.as_rule() != Rule::NEWLINE)
                    .map(walk_statement)
                    .collect::<Result<Vec<_>, _>>()?;
                cases.push(SwitchCase {
                    conditions: vec![], // empty = default
                    body: stmts });
            } else {
                let first = case_inner.next().ok_or("Empty switch case")?;
                let val = walk_expression(first)?;
                let stmts: Vec<Statement> = case_inner
                    .filter(|p| p.as_rule() != Rule::NEWLINE)
                    .map(walk_statement)
                    .collect::<Result<Vec<_>, _>>()?;
                cases.push(SwitchCase {
                    conditions: vec![CaseCondition::Value(val)],
                    body: stmts });
            }
        }
    }
    Ok(StmtKind::Switch {
        expr,
        cases,
        default: None })
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| !matches!(p.as_rule(), Rule::NEWLINE | Rule::return_kw))
        .map(walk_expression)
        .transpose()?;
    Ok(StmtKind::Return(expr))
}

fn walk_break(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let label = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string());
    Ok(StmtKind::Break(match label {
        Some(l) => BreakTarget::Label(l),
        None => BreakTarget::Implicit }))
}

fn walk_continue(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let label = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string());
    Ok(StmtKind::Continue(match label {
        Some(l) => ContinueTarget::Label(l),
        None => ContinueTarget::Implicit }))
}

fn walk_throw(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr_pair = pair
        .into_inner()
        .find(|p| !matches!(p.as_rule(), Rule::NEWLINE | Rule::throw_kw))
        .ok_or("throw: expected expression")?;
    let expr = walk_expression(expr_pair)?;
    Ok(StmtKind::Throw {
        expr: Some(expr),
        cause: None })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => body = walk_body_from_block(p)?,
            Rule::catch_clause => {
                let mut var_name = None;
                let mut catch_body = Vec::new();
                let mut destructure_prefix: Vec<Statement> = Vec::new();
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::ident_name => var_name = Some(cp.as_str().to_string()),
                        Rule::binding_pattern => {
                            // Destructuring catch: catch ({ message }) {}
                            // Desugar to: catch (__catch_tmp) { const { message } = __catch_tmp; }
                            let inner = cp.clone().into_inner().next();
                            match inner.as_ref().map(|p| p.as_rule()) {
                                Some(Rule::ident_name) => {
                                    var_name = Some(inner.unwrap().as_str().to_string());
                                }
                                _ => {
                                    let tmp = "__catch_tmp".to_string();
                                    var_name = Some(tmp.clone());
                                    let pattern = walk_binding_pattern(cp)?;
                                    destructure_prefix.push(Statement::new(StmtKind::VarDecl {
                                        declarations: vec![VarDeclarator {
                                            pattern,
                                            type_hint: None,
                                            init: Some(Expression::ident(&tmp)),
                                            array_bounds: None,
                                            with_events: false }],
                                        kind: VarDeclKind::Const }));
                                }
                            }
                        }
                        Rule::block_statement => {
                            catch_body = walk_body_from_block(cp)?;
                        }
                        _ => {}
                    }
                }
                if !destructure_prefix.is_empty() {
                    destructure_prefix.extend(catch_body);
                    catch_body = destructure_prefix;
                }
                catches.push(CatchClause {
                    types: Vec::new(),
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: None });
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_body_from_block(fp)?);
                    }
                }
            }
            _ => {}
        }
    }
    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally })
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let label = next_meaningful(&mut inner)?.as_str().to_string();
    let body = walk_statement(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::Labeled {
        label,
        body: Box::new(body) })
}

// ── Import / Export ─────────────────────────────────────────────────────────

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut source = String::new();
    let mut names = Vec::new();
    let mut default_name = None;
    let mut namespace_name = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::import_with => {} // ES2025 import attributes — ignored at AST level
            Rule::string_literal => source = unquote(p.as_str()),
            Rule::import_clause => {
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::default_import => {
                            default_name = Some(cp.as_str().to_string());
                        }
                        Rule::namespace_import => {
                            for np in cp.into_inner() {
                                if np.as_rule() == Rule::ident_name {
                                    namespace_name = Some(np.as_str().to_string());
                                }
                            }
                        }
                        Rule::named_imports => {
                            for sp in cp.into_inner() {
                                if sp.as_rule() == Rule::import_specifier {
                                    let mut parts = sp.into_inner();
                                    let first =
                                        parts.next().ok_or("import_specifier has no name")?;
                                    // ES2022: specifier name may be a string literal
                                    let name = if first.as_rule() == Rule::string_literal {
                                        unquote(first.as_str())
                                    } else {
                                        first.as_str().to_string()
                                    };
                                    let alias = parts.next().map(|p| p.as_str().to_string());
                                    names.push(ImportName { name, alias });
                                }
                            }
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let kind = if let Some(ns) = namespace_name {
        ImportKind::Wildcard {
            path: source,
            alias: Some(ns) }
    } else if let Some(def) = default_name {
        if names.is_empty() {
            ImportKind::Default {
                path: source,
                local: def }
        } else {
            // import default, { named } from "mod" — use Named with default as first
            let mut all_names = vec![ImportName {
                name: "default".into(),
                alias: Some(def) }];
            all_names.extend(names);
            ImportKind::Named {
                path: source,
                names: all_names,
                level: 0 }
        }
    } else if !names.is_empty() {
        ImportKind::Named {
            path: source,
            names,
            level: 0 }
    } else {
        ImportKind::Simple {
            path: source,
            alias: None }
    };

    Ok(Import { kind, span })
}

fn walk_export(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut declaration = None;
    let mut names = Vec::new();
    let mut default_expr = None;
    let mut from: Option<String> = None;
    let mut star = false;
    let mut star_alias: Option<String> = None;

    // Detect `export * [as n] from "m"` by looking at raw source — the
    // `*` token isn't captured as its own pair because pest matches
    // it as a literal in the rule. Scan the raw string for leading `*`.
    let raw = pair.as_str();
    let trimmed = raw.trim_start_matches("export").trim_start();
    if trimmed.starts_with('*') {
        star = true;
    }

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::import_with => {} // ES2025 import attributes — ignored at AST level
            Rule::function_declaration
            | Rule::async_function_declaration
            | Rule::class_declaration
            | Rule::variable_declaration => {
                declaration = Some(Box::new(walk_statement(p)?));
            }
            Rule::export_specifier => {
                let mut parts = p.into_inner();
                let first = parts.next().ok_or("export_specifier has no name")?;
                // ES2022: specifier name may be string literal
                let name = if first.as_rule() == Rule::string_literal {
                    unquote(first.as_str())
                } else {
                    first.as_str().to_string()
                };
                let alias = parts.next().map(|p| {
                    if p.as_rule() == Rule::string_literal {
                        unquote(p.as_str())
                    } else {
                        p.as_str().to_string()
                    }
                });
                names.push(ExportName { name, alias });
            }
            Rule::string_literal => {
                // The `from "m"` clause — a re-export source.
                from = Some(unquote(p.as_str()));
            }
            Rule::ident_name => {
                // `export * as n from "m"` — `n` captured as ident_name.
                if star {
                    star_alias = Some(p.as_str().to_string());
                }
            }
            _ => {
                // default expression
                if let Ok(expr) = walk_expression(p) {
                    default_expr = Some(Box::new(expr));
                }
            }
        }
    }

    // `export * as n from "m"` — expose the whole namespace under
    // local name `n`. Lower as a single ExportName with
    // `name = "*"` so the Linker recognizes the star-as-namespace
    // shape.
    if star {
        if let Some(n) = star_alias {
            names.push(ExportName {
                name: "*".into(),
                alias: Some(n) });
        }
    }

    Ok(StmtKind::Export {
        declaration,
        names,
        default: default_expr,
        from,
        star })
}

// ── Expressions ─────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(collapse_passthrough_expression(pair)?)?;
    Ok(normalize_optional_chain(Expression::with_span(kind, span)))
}

// ── Optional chain normalization — ECMA-262 §13.3 ───────────────────────────
//
// Once `?.` short-circuits on a nullish base, the ENTIRE remaining chain is
// skipped (member reads, index expressions, call arguments — no side
// effects). Normalize each chain whose spine contains an optional link into
// the spec's guard shape, splitting at the link nearest the top:
//
//   obj?.prop[sideEffect()]
//   → ((p) => p === null || p === undefined ? undefined : p.prop[sideEffect()])(obj)
//
// `delete obj?.prop` gets the same split with `true` as the short-circuit
// value and a real delete in the live branch (§13.5.1.2). Inner chains in
// arguments / indices are normalized by their own walk_expression calls;
// remaining optional links below the split point are handled by the
// recursive call on the head.

static OPTCHAIN_COUNTER: AtomicUsize = AtomicUsize::new(0);

/// Split the spine at the optional link nearest the top. Returns
/// `(head, rebuilt)` where `rebuilt` is the chain with that link made
/// non-optional and rooted at `Ident(param)`.
fn split_optional_spine(
    e: Expression,
    param: &str,
) -> Result<(Expression, Expression), Expression> {
    let span = e.span;
    match e.kind {
        ExprKind::Member {
            object,
            field,
            null_safe } => {
            if null_safe {
                let rebuilt = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(param)),
                    field,
                    null_safe: false });
                return Ok((*object, rebuilt));
            }
            match split_optional_spine(*object, param) {
                Ok((head, rebuilt_obj)) => Ok((
                    head,
                    Expression::new(ExprKind::Member {
                        object: Box::new(rebuilt_obj),
                        field,
                        null_safe: false }),
                )),
                Err(object) => Err(Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(object),
                        field,
                        null_safe: false },
                    span,
                )) }
        }
        ExprKind::Index {
            object,
            index,
            null_safe } => {
            if null_safe {
                let rebuilt = Expression::new(ExprKind::Index {
                    object: Box::new(Expression::ident(param)),
                    index,
                    null_safe: false });
                return Ok((*object, rebuilt));
            }
            match split_optional_spine(*object, param) {
                Ok((head, rebuilt_obj)) => Ok((
                    head,
                    Expression::new(ExprKind::Index {
                        object: Box::new(rebuilt_obj),
                        index,
                        null_safe: false }),
                )),
                Err(object) => Err(Expression::with_span(
                    ExprKind::Index {
                        object: Box::new(object),
                        index,
                        null_safe: false },
                    span,
                )) }
        }
        ExprKind::Call {
            callee,
            args,
            optional } => {
            // A *method* call (member/index callee) must keep its receiver
            // binding — splitting between the receiver and the call would
            // rebind `this`. Leave those to the compiler's existing
            // single-link handling, optional or not.
            let is_method_call = matches!(
                &callee.kind,
                ExprKind::Member { .. } | ExprKind::Index { .. }
            );
            if optional && !is_method_call {
                let rebuilt = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(param)),
                    args,
                    optional: false });
                return Ok((*callee, rebuilt));
            }
            if is_method_call {
                return Err(Expression::with_span(
                    ExprKind::Call {
                        callee,
                        args,
                        optional },
                    span,
                ));
            }
            match split_optional_spine(*callee, param) {
                Ok((head, rebuilt_callee)) => Ok((
                    head,
                    Expression::new(ExprKind::Call {
                        callee: Box::new(rebuilt_callee),
                        args,
                        optional }),
                )),
                Err(callee) => Err(Expression::with_span(
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args,
                        optional },
                    span,
                )) }
        }
        kind => Err(Expression::with_span(kind, span)) }
}

/// Does the member/index/call spine of this expression contain an optional
/// link? (Arguments and index expressions are separate chains — ignored.)
fn spine_has_optional(e: &Expression) -> bool {
    match &e.kind {
        ExprKind::Member {
            object, null_safe, ..
        } => *null_safe || spine_has_optional(object),
        ExprKind::Index {
            object, null_safe, ..
        } => *null_safe || spine_has_optional(object),
        ExprKind::Call {
            callee, optional, ..
        } => *optional || spine_has_optional(callee),
        _ => false }
}

fn normalize_optional_chain(e: Expression) -> Expression {
    if !spine_has_optional(&e) {
        return e;
    }
    let n = OPTCHAIN_COUNTER.fetch_add(1, Ordering::Relaxed);
    let param = format!("__vybe_optc_{n}");
    match split_optional_spine(e, &param) {
        Ok((head, rebuilt)) => {
            let head = normalize_optional_chain(head);
            optional_guard_iife(
                &param,
                head,
                rebuilt,
                Expression::new(ExprKind::Lit(Literal::Undefined)),
            )
        }
        Err(unchanged) => unchanged }
}

/// Is this expression the optional-guard IIFE `normalize_optional_chain`
/// produces? Used by the §13.15.1 invalid-assignment-target early error.
fn is_optional_chain_guard(e: &Expression) -> bool {
    let ExprKind::Call { callee, .. } = &e.kind else {
        return false;
    };
    let ExprKind::Lambda { params, .. } = &callee.kind else {
        return false;
    };
    params.len() == 1 && params[0].name.starts_with("__vybe_optc_")
}

/// Recognize the optional-guard IIFE produced by `normalize_optional_chain`
/// and rewrite it for `delete` semantics: short-circuit value becomes `true`,
/// the live chain gets a real `Delete`.
fn rewrite_optional_delete(operand: &Expression) -> Option<ExprKind> {
    let ExprKind::Call {
        callee,
        args,
        optional: false } = &operand.kind
    else {
        return None;
    };
    let ExprKind::Lambda {
        params,
        body: LambdaBody::Expr(body),
        ..
    } = &callee.kind
    else {
        return None;
    };
    if params.len() != 1 || !params[0].name.starts_with("__vybe_optc_") {
        return None;
    }
    let ExprKind::Ternary { cond, then, else_ } = &body.kind else {
        return None;
    };
    if !matches!(then.kind, ExprKind::Lit(Literal::Undefined)) {
        return None;
    }
    let live = Expression::new(ExprKind::Delete(else_.clone()));
    let guarded = Expression::new(ExprKind::Ternary {
        cond: cond.clone(),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        else_: Box::new(live) });
    Some(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: params.clone(),
            body: LambdaBody::Expr(Box::new(guarded)),
            is_async: false,
            captures: Vec::new() })),
        args: args.clone(),
        optional: false })
}

/// `((param) => param === null || param === undefined ? <short> : <live>)(head)`
fn optional_guard_iife(
    param: &str,
    head: Expression,
    live: Expression,
    short: Expression,
) -> Expression {
    let nullish = Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::ident(param)),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))) })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::StrictEq,
            left: Box::new(Expression::ident(param)),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))) })) });
    let body = Expression::new(ExprKind::Ternary {
        cond: Box::new(nullish),
        then: Box::new(short),
        else_: Box::new(live) });
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![Param {
                name: param.to_string(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false }],
            body: LambdaBody::Expr(Box::new(body)),
            is_async: false,
            captures: Vec::new() })),
        args: vec![Argument::positional(head)],
        optional: false })
}

fn collapse_passthrough_expression(mut pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    loop {
        let next = match pair.as_rule() {
            Rule::expression => {
                let mut inner = pair
                    .clone()
                    .into_inner()
                    .filter(|p| p.as_rule() != Rule::NEWLINE);
                match (inner.next(), inner.next()) {
                    (Some(first), None) => Some(first),
                    _ => None }
            }
            Rule::assignment_expression
            | Rule::conditional_expression
            | Rule::logical_expr
            | Rule::comparison
            | Rule::additive
            | Rule::multiplicative
            | Rule::call_chain
            | Rule::property_name
            | Rule::computed_property_name => {
                let mut inner = pair
                    .clone()
                    .into_inner()
                    .filter(|p| p.as_rule() != Rule::NEWLINE);
                match (inner.next(), inner.next()) {
                    (Some(first), None) => Some(first),
                    _ => None }
            }
            Rule::primary => match pair.as_str().trim() {
                "true" | "false" | "null" | "undefined" | "this" | "super" => None,
                _ => {
                    let mut inner = pair.clone().into_inner();
                    match (inner.next(), inner.next()) {
                        (Some(first), None) => Some(first),
                        _ => None }
                }
            },
            Rule::unary => {
                let mut inner = pair.clone().into_inner();
                match (inner.next(), inner.next()) {
                    (Some(first), None) if first.as_rule() == Rule::postfix => Some(first),
                    _ => None }
            }
            Rule::postfix => {
                let mut inner = pair.clone().into_inner();
                let first = inner.next();
                let has_postfix = inner.any(|p| p.as_rule() == Rule::postfix_op);
                match (first, has_postfix) {
                    (Some(first), false) => Some(first),
                    _ => None }
            }
            Rule::call_expression => {
                let mut inner = pair.clone().into_inner();
                let first = inner.next();
                let has_chain = inner.any(|p| p.as_rule() == Rule::call_chain);
                match (first, has_chain) {
                    (Some(first), false) => Some(first),
                    _ => None }
            }
            _ => None };

        match next {
            Some(next_pair) => pair = next_pair,
            None => return Ok(pair) }
    }
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // Literals
        Rule::bigint_literal => {
            let raw = pair.as_str();
            // strip `_` separators and trailing `n`
            let s_owned: String = raw.chars().filter(|c| *c != '_').collect();
            let s = s_owned.trim_end_matches('n');
            let parsed = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).ok()
            } else if s.starts_with("0o") || s.starts_with("0O") {
                i64::from_str_radix(&s[2..], 8).ok()
            } else if s.starts_with("0b") || s.starts_with("0B") {
                i64::from_str_radix(&s[2..], 2).ok()
            } else {
                s.parse().ok()
            };
            match parsed {
                Some(n) => Ok(ExprKind::Lit(Literal::BigInt(n))),
                // BigInt literals are ARBITRARY precision (§12.9.3) — a
                // literal past i64 normalizes to the spec-identical
                // `BigInt("digits")` constructor form (walker's job:
                // normalize syntax; the host parses exactly).
                None => Ok(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Ident("BigInt".into()))),
                    args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                        Literal::Str(s.to_string()),
                    )))],
                    optional: false }) }
        }
        Rule::numeric_literal => {
            // ES2021 numeric separator: strip `_` from digits before parsing
            let raw = pair.as_str();
            let s_owned: String = raw.chars().filter(|c| *c != '_').collect();
            let s = s_owned.as_str();
            if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?,
                )))
            } else if s.starts_with("0o") || s.starts_with("0O") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 8).map_err(|e| format!("{}", e))?,
                )))
            } else if s.starts_with("0b") || s.starts_with("0B") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?,
                )))
            } else if s.contains('.') || s.contains('e') || s.contains('E') {
                Ok(ExprKind::Lit(Literal::Float(
                    s.parse().map_err(|e| format!("{}", e))?,
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }
        Rule::string_literal => Ok(ExprKind::Lit(Literal::Str(unquote(pair.as_str())))),
        Rule::regex_literal => Ok(walk_regex_literal(pair.as_str())),
        Rule::import_meta => Ok(ExprKind::Ident("__js_import_meta".to_string())),
        Rule::new_target => Ok(ExprKind::Ident("__js_new_target".to_string())),
        Rule::dynamic_import => {
            let args = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::argument_list)
                .map(walk_arguments)
                .transpose()?
                .unwrap_or_default();
            Ok(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Ident(
                    "__js_dynamic_import".to_string(),
                ))),
                args,
                optional: false })
        }
        Rule::private_name => Ok(ExprKind::Ident(pair.as_str().to_string())),
        Rule::ident_name | Rule::ident_or_keyword => {
            let name = pair.as_str();
            match name {
                "true" => Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => Ok(ExprKind::Lit(Literal::Null)),
                "undefined" => Ok(ExprKind::Lit(Literal::Undefined)),
                "this" => Ok(ExprKind::This),
                "super" => Ok(ExprKind::Super),
                _ => Ok(ExprKind::Ident(name.to_string())) }
        }
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::undefined_kw => Ok(ExprKind::Lit(Literal::Undefined)),
        Rule::this_kw => Ok(ExprKind::This),
        Rule::super_kw => Ok(ExprKind::Super),

        // Sequence (comma expression)
        Rule::expression => {
            let mut inner: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() != Rule::NEWLINE)
                .collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else {
                let exprs: Vec<Expression> = inner
                    .into_iter()
                    .map(walk_expression)
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(ExprKind::Sequence(exprs))
            }
        }

        // Assignment
        Rule::assignment_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() == 3 {
                let left = walk_expression(inner.remove(0))?;
                // §13.15.1 early error: an optional chain is not a valid
                // assignment target. After normalization an optional-chain
                // target shows up as the guard IIFE.
                if is_optional_chain_guard(&left) {
                    return Err("Invalid left-hand side in assignment".to_string());
                }
                let op_str = inner.remove(0).as_str();
                let right = walk_expression(inner.remove(0))?;
                if op_str == "=" {
                    Ok(ExprKind::Assign {
                        target: Box::new(left),
                        value: Box::new(right) })
                } else {
                    // Compound assign — but this is expression level, wrap as assign
                    let op = match op_str {
                        "+=" => CompoundOp::Add,
                        "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul,
                        "/=" => CompoundOp::Div,
                        "%=" => CompoundOp::Mod,
                        "**=" => CompoundOp::Pow,
                        "&=" => CompoundOp::BitAnd,
                        "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor,
                        "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr,
                        ">>>=" => CompoundOp::UShr,
                        "&&=" => CompoundOp::And,
                        "||=" => CompoundOp::Or,
                        "??=" => CompoundOp::NullCoalesce,
                        _ => CompoundOp::Add };
                    Ok(ExprKind::Assign {
                        target: Box::new(left.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: compound_to_binop(op),
                            left: Box::new(left),
                            right: Box::new(right) })) })
                }
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // Ternary
        Rule::conditional_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() == 3 {
                let cond = walk_expression(inner.remove(0))?;
                let then = walk_expression(inner.remove(0))?;
                let else_ = walk_expression(inner.remove(0))?;
                Ok(ExprKind::Ternary {
                    cond: Box::new(cond),
                    then: Box::new(then),
                    else_: Box::new(else_) })
            } else {
                walk_expr_kind(inner.remove(0))
            }
        }

        // Binary chains
        Rule::logical_expr | Rule::comparison | Rule::additive | Rule::multiplicative => {
            walk_binary_chain(pair)
        }

        // Unary
        Rule::unary => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty unary")?;
            // If it's a postfix (no unary_op), delegate
            if first.as_rule() == Rule::postfix {
                return walk_expr_kind(first);
            }
            // unary_op ~ unary
            let op_str = first.as_str().trim();
            let operand = walk_expression(inner.next().ok_or("Missing unary operand")?)?;
            if op_str.starts_with("typeof") {
                return Ok(ExprKind::TypeOf(Box::new(operand)));
            }
            if op_str.starts_with("void") {
                return Ok(ExprKind::Void(Box::new(operand)));
            }
            if op_str.starts_with("delete") {
                // `delete varName` — deleting a bare variable always returns false
                // (var/let/const bindings are non-configurable). Only member/index
                // delete goes through the runtime property-deletion path.
                if matches!(operand.kind, ExprKind::Ident(_)) {
                    return Ok(ExprKind::Lit(vybe_ast::Literal::Bool(false)));
                }
                // `delete obj?.prop` — §13.5.1.2: a short-circuited chain
                // deletes nothing and yields true. walk_expression already
                // normalized the chain into the optional-guard IIFE; rewrite
                // its branches for delete semantics.
                if let Some(rewritten) = rewrite_optional_delete(&operand) {
                    return Ok(rewritten);
                }
                return Ok(ExprKind::Delete(Box::new(operand)));
            }
            if op_str.starts_with("await") {
                return Ok(ExprKind::Await(Box::new(operand)));
            }
            let op = match op_str {
                "-" => UnaryOp::Neg,
                "+" => UnaryOp::Pos,
                "!" => UnaryOp::Not,
                "~" => UnaryOp::BitNot,
                "++" => UnaryOp::PreInc,
                "--" => UnaryOp::PreDec,
                _ => UnaryOp::Neg };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand) })
        }

        // Postfix
        Rule::postfix => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let base = walk_expression(inner.remove(0))?;
            // Check for postfix_op (++/--)
            let has_postfix = inner.iter().any(|p| p.as_rule() == Rule::postfix_op);
            if !has_postfix {
                return Ok(base.kind);
            }
            let op_pair = inner
                .iter()
                .find(|p| p.as_rule() == Rule::postfix_op)
                .unwrap();
            let op = match op_pair.as_str() {
                "++" => UnaryOp::PostInc,
                "--" => UnaryOp::PostDec,
                _ => return Ok(base.kind) };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(base) })
        }

        // Call / member / index chain
        Rule::call_expression => walk_call_chain(pair),
        Rule::new_expression => {
            // new_expression = { "new" ~ primary ~ call_chain* }
            // Per JS spec: the FIRST `()` after `new` is the constructor args.
            // Any subsequent member/call/index chains are applied to the RESULT
            // of the construction (e.g. `new Foo().bar().baz`). The
            // word-boundary check happens at `call_expression`'s
            // `&new_keyword_lookahead` gate — see the grammar comment for
            // why the lookahead lives there instead of inside this rule.
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty new")?;
            let mut expr = walk_expression(first)?;
            let chains: Vec<Pair<Rule>> =
                inner.filter(|p| p.as_rule() == Rule::call_chain).collect();
            let mut new_consumed = false; // True after the first `(args)` is processed

            for chain in chains {
                let chain_src = chain.as_str().trim_start();
                let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

                if !new_consumed && chain_src.starts_with("(") {
                    // First parens — these are the constructor args.
                    let args = if let Some(arg_pair) = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::argument_list)
                    {
                        walk_arguments(arg_pair)?
                    } else {
                        Vec::new()
                    };
                    expr = Expression::new(ExprKind::New {
                        class: Box::new(expr),
                        args });
                    new_consumed = true;
                } else if !new_consumed && chain_src.starts_with(".") {
                    // Member access BEFORE constructor args: `new Foo.Bar(42)`.
                    let name = chain_inner
                        .into_iter()
                        .find(|p| {
                            p.as_rule() == Rule::ident_or_keyword || p.as_rule() == Rule::ident_name
                        })
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: false });
                } else {
                    // Chain AFTER `new X(...)` — applied to the constructed object.
                    // Handle: `(args)` call, `.member`, `?.member`, `?.()`, `[idx]`, tagged template.
                    if chain_src.starts_with("?.") {
                        if chain_inner.first().map_or(false, |p| {
                            p.as_rule() == Rule::argument_list || p.as_str().starts_with("(")
                        }) {
                            let args = if let Some(arg_pair) = chain_inner
                                .into_iter()
                                .find(|p| p.as_rule() == Rule::argument_list)
                            {
                                walk_arguments(arg_pair)?
                            } else {
                                Vec::new()
                            };
                            expr = Expression::new(ExprKind::Call {
                                callee: Box::new(expr),
                                args,
                                optional: true });
                        } else {
                            let name = chain_inner
                                .into_iter()
                                .find(|p| {
                                    p.as_rule() == Rule::ident_or_keyword
                                        || p.as_rule() == Rule::ident_name
                                        || p.as_rule() == Rule::private_name
                                })
                                .map(|p| p.as_str().to_string())
                                .unwrap_or_default();
                            expr = Expression::new(ExprKind::Member {
                                object: Box::new(expr),
                                field: name,
                                null_safe: true });
                        }
                    } else if chain_src.starts_with("(") {
                        let args = if let Some(arg_pair) = chain_inner
                            .into_iter()
                            .find(|p| p.as_rule() == Rule::argument_list)
                        {
                            walk_arguments(arg_pair)?
                        } else {
                            Vec::new()
                        };
                        expr = Expression::new(ExprKind::Call {
                            callee: Box::new(expr),
                            args,
                            optional: false });
                    } else if chain_src.starts_with(".") {
                        let name = chain_inner
                            .into_iter()
                            .find(|p| {
                                p.as_rule() == Rule::ident_or_keyword
                                    || p.as_rule() == Rule::ident_name
                                    || p.as_rule() == Rule::private_name
                            })
                            .map(|p| p.as_str().to_string())
                            .unwrap_or_default();
                        expr = canonicalize_member_access(expr, &name);
                    } else if chain_src.starts_with("[") {
                        let index_expr = chain_inner
                            .into_iter()
                            .find(|p| {
                                p.as_rule() == Rule::expression
                                    || matches!(
                                        p.as_rule(),
                                        Rule::assignment_expression
                                            | Rule::conditional_expression
                                            | Rule::ident_name
                                            | Rule::numeric_literal
                                            | Rule::string_literal
                                    )
                            })
                            .map(walk_expression)
                            .transpose()?
                            .unwrap_or(Expression::new(ExprKind::Lit(Literal::Int(0))));
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(index_expr),
                            null_safe: false });
                    }
                }
            }

            // If `new` had no `()` (e.g., `new X`), wrap the bare class.
            if !new_consumed {
                expr = Expression::new(ExprKind::New {
                    class: Box::new(expr),
                    args: Vec::new() });
            }
            Ok(expr.kind)
        }

        // Primary
        Rule::primary => {
            // Keyword literals (true/false/null/undefined/this/super) don't produce
            // inner pairs in pest — they're anonymous literals. Check as_str() first.
            let src = pair.as_str().trim();
            match src {
                "true" => return Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => return Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => return Ok(ExprKind::Lit(Literal::Null)),
                "undefined" => return Ok(ExprKind::Lit(Literal::Undefined)),
                "this" => return Ok(ExprKind::This),
                "super" => return Ok(ExprKind::Super),
                _ => {}
            }
            let inner = pair.into_inner().next().ok_or("Empty primary")?;
            walk_expr_kind(inner)
        }

        // Arrow functions
        Rule::yield_expression => {
            let mut inner = pair.into_inner();
            let mut is_yield_from = false;
            let mut value: Option<Expression> = None;
            while let Some(p) = inner.next() {
                match p.as_rule() {
                    Rule::yield_kw => {}
                    Rule::yield_delegate => {
                        is_yield_from = true;
                    }
                    _ if p.as_str() == "*" => {
                        is_yield_from = true;
                    }
                    _ => {
                        value = Some(walk_expression(p)?);
                    }
                }
            }
            if is_yield_from {
                Ok(ExprKind::YieldFrom(Box::new(
                    value.unwrap_or(Expression::null()),
                )))
            } else {
                Ok(ExprKind::Yield(value.map(Box::new)))
            }
        }
        Rule::arrow_function | Rule::async_arrow_function => {
            let is_async = pair.as_rule() == Rule::async_arrow_function;
            let pair_src = pair.as_str().trim_start();
            let mut params = Vec::new();
            let mut body =
                LambdaBody::Expr(Box::new(Expression::new(ExprKind::Lit(Literal::Null))));
            let mut param_prologue = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::ident_name if !pair_src.starts_with('(') => {
                        params = vec![Param {
                            name: p.as_str().to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false }]
                    }
                    Rule::param_list => {
                        let (parsed_params, prologue) = walk_params_with_prologue(p)?;
                        params = parsed_params;
                        param_prologue = prologue;
                    }
                    Rule::arrow_body => {
                        let inner = p.into_inner().next().ok_or("Empty arrow body")?;
                        body = match inner.as_rule() {
                            Rule::function_body => LambdaBody::Block(walk_body(inner)?),
                            _ => LambdaBody::Expr(Box::new(walk_expression(inner)?)) };
                    }
                    Rule::function_body => body = LambdaBody::Block(walk_body(p)?),
                    Rule::async_kw => {}
                    _ => {
                        // Could be direct expression or function_body
                        if let Ok(stmts) = walk_body(p.clone()) {
                            body = LambdaBody::Block(stmts);
                        } else {
                            body = LambdaBody::Expr(Box::new(walk_expression(p)?));
                        }
                    }
                }
            }
            // §15.3.1 early error: arrow parameter lists may NEVER
            // contain duplicate names (unlike sloppy-mode functions).
            {
                let mut seen = std::collections::HashSet::new();
                for param in &params {
                    if !param.name.starts_with("__") && !seen.insert(param.name.as_str()) {
                        return Err(format!(
                            "Duplicate parameter name not allowed in this context: '{}'",
                            param.name
                        ));
                    }
                }
            }
            if !param_prologue.is_empty() {
                body = match body {
                    LambdaBody::Expr(expr) => {
                        let mut stmts = param_prologue;
                        stmts.push(Statement::new(StmtKind::Return(Some(*expr))));
                        LambdaBody::Block(stmts)
                    }
                    LambdaBody::Block(stmts) => {
                        let mut full_body = param_prologue;
                        full_body.extend(stmts);
                        LambdaBody::Block(full_body)
                    }
                };
            }
            Ok(ExprKind::Lambda {
                params,
                body,
                is_async,
                captures: Vec::new() })
        }

        // Function expression
        Rule::function_expression | Rule::async_function_expression => {
            let mut stmt_kind = walk_func_decl(pair)?;
            // §15.2.5: a *named* function expression binds its name only
            // inside its own scope — recursion sees it, the enclosing scope
            // does not. Normalize to an IIFE holding the binding as a local
            // const (a nested FunctionDecl would register module-globally):
            //   (() => { const name = function (…) {…}; return name; })()
            // The const binding also drives the existing `.name` inference
            // for anonymous functions, so fn.name stays the source name.
            if let StmtKind::FunctionDecl { name, .. } = &mut stmt_kind {
                if !name.is_empty() {
                    let fn_name = std::mem::take(name);
                    let init = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
                        stmt_kind,
                    ))));
                    let decl = Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(fn_name.clone()),
                            type_hint: None,
                            init: Some(init),
                            array_bounds: None,
                            with_events: false }],
                        kind: VarDeclKind::Const });
                    let ret = Statement::new(StmtKind::Return(Some(Expression::ident(&fn_name))));
                    return Ok(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Lambda {
                            params: Vec::new(),
                            body: LambdaBody::Block(vec![decl, ret]),
                            is_async: false,
                            captures: Vec::new() })),
                        args: Vec::new(),
                        optional: false });
                }
            }
            Ok(ExprKind::FunctionExpr(Box::new(Statement::new(stmt_kind))))
        }

        // Class expression
        Rule::class_expression => {
            let mut name = None;
            let mut parent = None;
            let mut members = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::ident_name => name = Some(p.as_str().to_string()),
                    Rule::class_body => {
                        for m in p.into_inner() {
                            if m.as_rule() == Rule::class_member {
                                members.push(walk_class_member(m)?);
                            }
                        }
                    }
                    _ => parent = Some(Box::new(walk_expression(p)?)) }
            }
            Ok(ExprKind::ClassExpr {
                name,
                parent,
                interfaces: Vec::new(),
                members })
        }

        // Array literal
        Rule::array_literal => {
            let elements = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::array_slot)
                .map(|p| {
                    let mut inner = p.into_inner();
                    let Some(inner) = inner.next() else {
                        return Ok(js_array_elision_marker());
                    };
                    if inner.as_rule() == Rule::array_elision {
                        return Ok(js_array_elision_marker());
                    }
                    let src = inner.as_str();
                    let spread = src.trim_start().starts_with("...");
                    let value_pair = inner
                        .into_inner()
                        .next()
                        .ok_or("Empty array element".to_string())?;
                    let value = walk_expression(value_pair)?;
                    Ok(ArrayElement {
                        key: None,
                        value,
                        spread,
                        by_ref: false })
                })
                .collect::<Result<Vec<_>, String>>()?;
            Ok(ExprKind::Array(elements))
        }

        // Object literal
        Rule::object_literal => {
            let props = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::object_property)
                .map(walk_object_property)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(ExprKind::Object(props))
        }

        // Template literal
        Rule::template_literal => {
            let mut parts = Vec::new();
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::template_full => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(unescape_template(&s[1..s.len() - 1])));
                    }
                    Rule::template_head => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(unescape_template(&s[1..s.len() - 2])));
                    }
                    Rule::template_middle => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(unescape_template(&s[1..s.len() - 2])));
                    }
                    Rule::template_tail => {
                        let s = p.as_str();
                        parts.push(InterpolPart::Text(unescape_template(&s[1..s.len() - 1])));
                    }
                    _ => parts.push(InterpolPart::Expr(walk_expression(p)?)) }
            }
            Ok(ExprKind::Interpolation(parts))
        }

        // Spread
        Rule::argument => {
            let src = pair.as_str();
            let spread = src.trim_start().starts_with("...");
            let inner = pair.into_inner().next().ok_or("Empty argument")?;
            let expr = walk_expression(inner)?;
            if spread {
                Ok(ExprKind::Spread(Box::new(expr)))
            } else {
                Ok(expr.kind)
            }
        }

        // Passthrough wrappers
        Rule::call_chain | Rule::property_name | Rule::computed_property_name => {
            let inner = pair.into_inner().next().ok_or("Empty wrapper")?;
            walk_expr_kind(inner)
        }

        other => Err(format!("Unexpected expression rule: {:?}", other)) }
}

// ── AST-level normalization helpers ────────────────────────────────────────

/// Normalize `typeof x === "typename"` (and the commuted form) to
/// `ExprKind::IsType { expr: x, type_name: "typename" }`.
///
/// This is the ECMA-262 `typeof` type-guard pattern. Normalizing it at the
/// AST level means the IsType compiler arm handles it exactly the same way
/// as Python's `isinstance(x, str)` or VB's `TypeOf x Is String` — all of
/// which map to the same cross-language `IsType` node and produce
/// `Value::Bool` from the compiler (not raw `i32`).
///
/// Only normalizes the primitive types where the VM's `REF_IS_*` opcodes are
/// authoritative: "string", "number", "boolean", "undefined".
/// "function" is left as-is because `REF_IS_FUNC` misses `HostFunction`
/// objects — those are resolved correctly by `ecma:value.typeof` at runtime.
/// "object" is also left as-is because `typeof null === "object"` requires
/// the spec-precise host fn.
fn normalize_typeof_strict_eq(expr: Expression) -> Expression {
    if let ExprKind::Binary {
        op: BinOp::StrictEq,
        ref left,
        ref right } = expr.kind
    {
        // typeof x === "typename"
        if let ExprKind::TypeOf(inner) = &left.kind {
            if let ExprKind::Lit(vybe_ast::Literal::Str(typename)) = &right.kind {
                if matches!(
                    typename.as_str(),
                    "string" | "number" | "boolean" | "undefined" | "bigint" | "symbol"
                ) {
                    return Expression::new(ExprKind::IsType {
                        expr: inner.clone(),
                        type_name: typename.clone() });
                }
            }
        }
        // "typename" === typeof x  (commuted)
        if let ExprKind::Lit(vybe_ast::Literal::Str(typename)) = &left.kind {
            if let ExprKind::TypeOf(inner) = &right.kind {
                if matches!(
                    typename.as_str(),
                    "string" | "number" | "boolean" | "undefined" | "bigint" | "symbol"
                ) {
                    return Expression::new(ExprKind::IsType {
                        expr: inner.clone(),
                        type_name: typename.clone() });
                }
            }
        }
    }
    expr
}

// ── Binary chain walker ─────────────────────────────────────────────────────

fn walk_binary_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let _rule = pair.as_rule();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    // First operand
    let mut left = walk_expression(inner.remove(0))?;

    // Remaining: (op, operand) pairs
    let mut i = 0;
    while i + 1 < inner.len() {
        let op_pair = &inner[i];
        let op = match op_pair.as_rule() {
            Rule::logical_op
            | Rule::nullish_op
            | Rule::or_op
            | Rule::and_op
            | Rule::bitor_op
            | Rule::bitxor_op
            | Rule::bitand_op
            | Rule::comparison_op
            | Rule::equality_op
            | Rule::relational_op
            | Rule::shift_op
            | Rule::additive_op
            | Rule::multiplicative_op
            | Rule::mul_op
            | Rule::exp_op => op_pair.as_str().trim(),
            _ => op_pair.as_str().trim() };
        let right = walk_expression(inner[i + 1].clone())?;

        let bin_op = match op {
            "??" => BinOp::NullCoalesce,
            "||" => BinOp::Or,
            "&&" => BinOp::And,
            "|" => BinOp::BitOr,
            "^" => BinOp::BitXor,
            "&" => BinOp::BitAnd,
            "===" => BinOp::StrictEq,
            "!==" => BinOp::StrictNotEq,
            "==" => BinOp::Eq,
            "!=" => BinOp::NotEq,
            "<" => BinOp::Lt,
            ">" => BinOp::Gt,
            "<=" => BinOp::LtEq,
            ">=" => BinOp::GtEq,
            "instanceof" => BinOp::InstanceOf,
            "in" => BinOp::In,
            ">>>" => BinOp::UShr,
            ">>" => BinOp::Shr,
            "<<" => BinOp::Shl,
            "+" => BinOp::Add,
            "-" => BinOp::Sub,
            "*" => BinOp::Mul,
            "/" => BinOp::Div,
            "%" => BinOp::Mod,
            "**" => BinOp::Pow,
            _ => BinOp::Add };

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right) });

        // Normalize: `typeof x === "typename"` / `"typename" === typeof x`
        // → `IsType { expr: x, type_name: "typename" }`.
        // This is cross-language normalization: the ECMA typeof-guard pattern
        // maps to the same `IsType` AST node as Python's `isinstance` or
        // VB's `TypeOf x Is T`. The IsType compiler arm then produces
        // `Value::Bool` (not raw i32) for correct ECMA display semantics.
        left = normalize_typeof_strict_eq(left);

        i += 2;
    }

    Ok(left.kind)
}

// ── Call chain walker ───────────────────────────────────────────────────────

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty call expression")?;
    let mut expr = walk_expression(first)?;

    for chain in inner {
        if chain.as_rule() != Rule::call_chain {
            continue;
        }
        let chain_src = chain.as_str().trim_start();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with("?.") {
            // Optional chaining
            if chain_src.starts_with("?.[") {
                let index_expr = chain_inner
                    .into_iter()
                    .find(|p| {
                        p.as_rule() == Rule::expression
                            || matches!(
                                p.as_rule(),
                                Rule::assignment_expression
                                    | Rule::conditional_expression
                                    | Rule::ident_name
                                    | Rule::numeric_literal
                                    | Rule::string_literal
                            )
                    })
                    .map(walk_expression)
                    .transpose()?
                    .unwrap_or(Expression::new(ExprKind::Lit(Literal::Int(0))));
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index_expr),
                    null_safe: true });
            } else {
                // Detect optional call: ?.(...) — chain_inner may be empty (no args) or contain argument_list.
                // Use chain_src to detect the "(" after "?." since grammar literals aren't in chain_inner.
                let is_optional_call = chain_src.starts_with("?.(")
                    || chain_inner
                        .first()
                        .map_or(false, |p| p.as_rule() == Rule::argument_list);
                if is_optional_call {
                    let args = if let Some(arg_pair) = chain_inner
                        .into_iter()
                        .find(|p| p.as_rule() == Rule::argument_list)
                    {
                        walk_arguments(arg_pair)?
                    } else {
                        Vec::new()
                    };
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args,
                        optional: true });
                } else {
                    let name = chain_inner
                        .into_iter()
                        .find(|p| {
                            p.as_rule() == Rule::ident_or_keyword
                                || p.as_rule() == Rule::ident_name
                                || p.as_rule() == Rule::private_name
                        })
                        .map(|p| p.as_str().to_string())
                        .unwrap_or_default();
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: name,
                        null_safe: true });
                }
            }
        } else if chain_src.starts_with("(") {
            // Call
            let args = if let Some(arg_pair) = chain_inner
                .into_iter()
                .find(|p| p.as_rule() == Rule::argument_list)
            {
                walk_arguments(arg_pair)?
            } else {
                Vec::new()
            };
            expr = Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args,
                optional: false });
        } else if chain_src.starts_with(".") {
            // Member access — normalize JS .length to canonical __len__
            let name = chain_inner
                .into_iter()
                .find(|p| {
                    p.as_rule() == Rule::ident_or_keyword
                        || p.as_rule() == Rule::ident_name
                        || p.as_rule() == Rule::private_name
                })
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            expr = canonicalize_member_access(expr, &name);
        } else if chain_src.starts_with("[") {
            // Computed / index
            let index_expr = chain_inner
                .into_iter()
                .find(|p| {
                    p.as_rule() == Rule::expression
                        || matches!(
                            p.as_rule(),
                            Rule::assignment_expression
                                | Rule::conditional_expression
                                | Rule::ident_name
                                | Rule::numeric_literal
                                | Rule::string_literal
                        )
                })
                .map(walk_expression)
                .transpose()?
                .unwrap_or(Expression::new(ExprKind::Lit(Literal::Int(0))));
            // If the index is a well-known Symbol, lower to a Member access so
            // obj[Symbol.iterator]() compiles as a normal method call.
            if let Some(alias) = js_well_known_symbol_alias(&index_expr) {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: alias.to_string(),
                    null_safe: false });
            } else {
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index_expr),
                    null_safe: false });
            }
        } else if chain_src.starts_with("`") {
            // Tagged template: tag`parts...${expr}...`
            // Desugar to: tag(Object.assign(cooked, {raw: [raw...]}), expr0, expr1, ...)
            if let Some(tmpl) = chain_inner
                .into_iter()
                .find(|p| p.as_rule() == Rule::template_literal)
            {
                let (parts, raw_parts, exprs) = walk_template_parts(tmpl)?;
                let mut args: Vec<Argument> = Vec::new();
                let make_str_array = |ss: Vec<String>| {
                    Expression::new(ExprKind::Array(
                        ss.into_iter()
                            .map(|s| ArrayElement {
                                key: None,
                                value: Expression::new(ExprKind::Lit(Literal::Str(s))),
                                spread: false,
                                by_ref: false })
                            .collect(),
                    ))
                };
                let cooked_array = make_str_array(parts);
                let raw_array = make_str_array(raw_parts);
                // Object.assign(cooked, { raw: raw }) — sets .raw and returns the array
                let raw_obj = Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: Expression::string("raw"),
                    value: raw_array }]));
                let strings_with_raw = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Object")),
                        field: "assign".into(),
                        null_safe: false })),
                    args: vec![
                        Argument::positional(cooked_array),
                        Argument::positional(raw_obj),
                    ],
                    optional: false });
                // ECMA-262 §13.2.8.3: template objects are cached per call site.
                // Wrap in `__vybe_tmpl_N ?? (__vybe_tmpl_N = Object.assign(...))` so
                // the same object is returned on every invocation of this template site.
                let tmpl_id = TEMPLATE_COUNTER.fetch_add(1, Ordering::Relaxed);
                let tmpl_global = format!("__vybe_tmpl_{}", tmpl_id);
                let cached_template = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(Expression::ident(&tmpl_global)),
                    right: Box::new(Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::ident(&tmpl_global)),
                        value: Box::new(strings_with_raw) })) });
                args.push(Argument::positional(cached_template));
                for e in exprs {
                    args.push(Argument::positional(e));
                }
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args,
                    optional: false });
            }
        }
    }

    // Normalize variadic concat: `x.concat(a, b, c)` → `x.concat(a).concat(b).concat(c)`
    // The stdlib concat function is binary (receiver + 1 arg). For variadic calls,
    // desugar into a chain of binary concat calls. Works for both strings and arrays.
    expr = desugar_variadic_concat(expr);

    // Normalize `Array.isArray(x)` → `IsType { expr: x, type_name: "array" }`.
    // Keeps ECMA type-guard patterns at the AST level so the IsType compiler
    // arm can produce `Value::Bool` (not raw `i32` from `opcode:ref_is_array`).
    if let ExprKind::Call {
        ref callee,
        ref args,
        optional: false } = expr.kind
    {
        if let ExprKind::Member {
            ref object,
            ref field,
            null_safe: false } = callee.kind
        {
            if let ExprKind::Ident(ref name) = object.kind {
                if name == "Array" && field == "isArray" && args.len() == 1 {
                    let arg = args[0].value.clone();
                    return Ok(ExprKind::IsType {
                        expr: Box::new(arg),
                        type_name: "array".to_string() });
                }
            }
        }
    }

    Ok(expr.kind)
}

/// Check if a for-loop body contains closures (lambdas or function expressions)
/// that reference any of the given `let` variable names. Used to decide whether
/// to wrap the body in an IIFE for per-iteration binding.
fn body_contains_closure(stmts: &[Statement], _vars: &[String]) -> bool {
    // Simple heuristic: check if any lambda/function expression exists in the body.
    // A more precise check would verify the lambda references a let-var, but
    // the simple check is correct — IIFE is safe when there ARE closures, and
    // we skip it when there are none (preserving break/continue).
    fn has_closure_expr(expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) => true,
            ExprKind::Call { callee, args, .. } => {
                has_closure_expr(callee) || args.iter().any(|a| has_closure_expr(&a.value))
            }
            ExprKind::Member { object, .. } => has_closure_expr(object),
            ExprKind::Binary { left, right, .. } => {
                has_closure_expr(left) || has_closure_expr(right)
            }
            ExprKind::Unary { expr, .. } => has_closure_expr(expr),
            ExprKind::Ternary {
                cond, then, else_, ..
            } => has_closure_expr(cond) || has_closure_expr(then) || has_closure_expr(else_),
            ExprKind::Array(elems) => elems.iter().any(|e| has_closure_expr(&e.value)),
            ExprKind::Index { object, index, .. } => {
                has_closure_expr(object) || has_closure_expr(index)
            }
            ExprKind::Assign { target: _, value } => has_closure_expr(value),
            _ => false }
    }
    fn has_closure_stmt(stmt: &Statement) -> bool {
        match &stmt.kind {
            StmtKind::Expr(e) => has_closure_expr(e),
            StmtKind::VarDecl { declarations, .. } => declarations
                .iter()
                .any(|d| d.init.as_ref().map_or(false, |e| has_closure_expr(e))),
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
                ..
            } => {
                has_closure_expr(cond)
                    || then_body.iter().any(has_closure_stmt)
                    || elifs
                        .iter()
                        .any(|(c, b)| has_closure_expr(c) || b.iter().any(has_closure_stmt))
                    || else_body
                        .as_ref()
                        .map_or(false, |b| b.iter().any(has_closure_stmt))
            }
            StmtKind::Block(stmts) => stmts.iter().any(has_closure_stmt),
            StmtKind::Return(Some(e)) => has_closure_expr(e),
            _ => false }
    }
    stmts.iter().any(has_closure_stmt)
}

/// Desugar `x.concat(a, b, c)` into `x.concat(a).concat(b).concat(c)`.
fn desugar_variadic_concat(expr: Expression) -> Expression {
    if let ExprKind::Call {
        ref callee,
        ref args,
        optional } = expr.kind
    {
        if args.len() > 1 {
            if let ExprKind::Member {
                ref object,
                ref field,
                null_safe } = callee.kind
            {
                if field == "concat" {
                    // Chain: start with receiver.concat(args[0]), then .concat(args[1]), etc.
                    let mut result = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "concat".to_string(),
                            null_safe })),
                        args: vec![args[0].clone()],
                        optional });
                    for arg in &args[1..] {
                        result = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(result),
                                field: "concat".to_string(),
                                null_safe: false })),
                            args: vec![arg.clone()],
                            optional: false });
                    }
                    return result;
                }
            }
        }
    }
    expr
}

/// Walk a template_literal into (cooked_parts, raw_parts, expressions).
/// cooked has escape sequences processed; raw is the literal source text.
fn walk_template_parts(
    pair: Pair<Rule>,
) -> Result<(Vec<String>, Vec<String>, Vec<Expression>), String> {
    let mut cooked: Vec<String> = Vec::new();
    let mut raw: Vec<String> = Vec::new();
    let mut exprs: Vec<Expression> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::template_full => {
                let s = p.as_str();
                let inner = &s[1..s.len() - 1];
                raw.push(inner.to_string());
                cooked.push(unescape_template(inner));
            }
            Rule::template_head => {
                let s = p.as_str();
                let inner = &s[1..s.len() - 2];
                raw.push(inner.to_string());
                cooked.push(unescape_template(inner));
            }
            Rule::template_middle => {
                let s = p.as_str();
                let inner = &s[1..s.len() - 2];
                raw.push(inner.to_string());
                cooked.push(unescape_template(inner));
            }
            Rule::template_tail => {
                let s = p.as_str();
                let inner = &s[1..s.len() - 1];
                raw.push(inner.to_string());
                cooked.push(unescape_template(inner));
            }
            _ => {
                exprs.push(walk_expression(p)?);
            }
        }
    }
    Ok((cooked, raw, exprs))
}

fn unescape_template(s: &str) -> String {
    let mut result = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();
    while let Some(c) = chars.next() {
        if c != '\\' {
            result.push(c);
            continue;
        }
        match chars.peek().copied() {
            Some('n') => {
                chars.next();
                result.push('\n');
            }
            Some('t') => {
                chars.next();
                result.push('\t');
            }
            Some('r') => {
                chars.next();
                result.push('\r');
            }
            Some('0') => {
                chars.next();
                result.push('\0');
            }
            Some('\\') => {
                chars.next();
                result.push('\\');
            }
            Some('`') => {
                chars.next();
                result.push('`');
            }
            Some('$') => {
                chars.next();
                result.push('$');
            }
            Some('u') => {
                chars.next();
                // \u{HHHH} or \uHHHH
                let hex: String = if chars.peek() == Some(&'{') {
                    chars.next();
                    let h: String = chars.by_ref().take_while(|&ch| ch != '}').collect();
                    h
                } else {
                    chars.by_ref().take(4).collect()
                };
                if let Ok(n) = u32::from_str_radix(&hex, 16) {
                    // §11.8.4: surrogate-pair escapes combine into one
                    // supplementary code point.
                    if hex.len() == 4 && (0xD800..=0xDBFF).contains(&n) {
                        let mut probe = chars.clone();
                        if probe.next() == Some('\\') && probe.next() == Some('u') {
                            let lo_hex: String = probe.by_ref().take(4).collect();
                            if lo_hex.len() == 4 {
                                if let Ok(lo) = u32::from_str_radix(&lo_hex, 16) {
                                    if (0xDC00..=0xDFFF).contains(&lo) {
                                        let cp = 0x10000 + ((n - 0xD800) << 10) + (lo - 0xDC00);
                                        if let Some(ch) = char::from_u32(cp) {
                                            for _ in 0..6 {
                                                chars.next();
                                            }
                                            result.push(ch);
                                            continue;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if let Some(ch) = char::from_u32(n) {
                        result.push(ch);
                        continue;
                    }
                    // Unpaired surrogate escape → U+FFFD (see unquote).
                    if (0xD800..=0xDFFF).contains(&n) {
                        result.push('\u{FFFD}');
                        continue;
                    }
                }
                result.push('\\');
                result.push('u');
                result.push_str(&hex);
            }
            Some('x') => {
                chars.next();
                let hex: String = chars.by_ref().take(2).collect();
                if let Ok(n) = u8::from_str_radix(&hex, 16) {
                    result.push(n as char);
                } else {
                    result.push('\\');
                    result.push('x');
                    result.push_str(&hex);
                }
            }
            _ => result.push('\\') }
    }
    result
}

/// Canonicalize JS property access to unified AST representation.
/// `arr.length` → `Call(__len__, [arr])`
///
/// Note: only `.length` is normalized — `.size` is too generic in JS (could be a custom property).
fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    let canonical = match name {
        "length" => Some("__len__"),
        _ => None };
    if let Some(canonical_name) = canonical {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(canonical_name)),
            args: vec![Argument::positional(object)],
            optional: false })
    } else {
        Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false })
    }
}

fn js_well_known_symbol_alias_from_raw(name: &str) -> Option<&'static str> {
    match name {
        "Symbol.iterator" => Some("iterator"),
        "Symbol.asyncIterator" => Some("asyncIterator"),
        "Symbol.toPrimitive" => Some("toprimitive"),
        "Symbol.hasInstance" => Some("hasinstance"),
        "Symbol.toStringTag" => Some("tostringtag"),
        "Symbol.isConcatSpreadable" => Some("isconcatspreadable"),
        "Symbol.species" => Some("species"),
        "Symbol.match" => Some("symbolmatch"),
        "Symbol.matchAll" => Some("symbolmatchall"),
        "Symbol.replace" => Some("symbolreplace"),
        "Symbol.search" => Some("symbolsearch"),
        "Symbol.split" => Some("symbolsplit"),
        "Symbol.unscopables" => Some("unscopables"),
        _ => None }
}

fn js_well_known_symbol_alias(expr: &Expression) -> Option<&'static str> {
    let ExprKind::Member {
        object,
        field,
        null_safe } = &expr.kind
    else {
        return None;
    };
    if *null_safe {
        return None;
    }
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    if name != "Symbol" {
        return None;
    }
    let raw = format!("Symbol.{}", field);
    js_well_known_symbol_alias_from_raw(&raw)
}

// JS method call canonicalization is intentionally minimal:
// Methods like .toString() may be overridden on user classes, so we leave them as
// regular method calls and let the compiler dispatch via the class method binding.
// Only true builtin operations like .length (handled in canonicalize_member_access)
// are normalized to canonical builtins.

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::argument)
        .map(|p| {
            let spread = p.as_str().trim_start().starts_with("...");
            let inner = p.into_inner().next().ok_or("Empty argument".to_string())?;
            let value = walk_expression(inner)?;
            Ok(Argument {
                value,
                name: None,
                by_ref: false,
                spread })
        })
        .collect()
}

// ── Object property walker ──────────────────────────────────────────────────

fn walk_object_property(pair: Pair<Rule>) -> Result<ObjectProperty, String> {
    let src = pair.as_str().trim();
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    // Spread: { ...expr }
    if src.starts_with("...") {
        let expr = walk_expression(inner.remove(0))?;
        return Ok(ObjectProperty::Spread(expr));
    }

    // Computed: { [expr]: value }
    if inner
        .first()
        .map_or(false, |p| p.as_rule() == Rule::computed_property_name)
    {
        let key_pair = inner.remove(0);
        let key = walk_expression(key_pair.into_inner().next().ok_or("Empty computed key")?)?;
        let value = walk_expression(inner.remove(0))?;
        if let Some(alias) = js_well_known_symbol_alias(&key) {
            return Ok(ObjectProperty::KeyValue {
                key: Expression::string(alias),
                value });
        }
        return Ok(ObjectProperty::Computed { key, value });
    }

    // Method: { name() {} } or getter/setter
    if inner.len() >= 2 {
        let has_body = inner.iter().any(|p| p.as_rule() == Rule::function_body);
        if has_body {
            let trimmed = src.trim_start();
            let is_getter = trimmed.starts_with("get ") || trimmed.starts_with("get\t");
            let is_setter = trimmed.starts_with("set ") || trimmed.starts_with("set\t");
            if is_getter || is_setter {
                return walk_object_accessor(inner, is_getter);
            }
            return walk_object_method(inner);
        }
    }

    // Key: value or shorthand
    if inner.len() == 1 {
        return Ok(ObjectProperty::Shorthand(
            inner.remove(0).as_str().to_string(),
        ));
    }

    if inner.len() >= 2 {
        let key_pair = inner.remove(0);
        // Object keys: identifiers become string literals (JS object keys are always strings)
        let key = match key_pair.as_rule() {
            Rule::ident_name | Rule::ident_or_keyword | Rule::property_name => {
                let key_str = key_pair.as_str().to_string();
                // property_name may contain inner pairs (string/number/ident) — extract
                if let Some(inner_pair) = key_pair.into_inner().next() {
                    match inner_pair.as_rule() {
                        Rule::string_literal => walk_expression(inner_pair)?,
                        Rule::numeric_literal => walk_expression(inner_pair)?,
                        _ => Expression::string(&key_str) }
                } else {
                    Expression::string(&key_str)
                }
            }
            _ => walk_expression(key_pair)? };
        let value = walk_expression(inner.remove(0))?;
        return Ok(ObjectProperty::KeyValue { key, value });
    }

    Err("Could not parse object property".into())
}

// ── Helpers ─────────────────────────────────────────────────────────────────

fn to_span(pair: &Pair<Rule>) -> Span {
    let start = pair.as_span().start_pos().line_col();
    let end = pair.as_span().end_pos().line_col();
    Span {
        start_line: start.0 as u32 - 1,
        start_col: start.1 as u32 - 1,
        end_line: end.0 as u32 - 1,
        end_col: end.1 as u32 - 1 }
}

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    pair.into_inner()
        .find(|p| p.as_rule() != Rule::NEWLINE)
        .ok_or_else(|| "Expected inner pair".into())
}

fn next_meaningful<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if p.as_rule() != Rule::NEWLINE {
            return Ok(p);
        }
    }
    Err("Expected next pair".into())
}

fn next_rule<'a>(
    pairs: &mut impl Iterator<Item = Pair<'a, Rule>>,
    rule: Rule,
) -> Result<Pair<'a, Rule>, String> {
    for p in pairs {
        if p.as_rule() == rule {
            return Ok(p);
        }
    }
    Err(format!("Expected {:?}", rule))
}

fn walk_body_from_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() != Rule::NEWLINE)
        .map(walk_statement)
        .collect()
}

fn extract_ident_name(pair: &Pair<Rule>) -> String {
    pair.as_str().trim().to_string()
}

/// Resolve a `property_name` pair into a method/property name string.
/// Computed names like `[Symbol.iterator]` are recognised when the
/// expression is a known well-known-symbol member access — the
/// canonical resolver picks up `Symbol.iterator` / `Symbol.hasInstance`
/// / etc. and remaps to the cross-language method names. Other
/// computed names fall through as the raw bracketed text (caller can
/// detect and either lower to runtime install or error).
fn extract_property_name(pair: &Pair<Rule>) -> String {
    if pair.as_rule() == Rule::property_name {
        if let Some(inner) = pair.clone().into_inner().next() {
            if inner.as_rule() == Rule::computed_property_name {
                let inner_text = inner
                    .as_str()
                    .trim_start_matches('[')
                    .trim_end_matches(']')
                    .trim();
                if let Some(rest) = inner_text.strip_prefix("Symbol.") {
                    return format!("Symbol.{}", rest.trim());
                }
                return inner_text.to_string();
            }
        }
    }
    pair.as_str().trim().to_string()
}

/// Extract the loop variable and any destructuring prefix statements.
/// For `for (let x of arr)` returns ("x", []).
/// For `for (let [a, b] of arr)` returns ("__forof_tmp", [VarDecl let [a,b] = __forof_tmp])
fn extract_for_target(parts: &[Pair<Rule>]) -> Result<(String, Vec<Statement>), String> {
    let mut var_kind = VarDeclKind::Let;
    for p in parts {
        match p.as_rule() {
            Rule::var_kind => {
                var_kind = match p.as_str() {
                    "var" => VarDeclKind::Var,
                    "const" => VarDeclKind::Const,
                    _ => VarDeclKind::Let };
            }
            Rule::ident_name => {
                return Ok((p.as_str().to_string(), Vec::new()));
            }
            Rule::for_lhs_expr => {
                // Member/computed LHS: `for (obj.x in arr)` — walk as expression,
                // produce a synthetic assignment target name for the ForIn AST node.
                // The compiler will emit a store to the member at runtime.
                let expr_text = p.as_str().to_string();
                return Ok((expr_text, Vec::new()));
            }
            Rule::binding_pattern => {
                let inner = p
                    .clone()
                    .into_inner()
                    .next()
                    .ok_or("Empty binding pattern")?;
                if inner.as_rule() == Rule::ident_name {
                    return Ok((inner.as_str().to_string(), Vec::new()));
                }
                // Destructuring pattern — desugar to: let __forof_tmp; let [...] = __forof_tmp
                let pattern = walk_binding_pattern(p.clone())?;
                let tmp = "__forof_tmp".to_string();
                let prefix = Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern,
                        type_hint: None,
                        init: Some(Expression::ident(&tmp)),
                        array_bounds: None,
                        with_events: false }],
                    kind: var_kind });
                return Ok((tmp, vec![prefix]));
            }
            _ => continue }
    }
    Err("Expected identifier or binding pattern in for target".into())
}

/// `get name() {}` / `set name(v) {}` shorthand inside object literals.
/// Stored as a `__get_<name>` / `__set_<name>` synthetic key so the VM's
/// STRUCT_GET / STRUCT_SET accessor dispatch fires. A `this` param is
/// prepended so the body's `this` refs resolve via local-slot lookup
/// (the VM's getter dispatch passes the receiver as arg 0). Defined
/// out-of-line so walk_object_property's stack frame stays small.
fn walk_object_accessor(
    mut inner: Vec<Pair<Rule>>,
    is_getter: bool,
) -> Result<ObjectProperty, String> {
    let prop_pair = inner.remove(0);
    let prop_name = if prop_pair.as_rule() == Rule::computed_property_name {
        prop_pair
            .into_inner()
            .next()
            .map(|p| p.as_str().to_string())
            .unwrap_or_default()
    } else if prop_pair.as_rule() == Rule::property_name {
        let raw = prop_pair.as_str().to_string();
        let mut inner_pairs = prop_pair.into_inner();
        if let Some(inner) = inner_pairs.next() {
            if inner.as_rule() == Rule::computed_property_name {
                inner
                    .into_inner()
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or(raw)
            } else {
                raw
            }
        } else {
            raw
        }
    } else {
        prop_pair.as_str().to_string()
    };
    let mut params = Vec::new();
    let mut prologue = Vec::new();
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::param_list => (params, prologue) = walk_params_with_prologue(p)?,
            Rule::param => {
                let (param, init_stmt) = walk_param_with_prologue(p, 0)?;
                params = vec![param];
                prologue = init_stmt.into_iter().collect();
            }
            Rule::function_body => body = walk_body(p)?,
            _ => {}
        }
    }
    if !prologue.is_empty() {
        prologue.extend(body);
        body = prologue;
    }
    let mut full_params = vec![Param {
        name: "this".to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false }];
    full_params.extend(params);
    let storage_key = if is_getter {
        format!("__get_{}", prop_name)
    } else {
        format!("__set_{}", prop_name)
    };
    Ok(ObjectProperty::KeyValue {
        key: Expression::string(&storage_key),
        value: Expression::new(ExprKind::Lambda {
            params: full_params,
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new() }) })
}

/// Method shorthand `{ name() {} }` — emit as a key/value with a
/// FunctionDecl-wrapped lambda. Out-of-line for the same stack-frame
/// reason as walk_object_accessor.
fn walk_object_method(mut inner: Vec<Pair<Rule>>) -> Result<ObjectProperty, String> {
    let mut is_async = false;
    let mut has_generator_marker = false;
    if inner.first().is_some_and(|p| p.as_rule() == Rule::async_kw) {
        is_async = true;
        inner.remove(0);
    }
    if inner
        .first()
        .is_some_and(|p| p.as_rule() == Rule::generator_marker)
    {
        has_generator_marker = true;
        inner.remove(0);
    }
    let key_pair = inner.remove(0);

    // Detect computed method shorthand: `[expr]() {}` — key_pair is a
    // property_name whose inner is a computed_property_name.
    let computed_expr = if key_pair.as_rule() == Rule::property_name {
        if let Some(inner_p) = key_pair.clone().into_inner().next() {
            if inner_p.as_rule() == Rule::computed_property_name {
                // Peek at the raw text to see if it's a well-known Symbol alias
                let raw = inner_p
                    .as_str()
                    .trim_start_matches('[')
                    .trim_end_matches(']')
                    .trim();
                if js_well_known_symbol_alias_from_raw(raw).is_none() {
                    // Not a well-known symbol — treat key as a runtime expression
                    let key_inner = inner_p.into_inner().next().ok_or("Empty computed key")?;
                    Some(walk_expression(key_inner)?)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        }
    } else {
        None
    };

    // `[Symbol.iterator]() {…}` — rewrite to the canonical
    // cross-language method name (`iterator` / `toprimitive` / etc.)
    // so the iter-drain polyfill and to_primitive polyfill find the
    // method via the same key class declarations use.
    let raw_key = if key_pair.as_rule() == Rule::property_name {
        extract_property_name(&key_pair)
    } else {
        key_pair.as_str().to_string()
    };
    let key = js_well_known_symbol_alias_from_raw(&raw_key)
        .map(str::to_string)
        .unwrap_or(raw_key);
    let mut params = Vec::new();
    let mut prologue = Vec::new();
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::async_kw => is_async = true,
            Rule::param_list => (params, prologue) = walk_params_with_prologue(p)?,
            Rule::param => {
                let (param, init_stmt) = walk_param_with_prologue(p, 0)?;
                params = vec![param];
                prologue = init_stmt.into_iter().collect();
            }
            Rule::function_body => body = walk_body(p)?,
            _ => {}
        }
    }
    if !prologue.is_empty() {
        prologue.extend(body);
        body = prologue;
    }
    let is_generator = has_generator_marker || body_contains_yield(&body);

    // Computed method: return Computed { key: runtime_expr, value: function }.
    // A computed-key method is still a *method* — dynamic `this` per
    // §15.4.4 MethodDefinitionEvaluation — so emit a function expression,
    // not a Lambda (Lambdas compile with arrow-style lexical `this`).
    if let Some(key_expr) = computed_expr {
        let wrapped = wrap_generator_if_needed(String::new(), params, body, is_async, is_generator);
        let func = Statement::new(wrapped);
        return Ok(ObjectProperty::Computed {
            key: key_expr,
            value: Expression::new(ExprKind::FunctionExpr(Box::new(func))) });
    }

    let wrapped = wrap_generator_if_needed(key.clone(), params, body, is_async, is_generator);
    let func = Statement::new(wrapped);
    Ok(ObjectProperty::Method {
        key,
        value: Box::new(func) })
}

/// Translate a regex literal source `/pattern/flags` into the AST shape
/// `new RegExp("pattern", "flags")`. Defined out-of-line so the walker's
/// big match doesn't carry the construction's locals on every recursion
/// step (debug-build stack frames are ~bytes-per-arm sensitive).
fn walk_regex_literal(raw: &str) -> ExprKind {
    let (pattern, flags) = match raw
        .strip_prefix('/')
        .and_then(|s| s.rfind('/').map(|i| (&s[..i], &s[i + 1..])))
    {
        Some((p, f)) => (unescape_regex_literal_pattern(p), f.to_string()),
        None => (raw.to_string(), String::new()) };
    ExprKind::New {
        class: Box::new(Expression::ident("RegExp")),
        args: vec![
            Argument::positional(Expression::string(&pattern)),
            Argument::positional(Expression::string(&flags)),
        ] }
}

fn unescape_regex_literal_pattern(pattern: &str) -> String {
    let mut out = String::with_capacity(pattern.len());
    let mut chars = pattern.chars();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }
        match chars.next() {
            Some('/') => out.push('/'),
            Some(next) => {
                out.push('\\');
                out.push(next);
            }
            None => out.push('\\') }
    }
    out
}

fn unquote(s: &str) -> String {
    if s.len() < 2 {
        return s.to_string();
    }
    let inner = &s[1..s.len() - 1];
    // Single-pass escape processing — chained `replace` is wrong
    // because the second pass can re-process literal characters that
    // were already produced (e.g. `"\\n"` → first replace turns `\\`
    // into `\` leaving `\n` which the next replace then turns into
    // newline, losing the user's literal `\` + `n` input).
    let mut out = String::with_capacity(inner.len());
    let mut chars = inner.chars();
    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }
        match chars.next() {
            // ECMA-262 §12.8.4 SingleEscapeCharacter
            Some('n') => out.push('\n'),
            Some('t') => out.push('\t'),
            Some('r') => out.push('\r'),
            Some('b') => out.push('\x08'),
            Some('f') => out.push('\x0c'),
            Some('v') => out.push('\x0b'),
            Some('0') => out.push('\0'),
            Some('\'') => out.push('\''),
            Some('"') => out.push('"'),
            Some('\\') => out.push('\\'),
            Some('`') => out.push('`'),
            Some('$') => out.push('$'),
            // §12.8.4 HexEscapeSequence: \xHH
            Some('x') => {
                let hi = chars.next();
                let lo = chars.next();
                match (hi, lo) {
                    (Some(h), Some(l)) => {
                        let mut buf = [0u8; 2];
                        buf[0] = h as u8;
                        buf[1] = l as u8;
                        let s = std::str::from_utf8(&buf).unwrap_or("");
                        if let Ok(n) = u32::from_str_radix(s, 16) {
                            if let Some(c) = char::from_u32(n) {
                                out.push(c);
                                continue;
                            }
                        }
                        out.push('\\');
                        out.push('x');
                        if let Some(h) = hi {
                            out.push(h);
                        }
                        if let Some(l) = lo {
                            out.push(l);
                        }
                    }
                    _ => out.push('\\') }
            }
            // §12.8.4 UnicodeEscapeSequence: \uHHHH or \u{...}
            Some('u') => {
                let mut peek_iter = chars.clone();
                if peek_iter.next() == Some('{') {
                    chars.next(); // consume '{'
                    let mut hex = String::new();
                    while let Some(h) = chars.clone().next() {
                        if h == '}' {
                            chars.next();
                            break;
                        }
                        if h.is_ascii_hexdigit() {
                            hex.push(h);
                            chars.next();
                        } else {
                            break;
                        }
                    }
                    if let Ok(n) = u32::from_str_radix(&hex, 16) {
                        if let Some(c) = char::from_u32(n) {
                            out.push(c);
                            continue;
                        }
                    }
                    out.push('\\');
                    out.push('u');
                    out.push('{');
                    out.push_str(&hex);
                } else {
                    let mut hex = String::new();
                    for _ in 0..4 {
                        if let Some(h) = chars.clone().next() {
                            if h.is_ascii_hexdigit() {
                                hex.push(h);
                                chars.next();
                            } else {
                                break;
                            }
                        }
                    }
                    if hex.len() == 4 {
                        if let Ok(n) = u32::from_str_radix(&hex, 16) {
                            // §11.8.4: two \uHHHH escapes forming a surrogate
                            // pair encode one supplementary code point.
                            if (0xD800..=0xDBFF).contains(&n) {
                                let mut probe = chars.clone();
                                if probe.next() == Some('\\') && probe.next() == Some('u') {
                                    let mut lo_hex = String::new();
                                    for _ in 0..4 {
                                        match probe.next() {
                                            Some(h) if h.is_ascii_hexdigit() => lo_hex.push(h),
                                            _ => break }
                                    }
                                    if lo_hex.len() == 4 {
                                        if let Ok(lo) = u32::from_str_radix(&lo_hex, 16) {
                                            if (0xDC00..=0xDFFF).contains(&lo) {
                                                let cp =
                                                    0x10000 + ((n - 0xD800) << 10) + (lo - 0xDC00);
                                                if let Some(c) = char::from_u32(cp) {
                                                    for _ in 0..6 {
                                                        chars.next();
                                                    }
                                                    out.push(c);
                                                    continue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            if let Some(c) = char::from_u32(n) {
                                out.push(c);
                                continue;
                            }
                            // Unpaired surrogate escape: UTF-8 storage
                            // cannot hold it — U+FFFD is the closest
                            // faithful value (what toWellFormed yields).
                            if (0xD800..=0xDFFF).contains(&n) {
                                out.push('\u{FFFD}');
                                continue;
                            }
                        }
                    }
                    out.push('\\');
                    out.push('u');
                    out.push_str(&hex);
                }
            }
            Some(other) => {
                out.push('\\');
                out.push(other);
            }
            None => out.push('\\') }
    }
    out
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
        CompoundOp::UShr => BinOp::UShr,
        CompoundOp::And => BinOp::And,
        CompoundOp::Or => BinOp::Or,
        CompoundOp::NullCoalesce => BinOp::NullCoalesce,
        CompoundOp::IDiv => BinOp::IDiv,
        CompoundOp::Concat => BinOp::Concat }
}

fn rewrite_this_in_generator_body(stmts: &mut [Statement]) {
    fn rwe(e: &mut Expression) {
        match &mut e.kind {
            ExprKind::This => {
                e.kind = ExprKind::Ident("__self_captured".to_string());
            }
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => {
                // Scope boundary — separate this context
            }
            ExprKind::RefOf(place) => match place.as_mut() {
                PlaceExpr::Ident(_) => {}
                PlaceExpr::Member { object, .. } => rwe(object),
                PlaceExpr::Index { object, index, .. } => {
                    rwe(object);
                    rwe(index);
                }
                PlaceExpr::Deref(expr) => rwe(expr) },
            ExprKind::Unary { expr: i, .. }
            | ExprKind::RefLoad(i)
            | ExprKind::IsType { expr: i, .. }
            | ExprKind::Cast { expr: i, .. }
            | ExprKind::TypeOf(i)
            | ExprKind::Spread(i)
            | ExprKind::Await(i)
            | ExprKind::Void(i)
            | ExprKind::Delete(i) => rwe(i),
            ExprKind::Binary {
                left: a, right: b, ..
            }
            | ExprKind::NullCoalesce { left: a, right: b }
            | ExprKind::Assign {
                target: a,
                value: b }
            | ExprKind::Walrus {
                target: a,
                value: b }
            | ExprKind::Range {
                start: a, end: b, ..
            } => {
                rwe(a);
                rwe(b);
            }
            ExprKind::StaticAccess {
                class: a,
                member: b } => {
                rwe(a);
                rwe(b);
            }
            ExprKind::Ternary { cond, then, else_ } => {
                rwe(cond);
                rwe(then);
                rwe(else_);
            }
            ExprKind::Member { object, .. } => rwe(object),
            ExprKind::Index { object, index, .. } => {
                rwe(object);
                rwe(index);
            }
            ExprKind::Call { callee, args, .. } => {
                rwe(callee);
                for arg in args {
                    rwe(&mut arg.value);
                }
            }
            ExprKind::New { class, args } => {
                rwe(class);
                for arg in args {
                    rwe(&mut arg.value);
                }
            }
            ExprKind::SuperCall { args, .. } => {
                for arg in args {
                    rwe(&mut arg.value);
                }
            }
            ExprKind::Array(elems) => {
                for el in elems {
                    rwe(&mut el.value);
                    if let Some(k) = &mut el.key {
                        rwe(k);
                    }
                }
            }
            ExprKind::Tuple(es) | ExprKind::Set(es) | ExprKind::Sequence(es) => {
                for x in es {
                    rwe(x);
                }
            }
            ExprKind::Object(props) => {
                for p in props {
                    match p {
                        ObjectProperty::KeyValue { key, value }
                        | ObjectProperty::Computed { key, value } => {
                            rwe(key);
                            rwe(value);
                        }
                        ObjectProperty::Spread(x) => rwe(x),
                        _ => {}
                    }
                }
            }
            ExprKind::Interpolation(parts) => {
                for p in parts {
                    match p {
                        InterpolPart::Expr(x) | InterpolPart::Formatted(x, _) => rwe(x),
                        _ => {}
                    }
                }
            }
            ExprKind::Match { subject, arms } => {
                rwe(subject);
                for a in arms {
                    if let Some(cs) = &mut a.conditions {
                        for c in cs {
                            rwe(c);
                        }
                    }
                    rwe(&mut a.body);
                }
            }
            ExprKind::Comprehension {
                element,
                generators,
                ..
            } => {
                rwe(element);
                for g in generators {
                    rwe(&mut g.iter);
                    for c in &mut g.conditions {
                        rwe(c);
                    }
                }
            }
            ExprKind::Slice { lower, upper, step } => {
                if let Some(x) = lower {
                    rwe(x);
                }
                if let Some(x) = upper {
                    rwe(x);
                }
                if let Some(x) = step {
                    rwe(x);
                }
            }
            ExprKind::Yield(x) => {
                if let Some(y) = x {
                    rwe(y);
                }
            }
            ExprKind::YieldFrom(x) => rwe(x),
            _ => {}
        }
    }
    fn rws(s: &mut Statement) {
        match &mut s.kind {
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {
                // Scope boundary — separate this context
            }
            StmtKind::Expr(e) => rwe(e),
            StmtKind::Block(ss) => {
                for s in ss {
                    rws(s);
                }
            }
            StmtKind::VarDecl { declarations, .. } => {
                for d in declarations {
                    if let Some(e) = &mut d.init {
                        rwe(e);
                    }
                }
            }
            StmtKind::Return(e) => {
                if let Some(x) = e {
                    rwe(x);
                }
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body } => {
                rwe(cond);
                for s in then_body {
                    rws(s);
                }
                for (c, b) in elifs {
                    rwe(c);
                    for s in b {
                        rws(s);
                    }
                }
                if let Some(b) = else_body {
                    for s in b {
                        rws(s);
                    }
                }
            }
            StmtKind::While {
                cond,
                body,
                else_body } => {
                rwe(cond);
                for s in body {
                    rws(s);
                }
                if let Some(b) = else_body {
                    for s in b {
                        rws(s);
                    }
                }
            }
            StmtKind::DoWhile { body, cond, .. } => {
                for s in body {
                    rws(s);
                }
                rwe(cond);
            }
            StmtKind::For {
                init,
                cond,
                update,
                body } => {
                if let Some(s) = init {
                    rws(s);
                }
                if let Some(e) = cond {
                    rwe(e);
                }
                if let Some(e) = update {
                    rwe(e);
                }
                for s in body {
                    rws(s);
                }
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                rwe(iter);
                for s in body {
                    rws(s);
                }
                if let Some(b) = else_body {
                    for s in b {
                        rws(s);
                    }
                }
            }
            StmtKind::Switch {
                expr,
                cases,
                default } => {
                rwe(expr);
                for c in cases {
                    for cond in &mut c.conditions {
                        match cond {
                            CaseCondition::Value(e) => rwe(e),
                            CaseCondition::Range { from, to } => {
                                rwe(from);
                                rwe(to);
                            }
                            CaseCondition::Comparison { expr, .. } => rwe(expr) }
                    }
                    for s in &mut c.body {
                        rws(s);
                    }
                }
                if let Some(b) = default {
                    for s in b {
                        rws(s);
                    }
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally } => {
                for s in body {
                    rws(s);
                }
                for c in catches {
                    for s in &mut c.body {
                        rws(s);
                    }
                }
                if let Some(b) = else_body {
                    for s in b {
                        rws(s);
                    }
                }
                if let Some(b) = finally {
                    for s in b {
                        rws(s);
                    }
                }
            }
            StmtKind::Assign { targets, value, .. } => {
                for e in targets {
                    rwe(e);
                }
                rwe(value);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                rwe(target);
                rwe(value);
            }
            StmtKind::Throw { expr, cause } => {
                if let Some(e) = expr {
                    rwe(e);
                }
                if let Some(e) = cause {
                    rwe(e);
                }
            }
            StmtKind::Labeled { body, .. } => rws(body),
            StmtKind::Echo(es) | StmtKind::Delete(es) => {
                for e in es {
                    rwe(e);
                }
            }
            StmtKind::Export {
                declaration,
                default,
                ..
            } => {
                if let Some(s) = declaration {
                    rws(s);
                }
                if let Some(e) = default {
                    rwe(e);
                }
            }
            StmtKind::With { body, .. }
            | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. }
            | StmtKind::NamespaceDecl { body, .. } => {
                for s in body {
                    rws(s);
                }
            }
            StmtKind::MatchStatement { subject, cases } => {
                rwe(subject);
                for c in cases {
                    if let Some(e) = &mut c.guard {
                        rwe(e);
                    }
                    for s in &mut c.body {
                        rws(s);
                    }
                }
            }
            StmtKind::Assert { test, msg } => {
                rwe(test);
                if let Some(e) = msg {
                    rwe(e);
                }
            }
            _ => {}
        }
    }
    for stmt in stmts {
        rws(stmt);
    }
}

fn wrap_generator(
    name: String,
    params: Vec<Param>,
    body: Vec<Statement>,
    is_async: bool,
) -> StmtKind {
    let mut body = body;
    rewrite_this_in_generator_body(&mut body);

    let gen_func_expr = Expression::new(ExprKind::FunctionExpr(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: String::new(),
            params: Vec::new(),
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async,
            is_generator: true,
            is_sub: false },
    ))));

    let gen_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__gen_fn".to_string()),
            type_hint: None,
            init: Some(gen_func_expr),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Const });

    let self_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__self_captured".to_string()),
            type_hint: None,
            init: Some(Expression::new(ExprKind::This)),
            array_bounds: None,
            with_events: false }],
        kind: VarDeclKind::Const });

    let callee = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident("__gen_fn")),
        field: "call".to_string(),
        null_safe: false });
    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: vec![Argument {
            name: None,
            value: Expression::ident("__self_captured"),
            by_ref: false,
            spread: false }],
        optional: false });
    let ret_stmt = Statement::new(StmtKind::Return(Some(call_expr)));

    let outer_body = vec![self_decl, gen_decl, ret_stmt];

    StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body: outer_body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false }
}

fn wrap_generator_if_needed(
    name: String,
    params: Vec<Param>,
    body: Vec<Statement>,
    is_async: bool,
    is_generator: bool,
) -> StmtKind {
    if is_generator {
        wrap_generator(name, params, body, is_async)
    } else {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async,
            is_generator: false,
            is_sub: false }
    }
}

/// Normalize the STATIC `Promise.*` combinators into the common async model.
///
/// JS is the language whose spellings already ARE the §27.2 surface, so this
/// changes nothing observable — it routes the statics through the same
/// vocabulary and single lowering every other language uses (a direct
/// `ecma:promise` import instead of a runtime member walk on the `Promise`
/// global). Instance combinators (`.then`/`.catch`/`.finally`) are
/// DELIBERATELY not normalized: any object may carry a user `then` (that is
/// what makes a thenable a thenable, §27.2.4.7 step 8), so only runtime
/// dispatch can decide — a syntactic rewrite would hijack user objects.
fn normalize_promise_statics(expr: &mut Expression) {
    use vybe_ast::{AsyncOp, JoinMode};

    let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
        return;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return;
    };
    if !matches!(&object.kind, ExprKind::Ident(id) if id == "Promise") {
        return;
    }
    fn only_arg(args: &mut Vec<vybe_ast::Argument>) -> Option<Expression> {
        (args.len() == 1 && args[0].name.is_none()).then(|| args.remove(0).value)
    }
    // The combinators take one ITERABLE argument; only the ARRAY-LITERAL
    // spelling is normalized (its elements become Join sources directly).
    // `Promise.all(someIterable)` keeps the runtime path — the vocabulary's
    // Join is variadic-by-value, and re-spreading an arbitrary iterable
    // syntactically would change evaluation order.
    fn array_literal_sources(args: &mut Vec<vybe_ast::Argument>) -> Option<Vec<Expression>> {
        if args.len() != 1 || args[0].name.is_some() {
            return None;
        }
        match &args[0].value.kind {
            ExprKind::Array(items)
                if items.iter().all(|i| i.key.is_none() && !i.spread) =>
            {
                let ExprKind::Array(items) =
                    std::mem::replace(&mut args[0].value.kind, ExprKind::Lit(Literal::Null))
                else {
                    unreachable!()
                };
                Some(items.into_iter().map(|i| i.value).collect())
            }
            _ => None }
    }
    let join = |mode| {
        move |sources| AsyncOp::Join { mode, sources }
    };
    let op = match field.as_str() {
        "resolve" => only_arg(args).map(|v| AsyncOp::Resolved(Box::new(v))),
        "reject" => only_arg(args).map(|r| AsyncOp::Rejected(Box::new(r))),
        "all" => array_literal_sources(args).map(join(JoinMode::All)),
        "allSettled" => array_literal_sources(args).map(join(JoinMode::AllSettled)),
        "race" => array_literal_sources(args).map(join(JoinMode::Race)),
        "any" => array_literal_sources(args).map(join(JoinMode::Any)),
        _ => None };
    if let Some(op) = op {
        expr.kind = ExprKind::Async(op);
    }
}
