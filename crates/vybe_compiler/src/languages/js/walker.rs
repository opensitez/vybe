use super::{JsParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;
use std::sync::atomic::{AtomicUsize, Ordering};

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
                _ => body.push(walk_statement(pair)?),
            }
        }
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

    Ok(Module {
        name: "main".into(),
        language: Lang::JavaScript,
        body,
        imports,
    })
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

fn js_array_elision_marker() -> ArrayElement {
    ArrayElement {
        key: Some(Expression::int(-1)),
        value: Expression::new(ExprKind::Lit(Literal::Undefined)),
        spread: false,
        by_ref: false,
    }
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
            ClassMember::Event { .. } => {}
        }
    }
}

fn rewrite_expression_keys(
    expr: &mut Expression,
    consts: &std::collections::HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::Walrus {
            target: left,
            value: right,
        }
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
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items.iter_mut() {
                rewrite_expression_keys(item, consts);
            }
        }
        ExprKind::Object(props) => {
            for prop in props.iter_mut() {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_expression_keys(key, consts);
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
                                    value: value.clone(),
                                };
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
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
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
        _ => VarDeclKind::Let,
    };
    let mut declarations = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::var_declarator {
            declarations.push(walk_var_declarator(p)?);
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: var_kind,
    })
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
        with_events: false,
    })
}

// ES2025 `using x = expr` / `await using x = expr` — normalize as `const`
// declarations. The compiler treats them identically for now; disposal
// semantics would require finalizer support in the VM.
fn walk_using_decl(pair: Pair<Rule>, _is_await: bool) -> Result<StmtKind, String> {
    let mut declarations = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::using_declarator {
            declarations.push(walk_var_declarator(p)?);
        }
    }
    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const,
    })
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
        other => Err(format!("Unexpected binding pattern: {:?}", other)),
    }
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
            is_rest: true,
        });
    }
    let mut value = None;
    let mut default = None;
    for p in inner {
        match p.as_rule() {
            Rule::binding_pattern => value = Some(walk_binding_pattern(p)?),
            _ => default = Some(walk_expression(p)?),
        }
    }
    Ok(ObjectPatternProp {
        key,
        value,
        default,
        is_rest: false,
    })
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
        other => Err(format!("Unexpected array pattern elem: {:?}", other)),
    }
}

// ── Function declaration ────────────────────────────────────────────────────

/// Recursively scan a function body for `yield` / `yield from` expressions.
/// Does NOT descend into nested function/closure/class bodies — those are
/// their own generator scope.
fn body_contains_yield(stmts: &[Statement]) -> bool {
    fn ey(e: &Expression) -> bool {
        match &e.kind {
            ExprKind::Yield(_) | ExprKind::YieldFrom(_) => true,
            // Scope boundaries — separate generator context
            ExprKind::Lambda { .. } | ExprKind::FunctionExpr(_) | ExprKind::ClassExpr { .. } => {
                false
            }
            ExprKind::RefOf(place) => match place.as_ref() {
                PlaceExpr::Ident(_) => false,
                PlaceExpr::Member { object, .. } => ey(object),
                PlaceExpr::Index { object, index, .. } => ey(object) || ey(index),
                PlaceExpr::Deref(expr) => ey(expr),
            },
            // Leaves
            ExprKind::Lit(_)
            | ExprKind::Ident(_)
            | ExprKind::DefaultOf(_)
            | ExprKind::This
            | ExprKind::Super
            | ExprKind::AddressOf(_)
            | ExprKind::Destructure(_) => false,
            // Unary wrappers
            ExprKind::Unary { expr: i, .. }
            | ExprKind::RefLoad(i)
            | ExprKind::IsType { expr: i, .. }
            | ExprKind::Cast { expr: i, .. }
            | ExprKind::TypeOf(i)
            | ExprKind::Spread(i)
            | ExprKind::Await(i)
            | ExprKind::Void(i)
            | ExprKind::Delete(i) => ey(i),
            // Binary / two-child
            ExprKind::Binary {
                left: a, right: b, ..
            }
            | ExprKind::NullCoalesce { left: a, right: b }
            | ExprKind::Assign {
                target: a,
                value: b,
            }
            | ExprKind::Walrus {
                target: a,
                value: b,
            }
            | ExprKind::Range {
                start: a, end: b, ..
            } => ey(a) || ey(b),
            ExprKind::StaticAccess {
                class: a,
                member: b,
            } => ey(a) || ey(b),
            ExprKind::Ternary { cond, then, else_ } => ey(cond) || ey(then) || ey(else_),
            ExprKind::Member { object, .. } => ey(object),
            ExprKind::Index { object, index, .. } => ey(object) || ey(index),
            ExprKind::Call { callee, args, .. } => ey(callee) || args.iter().any(|a| ey(&a.value)),
            ExprKind::New { class, args } => ey(class) || args.iter().any(|a| ey(&a.value)),
            ExprKind::SuperCall { args, .. } => args.iter().any(|a| ey(&a.value)),
            ExprKind::Array(elems) => elems
                .iter()
                .any(|el| ey(&el.value) || el.key.as_ref().map_or(false, |k| ey(k))),
            ExprKind::Tuple(es) | ExprKind::Set(es) | ExprKind::Sequence(es) => {
                es.iter().any(|x| ey(x))
            }
            ExprKind::Object(props) => props.iter().any(|p| match p {
                ObjectProperty::KeyValue { key, value }
                | ObjectProperty::Computed { key, value } => ey(key) || ey(value),
                ObjectProperty::Spread(x) => ey(x),
                _ => false,
            }),
            ExprKind::Interpolation(parts) => parts.iter().any(|p| match p {
                InterpolPart::Expr(x) | InterpolPart::Formatted(x, _) => ey(x),
                _ => false,
            }),
            ExprKind::Match { subject, arms } => {
                ey(subject)
                    || arms.iter().any(|a| {
                        a.conditions
                            .as_ref()
                            .map_or(false, |cs| cs.iter().any(|c| ey(c)))
                            || ey(&a.body)
                    })
            }
            ExprKind::Comprehension {
                element,
                generators,
                ..
            } => {
                ey(element)
                    || generators
                        .iter()
                        .any(|g| ey(&g.iter) || g.conditions.iter().any(|c| ey(c)))
            }
            ExprKind::Slice { lower, upper, step } => [lower, upper, step]
                .iter()
                .any(|o| o.as_ref().map_or(false, |x| ey(x))),
        }
    }
    fn sy(s: &Statement) -> bool {
        match &s.kind {
            StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => false,
            StmtKind::Expr(e) => ey(e),
            StmtKind::Block(ss) => ss.iter().any(|s| sy(s)),
            StmtKind::VarDecl { declarations, .. } => declarations
                .iter()
                .any(|d| d.init.as_ref().map_or(false, |e| ey(e))),
            StmtKind::Return(e) => e.as_ref().map_or(false, |e| ey(e)),
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                ey(cond)
                    || then_body.iter().any(|s| sy(s))
                    || elifs.iter().any(|(c, b)| ey(c) || b.iter().any(|s| sy(s)))
                    || else_body
                        .as_ref()
                        .map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::While {
                cond,
                body,
                else_body,
            } => {
                ey(cond)
                    || body.iter().any(|s| sy(s))
                    || else_body
                        .as_ref()
                        .map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::DoWhile { body, cond, .. } => body.iter().any(|s| sy(s)) || ey(cond),
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                init.as_ref().map_or(false, |s| sy(s))
                    || cond.as_ref().map_or(false, |e| ey(e))
                    || update.as_ref().map_or(false, |e| ey(e))
                    || body.iter().any(|s| sy(s))
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                ey(iter)
                    || body.iter().any(|s| sy(s))
                    || else_body
                        .as_ref()
                        .map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Switch {
                expr,
                cases,
                default,
            } => {
                ey(expr)
                    || cases.iter().any(|c| c.body.iter().any(|s| sy(s)))
                    || default.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                body.iter().any(|s| sy(s))
                    || catches.iter().any(|c| c.body.iter().any(|s| sy(s)))
                    || else_body
                        .as_ref()
                        .map_or(false, |b| b.iter().any(|s| sy(s)))
                    || finally.as_ref().map_or(false, |b| b.iter().any(|s| sy(s)))
            }
            StmtKind::Assign { targets, value } => targets.iter().any(|e| ey(e)) || ey(value),
            StmtKind::CompoundAssign { target, value, .. } => ey(target) || ey(value),
            StmtKind::Throw { expr, cause } => {
                expr.as_ref().map_or(false, |e| ey(e)) || cause.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::Labeled { body, .. } => sy(body),
            StmtKind::Echo(es) | StmtKind::Delete(es) => es.iter().any(|e| ey(e)),
            StmtKind::Export {
                declaration,
                default,
                ..
            } => {
                declaration.as_ref().map_or(false, |s| sy(s))
                    || default.as_ref().map_or(false, |e| ey(e))
            }
            StmtKind::With { body, .. }
            | StmtKind::Using { body, .. }
            | StmtKind::Lock { body, .. }
            | StmtKind::NamespaceDecl { body, .. } => body.iter().any(|s| sy(s)),
            StmtKind::MatchStatement { subject, cases } => {
                ey(subject)
                    || cases.iter().any(|c| {
                        c.guard.as_ref().map_or(false, |e| ey(e)) || c.body.iter().any(|s| sy(s))
                    })
            }
            StmtKind::Assert { test, msg } => ey(test) || msg.as_ref().map_or(false, |e| ey(e)),
            _ => false,
        }
    }
    stmts.iter().any(|s| sy(s))
}

fn walk_func_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let is_async = pair.as_rule() == Rule::async_function_declaration;
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
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub: false,
    })
}

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    walk_params_with_prologue(pair).map(|(params, _)| params)
}

fn walk_param(pair: Pair<Rule>) -> Result<Param, String> {
    walk_param_with_prologue(pair, 0).map(|(param, _)| param)
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
            _ => default = Some(walk_expression(p)?),
        }
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
                is_nullable: false,
            },
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
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            });
            Ok((
                Param {
                    name: temp_name,
                    type_hint: None,
                    default,
                    pass_by: PassBy::Value,
                    is_rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                },
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
                let raw = extract_ident_name(&p);
                let is_simple = !raw.contains('(')
                    && !raw.contains('.')
                    && !raw.contains(' ')
                    && !raw.contains('[')
                    && !raw.is_empty();
                if is_simple {
                    parents.push(raw);
                } else {
                    let synth = format!("__extends_{}_{}", name, parents.len());
                    let init = walk_expression(p)?;
                    pre_class_stmts.push(Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(synth.clone()),
                            type_hint: None,
                            init: Some(init),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Var,
                    }));
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

    // Extract static block bodies — these run immediately after class definition.
    // Collect them and emit as post-class statements in a wrapping Block.
    let mut static_init_stmts: Vec<Statement> = Vec::new();
    members.retain(|m| {
        if let ClassMember::Method(func) = m {
            if let StmtKind::FunctionDecl {
                name: ref mname,
                ref body,
                ref modifiers,
                ..
            } = func.kind
            {
                if mname == "__static_init" && modifiers.is_static {
                    static_init_stmts.extend(body.iter().cloned());
                    return false; // remove from members
                }
            }
        }
        true
    });

    let class_stmt = StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: ClassModifiers::default(),
        decorators: vec![],
    };

    if static_init_stmts.is_empty() && pre_class_stmts.is_empty() {
        Ok(class_stmt)
    } else {
        // Wrap: pre-class extends bindings, then class declaration,
        // then static init statements, all in a Block.
        let mut block = pre_class_stmts;
        block.push(Statement::new(class_stmt));
        block.extend(static_init_stmts);
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
            is_sub: true,
        });
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
                },
            })
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
                is_nullable: false,
            };
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
                },
            })
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
                    params,
                    body,
                    base_args: None,
                    initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public,
                })
            } else {
                if !is_generator {
                    is_generator = body_contains_yield(&body);
                }
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
                        is_sub: false,
                    },
                ))))
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
                    _ => init = Some(walk_expression(p)?),
                }
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
                array_bounds: None,
            })
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
                    _ => init = Some(walk_expression(p)?),
                }
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
                array_bounds: None,
            })
        }
        other => Err(format!("Unexpected class member: {:?}", other)),
    }
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
        Err(_) => None,
    };
    Ok(StmtKind::If {
        cond,
        then_body: vec![then_stmt],
        elifs: Vec::new(),
        else_body,
    })
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
                is_async: is_for_await,
            })
        }
        Rule::for_of_header => {
            let parts: Vec<Pair<Rule>> = header_inner.into_inner().collect();
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
                    is_nullable: false,
                }];
                let args = vec![Argument::positional(Expression::ident(&var))];
                let iife = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params,
                        body: LambdaBody::Block(full_body),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    args,
                    optional: false,
                });
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
                is_async: is_for_await,
            })
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
                                    _ => VarDeclKind::Let,
                                };
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
                                    kind: var_kind,
                                })));
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
                        is_nullable: false,
                    })
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
                        captures: Vec::new(),
                    })),
                    args,
                    optional: false,
                });
                vec![Statement::new(StmtKind::Expr(iife))]
            } else {
                body
            };

            Ok(StmtKind::For {
                init,
                cond,
                update,
                body,
            })
        }
        other => Err(format!("Unexpected for header: {:?}", other)),
    }
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body = vec![walk_statement(next_meaningful(&mut inner)?)?];
    let cond = walk_expression(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    })
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
                    body: stmts,
                });
            } else {
                let first = case_inner.next().ok_or("Empty switch case")?;
                let val = walk_expression(first)?;
                let stmts: Vec<Statement> = case_inner
                    .filter(|p| p.as_rule() != Rule::NEWLINE)
                    .map(walk_statement)
                    .collect::<Result<Vec<_>, _>>()?;
                cases.push(SwitchCase {
                    conditions: vec![CaseCondition::Value(val)],
                    body: stmts,
                });
            }
        }
    }
    Ok(StmtKind::Switch {
        expr,
        cases,
        default: None,
    })
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| p.as_rule() != Rule::NEWLINE)
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
        None => BreakTarget::Implicit,
    }))
}

fn walk_continue(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let label = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::ident_name)
        .map(|p| p.as_str().to_string());
    Ok(StmtKind::Continue(match label {
        Some(l) => ContinueTarget::Label(l),
        None => ContinueTarget::Implicit,
    }))
}

fn walk_throw(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = walk_expression(first_meaningful(pair)?)?;
    Ok(StmtKind::Throw {
        expr: Some(expr),
        cause: None,
    })
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
                                            with_events: false,
                                        }],
                                        kind: VarDeclKind::Const,
                                    }));
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
                    when_clause: None,
                });
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
        finally,
    })
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let label = next_meaningful(&mut inner)?.as_str().to_string();
    let body = walk_statement(next_meaningful(&mut inner)?)?;
    Ok(StmtKind::Labeled {
        label,
        body: Box::new(body),
    })
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
            alias: Some(ns),
        }
    } else if let Some(def) = default_name {
        if names.is_empty() {
            ImportKind::Default {
                path: source,
                local: def,
            }
        } else {
            // import default, { named } from "mod" — use Named with default as first
            let mut all_names = vec![ImportName {
                name: "default".into(),
                alias: Some(def),
            }];
            all_names.extend(names);
            ImportKind::Named {
                path: source,
                names: all_names,
                level: 0,
            }
        }
    } else if !names.is_empty() {
        ImportKind::Named {
            path: source,
            names,
            level: 0,
        }
    } else {
        ImportKind::Simple {
            path: source,
            alias: None,
        }
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
                alias: Some(n),
            });
        }
    }

    Ok(StmtKind::Export {
        declaration,
        names,
        default: default_expr,
        from,
        star,
    })
}

// ── Expressions ─────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(collapse_passthrough_expression(pair)?)?;
    Ok(Expression::with_span(kind, span))
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
                    _ => None,
                }
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
                    _ => None,
                }
            }
            Rule::primary => match pair.as_str().trim() {
                "true" | "false" | "null" | "undefined" | "this" | "super" => None,
                _ => {
                    let mut inner = pair.clone().into_inner();
                    match (inner.next(), inner.next()) {
                        (Some(first), None) => Some(first),
                        _ => None,
                    }
                }
            },
            Rule::unary => {
                let mut inner = pair.clone().into_inner();
                match (inner.next(), inner.next()) {
                    (Some(first), None) if first.as_rule() == Rule::postfix => Some(first),
                    _ => None,
                }
            }
            Rule::postfix => {
                let mut inner = pair.clone().into_inner();
                let first = inner.next();
                let has_postfix = inner.any(|p| p.as_rule() == Rule::postfix_op);
                match (first, has_postfix) {
                    (Some(first), false) => Some(first),
                    _ => None,
                }
            }
            Rule::call_expression => {
                let mut inner = pair.clone().into_inner();
                let first = inner.next();
                let has_chain = inner.any(|p| p.as_rule() == Rule::call_chain);
                match (first, has_chain) {
                    (Some(first), false) => Some(first),
                    _ => None,
                }
            }
            _ => None,
        };

        match next {
            Some(next_pair) => pair = next_pair,
            None => return Ok(pair),
        }
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
            let n = if s.starts_with("0x") || s.starts_with("0X") {
                i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?
            } else if s.starts_with("0o") || s.starts_with("0O") {
                i64::from_str_radix(&s[2..], 8).map_err(|e| format!("{}", e))?
            } else if s.starts_with("0b") || s.starts_with("0B") {
                i64::from_str_radix(&s[2..], 2).map_err(|e| format!("{}", e))?
            } else {
                s.parse().unwrap_or(0)
            };
            Ok(ExprKind::Lit(Literal::BigInt(n)))
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
                optional: false,
            })
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
                _ => Ok(ExprKind::Ident(name.to_string())),
            }
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
                let op_str = inner.remove(0).as_str();
                let right = walk_expression(inner.remove(0))?;
                if op_str == "=" {
                    Ok(ExprKind::Assign {
                        target: Box::new(left),
                        value: Box::new(right),
                    })
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
                        _ => CompoundOp::Add,
                    };
                    Ok(ExprKind::Assign {
                        target: Box::new(left.clone()),
                        value: Box::new(Expression::new(ExprKind::Binary {
                            op: compound_to_binop(op),
                            left: Box::new(left),
                            right: Box::new(right),
                        })),
                    })
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
                    else_: Box::new(else_),
                })
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
                    return Ok(ExprKind::Lit(crate::ast::Literal::Bool(false)));
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
                _ => UnaryOp::Neg,
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(operand),
            })
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
                _ => return Ok(base.kind),
            };
            Ok(ExprKind::Unary {
                op,
                expr: Box::new(base),
            })
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
                        args,
                    });
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
                        null_safe: false,
                    });
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
                                optional: true,
                            });
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
                                null_safe: true,
                            });
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
                            optional: false,
                        });
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
                            null_safe: false,
                        });
                    }
                }
            }

            // If `new` had no `()` (e.g., `new X`), wrap the bare class.
            if !new_consumed {
                expr = Expression::new(ExprKind::New {
                    class: Box::new(expr),
                    args: Vec::new(),
                });
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
                            is_nullable: false,
                        }]
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
                            _ => LambdaBody::Expr(Box::new(walk_expression(inner)?)),
                        };
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
                captures: Vec::new(),
            })
        }

        // Function expression
        Rule::function_expression | Rule::async_function_expression => {
            let stmt_kind = walk_func_decl(pair)?;
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
                    _ => parent = Some(Box::new(walk_expression(p)?)),
                }
            }
            Ok(ExprKind::ClassExpr {
                name,
                parent,
                members,
            })
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
                        by_ref: false,
                    })
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
                    _ => parts.push(InterpolPart::Expr(walk_expression(p)?)),
                }
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

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
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
            _ => op_pair.as_str().trim(),
        };
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
            _ => BinOp::Add,
        };

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
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
                    null_safe: true,
                });
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
                        optional: true,
                    });
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
                        null_safe: true,
                    });
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
                optional: false,
            });
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
                    null_safe: false,
                });
            } else {
                expr = Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(index_expr),
                    null_safe: false,
                });
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
                                by_ref: false,
                            })
                            .collect(),
                    ))
                };
                let cooked_array = make_str_array(parts);
                let raw_array = make_str_array(raw_parts);
                // Object.assign(cooked, { raw: raw }) — sets .raw and returns the array
                let raw_obj = Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: Expression::string("raw"),
                    value: raw_array,
                }]));
                let strings_with_raw = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Object")),
                        field: "assign".into(),
                        null_safe: false,
                    })),
                    args: vec![
                        Argument::positional(cooked_array),
                        Argument::positional(raw_obj),
                    ],
                    optional: false,
                });
                // ECMA-262 §13.2.8.3: template objects are cached per call site.
                // Wrap in `__vybe_tmpl_N ?? (__vybe_tmpl_N = Object.assign(...))` so
                // the same object is returned on every invocation of this template site.
                let tmpl_id = TEMPLATE_COUNTER.fetch_add(1, Ordering::Relaxed);
                let tmpl_global = format!("__vybe_tmpl_{}", tmpl_id);
                let cached_template = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(Expression::ident(&tmpl_global)),
                    right: Box::new(Expression::new(ExprKind::Assign {
                        target: Box::new(Expression::ident(&tmpl_global)),
                        value: Box::new(strings_with_raw),
                    })),
                });
                args.push(Argument::positional(cached_template));
                for e in exprs {
                    args.push(Argument::positional(e));
                }
                expr = Expression::new(ExprKind::Call {
                    callee: Box::new(expr),
                    args,
                    optional: false,
                });
            }
        }
    }

    // Normalize variadic concat: `x.concat(a, b, c)` → `x.concat(a).concat(b).concat(c)`
    // The stdlib concat function is binary (receiver + 1 arg). For variadic calls,
    // desugar into a chain of binary concat calls. Works for both strings and arrays.
    expr = desugar_variadic_concat(expr);

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
            _ => false,
        }
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
            _ => false,
        }
    }
    stmts.iter().any(has_closure_stmt)
}

/// Desugar `x.concat(a, b, c)` into `x.concat(a).concat(b).concat(c)`.
fn desugar_variadic_concat(expr: Expression) -> Expression {
    if let ExprKind::Call {
        ref callee,
        ref args,
        optional,
    } = expr.kind
    {
        if args.len() > 1 {
            if let ExprKind::Member {
                ref object,
                ref field,
                null_safe,
            } = callee.kind
            {
                if field == "concat" {
                    // Chain: start with receiver.concat(args[0]), then .concat(args[1]), etc.
                    let mut result = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "concat".to_string(),
                            null_safe,
                        })),
                        args: vec![args[0].clone()],
                        optional,
                    });
                    for arg in &args[1..] {
                        result = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(result),
                                field: "concat".to_string(),
                                null_safe: false,
                            })),
                            args: vec![arg.clone()],
                            optional: false,
                        });
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
                    if let Some(ch) = char::from_u32(n) {
                        result.push(ch);
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
            _ => result.push('\\'),
        }
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
        _ => None,
    };
    if let Some(canonical_name) = canonical {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(canonical_name)),
            args: vec![Argument::positional(object)],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: name.to_string(),
            null_safe: false,
        })
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
        _ => None,
    }
}

fn js_well_known_symbol_alias(expr: &Expression) -> Option<&'static str> {
    let ExprKind::Member {
        object,
        field,
        null_safe,
    } = &expr.kind
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
                spread,
            })
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
                value,
            });
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
                        _ => Expression::string(&key_str),
                    }
                } else {
                    Expression::string(&key_str)
                }
            }
            _ => walk_expression(key_pair)?,
        };
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
        end_col: end.1 as u32 - 1,
    }
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
                    _ => VarDeclKind::Let,
                };
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
                        with_events: false,
                    }],
                    kind: var_kind,
                });
                return Ok((tmp, vec![prefix]));
            }
            _ => continue,
        }
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
    let prop_name = inner.remove(0).as_str().to_string();
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::param_list => params = walk_params(p)?,
            Rule::param => params = vec![walk_param(p)?],
            Rule::function_body => body = walk_body(p)?,
            _ => {}
        }
    }
    let mut full_params = vec![Param {
        name: "this".to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }];
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
            captures: Vec::new(),
        }),
    })
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
    let mut body = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::async_kw => is_async = true,
            Rule::param_list => params = walk_params(p)?,
            Rule::param => params = vec![walk_param(p)?],
            Rule::function_body => body = walk_body(p)?,
            _ => {}
        }
    }
    let is_generator = has_generator_marker || body_contains_yield(&body);

    // Computed method: return Computed { key: runtime_expr, value: lambda }
    if let Some(key_expr) = computed_expr {
        let lambda = Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Block(body),
            is_async,
            captures: Vec::new(),
        });
        return Ok(ObjectProperty::Computed {
            key: key_expr,
            value: lambda,
        });
    }

    let func = Statement::new(StmtKind::FunctionDecl {
        name: key.clone(),
        params,
        return_type: None,
        body,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async,
        is_generator,
        is_sub: false,
    });
    Ok(ObjectProperty::Method {
        key,
        value: Box::new(func),
    })
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
        Some((p, f)) => (p.to_string(), f.to_string()),
        None => (raw.to_string(), String::new()),
    };
    ExprKind::New {
        class: Box::new(Expression::ident("RegExp")),
        args: vec![
            Argument::positional(Expression::string(&pattern)),
            Argument::positional(Expression::string(&flags)),
        ],
    }
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
                    _ => out.push('\\'),
                }
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
                            if let Some(c) = char::from_u32(n) {
                                out.push(c);
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
            None => out.push('\\'),
        }
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
        CompoundOp::Concat => BinOp::Concat,
    }
}
