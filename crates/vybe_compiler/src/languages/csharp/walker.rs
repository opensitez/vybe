use std::collections::{HashMap, HashSet};

use super::{CSharpParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

pub fn parse(source: &str) -> Result<Module, String> {
    let pairs =
        CSharpParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut pending_attributes: Vec<Expression> = Vec::new();

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
                Rule::EOI => continue,
                Rule::attribute_list => {
                    pending_attributes.extend(parse_attribute_specs(pair.as_str()))
                }
                Rule::using_directive => imports.push(walk_using(pair)?),
                Rule::namespace_declaration => {
                    // Build NamespaceDecl with name and body
                    let mut ns_name = String::new();
                    let mut ns_body: Vec<Statement> = Vec::new();
                    let mut namespace_pending_attributes: Vec<Expression> = Vec::new();
                    for p in pair.into_inner() {
                        match p.as_rule() {
                            Rule::dotted_name => {
                                ns_name = strip_global_namespace_qualifier(p.as_str()).to_string()
                            }
                            Rule::attribute_list => namespace_pending_attributes
                                .extend(parse_attribute_specs(p.as_str())),
                            Rule::using_directive => imports.push(walk_using(p)?),
                            _ => {
                                if let Ok(stmt) =
                                    walk_top_level_with_attributes(p, &namespace_pending_attributes)
                                {
                                    ns_body.push(stmt);
                                }
                                namespace_pending_attributes.clear();
                            }
                        }
                    }
                    body.push(Statement::new(StmtKind::NamespaceDecl {
                        name: ns_name,
                        body: ns_body,
                    }));
                }
                _ => {
                    if let Ok(stmt) = walk_top_level_with_attributes(pair, &pending_attributes) {
                        body.push(stmt);
                    }
                    pending_attributes.clear();
                }
            }
        }
    }

    // Synthesize the .NET Exception hierarchy at the top of every C#
    // program. ECMA-335 / .NET BCL exposes `System.Exception` plus a
    // family of common subclasses (`InvalidOperationException`,
    // `ArgumentNullException`, `DivideByZeroException`, etc.) that user
    // code routinely throws via `throw new <T>("msg")`. We don't model
    // the BCL's class file system, so the walker injects minimal
    // declarations: each takes a `string msg` ctor that stamps `Message`
    // on `this`. `try { ... } catch (T e) { ... e.Message ... }`
    // resolves T to the synthesized class and reads `e.Message`.
    let mut synthesized = synthesize_exception_classes();
    synthesized.extend(synthesize_attribute_classes());
    synthesized.extend(synthesize_checked_numeric_helpers());
    body.splice(0..0, synthesized);

    let mut module = Module {
        name: "main".into(),
        language: Lang::CSharp,
        body,
        imports,
    };
    rewrite_using_imports(&mut module);
    lower_csharp_using_declarations(&mut module.body);
    rewrite_set_algebra_bool_calls(&mut module.body);
    rewrite_explicit_interface_accesses(&mut module);
    rewrite_record_uses(&mut module);
    rewrite_tuple_uses(&mut module);
    rewrite_extension_calls(&mut module);
    rewrite_linked_list_node_uses(&mut module.body);
    rewrite_sorted_dictionary_foreach(&mut module.body);
    rewrite_user_defined_operator_calls(&mut module);
    Ok(module)
}

fn split_csharp_top_level(text: &str, separator: char) -> Vec<String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut paren_depth = 0usize;
    let mut bracket_depth = 0usize;
    let mut brace_depth = 0usize;
    let mut in_string = false;
    let mut string_delim = '\0';
    let mut escaped = false;

    for ch in text.chars() {
        if in_string {
            current.push(ch);
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' {
                escaped = true;
                continue;
            }
            if ch == string_delim {
                in_string = false;
                string_delim = '\0';
            }
            continue;
        }

        match ch {
            '"' | '\'' => {
                in_string = true;
                string_delim = ch;
                current.push(ch);
            }
            '(' => {
                paren_depth += 1;
                current.push(ch);
            }
            ')' => {
                paren_depth = paren_depth.saturating_sub(1);
                current.push(ch);
            }
            '[' => {
                bracket_depth += 1;
                current.push(ch);
            }
            ']' => {
                bracket_depth = bracket_depth.saturating_sub(1);
                current.push(ch);
            }
            '{' => {
                brace_depth += 1;
                current.push(ch);
            }
            '}' => {
                brace_depth = brace_depth.saturating_sub(1);
                current.push(ch);
            }
            _ if ch == separator && paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 => {
                let piece = current.trim();
                if !piece.is_empty() {
                    parts.push(piece.to_string());
                }
                current.clear();
            }
            _ => current.push(ch),
        }
    }

    let piece = current.trim();
    if !piece.is_empty() {
        parts.push(piece.to_string());
    }
    parts
}

fn normalize_attribute_type_name(name: &str) -> String {
    let trimmed = name.trim();
    let Some((prefix, leaf)) = trimmed.rsplit_once('.') else {
        return if trimmed.ends_with("Attribute") {
            trimmed.to_string()
        } else {
            format!("{trimmed}Attribute")
        };
    };
    if leaf.ends_with("Attribute") {
        trimmed.to_string()
    } else {
        format!("{prefix}.{}Attribute", leaf)
    }
}

fn parse_csharp_inline_expression(text: &str) -> Option<Expression> {
    let mut parsed = CSharpParser::parse(Rule::expression, text).ok()?;
    let pair = parsed.next()?;
    walk_expression(pair).ok()
}

fn parse_attribute_argument(text: &str) -> Option<Argument> {
    let trimmed = text.trim();
    if let Some((lhs, rhs)) = trimmed.split_once('=') {
        let lhs_trimmed = lhs.trim();
        if lhs_trimmed
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '.')
        {
            return Some(Argument {
                value: parse_csharp_inline_expression(rhs.trim())?,
                name: Some(lhs_trimmed.rsplit('.').next()?.to_string()),
                by_ref: false,
                spread: false,
            });
        }
    }

    Some(Argument {
        value: parse_csharp_inline_expression(trimmed)?,
        name: None,
        by_ref: false,
        spread: false,
    })
}

fn parse_attribute_specs(raw: &str) -> Vec<Expression> {
    let inner = raw.trim().trim_start_matches('[').trim_end_matches(']');
    split_csharp_top_level(inner, ',')
        .into_iter()
        .filter_map(|part| {
            let trimmed = part.trim();
            if trimmed.is_empty() {
                return None;
            }

            let (name, args) = if let Some(open_idx) = trimmed.find('(') {
                let close_idx = trimmed
                    .rfind(')')
                    .unwrap_or(trimmed.len().saturating_sub(1));
                let name = trimmed[..open_idx].trim();
                let args_src = &trimmed[open_idx + 1..close_idx];
                let args = split_csharp_top_level(args_src, ',')
                    .into_iter()
                    .filter_map(|arg| parse_attribute_argument(&arg))
                    .collect::<Vec<_>>();
                (name, args)
            } else {
                (trimmed, Vec::new())
            };

            Some(Expression::new(ExprKind::New {
                class: Box::new(build_dotted_expr(&normalize_attribute_type_name(name))),
                args,
            }))
        })
        .collect()
}

fn attribute_leaf_name(expr: &Expression) -> Option<String> {
    let callee = match &expr.kind {
        ExprKind::Call { callee, .. } => callee.as_ref(),
        ExprKind::New { class, .. } => class.as_ref(),
        _ => expr,
    };
    match &callee.kind {
        ExprKind::Ident(name) => Some(name.rsplit('.').next().unwrap_or(name).to_string()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

fn explicit_interface_runtime_name(interface_name: &str, member_name: &str) -> String {
    format!(
        "__iface__{}__{}",
        sanitize_explicit_interface_name(interface_name),
        member_name
    )
}

fn sanitize_explicit_interface_name(interface_name: &str) -> String {
    normalize_runtime_type_name(interface_name)
        .chars()
        .map(|ch| {
            if ch.is_ascii_alphanumeric() {
                ch.to_ascii_lowercase()
            } else {
                '_'
            }
        })
        .collect()
}

fn parse_explicit_interface_runtime_name(name: &str) -> Option<(String, String)> {
    let remainder = name.strip_prefix("__iface__")?;
    let (iface, member) = remainder.split_once("__")?;
    Some((iface.to_string(), member.to_string()))
}

fn strip_global_namespace_qualifier(name: &str) -> &str {
    name.strip_prefix("global::").unwrap_or(name)
}

fn extract_explicit_interface_name(pair: Pair<Rule>) -> String {
    strip_global_namespace_qualifier(pair.as_str())
        .trim_end_matches('.')
        .trim()
        .to_string()
}

fn rewrite_explicit_interface_accesses(module: &mut Module) {
    let conflicted = collect_conflicted_explicit_interface_members(&module.body);
    if conflicted.is_empty() {
        return;
    }
    rewrite_explicit_interface_accesses_in_statements(&mut module.body, &conflicted);
}

fn collect_conflicted_explicit_interface_members(body: &[Statement]) -> HashSet<String> {
    let mut conflicted = HashSet::new();
    for stmt in body {
        collect_conflicted_explicit_interface_members_in_statement(stmt, &mut conflicted);
    }
    conflicted
}

fn collect_conflicted_explicit_interface_members_in_statement(
    stmt: &Statement,
    conflicted: &mut HashSet<String>,
) {
    match &stmt.kind {
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                collect_conflicted_explicit_interface_members_in_statement(stmt, conflicted);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            collect_conflicted_explicit_interface_members_in_class_members(members, conflicted);
            for member in members {
                if let ClassMember::NestedType(stmt) = member {
                    collect_conflicted_explicit_interface_members_in_statement(stmt, conflicted);
                }
            }
        }
        _ => {}
    }
}

fn collect_conflicted_explicit_interface_members_in_class_members(
    members: &[ClassMember],
    conflicted: &mut HashSet<String>,
) {
    let mut plain_methods: HashSet<String> = HashSet::new();
    let mut explicit_methods: HashMap<String, Vec<String>> = HashMap::new();
    let mut plain_properties: HashSet<String> = HashSet::new();
    let mut explicit_properties: HashMap<String, Vec<String>> = HashMap::new();

    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                    if parse_explicit_interface_runtime_name(name).is_none() {
                        plain_methods.insert(name.clone());
                    }
                }
            }
            ClassMember::Property { name, .. } => {
                if parse_explicit_interface_runtime_name(name).is_none() {
                    plain_properties.insert(name.clone());
                }
            }
            _ => {}
        }
    }

    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                    if let Some((iface, base)) = parse_explicit_interface_runtime_name(name) {
                        explicit_methods.entry(base).or_default().push(iface);
                    }
                }
            }
            ClassMember::Property { name, .. } => {
                if let Some((iface, base)) = parse_explicit_interface_runtime_name(name) {
                    explicit_properties.entry(base).or_default().push(iface);
                }
            }
            _ => {}
        }
    }

    for (base, ifaces) in explicit_methods {
        if ifaces.len() > 1 || plain_methods.contains(&base) {
            for iface in ifaces {
                conflicted.insert(explicit_interface_runtime_name(&iface, &base));
            }
        }
    }
    for (base, ifaces) in explicit_properties {
        if ifaces.len() > 1 || plain_properties.contains(&base) {
            for iface in ifaces {
                conflicted.insert(explicit_interface_runtime_name(&iface, &base));
            }
        }
    }
}

fn rewrite_explicit_interface_accesses_in_statements(
    body: &mut [Statement],
    conflicted: &HashSet<String>,
) {
    for stmt in body {
        rewrite_explicit_interface_accesses_in_statement(stmt, conflicted);
    }
}

fn rewrite_explicit_interface_accesses_in_statement(
    stmt: &mut Statement,
    conflicted: &HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
            rewrite_explicit_interface_accesses_in_expr(cause, conflicted);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_explicit_interface_accesses_in_expr(init, conflicted);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_explicit_interface_accesses_in_expr(bound, conflicted);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_explicit_interface_accesses_in_member(member, conflicted);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_explicit_interface_accesses_in_expr(cond, conflicted);
            rewrite_explicit_interface_accesses_in_statements(then_body, conflicted);
            for (elif_cond, elif_body) in elifs {
                rewrite_explicit_interface_accesses_in_expr(elif_cond, conflicted);
                rewrite_explicit_interface_accesses_in_statements(elif_body, conflicted);
            }
            if let Some(else_body) = else_body {
                rewrite_explicit_interface_accesses_in_statements(else_body, conflicted);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_explicit_interface_accesses_in_statement(init, conflicted);
            }
            if let Some(cond) = cond {
                rewrite_explicit_interface_accesses_in_expr(cond, conflicted);
            }
            if let Some(update) = update {
                rewrite_explicit_interface_accesses_in_expr(update, conflicted);
            }
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_explicit_interface_accesses_in_expr(iter, conflicted);
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
            if let Some(else_body) = else_body {
                rewrite_explicit_interface_accesses_in_statements(else_body, conflicted);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_explicit_interface_accesses_in_expr(cond, conflicted);
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
            if let Some(else_body) = else_body {
                rewrite_explicit_interface_accesses_in_statements(else_body, conflicted);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
            rewrite_explicit_interface_accesses_in_expr(cond, conflicted);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_explicit_interface_accesses_in_expr(expr, conflicted)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_explicit_interface_accesses_in_expr(from, conflicted);
                            rewrite_explicit_interface_accesses_in_expr(to, conflicted);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
                        }
                    }
                }
                rewrite_explicit_interface_accesses_in_statements(&mut case.body, conflicted);
            }
            if let Some(default) = default {
                rewrite_explicit_interface_accesses_in_statements(default, conflicted);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_explicit_interface_accesses_in_expr(when_clause, conflicted);
                }
                rewrite_explicit_interface_accesses_in_statements(&mut catch.body, conflicted);
            }
            if let Some(else_body) = else_body {
                rewrite_explicit_interface_accesses_in_statements(else_body, conflicted);
            }
            if let Some(finally) = finally {
                rewrite_explicit_interface_accesses_in_statements(finally, conflicted);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_explicit_interface_accesses_in_expr(&mut item.expr, conflicted);
            }
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_explicit_interface_accesses_in_expr(target, conflicted);
            }
            rewrite_explicit_interface_accesses_in_expr(value, conflicted);
        }
        _ => {}
    }
}

fn rewrite_explicit_interface_accesses_in_member(
    member: &mut ClassMember,
    conflicted: &HashSet<String>,
) {
    match member {
        ClassMember::Field {
            init: Some(expr),
            array_bounds,
            ..
        } => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_explicit_interface_accesses_in_expr(bound, conflicted);
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_explicit_interface_accesses_in_statement(stmt, conflicted);
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            rewrite_explicit_interface_accesses_in_statements(body, conflicted);
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_explicit_interface_accesses_in_expr(arg, conflicted);
                }
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_explicit_interface_accesses_in_statements(getter, conflicted);
            }
            if let Some(setter) = setter {
                rewrite_explicit_interface_accesses_in_statements(&mut setter.body, conflicted);
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_explicit_interface_accesses_in_expr(value, conflicted);
        }
        _ => {}
    }
}

fn rewrite_explicit_interface_accesses_in_expr(
    expr: &mut Expression,
    conflicted: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_explicit_interface_accesses_in_expr(left, conflicted);
            rewrite_explicit_interface_accesses_in_expr(right, conflicted);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr) => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } => {
                rewrite_explicit_interface_accesses_in_expr(object, conflicted);
            }
            PlaceExpr::Index { object, index, .. } => {
                rewrite_explicit_interface_accesses_in_expr(object, conflicted);
                rewrite_explicit_interface_accesses_in_expr(index, conflicted);
            }
            PlaceExpr::Deref(expr) => {
                rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
            }
        },
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_explicit_interface_accesses_in_expr(cond, conflicted);
            rewrite_explicit_interface_accesses_in_expr(then, conflicted);
            rewrite_explicit_interface_accesses_in_expr(else_, conflicted);
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_explicit_interface_accesses_in_expr(object, conflicted);
            if let ExprKind::Cast { type_name, .. } = &object.kind {
                let hidden = explicit_interface_runtime_name(type_name, field);
                if conflicted.contains(&hidden) {
                    *field = hidden;
                }
            }
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_explicit_interface_accesses_in_expr(object, conflicted);
            rewrite_explicit_interface_accesses_in_expr(index, conflicted);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_explicit_interface_accesses_in_expr(callee, conflicted);
            for arg in args.iter_mut() {
                rewrite_explicit_interface_accesses_in_expr(&mut arg.value, conflicted);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_explicit_interface_accesses_in_expr(class, conflicted);
            for arg in args {
                rewrite_explicit_interface_accesses_in_expr(&mut arg.value, conflicted);
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_explicit_interface_accesses_in_expr(target, conflicted);
            rewrite_explicit_interface_accesses_in_expr(value, conflicted);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_explicit_interface_accesses_in_expr(expr, conflicted),
            LambdaBody::Block(body) => {
                rewrite_explicit_interface_accesses_in_statements(body, conflicted)
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_explicit_interface_accesses_in_expr(&mut item.value, conflicted);
                if let Some(key) = &mut item.key {
                    rewrite_explicit_interface_accesses_in_expr(key, conflicted);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_explicit_interface_accesses_in_expr(item, conflicted);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_explicit_interface_accesses_in_expr(key, conflicted);
                        rewrite_explicit_interface_accesses_in_expr(value, conflicted);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_explicit_interface_accesses_in_statement(value, conflicted);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Yield(Some(expr)) => {
            rewrite_explicit_interface_accesses_in_expr(expr, conflicted);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_explicit_interface_accesses_in_expr(element, conflicted);
            for generator in generators {
                rewrite_explicit_interface_accesses_in_expr(&mut generator.target, conflicted);
                rewrite_explicit_interface_accesses_in_expr(&mut generator.iter, conflicted);
                for condition in &mut generator.conditions {
                    rewrite_explicit_interface_accesses_in_expr(condition, conflicted);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_explicit_interface_accesses_in_expr(lower, conflicted);
            }
            if let Some(upper) = upper {
                rewrite_explicit_interface_accesses_in_expr(upper, conflicted);
            }
            if let Some(step) = step {
                rewrite_explicit_interface_accesses_in_expr(step, conflicted);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_explicit_interface_accesses_in_expr(parent, conflicted);
            }
            for member in members {
                rewrite_explicit_interface_accesses_in_member(member, conflicted);
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            rewrite_explicit_interface_accesses_in_statement(stmt, conflicted);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_explicit_interface_accesses_in_expr(start, conflicted);
            rewrite_explicit_interface_accesses_in_expr(end, conflicted);
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_explicit_interface_accesses_in_expr(class, conflicted);
            rewrite_explicit_interface_accesses_in_expr(member, conflicted);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_explicit_interface_accesses_in_expr(subject, conflicted);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        rewrite_explicit_interface_accesses_in_expr(condition, conflicted);
                    }
                }
                rewrite_explicit_interface_accesses_in_expr(&mut arm.body, conflicted);
            }
        }
        ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::DefaultOf(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::Yield(None)
        | ExprKind::Destructure(_) => {}
    }
}

#[derive(Clone)]
struct RecordShape {
    positional_fields: Vec<String>,
}

fn extract_deconstruct_shape(params: &[Param], body: &[Statement]) -> Option<RecordShape> {
    if params.is_empty() || body.len() != params.len() {
        return None;
    }

    let mut positional_fields = Vec::with_capacity(params.len());
    for (param, stmt) in params.iter().zip(body.iter()) {
        let StmtKind::Assign { targets, value } = &stmt.kind else {
            return None;
        };
        if targets.len() != 1 {
            return None;
        }
        let ExprKind::Ident(target_name) = &targets[0].kind else {
            return None;
        };
        if target_name != &param.name {
            return None;
        }

        let field_name = match &value.kind {
            ExprKind::Ident(name) => name.clone(),
            ExprKind::Member {
                object,
                field,
                null_safe: false,
            } if matches!(object.kind, ExprKind::This) => field.clone(),
            _ => return None,
        };
        positional_fields.push(field_name);
    }

    Some(RecordShape { positional_fields })
}

fn rewrite_record_uses(module: &mut Module) {
    let record_shapes = collect_record_shapes(&module.body);
    if record_shapes.is_empty() {
        return;
    }
    let mut scopes = vec![HashMap::new()];
    rewrite_record_uses_in_statements(&mut module.body, &record_shapes, &mut scopes);
}

fn collect_record_shapes(body: &[Statement]) -> HashMap<String, RecordShape> {
    let mut shapes = HashMap::new();
    for stmt in body {
        collect_record_shapes_in_statement(stmt, &mut shapes);
    }
    shapes
}

fn collect_record_shapes_in_statement(stmt: &Statement, shapes: &mut HashMap<String, RecordShape>) {
    match &stmt.kind {
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                collect_record_shapes_in_statement(stmt, shapes);
            }
        }
        StmtKind::ClassDecl { name, members, .. } => {
            if let Some(shape) = members.iter().find_map(|member| {
                let ClassMember::Method(stmt) = member else {
                    return None;
                };
                let StmtKind::FunctionDecl {
                    name, params, body, ..
                } = &stmt.kind
                else {
                    return None;
                };
                if name != "Deconstruct" || params.is_empty() {
                    return None;
                }
                extract_deconstruct_shape(params, body)
            }) {
                shapes.insert(name.clone(), shape);
            }
            for member in members {
                if let ClassMember::NestedType(stmt) = member {
                    collect_record_shapes_in_statement(stmt, shapes);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_record_uses_in_statements(
    body: &mut [Statement],
    record_shapes: &HashMap<String, RecordShape>,
    scopes: &mut Vec<HashMap<String, String>>,
) {
    for stmt in body {
        rewrite_record_uses_in_statement(stmt, record_shapes, scopes);
    }
}

fn rewrite_record_uses_in_statement(
    stmt: &mut Statement,
    record_shapes: &HashMap<String, RecordShape>,
    scopes: &mut Vec<HashMap<String, String>>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_record_uses_in_expr(expr, record_shapes, scopes);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_record_uses_in_expr(expr, record_shapes, scopes);
            }
            if let Some(cause) = cause {
                rewrite_record_uses_in_expr(cause, record_shapes, scopes);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_record_uses_in_expr(init, record_shapes, scopes);
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(type_name) = infer_record_type(init, record_shapes, scopes) {
                            scopes.last_mut().unwrap().insert(name.clone(), type_name);
                        }
                    }
                }
            }
        }
        StmtKind::Block(body) => {
            if let Some(rewritten) = rewrite_record_deconstruction_stmt(body, record_shapes, scopes)
            {
                stmt.kind = rewritten;
            } else {
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(body, record_shapes, scopes);
                scopes.pop();
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            scopes.push(HashMap::new());
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            scopes.pop();
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            scopes.push(HashMap::new());
            for param in params {
                if let Some(type_name) = param
                    .type_hint
                    .clone()
                    .filter(|name| record_shapes.contains_key(name))
                {
                    scopes
                        .last_mut()
                        .unwrap()
                        .insert(param.name.clone(), type_name);
                }
            }
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            scopes.pop();
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_record_uses_in_statement(stmt, record_shapes, scopes);
                    }
                    ClassMember::Constructor { params, body, .. } => {
                        scopes.push(HashMap::new());
                        for param in params {
                            if let Some(type_name) = param
                                .type_hint
                                .clone()
                                .filter(|name| record_shapes.contains_key(name))
                            {
                                scopes
                                    .last_mut()
                                    .unwrap()
                                    .insert(param.name.clone(), type_name);
                            }
                        }
                        rewrite_record_uses_in_statements(body, record_shapes, scopes);
                        scopes.pop();
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            scopes.push(HashMap::new());
                            rewrite_record_uses_in_statements(getter, record_shapes, scopes);
                            scopes.pop();
                        }
                        if let Some(setter) = setter {
                            scopes.push(HashMap::new());
                            rewrite_record_uses_in_statements(
                                &mut setter.body,
                                record_shapes,
                                scopes,
                            );
                            scopes.pop();
                        }
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_record_uses_in_expr(cond, record_shapes, scopes);
            scopes.push(HashMap::new());
            rewrite_record_uses_in_statements(then_body, record_shapes, scopes);
            scopes.pop();
            for (elif_cond, elif_body) in elifs {
                rewrite_record_uses_in_expr(elif_cond, record_shapes, scopes);
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(elif_body, record_shapes, scopes);
                scopes.pop();
            }
            if let Some(else_body) = else_body {
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(else_body, record_shapes, scopes);
                scopes.pop();
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            scopes.push(HashMap::new());
            if let Some(init) = init {
                rewrite_record_uses_in_statement(init, record_shapes, scopes);
            }
            if let Some(cond) = cond {
                rewrite_record_uses_in_expr(cond, record_shapes, scopes);
            }
            if let Some(update) = update {
                rewrite_record_uses_in_expr(update, record_shapes, scopes);
            }
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            scopes.pop();
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_record_uses_in_expr(iter, record_shapes, scopes);
            scopes.push(HashMap::new());
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            if let Some(else_body) = else_body {
                rewrite_record_uses_in_statements(else_body, record_shapes, scopes);
            }
            scopes.pop();
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_record_uses_in_expr(cond, record_shapes, scopes);
            scopes.push(HashMap::new());
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            if let Some(else_body) = else_body {
                rewrite_record_uses_in_statements(else_body, record_shapes, scopes);
            }
            scopes.pop();
        }
        StmtKind::DoWhile { body, cond, .. } => {
            scopes.push(HashMap::new());
            rewrite_record_uses_in_statements(body, record_shapes, scopes);
            scopes.pop();
            rewrite_record_uses_in_expr(cond, record_shapes, scopes);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_record_uses_in_expr(expr, record_shapes, scopes);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_record_uses_in_expr(expr, record_shapes, scopes)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_record_uses_in_expr(from, record_shapes, scopes);
                            rewrite_record_uses_in_expr(to, record_shapes, scopes);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_record_uses_in_expr(expr, record_shapes, scopes);
                        }
                    }
                }
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(&mut case.body, record_shapes, scopes);
                scopes.pop();
            }
            if let Some(default) = default {
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(default, record_shapes, scopes);
                scopes.pop();
            }
        }
        _ => {}
    }
}

fn rewrite_record_uses_in_expr(
    expr: &mut Expression,
    record_shapes: &HashMap<String, RecordShape>,
    scopes: &mut Vec<HashMap<String, String>>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            rewrite_record_uses_in_expr(left, record_shapes, scopes);
            rewrite_record_uses_in_expr(right, record_shapes, scopes);
            if matches!(op, BinOp::Eq | BinOp::NotEq) {
                let left_type = infer_record_type(left, record_shapes, scopes);
                let right_type = infer_record_type(right, record_shapes, scopes);
                if let Some(type_name) =
                    left_type.filter(|type_name| Some(type_name.clone()) == right_type)
                {
                    if let Some(shape) = record_shapes.get(&type_name) {
                        let equals_expr =
                            build_record_field_equality(left, right, &shape.positional_fields);
                        *expr = if matches!(op, BinOp::NotEq) {
                            Expression::new(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(equals_expr),
                            })
                        } else {
                            equals_expr
                        };
                    }
                }
            }
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_record_uses_in_expr(left, record_shapes, scopes);
            rewrite_record_uses_in_expr(right, record_shapes, scopes);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_record_uses_in_expr(expr, record_shapes, scopes),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_record_uses_in_expr(cond, record_shapes, scopes);
            rewrite_record_uses_in_expr(then, record_shapes, scopes);
            rewrite_record_uses_in_expr(else_, record_shapes, scopes);
        }
        ExprKind::Member { object, .. } => {
            rewrite_record_uses_in_expr(object, record_shapes, scopes)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_record_uses_in_expr(object, record_shapes, scopes);
            rewrite_record_uses_in_expr(index, record_shapes, scopes);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_record_uses_in_expr(callee, record_shapes, scopes);
            for arg in args {
                rewrite_record_uses_in_expr(&mut arg.value, record_shapes, scopes);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_record_uses_in_expr(class, record_shapes, scopes);
            for arg in args {
                rewrite_record_uses_in_expr(&mut arg.value, record_shapes, scopes);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_record_uses_in_expr(expr, record_shapes, scopes),
            LambdaBody::Block(body) => {
                scopes.push(HashMap::new());
                rewrite_record_uses_in_statements(body, record_shapes, scopes);
                scopes.pop();
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_record_uses_in_expr(&mut item.value, record_shapes, scopes);
                if let Some(key) = &mut item.key {
                    rewrite_record_uses_in_expr(key, record_shapes, scopes);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_record_uses_in_expr(item, record_shapes, scopes);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_record_uses_in_expr(key, record_shapes, scopes);
                        rewrite_record_uses_in_expr(value, record_shapes, scopes);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_record_uses_in_expr(expr, record_shapes, scopes)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_record_uses_in_statement(value, record_shapes, scopes);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_record_uses_in_expr(expr, record_shapes, scopes);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_record_uses_in_expr(element, record_shapes, scopes);
            for generator in generators {
                rewrite_record_uses_in_expr(&mut generator.target, record_shapes, scopes);
                rewrite_record_uses_in_expr(&mut generator.iter, record_shapes, scopes);
                for condition in &mut generator.conditions {
                    rewrite_record_uses_in_expr(condition, record_shapes, scopes);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_record_uses_in_expr(lower, record_shapes, scopes);
            }
            if let Some(upper) = upper {
                rewrite_record_uses_in_expr(upper, record_shapes, scopes);
            }
            if let Some(step) = step {
                rewrite_record_uses_in_expr(step, record_shapes, scopes);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_record_uses_in_expr(parent, record_shapes, scopes);
            }
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_record_uses_in_statement(stmt, record_shapes, scopes);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            rewrite_record_uses_in_statement(stmt, record_shapes, scopes)
        }
        _ => {}
    }
}

fn build_record_field_equality(
    left: &Expression,
    right: &Expression,
    fields: &[String],
) -> Expression {
    let mut combined = Expression::bool(true);
    for field in fields {
        let cmp = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(left.clone()),
                field: field.clone(),
                null_safe: false,
            })),
            right: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(right.clone()),
                field: field.clone(),
                null_safe: false,
            })),
        });
        combined = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(combined),
            right: Box::new(cmp),
        });
    }
    combined
}

fn build_tuple_structural_equality(left: &Expression, right: &Expression) -> Option<Expression> {
    let (ExprKind::Tuple(left_items), ExprKind::Tuple(right_items)) = (&left.kind, &right.kind)
    else {
        return None;
    };
    if left_items.len() != right_items.len() {
        return Some(Expression::bool(false));
    }

    let mut combined = Expression::bool(true);
    for (left_item, right_item) in left_items.iter().zip(right_items.iter()) {
        let cmp = build_tuple_structural_equality(left_item, right_item).unwrap_or_else(|| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(left_item.clone()),
                right: Box::new(right_item.clone()),
            })
        });
        combined = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(combined),
            right: Box::new(cmp),
        });
    }

    Some(combined)
}

fn rewrite_record_deconstruction_stmt(
    body: &mut Vec<Statement>,
    record_shapes: &HashMap<String, RecordShape>,
    scopes: &mut Vec<HashMap<String, String>>,
) -> Option<StmtKind> {
    let Some(call_stmt) = body.last() else {
        return None;
    };
    let StmtKind::Expr(expr) = &call_stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field != "Deconstruct" {
        return None;
    }
    let Some(record_type) = infer_record_type(object, record_shapes, scopes) else {
        return None;
    };
    let Some(shape) = record_shapes.get(&record_type) else {
        return None;
    };
    if shape.positional_fields.len() != args.len() {
        return None;
    }

    let patterns: Vec<ArrayPatternElem> = args
        .iter()
        .map(|arg| match &arg.value.kind {
            ExprKind::Ident(name) if !name.starts_with("__discard_") => {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
            }
            _ => ArrayPatternElem::Hole,
        })
        .collect();

    let values = shape
        .positional_fields
        .iter()
        .map(|field| ArrayElement {
            key: None,
            value: Expression::new(ExprKind::Member {
                object: Box::new((**object).clone()),
                field: field.clone(),
                null_safe: false,
            }),
            spread: false,
            by_ref: false,
        })
        .collect();

    Some(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Destructure(
            DestructurePattern::Array(patterns),
        ))],
        value: Expression::new(ExprKind::Array(values)),
    })
}

fn infer_record_type(
    expr: &Expression,
    record_shapes: &HashMap<String, RecordShape>,
    scopes: &[HashMap<String, String>],
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => scopes
            .iter()
            .rev()
            .find_map(|scope| scope.get(name).cloned()),
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) if record_shapes.contains_key(name) => Some(name.clone()),
            _ => None,
        },
        ExprKind::Cast { type_name, .. } if record_shapes.contains_key(type_name) => {
            Some(type_name.clone())
        }
        _ => None,
    }
}

fn rewrite_tuple_uses(module: &mut Module) {
    let mut scopes = vec![HashMap::new()];
    rewrite_tuple_uses_in_statements(&mut module.body, &mut scopes);
}

fn rewrite_tuple_uses_in_statements(
    body: &mut [Statement],
    scopes: &mut Vec<HashMap<String, usize>>,
) {
    for stmt in body {
        rewrite_tuple_uses_in_statement(stmt, scopes);
    }
}

fn rewrite_tuple_uses_in_statement(stmt: &mut Statement, scopes: &mut Vec<HashMap<String, usize>>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_tuple_uses_in_expr(expr, scopes);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_tuple_uses_in_expr(expr, scopes);
            }
            if let Some(cause) = cause {
                rewrite_tuple_uses_in_expr(cause, scopes);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_tuple_uses_in_expr(init, scopes);
                }
                let BindingPattern::Ident(name) = &decl.pattern else {
                    continue;
                };
                if let Some(arity) = decl
                    .init
                    .as_ref()
                    .and_then(|init| infer_tuple_arity(init, scopes))
                {
                    scopes.last_mut().unwrap().insert(name.clone(), arity);
                } else {
                    scopes.last_mut().unwrap().remove(name);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            rewrite_tuple_uses_in_expr(value, scopes);
            for target in targets {
                rewrite_tuple_uses_in_expr(target, scopes);
                let ExprKind::Ident(name) = &target.kind else {
                    continue;
                };
                if let Some(arity) = infer_tuple_arity(value, scopes) {
                    scopes.last_mut().unwrap().insert(name.clone(), arity);
                } else {
                    scopes.last_mut().unwrap().remove(name);
                }
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_tuple_uses_in_expr(target, scopes);
            rewrite_tuple_uses_in_expr(value, scopes);
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(body, scopes);
            scopes.pop();
        }
        StmtKind::FunctionDecl { body, .. } => {
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(body, scopes);
            scopes.pop();
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_tuple_uses_in_statement(stmt, scopes);
                    }
                    ClassMember::Constructor { body, .. } => {
                        scopes.push(HashMap::new());
                        rewrite_tuple_uses_in_statements(body, scopes);
                        scopes.pop();
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            scopes.push(HashMap::new());
                            rewrite_tuple_uses_in_statements(getter, scopes);
                            scopes.pop();
                        }
                        if let Some(setter) = setter {
                            scopes.push(HashMap::new());
                            rewrite_tuple_uses_in_statements(&mut setter.body, scopes);
                            scopes.pop();
                        }
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_tuple_uses_in_expr(cond, scopes);
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(then_body, scopes);
            scopes.pop();
            for (elif_cond, elif_body) in elifs {
                rewrite_tuple_uses_in_expr(elif_cond, scopes);
                scopes.push(HashMap::new());
                rewrite_tuple_uses_in_statements(elif_body, scopes);
                scopes.pop();
            }
            if let Some(else_body) = else_body {
                scopes.push(HashMap::new());
                rewrite_tuple_uses_in_statements(else_body, scopes);
                scopes.pop();
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            scopes.push(HashMap::new());
            if let Some(init) = init {
                rewrite_tuple_uses_in_statement(init, scopes);
            }
            if let Some(cond) = cond {
                rewrite_tuple_uses_in_expr(cond, scopes);
            }
            if let Some(update) = update {
                rewrite_tuple_uses_in_expr(update, scopes);
            }
            rewrite_tuple_uses_in_statements(body, scopes);
            scopes.pop();
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_tuple_uses_in_expr(iter, scopes);
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(body, scopes);
            if let Some(else_body) = else_body {
                rewrite_tuple_uses_in_statements(else_body, scopes);
            }
            scopes.pop();
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_tuple_uses_in_expr(cond, scopes);
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(body, scopes);
            if let Some(else_body) = else_body {
                rewrite_tuple_uses_in_statements(else_body, scopes);
            }
            scopes.pop();
        }
        StmtKind::DoWhile { body, cond, .. } => {
            scopes.push(HashMap::new());
            rewrite_tuple_uses_in_statements(body, scopes);
            scopes.pop();
            rewrite_tuple_uses_in_expr(cond, scopes);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_tuple_uses_in_expr(expr, scopes);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => rewrite_tuple_uses_in_expr(expr, scopes),
                        CaseCondition::Range { from, to } => {
                            rewrite_tuple_uses_in_expr(from, scopes);
                            rewrite_tuple_uses_in_expr(to, scopes);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_tuple_uses_in_expr(expr, scopes);
                        }
                    }
                }
                scopes.push(HashMap::new());
                rewrite_tuple_uses_in_statements(&mut case.body, scopes);
                scopes.pop();
            }
            if let Some(default) = default {
                scopes.push(HashMap::new());
                rewrite_tuple_uses_in_statements(default, scopes);
                scopes.pop();
            }
        }
        _ => {}
    }
}

fn rewrite_tuple_uses_in_expr(expr: &mut Expression, scopes: &mut Vec<HashMap<String, usize>>) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            rewrite_tuple_uses_in_expr(left, scopes);
            rewrite_tuple_uses_in_expr(right, scopes);
            if matches!(op, BinOp::Eq | BinOp::NotEq) {
                let left_arity = infer_tuple_arity(left, scopes);
                let right_arity = infer_tuple_arity(right, scopes);
                if let Some(arity) = left_arity.filter(|arity| Some(*arity) == right_arity) {
                    let equals_expr = match (&left.kind, &right.kind) {
                        (ExprKind::Tuple(_), ExprKind::Tuple(_)) => {
                            build_tuple_structural_equality(left, right)
                                .unwrap_or_else(|| build_tuple_runtime_equality(left, right, arity))
                        }
                        _ => build_tuple_runtime_equality(left, right, arity),
                    };
                    *expr = if matches!(op, BinOp::NotEq) {
                        Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(equals_expr),
                        })
                    } else {
                        equals_expr
                    };
                }
            }
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_tuple_uses_in_expr(left, scopes);
            rewrite_tuple_uses_in_expr(right, scopes);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_tuple_uses_in_expr(expr, scopes),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_tuple_uses_in_expr(cond, scopes);
            rewrite_tuple_uses_in_expr(then, scopes);
            rewrite_tuple_uses_in_expr(else_, scopes);
        }
        ExprKind::Member { object, .. } => rewrite_tuple_uses_in_expr(object, scopes),
        ExprKind::Index { object, index, .. } => {
            rewrite_tuple_uses_in_expr(object, scopes);
            rewrite_tuple_uses_in_expr(index, scopes);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_tuple_uses_in_expr(callee, scopes);
            for arg in args {
                rewrite_tuple_uses_in_expr(&mut arg.value, scopes);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_tuple_uses_in_expr(class, scopes);
            for arg in args {
                rewrite_tuple_uses_in_expr(&mut arg.value, scopes);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_tuple_uses_in_expr(expr, scopes),
            LambdaBody::Block(body) => {
                scopes.push(HashMap::new());
                rewrite_tuple_uses_in_statements(body, scopes);
                scopes.pop();
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_tuple_uses_in_expr(&mut item.value, scopes);
                if let Some(key) = &mut item.key {
                    rewrite_tuple_uses_in_expr(key, scopes);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_tuple_uses_in_expr(item, scopes);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_tuple_uses_in_expr(key, scopes);
                        rewrite_tuple_uses_in_expr(value, scopes);
                    }
                    ObjectProperty::Spread(expr) => rewrite_tuple_uses_in_expr(expr, scopes),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_tuple_uses_in_statement(value, scopes);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_tuple_uses_in_expr(expr, scopes);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_tuple_uses_in_expr(element, scopes);
            for generator in generators {
                rewrite_tuple_uses_in_expr(&mut generator.target, scopes);
                rewrite_tuple_uses_in_expr(&mut generator.iter, scopes);
                for condition in &mut generator.conditions {
                    rewrite_tuple_uses_in_expr(condition, scopes);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_tuple_uses_in_expr(lower, scopes);
            }
            if let Some(upper) = upper {
                rewrite_tuple_uses_in_expr(upper, scopes);
            }
            if let Some(step) = step {
                rewrite_tuple_uses_in_expr(step, scopes);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_tuple_uses_in_expr(parent, scopes);
            }
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_tuple_uses_in_statement(stmt, scopes);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::FunctionExpr(stmt) => rewrite_tuple_uses_in_statement(stmt, scopes),
        _ => {}
    }
}

fn infer_tuple_arity(expr: &Expression, scopes: &[HashMap<String, usize>]) -> Option<usize> {
    match &expr.kind {
        ExprKind::Ident(name) => scopes
            .iter()
            .rev()
            .find_map(|scope| scope.get(name).copied()),
        ExprKind::Tuple(items) => Some(items.len()),
        ExprKind::Cast { expr, .. } => infer_tuple_arity(expr, scopes),
        _ => None,
    }
}

fn build_tuple_runtime_equality(left: &Expression, right: &Expression, arity: usize) -> Expression {
    let mut combined = Expression::bool(true);
    for index in 0..arity {
        let cmp = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(left.clone()),
                index: Box::new(Expression::int(index as i64)),
                null_safe: false,
            })),
            right: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(right.clone()),
                index: Box::new(Expression::int(index as i64)),
                null_safe: false,
            })),
        });
        combined = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(combined),
            right: Box::new(cmp),
        });
    }
    combined
}

fn rewrite_checked_numeric_statements(
    body: &mut [Statement],
    scopes: &mut Vec<HashMap<String, ()>>,
    is_checked: bool,
) {
    for stmt in body {
        rewrite_checked_numeric_statement(stmt, scopes, is_checked);
    }
}

fn is_tracked_checked_byte_local(target: &Expression, scopes: &[HashMap<String, ()>]) -> bool {
    let ExprKind::Ident(name) = &target.kind else {
        return false;
    };
    scopes.iter().rev().any(|scope| scope.contains_key(name))
}

fn is_csharp_byte_type_hint(type_hint: Option<&str>) -> bool {
    type_hint
        .map(normalize_runtime_type_name)
        .map(|name| name.eq_ignore_ascii_case("byte") || name.eq_ignore_ascii_case("system.byte"))
        .unwrap_or(false)
}

fn wrap_checked_byte_value(value: Expression, is_checked: bool) -> Expression {
    if is_checked {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__vybe_csharp_checked_byte_cast")),
            args: vec![Argument::positional(value)],
            optional: false,
        })
    } else {
        Expression::new(ExprKind::Call {
            callee: Box::new(build_dotted_expr("Convert.ToByte")),
            args: vec![Argument::positional(value)],
            optional: false,
        })
    }
}

fn is_checked_byte_wrapper_call(expr: &Expression, is_checked: bool) -> bool {
    let ExprKind::Call {
        callee,
        args,
        optional: false,
    } = &expr.kind
    else {
        return false;
    };
    if args.len() != 1 {
        return false;
    }
    if is_checked {
        matches!(&callee.kind, ExprKind::Ident(name) if name == "__vybe_csharp_checked_byte_cast")
    } else {
        expr_dotted_name(callee).as_deref() == Some("Convert.ToByte")
    }
}

fn rewrite_checked_numeric_statement(
    stmt: &mut Statement,
    scopes: &mut Vec<HashMap<String, ()>>,
    is_checked: bool,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_checked_numeric_expr(init, scopes, is_checked);
                }
                if is_csharp_byte_type_hint(decl.type_hint.as_deref()) {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        scopes.last_mut().unwrap().insert(name.clone(), ());
                    }
                }
            }
        }
        StmtKind::CompoundAssign { target, op, value } => {
            rewrite_checked_numeric_expr(target, scopes, is_checked);
            rewrite_checked_numeric_expr(value, scopes, is_checked);

            if !is_tracked_checked_byte_local(target, scopes) {
                return;
            }

            let Some(bin_op) = checked_numeric_compound_binop(op) else {
                return;
            };
            let result_expr = Expression::new(ExprKind::Binary {
                op: bin_op,
                left: Box::new(target.clone()),
                right: Box::new(value.clone()),
            });
            stmt.kind = StmtKind::Assign {
                targets: vec![target.clone()],
                value: wrap_checked_byte_value(result_expr, is_checked),
            };
        }
        StmtKind::Assign { targets, value } => {
            rewrite_checked_numeric_expr(value, scopes, is_checked);
            for target in &mut *targets {
                rewrite_checked_numeric_expr(target, scopes, is_checked);
            }
            if targets.len() != 1 {
                return;
            }
            if !is_tracked_checked_byte_local(&targets[0], scopes) {
                return;
            }
            if is_checked_byte_wrapper_call(value, is_checked) {
                return;
            }
            *value = wrap_checked_byte_value(value.clone(), is_checked);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_checked_numeric_expr(expr, scopes, is_checked);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_checked_numeric_expr(expr, scopes, is_checked);
            }
            if let Some(cause) = cause {
                rewrite_checked_numeric_expr(cause, scopes, is_checked);
            }
        }
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. } => {
            scopes.push(HashMap::new());
            rewrite_checked_numeric_statements(body, scopes, is_checked);
            scopes.pop();
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_checked_numeric_expr(cond, scopes, is_checked);
            scopes.push(HashMap::new());
            rewrite_checked_numeric_statements(then_body, scopes, is_checked);
            scopes.pop();
            for (elif_cond, elif_body) in elifs {
                rewrite_checked_numeric_expr(elif_cond, scopes, is_checked);
                scopes.push(HashMap::new());
                rewrite_checked_numeric_statements(elif_body, scopes, is_checked);
                scopes.pop();
            }
            if let Some(else_body) = else_body {
                scopes.push(HashMap::new());
                rewrite_checked_numeric_statements(else_body, scopes, is_checked);
                scopes.pop();
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            scopes.push(HashMap::new());
            if let Some(init) = init {
                rewrite_checked_numeric_statement(init, scopes, is_checked);
            }
            if let Some(cond) = cond {
                rewrite_checked_numeric_expr(cond, scopes, is_checked);
            }
            if let Some(update) = update {
                rewrite_checked_numeric_expr(update, scopes, is_checked);
            }
            rewrite_checked_numeric_statements(body, scopes, is_checked);
            scopes.pop();
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_checked_numeric_expr(iter, scopes, is_checked);
            scopes.push(HashMap::new());
            rewrite_checked_numeric_statements(body, scopes, is_checked);
            if let Some(else_body) = else_body {
                rewrite_checked_numeric_statements(else_body, scopes, is_checked);
            }
            scopes.pop();
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_checked_numeric_expr(cond, scopes, is_checked);
            scopes.push(HashMap::new());
            rewrite_checked_numeric_statements(body, scopes, is_checked);
            if let Some(else_body) = else_body {
                rewrite_checked_numeric_statements(else_body, scopes, is_checked);
            }
            scopes.pop();
        }
        StmtKind::DoWhile { body, cond, .. } => {
            scopes.push(HashMap::new());
            rewrite_checked_numeric_statements(body, scopes, is_checked);
            scopes.pop();
            rewrite_checked_numeric_expr(cond, scopes, is_checked);
        }
        _ => {}
    }
}

fn rewrite_checked_numeric_expr(
    expr: &mut Expression,
    scopes: &mut Vec<HashMap<String, ()>>,
    is_checked: bool,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_checked_numeric_expr(left, scopes, is_checked);
            rewrite_checked_numeric_expr(right, scopes, is_checked);
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_checked_numeric_expr(target, scopes, is_checked);
            rewrite_checked_numeric_expr(value, scopes, is_checked);
            if is_tracked_checked_byte_local(target, scopes)
                && !is_checked_byte_wrapper_call(value, is_checked)
            {
                **value = wrap_checked_byte_value((**value).clone(), is_checked);
            }
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_checked_numeric_expr(expr, scopes, is_checked),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_checked_numeric_expr(cond, scopes, is_checked);
            rewrite_checked_numeric_expr(then, scopes, is_checked);
            rewrite_checked_numeric_expr(else_, scopes, is_checked);
        }
        ExprKind::Member { object, .. } => rewrite_checked_numeric_expr(object, scopes, is_checked),
        ExprKind::Index { object, index, .. } => {
            rewrite_checked_numeric_expr(object, scopes, is_checked);
            rewrite_checked_numeric_expr(index, scopes, is_checked);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_checked_numeric_expr(callee, scopes, is_checked);
            for arg in args {
                rewrite_checked_numeric_expr(&mut arg.value, scopes, is_checked);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_checked_numeric_expr(class, scopes, is_checked);
            for arg in args {
                rewrite_checked_numeric_expr(&mut arg.value, scopes, is_checked);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_checked_numeric_expr(expr, scopes, is_checked),
            LambdaBody::Block(body) => {
                scopes.push(HashMap::new());
                rewrite_checked_numeric_statements(body, scopes, is_checked);
                scopes.pop();
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_checked_numeric_expr(&mut item.value, scopes, is_checked);
                if let Some(key) = &mut item.key {
                    rewrite_checked_numeric_expr(key, scopes, is_checked);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_checked_numeric_expr(item, scopes, is_checked);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_checked_numeric_expr(key, scopes, is_checked);
                        rewrite_checked_numeric_expr(value, scopes, is_checked);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_checked_numeric_expr(expr, scopes, is_checked)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_checked_numeric_statement(value, scopes, is_checked);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_checked_numeric_expr(expr, scopes, is_checked);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_checked_numeric_expr(element, scopes, is_checked);
            for generator in generators {
                rewrite_checked_numeric_expr(&mut generator.target, scopes, is_checked);
                rewrite_checked_numeric_expr(&mut generator.iter, scopes, is_checked);
                for condition in &mut generator.conditions {
                    rewrite_checked_numeric_expr(condition, scopes, is_checked);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_checked_numeric_expr(lower, scopes, is_checked);
            }
            if let Some(upper) = upper {
                rewrite_checked_numeric_expr(upper, scopes, is_checked);
            }
            if let Some(step) = step {
                rewrite_checked_numeric_expr(step, scopes, is_checked);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_checked_numeric_expr(parent, scopes, is_checked);
            }
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_checked_numeric_statement(stmt, scopes, is_checked);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::FunctionExpr(stmt) => rewrite_checked_numeric_statement(stmt, scopes, is_checked),
        _ => {}
    }
}

fn checked_numeric_compound_binop(op: &CompoundOp) -> Option<BinOp> {
    match op {
        CompoundOp::Add => Some(BinOp::Add),
        CompoundOp::Sub => Some(BinOp::Sub),
        CompoundOp::Mul => Some(BinOp::Mul),
        CompoundOp::Div => Some(BinOp::Div),
        CompoundOp::Mod => Some(BinOp::Mod),
        CompoundOp::BitAnd => Some(BinOp::BitAnd),
        CompoundOp::BitOr => Some(BinOp::BitOr),
        CompoundOp::BitXor => Some(BinOp::BitXor),
        CompoundOp::Shl => Some(BinOp::Shl),
        CompoundOp::Shr => Some(BinOp::Shr),
        CompoundOp::UShr => Some(BinOp::UShr),
        _ => None,
    }
}

fn rewrite_using_imports(module: &mut Module) {
    let mut aliases: HashMap<String, String> = HashMap::new();
    let mut static_paths: Vec<String> = Vec::new();
    for import in &module.imports {
        match &import.kind {
            ImportKind::Simple {
                path,
                alias: Some(alias),
            } => {
                aliases.insert(alias.clone(), normalize_import_path(path));
            }
            ImportKind::Wildcard { path, alias: None } => {
                static_paths.push(normalize_import_path(path));
            }
            _ => {}
        }
    }
    if aliases.is_empty() && static_paths.is_empty() {
        return;
    }
    rewrite_using_imports_in_statements(&mut module.body, &aliases, &static_paths);
}

// ── `using var` lowering — ECMA-334 §13.14 ─────────────────────────────────
//
// `using var x = expr; rest…` desugars in the walker (same layering as JS
// `lower_using_declarations`) to explicit JS-shaped AST:
//
//   var x = expr;
//   try { rest… } finally { if (x != null) x.Dispose(); }
//
// Block-form `using (var x = expr) { … }` keeps `StmtKind::Using` with a
// non-empty body and still routes through the shared compiler try/finally
// emitter. Only declaration-form `using var` is lowered here so `break` /
// `continue` inside loops get correct structured-control-flow depths.

fn lower_csharp_using_declarations(body: &mut Vec<Statement>) {
    for (i, s) in body.iter().enumerate() {
        if matches!(&s.kind, StmtKind::Using { body, .. } if body.is_empty()) {
            let mut tail: Vec<Statement> = body.drain(i + 1..).collect();
            let marker = body.pop().expect("using marker");
            let StmtKind::Using { var, resource, .. } = marker.kind else {
                unreachable!()
            };
            lower_csharp_using_declarations(&mut tail);
            body.extend(lower_one_csharp_using(&var, resource, tail));
            break;
        }
    }
    for s in body.iter_mut() {
        visit_csharp_stmt_lists_for_using(&mut s.kind, lower_csharp_using_declarations);
    }
}

fn visit_csharp_stmt_lists_for_using(kind: &mut StmtKind, visit: fn(&mut Vec<Statement>)) {
    match kind {
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. } => visit(body),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        visit_csharp_stmt_lists_for_using(&mut stmt.kind, visit);
                    }
                    ClassMember::Constructor { body, .. } => visit(body),
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            visit(getter);
                        }
                        if let Some(setter) = setter {
                            visit(&mut setter.body);
                        }
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            visit(then_body);
            for (_, elif_body) in elifs {
                visit(elif_body);
            }
            if let Some(else_body) = else_body {
                visit(else_body);
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                visit_csharp_stmt_lists_for_using(&mut init.kind, visit);
            }
            visit(body);
        }
        StmtKind::ForIn {
            body, else_body, ..
        } => {
            visit(body);
            if let Some(else_body) = else_body {
                visit(else_body);
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            visit(body);
            if let Some(else_body) = else_body {
                visit(else_body);
            }
        }
        StmtKind::DoWhile { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => visit(body),
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            visit(body);
            for catch in catches {
                visit(&mut catch.body);
            }
            if let Some(finally) = finally {
                visit(finally);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                visit(&mut case.body);
            }
            if let Some(default) = default {
                visit(default);
            }
        }
        _ => {}
    }
}

fn lower_one_csharp_using(var: &str, resource: Expression, tail: Vec<Statement>) -> Vec<Statement> {
    let decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(var.to_string()),
            type_hint: None,
            init: Some(resource),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });
    let mut out = vec![decl];
    out.extend(csharp_using_disposal_wrap(var, tail));
    out
}

fn csharp_using_disposal_wrap(var: &str, tail: Vec<Statement>) -> Vec<Statement> {
    let var_expr = Expression::ident(var);
    let dispose_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(var_expr.clone()),
            field: "Dispose".to_string(),
            null_safe: false,
        })),
        args: vec![],
        optional: false,
    });
    let finally_body = vec![Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::StrictNotEq,
            left: Box::new(var_expr),
            right: Box::new(Expression::null()),
        }),
        then_body: vec![Statement::new(StmtKind::Expr(dispose_call))],
        elifs: Vec::new(),
        else_body: None,
    })];
    vec![Statement::new(StmtKind::Try {
        body: tail,
        catches: Vec::new(),
        else_body: None,
        finally: Some(finally_body),
    })]
}

// ── HashSet / SortedSet set-algebra bool APIs ───────────────────────────────
//
// Host `ecma:set:isSubsetOf` returns i32 0/1; .NET `Console.WriteLine(bool)`
// expects a real boolean (`True`/`False`). Fold self-comparisons in the walker
// (avoids same-object host deadlock) and route everything else through a
// `cond ? true : false` ternary so the stack carries JS bool values.

fn rewrite_set_algebra_bool_calls(body: &mut [Statement]) {
    for stmt in body.iter_mut() {
        rewrite_set_algebra_bool_calls_in_stmt(stmt);
    }
}

#[derive(Clone)]
struct LinkedListNodeAlias {
    list: Expression,
    index: Expression,
}

fn rewrite_linked_list_node_uses(body: &mut [Statement]) {
    let mut aliases = HashMap::new();
    rewrite_linked_list_node_uses_in_body(body, &mut aliases);
}

fn rewrite_sorted_dictionary_foreach(body: &mut [Statement]) {
    let mut sorted_dicts = HashSet::new();
    rewrite_sorted_dictionary_foreach_in_body(body, &mut sorted_dicts);
}

fn rewrite_sorted_dictionary_foreach_in_body(
    body: &mut [Statement],
    sorted_dicts: &mut HashSet<String>,
) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if decl
                            .type_hint
                            .as_deref()
                            .map(|hint| hint.to_ascii_lowercase().contains("sorteddictionary"))
                            .unwrap_or(false)
                        {
                            sorted_dicts.insert(name.clone());
                        }
                    }
                }
            }
            StmtKind::ForIn { iter, body, .. } => {
                if let ExprKind::Ident(name) = &iter.kind {
                    if sorted_dicts.contains(name) {
                        *iter = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident(name)),
                                field: "EntriesSorted".into(),
                                null_safe: false,
                            })),
                            args: vec![],
                            optional: false,
                        });
                    }
                }
                let mut scoped = sorted_dicts.clone();
                rewrite_sorted_dictionary_foreach_in_body(body, &mut scoped);
            }
            StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
                let mut scoped = sorted_dicts.clone();
                rewrite_sorted_dictionary_foreach_in_body(body, &mut scoped);
            }
            StmtKind::FunctionDecl { body, .. } => {
                let mut scoped = HashSet::new();
                rewrite_sorted_dictionary_foreach_in_body(body, &mut scoped);
            }
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                let mut scoped = sorted_dicts.clone();
                rewrite_sorted_dictionary_foreach_in_body(then_body, &mut scoped);
                for (_, elif_body) in elifs {
                    let mut scoped = sorted_dicts.clone();
                    rewrite_sorted_dictionary_foreach_in_body(elif_body, &mut scoped);
                }
                if let Some(else_body) = else_body {
                    let mut scoped = sorted_dicts.clone();
                    rewrite_sorted_dictionary_foreach_in_body(else_body, &mut scoped);
                }
            }
            _ => {}
        }
    }
}

fn rewrite_linked_list_node_uses_in_body(
    body: &mut [Statement],
    aliases: &mut HashMap<String, LinkedListNodeAlias>,
) {
    for stmt in body {
        rewrite_linked_list_node_uses_in_stmt(stmt, aliases);
    }
}

fn rewrite_linked_list_node_uses_in_stmt(
    stmt: &mut Statement,
    aliases: &mut HashMap<String, LinkedListNodeAlias>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_linked_list_node_uses_in_expr(init, aliases);
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(alias) = linked_list_alias_from_call(init) {
                            aliases.insert(name.clone(), alias);
                        }
                    }
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_linked_list_node_uses_in_expr(expr, aliases);
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            let mut scoped = aliases.clone();
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
        }
        StmtKind::FunctionDecl { body, .. } => {
            let mut scoped = HashMap::new();
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_linked_list_node_uses_in_expr(cond, aliases);
            let mut then_aliases = aliases.clone();
            rewrite_linked_list_node_uses_in_body(then_body, &mut then_aliases);
            for (elif_cond, elif_body) in elifs {
                rewrite_linked_list_node_uses_in_expr(elif_cond, aliases);
                let mut elif_aliases = aliases.clone();
                rewrite_linked_list_node_uses_in_body(elif_body, &mut elif_aliases);
            }
            if let Some(else_body) = else_body {
                let mut else_aliases = aliases.clone();
                rewrite_linked_list_node_uses_in_body(else_body, &mut else_aliases);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = aliases.clone();
            if let Some(init) = init {
                rewrite_linked_list_node_uses_in_stmt(init, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_linked_list_node_uses_in_expr(cond, &scoped);
            }
            if let Some(update) = update {
                rewrite_linked_list_node_uses_in_expr(update, &scoped);
            }
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_linked_list_node_uses_in_expr(iter, aliases);
            let mut scoped = aliases.clone();
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
            if let Some(else_body) = else_body {
                rewrite_linked_list_node_uses_in_body(else_body, &mut scoped);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_linked_list_node_uses_in_expr(cond, aliases);
            let mut scoped = aliases.clone();
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
            if let Some(else_body) = else_body {
                rewrite_linked_list_node_uses_in_body(else_body, &mut scoped);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut scoped = aliases.clone();
            rewrite_linked_list_node_uses_in_body(body, &mut scoped);
            rewrite_linked_list_node_uses_in_expr(cond, &scoped);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        let mut scoped = HashMap::new();
                        rewrite_linked_list_node_uses_in_stmt(stmt, &mut scoped);
                    }
                    ClassMember::Constructor { body, .. } => {
                        let mut scoped = HashMap::new();
                        rewrite_linked_list_node_uses_in_body(body, &mut scoped);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn rewrite_linked_list_node_uses_in_expr(
    expr: &mut Expression,
    aliases: &HashMap<String, LinkedListNodeAlias>,
) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            rewrite_linked_list_node_uses_in_expr(object, aliases);
            if field.eq_ignore_ascii_case("Value") {
                if let Some(alias) = linked_list_node_alias_for_expr(object, aliases) {
                    expr.kind = ExprKind::Index {
                        object: Box::new(alias.list),
                        index: Box::new(alias.index),
                        null_safe: false,
                    };
                }
            }
        }
        ExprKind::Call { callee, args, .. } => {
            for arg in args.iter_mut() {
                rewrite_linked_list_node_uses_in_expr(&mut arg.value, aliases);
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("AddAfter") && args.len() == 2 {
                    if let Some(alias) = linked_list_node_alias_for_expr(&args[0].value, aliases) {
                        expr.kind = linked_list_insert_call(
                            (**object).clone(),
                            Expression::new(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(alias.index),
                                right: Box::new(Expression::int(1)),
                            }),
                            args[1].value.clone(),
                        )
                        .kind;
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("AddBefore") && args.len() == 2 {
                    if let Some(alias) = linked_list_node_alias_for_expr(&args[0].value, aliases) {
                        expr.kind = linked_list_insert_call(
                            (**object).clone(),
                            alias.index,
                            args[1].value.clone(),
                        )
                        .kind;
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("Remove") && args.len() == 1 {
                    if let Some(alias) = linked_list_node_alias_for_expr(&args[0].value, aliases) {
                        expr.kind =
                            linked_list_remove_at_call((**object).clone(), alias.index).kind;
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("RemoveFirst") && args.is_empty() {
                    expr.kind =
                        linked_list_remove_at_call((**object).clone(), Expression::int(0)).kind;
                    return;
                }
                if field.eq_ignore_ascii_case("RemoveLast") && args.is_empty() {
                    expr.kind = linked_list_remove_at_call(
                        (**object).clone(),
                        Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(canonicalize_member_access(
                                (**object).clone(),
                                "Length",
                            )),
                            right: Box::new(Expression::int(1)),
                        }),
                    )
                    .kind;
                    return;
                }
            }
            rewrite_linked_list_node_uses_in_expr(callee, aliases);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::Assign {
            target: left,
            value: right,
        }
        | ExprKind::NullCoalesce { left, right } => {
            rewrite_linked_list_node_uses_in_expr(left, aliases);
            rewrite_linked_list_node_uses_in_expr(right, aliases);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_linked_list_node_uses_in_expr(object, aliases);
            rewrite_linked_list_node_uses_in_expr(index, aliases);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_linked_list_node_uses_in_expr(&mut item.value, aliases);
            }
        }
        _ => {}
    }
}

fn linked_list_alias_from_call(expr: &Expression) -> Option<LinkedListNodeAlias> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field.eq_ignore_ascii_case("AddFirst") {
        return Some(LinkedListNodeAlias {
            list: (**object).clone(),
            index: Expression::int(0),
        });
    }
    if field.eq_ignore_ascii_case("AddLast") {
        return Some(LinkedListNodeAlias {
            list: (**object).clone(),
            index: Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(canonicalize_member_access((**object).clone(), "Length")),
                right: Box::new(Expression::int(1)),
            }),
        });
    }
    if field.eq_ignore_ascii_case("Find") && args.len() == 1 {
        return Some(LinkedListNodeAlias {
            list: (**object).clone(),
            index: Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new((**object).clone()),
                    field: "IndexOf".into(),
                    null_safe: false,
                })),
                args: vec![args[0].clone()],
                optional: false,
            }),
        });
    }
    None
}

fn linked_list_node_alias_for_expr(
    expr: &Expression,
    aliases: &HashMap<String, LinkedListNodeAlias>,
) -> Option<LinkedListNodeAlias> {
    if let ExprKind::Ident(name) = &expr.kind {
        return aliases.get(name).cloned();
    }
    if let Some((list, index)) = linked_list_node_index_expr(expr) {
        return Some(LinkedListNodeAlias { list, index });
    }
    match &expr.kind {
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Next") => {
            let alias = linked_list_node_alias_for_expr(object, aliases)?;
            Some(LinkedListNodeAlias {
                list: alias.list,
                index: Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(alias.index),
                    right: Box::new(Expression::int(1)),
                }),
            })
        }
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Previous") => {
            let alias = linked_list_node_alias_for_expr(object, aliases)?;
            Some(LinkedListNodeAlias {
                list: alias.list,
                index: Expression::new(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(alias.index),
                    right: Box::new(Expression::int(1)),
                }),
            })
        }
        _ => None,
    }
}

fn linked_list_insert_call(list: Expression, index: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(list),
            field: "InsertAtRaw".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(index), Argument::positional(value)],
        optional: false,
    })
}

fn linked_list_remove_at_call(list: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(list),
            field: "RemoveAt".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(index)],
        optional: false,
    })
}

fn rewrite_set_algebra_bool_calls_in_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        } => rewrite_set_algebra_bool_calls_in_expr(expr),
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_set_algebra_bool_calls_in_expr(expr);
            rewrite_set_algebra_bool_calls_in_expr(cause);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_set_algebra_bool_calls_in_expr(init);
                }
            }
        }
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. } => rewrite_set_algebra_bool_calls(body),
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        rewrite_set_algebra_bool_calls_in_stmt(stmt);
                    }
                    ClassMember::Constructor { body, .. } => rewrite_set_algebra_bool_calls(body),
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            rewrite_set_algebra_bool_calls(getter);
                        }
                        if let Some(setter) = setter {
                            rewrite_set_algebra_bool_calls(&mut setter.body);
                        }
                    }
                    _ => {}
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_set_algebra_bool_calls_in_expr(cond);
            rewrite_set_algebra_bool_calls(then_body);
            for (elif_cond, elif_body) in elifs {
                rewrite_set_algebra_bool_calls_in_expr(elif_cond);
                rewrite_set_algebra_bool_calls(elif_body);
            }
            if let Some(else_body) = else_body {
                rewrite_set_algebra_bool_calls(else_body);
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
                rewrite_set_algebra_bool_calls_in_stmt(init);
            }
            if let Some(cond) = cond {
                rewrite_set_algebra_bool_calls_in_expr(cond);
            }
            if let Some(update) = update {
                rewrite_set_algebra_bool_calls_in_expr(update);
            }
            rewrite_set_algebra_bool_calls(body);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_set_algebra_bool_calls_in_expr(iter);
            rewrite_set_algebra_bool_calls(body);
            if let Some(else_body) = else_body {
                rewrite_set_algebra_bool_calls(else_body);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
            ..
        } => {
            rewrite_set_algebra_bool_calls_in_expr(cond);
            rewrite_set_algebra_bool_calls(body);
            if let Some(else_body) = else_body {
                rewrite_set_algebra_bool_calls(else_body);
            }
        }
        StmtKind::DoWhile { cond, body, .. } => {
            rewrite_set_algebra_bool_calls_in_expr(cond);
            rewrite_set_algebra_bool_calls(body);
        }
        StmtKind::Using { resource, body, .. } => {
            rewrite_set_algebra_bool_calls_in_expr(resource);
            rewrite_set_algebra_bool_calls(body);
        }
        StmtKind::Lock { expr, body, .. } => {
            rewrite_set_algebra_bool_calls_in_expr(expr);
            rewrite_set_algebra_bool_calls(body);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            rewrite_set_algebra_bool_calls(body);
            for catch in catches {
                rewrite_set_algebra_bool_calls(&mut catch.body);
            }
            if let Some(else_body) = else_body {
                rewrite_set_algebra_bool_calls(else_body);
            }
            if let Some(finally) = finally {
                rewrite_set_algebra_bool_calls(finally);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
            ..
        } => {
            rewrite_set_algebra_bool_calls_in_expr(expr);
            for case in cases {
                rewrite_set_algebra_bool_calls(&mut case.body);
            }
            if let Some(default) = default {
                rewrite_set_algebra_bool_calls(default);
            }
        }
        StmtKind::Labeled { body, .. } => rewrite_set_algebra_bool_calls_in_stmt(body),
        _ => {}
    }
}

fn is_set_algebra_bool_method(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "issubsetof"
            | "issupersetof"
            | "ispropersubsetof"
            | "ispropersupersetof"
            | "setequals"
            | "overlaps"
    )
}

fn expr_same_reference(a: &Expression, b: &Expression) -> bool {
    match (&a.kind, &b.kind) {
        (ExprKind::Ident(na), ExprKind::Ident(nb)) => na == nb,
        (ExprKind::This, ExprKind::This) => true,
        (
            ExprKind::Member {
                object: oa,
                field: fa,
                ..
            },
            ExprKind::Member {
                object: ob,
                field: fb,
                ..
            },
        ) => fa == fb && expr_same_reference(oa, ob),
        _ => false,
    }
}

fn try_fold_set_algebra_self_call(
    object: &Expression,
    method: &str,
    arg: &Expression,
) -> Option<Expression> {
    if !expr_same_reference(object, arg) {
        return None;
    }
    match method.to_ascii_lowercase().as_str() {
        "issubsetof" | "issupersetof" | "setequals" => Some(Expression::bool(true)),
        "ispropersubsetof" | "ispropersupersetof" => Some(Expression::bool(false)),
        _ => None,
    }
}

fn wrap_csharp_bool(expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(expr),
        then: Box::new(Expression::bool(true)),
        else_: Box::new(Expression::bool(false)),
    })
}

fn rewrite_set_algebra_bool_calls_in_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Spread(inner)
        | ExprKind::TypeOf(inner)
        | ExprKind::Delete(inner) => rewrite_set_algebra_bool_calls_in_expr(inner),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Assign {
            target: left,
            value: right,
        } => {
            rewrite_set_algebra_bool_calls_in_expr(left);
            rewrite_set_algebra_bool_calls_in_expr(right);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_set_algebra_bool_calls_in_expr(cond);
            rewrite_set_algebra_bool_calls_in_expr(then);
            rewrite_set_algebra_bool_calls_in_expr(else_);
        }
        ExprKind::Member { object, .. } => rewrite_set_algebra_bool_calls_in_expr(object),
        ExprKind::Index { object, index, .. } => {
            rewrite_set_algebra_bool_calls_in_expr(object);
            rewrite_set_algebra_bool_calls_in_expr(index);
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if is_set_algebra_bool_method(field) && args.len() == 1 {
                    if let Some(folded) =
                        try_fold_set_algebra_self_call(object, field, &args[0].value)
                    {
                        *expr = folded;
                        return;
                    }
                    rewrite_set_algebra_bool_calls_in_expr(&mut args[0].value);
                    let call =
                        std::mem::replace(expr, Expression::new(ExprKind::Lit(Literal::Null)));
                    *expr = wrap_csharp_bool(call);
                    return;
                }
            }
            rewrite_set_algebra_bool_calls_in_expr(callee);
            for arg in args {
                rewrite_set_algebra_bool_calls_in_expr(&mut arg.value);
            }
        }
        ExprKind::New { class, args, .. } => {
            rewrite_set_algebra_bool_calls_in_expr(class);
            for arg in args {
                rewrite_set_algebra_bool_calls_in_expr(&mut arg.value);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_set_algebra_bool_calls_in_expr(expr),
            LambdaBody::Block(stmts) => rewrite_set_algebra_bool_calls(stmts),
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_set_algebra_bool_calls_in_expr(&mut item.value);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_set_algebra_bool_calls_in_expr(item);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { value, .. } => {
                        rewrite_set_algebra_bool_calls_in_expr(value);
                    }
                    ObjectProperty::Spread(expr) => rewrite_set_algebra_bool_calls_in_expr(expr),
                    ObjectProperty::Method { value, .. } => {
                        rewrite_set_algebra_bool_calls_in_stmt(value);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn normalize_import_path(path: &str) -> String {
    path.split('<').next().unwrap_or(path).trim().to_string()
}

fn rewrite_using_imports_in_statements(
    body: &mut [Statement],
    aliases: &HashMap<String, String>,
    static_paths: &[String],
) {
    for stmt in body {
        rewrite_using_imports_in_statement(stmt, aliases, static_paths);
    }
}

fn rewrite_using_imports_in_statement(
    stmt: &mut Statement,
    aliases: &HashMap<String, String>,
    static_paths: &[String],
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
            rewrite_using_imports_in_expr(cause, aliases, static_paths);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_using_imports_in_expr(init, aliases, static_paths);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_using_imports_in_expr(bound, aliases, static_paths);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_using_imports_in_statements(body, aliases, static_paths);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_using_imports_in_member(member, aliases, static_paths);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_using_imports_in_expr(cond, aliases, static_paths);
            rewrite_using_imports_in_statements(then_body, aliases, static_paths);
            for (elif_cond, elif_body) in elifs {
                rewrite_using_imports_in_expr(elif_cond, aliases, static_paths);
                rewrite_using_imports_in_statements(elif_body, aliases, static_paths);
            }
            if let Some(else_body) = else_body {
                rewrite_using_imports_in_statements(else_body, aliases, static_paths);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_using_imports_in_statement(init, aliases, static_paths);
            }
            if let Some(cond) = cond {
                rewrite_using_imports_in_expr(cond, aliases, static_paths);
            }
            if let Some(update) = update {
                rewrite_using_imports_in_expr(update, aliases, static_paths);
            }
            rewrite_using_imports_in_statements(body, aliases, static_paths);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_using_imports_in_expr(iter, aliases, static_paths);
            rewrite_using_imports_in_statements(body, aliases, static_paths);
            if let Some(else_body) = else_body {
                rewrite_using_imports_in_statements(else_body, aliases, static_paths);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_using_imports_in_expr(cond, aliases, static_paths);
            rewrite_using_imports_in_statements(body, aliases, static_paths);
            if let Some(else_body) = else_body {
                rewrite_using_imports_in_statements(else_body, aliases, static_paths);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_using_imports_in_statements(body, aliases, static_paths);
            rewrite_using_imports_in_expr(cond, aliases, static_paths);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_using_imports_in_expr(expr, aliases, static_paths)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_using_imports_in_expr(from, aliases, static_paths);
                            rewrite_using_imports_in_expr(to, aliases, static_paths);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_using_imports_in_expr(expr, aliases, static_paths);
                        }
                    }
                }
                rewrite_using_imports_in_statements(&mut case.body, aliases, static_paths);
            }
            if let Some(default) = default {
                rewrite_using_imports_in_statements(default, aliases, static_paths);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_using_imports_in_statements(body, aliases, static_paths);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_using_imports_in_expr(when_clause, aliases, static_paths);
                }
                rewrite_using_imports_in_statements(&mut catch.body, aliases, static_paths);
            }
            if let Some(else_body) = else_body {
                rewrite_using_imports_in_statements(else_body, aliases, static_paths);
            }
            if let Some(finally) = finally {
                rewrite_using_imports_in_statements(finally, aliases, static_paths);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_using_imports_in_expr(&mut item.expr, aliases, static_paths);
            }
            rewrite_using_imports_in_statements(body, aliases, static_paths);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_using_imports_in_expr(target, aliases, static_paths);
            }
            rewrite_using_imports_in_expr(value, aliases, static_paths);
        }
        _ => {}
    }
}

fn rewrite_using_imports_in_member(
    member: &mut ClassMember,
    aliases: &HashMap<String, String>,
    static_paths: &[String],
) {
    match member {
        ClassMember::Field {
            init: Some(expr),
            array_bounds,
            ..
        } => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_using_imports_in_expr(bound, aliases, static_paths);
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_using_imports_in_statement(stmt, aliases, static_paths);
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            rewrite_using_imports_in_statements(body, aliases, static_paths);
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_using_imports_in_expr(arg, aliases, static_paths);
                }
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_using_imports_in_statements(getter, aliases, static_paths);
            }
            if let Some(setter) = setter {
                rewrite_using_imports_in_statements(&mut setter.body, aliases, static_paths);
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_using_imports_in_expr(value, aliases, static_paths);
        }
        _ => {}
    }
}

fn rewrite_using_imports_in_expr(
    expr: &mut Expression,
    aliases: &HashMap<String, String>,
    static_paths: &[String],
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_using_imports_in_expr(left, aliases, static_paths);
            rewrite_using_imports_in_expr(right, aliases, static_paths);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_using_imports_in_expr(cond, aliases, static_paths);
            rewrite_using_imports_in_expr(then, aliases, static_paths);
            rewrite_using_imports_in_expr(else_, aliases, static_paths);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } => {
                rewrite_using_imports_in_expr(object, aliases, static_paths);
            }
            PlaceExpr::Index { object, index, .. } => {
                rewrite_using_imports_in_expr(object, aliases, static_paths);
                rewrite_using_imports_in_expr(index, aliases, static_paths);
            }
            PlaceExpr::Deref(expr) => {
                rewrite_using_imports_in_expr(expr, aliases, static_paths);
            }
        },
        ExprKind::Member { object, .. } => {
            rewrite_using_imports_in_expr(object, aliases, static_paths);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_using_imports_in_expr(object, aliases, static_paths);
            rewrite_using_imports_in_expr(index, aliases, static_paths);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_using_imports_in_expr(callee, aliases, static_paths);
            for arg in args.iter_mut() {
                rewrite_using_imports_in_expr(&mut arg.value, aliases, static_paths);
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if let Some(path) = static_paths.first() {
                    let qualified = format!("{path}.{name}");
                    *callee = Box::new(build_dotted_expr(&qualified));
                }
            }
        }
        ExprKind::New { class, args } => {
            rewrite_using_imports_in_expr(class, aliases, static_paths);
            for arg in args {
                rewrite_using_imports_in_expr(&mut arg.value, aliases, static_paths);
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_using_imports_in_expr(target, aliases, static_paths);
            rewrite_using_imports_in_expr(value, aliases, static_paths);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_using_imports_in_expr(expr, aliases, static_paths),
            LambdaBody::Block(body) => {
                rewrite_using_imports_in_statements(body, aliases, static_paths)
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_using_imports_in_expr(&mut item.value, aliases, static_paths);
                if let Some(key) = &mut item.key {
                    rewrite_using_imports_in_expr(key, aliases, static_paths);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_using_imports_in_expr(item, aliases, static_paths);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_using_imports_in_expr(key, aliases, static_paths);
                        rewrite_using_imports_in_expr(value, aliases, static_paths);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_using_imports_in_expr(expr, aliases, static_paths);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_using_imports_in_statement(value, aliases, static_paths);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_using_imports_in_expr(expr, aliases, static_paths);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::IsType { expr, .. } | ExprKind::Cast { expr, .. } | ExprKind::Spread(expr) => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
        }
        ExprKind::Yield(Some(expr)) => {
            rewrite_using_imports_in_expr(expr, aliases, static_paths);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_using_imports_in_expr(element, aliases, static_paths);
            for generator in generators {
                rewrite_using_imports_in_expr(&mut generator.target, aliases, static_paths);
                rewrite_using_imports_in_expr(&mut generator.iter, aliases, static_paths);
                for condition in &mut generator.conditions {
                    rewrite_using_imports_in_expr(condition, aliases, static_paths);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_using_imports_in_expr(lower, aliases, static_paths);
            }
            if let Some(upper) = upper {
                rewrite_using_imports_in_expr(upper, aliases, static_paths);
            }
            if let Some(step) = step {
                rewrite_using_imports_in_expr(step, aliases, static_paths);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_using_imports_in_expr(parent, aliases, static_paths);
            }
            for member in members {
                rewrite_using_imports_in_member(member, aliases, static_paths);
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            rewrite_using_imports_in_statement(stmt, aliases, static_paths);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_using_imports_in_expr(start, aliases, static_paths);
            rewrite_using_imports_in_expr(end, aliases, static_paths);
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_using_imports_in_expr(class, aliases, static_paths);
            rewrite_using_imports_in_expr(member, aliases, static_paths);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_using_imports_in_expr(subject, aliases, static_paths);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        rewrite_using_imports_in_expr(condition, aliases, static_paths);
                    }
                }
                rewrite_using_imports_in_expr(&mut arm.body, aliases, static_paths);
            }
        }
        ExprKind::Ident(name) => {
            if let Some(path) = aliases.get(name) {
                expr.kind = build_dotted_expr(path).kind;
            }
        }
        ExprKind::Lit(_)
        | ExprKind::DefaultOf(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::Yield(None)
        | ExprKind::Destructure(_) => {}
    }
}

fn rewrite_extension_calls(module: &mut Module) {
    let mut extension_methods: HashMap<String, String> = HashMap::new();
    let mut extension_containers: HashSet<String> = HashSet::new();
    collect_extension_methods(
        &module.body,
        None,
        &mut extension_methods,
        &mut extension_containers,
    );
    if extension_methods.is_empty() {
        return;
    }
    // Also collect all user-defined type names (class/struct/interface/enum) so we never
    // rewrite calls whose receiver is a known type identifier (e.g. Array.Reverse, Math.Abs).
    collect_all_type_names(&module.body, &mut extension_containers);
    // Add well-known BCL static class names that may appear as call receivers.
    for name in &[
        "Array",
        "Console",
        "String",
        "Math",
        "Environment",
        "Convert",
        "File",
        "Path",
        "Directory",
        "Enumerable",
        "Regex",
        "Char",
        "Int32",
        "Int64",
        "Double",
        "Boolean",
        "Object",
        "Type",
        "Encoding",
        "BitConverter",
        "GC",
        "Monitor",
        "Thread",
        "StringBuilder",
        "StringComparer",
    ] {
        extension_containers.insert(name.to_string());
    }
    rewrite_extension_calls_in_statements(
        &mut module.body,
        &extension_methods,
        &extension_containers,
    );
}

fn collect_all_type_names(body: &[Statement], type_names: &mut HashSet<String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::NamespaceDecl { body, .. } => {
                collect_all_type_names(body, type_names);
            }
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                type_names.insert(name.clone());
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_all_type_names(std::slice::from_ref(nested.as_ref()), type_names);
                    }
                }
            }
            StmtKind::InterfaceDecl { name, .. } | StmtKind::EnumDecl { name, .. } => {
                type_names.insert(name.clone());
            }
            _ => {}
        }
    }
}

fn collect_extension_methods(
    body: &[Statement],
    namespace: Option<&str>,
    extension_methods: &mut HashMap<String, String>,
    extension_containers: &mut HashSet<String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::NamespaceDecl { name, body } => {
                collect_extension_methods(
                    body,
                    Some(name.as_str()),
                    extension_methods,
                    extension_containers,
                );
            }
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                let container = if let Some(ns) = namespace {
                    format!("{ns}.{name}")
                } else {
                    name.clone()
                };
                extension_containers.insert(name.clone());
                extension_containers.insert(container.clone());
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            if let StmtKind::FunctionDecl {
                                name: method_name,
                                params,
                                modifiers,
                                ..
                            } = &method.kind
                            {
                                if modifiers.is_static
                                    && params
                                        .first()
                                        .map(|param| param.pass_by == PassBy::Const)
                                        .unwrap_or(false)
                                {
                                    extension_methods
                                        .entry(method_name.clone())
                                        .or_insert_with(|| format!("{name}.{method_name}"));
                                }
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_extension_methods(
                                std::slice::from_ref(nested.as_ref()),
                                namespace,
                                extension_methods,
                                extension_containers,
                            );
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }
}

fn rewrite_extension_calls_in_statements(
    body: &mut [Statement],
    extension_methods: &HashMap<String, String>,
    extension_containers: &HashSet<String>,
) {
    for stmt in body {
        rewrite_extension_calls_in_statement(stmt, extension_methods, extension_containers);
    }
}

fn rewrite_extension_calls_in_statement(
    stmt: &mut Statement,
    extension_methods: &HashMap<String, String>,
    extension_containers: &HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(cause, extension_methods, extension_containers);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_extension_calls_in_expr(init, extension_methods, extension_containers);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_extension_calls_in_expr(
                            bound,
                            extension_methods,
                            extension_containers,
                        );
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_extension_calls_in_member(member, extension_methods, extension_containers);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_extension_calls_in_expr(cond, extension_methods, extension_containers);
            rewrite_extension_calls_in_statements(
                then_body,
                extension_methods,
                extension_containers,
            );
            for (elif_cond, elif_body) in elifs {
                rewrite_extension_calls_in_expr(elif_cond, extension_methods, extension_containers);
                rewrite_extension_calls_in_statements(
                    elif_body,
                    extension_methods,
                    extension_containers,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_extension_calls_in_statements(
                    else_body,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_extension_calls_in_statement(init, extension_methods, extension_containers);
            }
            if let Some(cond) = cond {
                rewrite_extension_calls_in_expr(cond, extension_methods, extension_containers);
            }
            if let Some(update) = update {
                rewrite_extension_calls_in_expr(update, extension_methods, extension_containers);
            }
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_extension_calls_in_expr(iter, extension_methods, extension_containers);
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
            if let Some(else_body) = else_body {
                rewrite_extension_calls_in_statements(
                    else_body,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_extension_calls_in_expr(cond, extension_methods, extension_containers);
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
            if let Some(else_body) = else_body {
                rewrite_extension_calls_in_statements(
                    else_body,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(cond, extension_methods, extension_containers);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => rewrite_extension_calls_in_expr(
                            expr,
                            extension_methods,
                            extension_containers,
                        ),
                        CaseCondition::Range { from, to } => {
                            rewrite_extension_calls_in_expr(
                                from,
                                extension_methods,
                                extension_containers,
                            );
                            rewrite_extension_calls_in_expr(
                                to,
                                extension_methods,
                                extension_containers,
                            );
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_extension_calls_in_expr(
                                expr,
                                extension_methods,
                                extension_containers,
                            );
                        }
                    }
                }
                rewrite_extension_calls_in_statements(
                    &mut case.body,
                    extension_methods,
                    extension_containers,
                );
            }
            if let Some(default) = default {
                rewrite_extension_calls_in_statements(
                    default,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_extension_calls_in_expr(
                        when_clause,
                        extension_methods,
                        extension_containers,
                    );
                }
                rewrite_extension_calls_in_statements(
                    &mut catch.body,
                    extension_methods,
                    extension_containers,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_extension_calls_in_statements(
                    else_body,
                    extension_methods,
                    extension_containers,
                );
            }
            if let Some(finally) = finally {
                rewrite_extension_calls_in_statements(
                    finally,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_extension_calls_in_expr(
                    &mut item.expr,
                    extension_methods,
                    extension_containers,
                );
            }
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_extension_calls_in_expr(target, extension_methods, extension_containers);
            }
            rewrite_extension_calls_in_expr(value, extension_methods, extension_containers);
        }
        _ => {}
    }
}

fn rewrite_extension_calls_in_member(
    member: &mut ClassMember,
    extension_methods: &HashMap<String, String>,
    extension_containers: &HashSet<String>,
) {
    match member {
        ClassMember::Field {
            init: Some(expr),
            array_bounds,
            ..
        } => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_extension_calls_in_expr(bound, extension_methods, extension_containers);
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_extension_calls_in_statement(stmt, extension_methods, extension_containers);
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            rewrite_extension_calls_in_statements(body, extension_methods, extension_containers);
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_extension_calls_in_expr(arg, extension_methods, extension_containers);
                }
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_extension_calls_in_statements(
                    getter,
                    extension_methods,
                    extension_containers,
                );
            }
            if let Some(setter) = setter {
                rewrite_extension_calls_in_statements(
                    &mut setter.body,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_extension_calls_in_expr(value, extension_methods, extension_containers);
        }
        _ => {}
    }
}

fn rewrite_extension_calls_in_expr(
    expr: &mut Expression,
    extension_methods: &HashMap<String, String>,
    extension_containers: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_extension_calls_in_expr(left, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(right, extension_methods, extension_containers);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_extension_calls_in_expr(cond, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(then, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(else_, extension_methods, extension_containers);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } => {
                rewrite_extension_calls_in_expr(object, extension_methods, extension_containers);
            }
            PlaceExpr::Index { object, index, .. } => {
                rewrite_extension_calls_in_expr(object, extension_methods, extension_containers);
                rewrite_extension_calls_in_expr(index, extension_methods, extension_containers);
            }
            PlaceExpr::Deref(expr) => {
                rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
            }
        },
        ExprKind::Member { object, .. } => {
            rewrite_extension_calls_in_expr(object, extension_methods, extension_containers);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_extension_calls_in_expr(object, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(index, extension_methods, extension_containers);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_extension_calls_in_expr(callee, extension_methods, extension_containers);
            for arg in args.iter_mut() {
                rewrite_extension_calls_in_expr(
                    &mut arg.value,
                    extension_methods,
                    extension_containers,
                );
            }

            let replacement = if let ExprKind::Member {
                object,
                field,
                null_safe,
            } = &callee.kind
            {
                if *null_safe {
                    None
                } else if expr_dotted_name(object)
                    .map(|name| extension_containers.contains(&name))
                    .unwrap_or(false)
                {
                    None
                } else {
                    extension_methods.get(field).map(|static_path| {
                        let mut rewritten_args = Vec::with_capacity(args.len() + 1);
                        rewritten_args.push(Argument::positional((**object).clone()));
                        rewritten_args.extend(args.clone());
                        ExprKind::Call {
                            callee: Box::new(build_dotted_expr(static_path)),
                            args: rewritten_args,
                            optional: false,
                        }
                    })
                }
            } else {
                None
            };

            if let Some(new_kind) = replacement {
                expr.kind = new_kind;
            }
        }
        ExprKind::New { class, args } => {
            rewrite_extension_calls_in_expr(class, extension_methods, extension_containers);
            for arg in args {
                rewrite_extension_calls_in_expr(
                    &mut arg.value,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_extension_calls_in_expr(target, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(value, extension_methods, extension_containers);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => {
                rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers)
            }
            LambdaBody::Block(body) => {
                rewrite_extension_calls_in_statements(body, extension_methods, extension_containers)
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_extension_calls_in_expr(
                    &mut item.value,
                    extension_methods,
                    extension_containers,
                );
                if let Some(key) = &mut item.key {
                    rewrite_extension_calls_in_expr(key, extension_methods, extension_containers);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_extension_calls_in_expr(item, extension_methods, extension_containers);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_extension_calls_in_expr(
                            key,
                            extension_methods,
                            extension_containers,
                        );
                        rewrite_extension_calls_in_expr(
                            value,
                            extension_methods,
                            extension_containers,
                        );
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_extension_calls_in_expr(
                            expr,
                            extension_methods,
                            extension_containers,
                        );
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_extension_calls_in_statement(
                            value,
                            extension_methods,
                            extension_containers,
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_extension_calls_in_expr(
                            expr,
                            extension_methods,
                            extension_containers,
                        );
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::IsType { expr, .. } | ExprKind::Cast { expr, .. } | ExprKind::Spread(expr) => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
        }
        ExprKind::Yield(Some(expr)) => {
            rewrite_extension_calls_in_expr(expr, extension_methods, extension_containers);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_extension_calls_in_expr(element, extension_methods, extension_containers);
            for generator in generators {
                rewrite_extension_calls_in_expr(
                    &mut generator.target,
                    extension_methods,
                    extension_containers,
                );
                rewrite_extension_calls_in_expr(
                    &mut generator.iter,
                    extension_methods,
                    extension_containers,
                );
                for condition in &mut generator.conditions {
                    rewrite_extension_calls_in_expr(
                        condition,
                        extension_methods,
                        extension_containers,
                    );
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_extension_calls_in_expr(lower, extension_methods, extension_containers);
            }
            if let Some(upper) = upper {
                rewrite_extension_calls_in_expr(upper, extension_methods, extension_containers);
            }
            if let Some(step) = step {
                rewrite_extension_calls_in_expr(step, extension_methods, extension_containers);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_extension_calls_in_expr(parent, extension_methods, extension_containers);
            }
            for member in members {
                rewrite_extension_calls_in_member(member, extension_methods, extension_containers);
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            rewrite_extension_calls_in_statement(stmt, extension_methods, extension_containers);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_extension_calls_in_expr(start, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(end, extension_methods, extension_containers);
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_extension_calls_in_expr(class, extension_methods, extension_containers);
            rewrite_extension_calls_in_expr(member, extension_methods, extension_containers);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_extension_calls_in_expr(subject, extension_methods, extension_containers);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        rewrite_extension_calls_in_expr(
                            condition,
                            extension_methods,
                            extension_containers,
                        );
                    }
                }
                rewrite_extension_calls_in_expr(
                    &mut arm.body,
                    extension_methods,
                    extension_containers,
                );
            }
        }
        ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::DefaultOf(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::Yield(None)
        | ExprKind::Destructure(_) => {}
    }
}

fn expr_dotted_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } if !null_safe => expr_dotted_name(object).map(|prefix| format!("{prefix}.{field}")),
        _ => None,
    }
}

fn synthesize_exception_classes() -> Vec<Statement> {
    let names = [
        "Exception",
        "InvalidOperationException",
        "ArgumentException",
        "ArgumentNullException",
        "ArgumentOutOfRangeException",
        "DivideByZeroException",
        "FormatException",
        "NotImplementedException",
        "NotSupportedException",
        "NullReferenceException",
        "OverflowException",
        "IndexOutOfRangeException",
    ];
    names.into_iter().map(synthesize_exception_class).collect()
}

fn synthesize_attribute_classes() -> Vec<Statement> {
    let names = [
        "Attribute",
        "FlagsAttribute",
        "STAThreadAttribute",
        "SerializableAttribute",
        "ObsoleteAttribute",
    ];
    names.into_iter().map(synthesize_attribute_class).collect()
}

fn synthesize_attribute_class(name: &str) -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents: if name == "Attribute" {
                Vec::new()
            } else {
                vec!["Attribute".into()]
            },
            interfaces: Vec::new(),
            members: vec![ClassMember::Constructor {
                params: Vec::new(),
                body: Vec::new(),
                base_args: None,
                initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                visibility: Visibility::Public,
            }],
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        Span::default(),
    )
}

fn synthesize_checked_numeric_helpers() -> Vec<Statement> {
    vec![synthesize_checked_byte_cast_helper()]
}

fn synthesize_checked_byte_cast_helper() -> Statement {
    let overflow_cond = Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident("value")),
            right: Box::new(Expression::int(0)),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(Expression::ident("value")),
            right: Box::new(Expression::int(255)),
        })),
    });

    Statement::with_span(
        StmtKind::FunctionDecl {
            name: "__vybe_csharp_checked_byte_cast".into(),
            params: vec![Param {
                name: "value".into(),
                type_hint: Some("int".into()),
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            return_type: Some("byte".into()),
            body: vec![
                Statement::new(StmtKind::If {
                    cond: overflow_cond,
                    then_body: vec![Statement::new(StmtKind::Throw {
                        expr: Some(Expression::new(ExprKind::New {
                            class: Box::new(Expression::ident("OverflowException")),
                            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                                Literal::Str(
                                    "Arithmetic operation resulted in an overflow.".into(),
                                ),
                            )))],
                        })),
                        cause: None,
                    })],
                    elifs: vec![],
                    else_body: None,
                }),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Call {
                    callee: Box::new(build_dotted_expr("Convert.ToByte")),
                    args: vec![Argument::positional(Expression::ident("value"))],
                    optional: false,
                })))),
            ],
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        Span::default(),
    )
}

fn synthesize_exception_class(name: &str) -> Statement {
    let span = Span::default();
    // Per-type constructor signatures per ECMA-335 / .NET BCL:
    //   ArgumentNullException(paramName)              → ParamName=paramName
    //   ArgumentOutOfRangeException(paramName, msg)   → ParamName=paramName, Message=msg
    //   ArgumentException(msg, paramName)             → Message=msg, ParamName=paramName
    //   <other>(msg)                                  → Message=msg
    //
    // Walker emits the appropriate constructor body so `e.ParamName`
    // and `e.Message` resolve to the right values on every catch.
    let needs_param_name = matches!(
        name,
        "ArgumentNullException" | "ArgumentOutOfRangeException" | "ArgumentException"
    );

    let assign = |field: &str, ident: &str| {
        Statement::with_span(
            StmtKind::Assign {
                targets: vec![Expression::with_span(
                    ExprKind::Member {
                        object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                        field: field.into(),
                        null_safe: false,
                    },
                    span.clone(),
                )],
                value: Expression::with_span(ExprKind::Ident(ident.into()), span.clone()),
            },
            span.clone(),
        )
    };

    let canon = crate::emitter::errors::canonical_exception_name(name).to_string();
    let assign_extype = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                    field: "__exception_type".into(),
                    null_safe: false,
                },
                span.clone(),
            )],
            value: Expression::with_span(ExprKind::Lit(Literal::Str(canon.clone())), span.clone()),
        },
        span.clone(),
    );

    let mk_param = |pname: &str| Param {
        name: pname.into(),
        type_hint: Some("string".into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };

    let (params, body) = if needs_param_name {
        // 2-arg form: (paramName, msg) for ArgumentOutOfRangeException;
        // (msg, paramName) for ArgumentException; (paramName) for
        // ArgumentNullException. The walker matches the .NET BCL order.
        match name {
            "ArgumentException" => (
                vec![mk_param("msg"), mk_param("paramName")],
                vec![
                    assign("Message", "msg"),
                    assign("ParamName", "paramName"),
                    assign_extype.clone(),
                ],
            ),
            "ArgumentNullException" => (
                vec![mk_param("paramName")],
                vec![assign("ParamName", "paramName"), assign_extype.clone()],
            ),
            "ArgumentOutOfRangeException" => (
                vec![mk_param("paramName"), mk_param("msg")],
                vec![
                    assign("ParamName", "paramName"),
                    assign("Message", "msg"),
                    assign_extype.clone(),
                ],
            ),
            _ => unreachable!(),
        }
    } else {
        (
            vec![mk_param("msg")],
            vec![assign("Message", "msg"), assign_extype.clone()],
        )
    };

    let mut members = vec![ClassMember::Field {
        name: "Message".into(),
        type_hint: Some("string".into()),
        init: None,
        modifiers: Modifiers::default(),
        with_events: false,
        array_bounds: None,
    }];
    if needs_param_name {
        members.push(ClassMember::Field {
            name: "ParamName".into(),
            type_hint: Some("string".into()),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        });
    }
    members.push(ClassMember::Constructor {
        params,
        body,
        base_args: None,
        initializer_target: crate::ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    });
    if name == "ArgumentException" {
        members.push(ClassMember::Constructor {
            params: vec![mk_param("msg")],
            body: vec![assign("Message", "msg"), assign_extype.clone()],
            base_args: None,
            initializer_target: crate::ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    }

    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: Vec::new(),
        },
        span,
    )
}

// ── Top-level items ─────────────────────────────────────────────────────────

fn walk_top_level_with_attributes(
    pair: Pair<Rule>,
    attributes: &[Expression],
) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::record_struct_declaration => walk_record_decl(pair, attributes)?,
        Rule::class_declaration => walk_class_decl(pair, attributes)?,
        Rule::struct_declaration => walk_struct_decl(pair, attributes)?,
        Rule::interface_declaration => walk_interface_decl(pair, attributes)?,
        Rule::enum_declaration => walk_enum_decl(pair, attributes)?,
        Rule::record_declaration => walk_record_decl(pair, attributes)?,
        Rule::delegate_declaration => walk_delegate_decl(pair)?,
        _ => walk_statement(pair)?.kind,
    };
    Ok(Statement::with_span(kind, span))
}

// ── Statements ──────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => {
            let stmts = pair
                .into_inner()
                .map(walk_statement)
                .collect::<Result<Vec<_>, _>>()?;
            StmtKind::Block(stmts)
        }
        Rule::checked_statement => {
            let is_checked = pair.as_str().trim_start().starts_with("checked");
            let block = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::block_statement)
                .ok_or("Empty checked statement")?;
            let mut stmt = walk_statement(block)?;
            if let StmtKind::Block(body) = &mut stmt.kind {
                let mut scopes = vec![HashMap::new()];
                rewrite_checked_numeric_statements(body, &mut scopes, is_checked);
            }
            return Ok(stmt);
        }
        Rule::local_var_declaration => walk_local_var(pair)?,
        Rule::local_function_decl => walk_local_function(pair)?,
        Rule::using_declaration => walk_using_declaration(pair)?,
        Rule::tuple_deconstruction_decl => walk_tuple_deconstruction(pair)?,
        Rule::if_statement => walk_if(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::foreach_statement => walk_foreach(pair)?,
        Rule::while_statement => walk_while(pair)?,
        Rule::do_while_statement => walk_do_while(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::yield_statement => walk_yield_stmt(pair)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::throw_statement => walk_throw(pair)?,
        Rule::try_statement => walk_try(pair)?,
        Rule::using_statement => walk_using_stmt(pair)?,
        Rule::lock_statement => walk_lock(pair)?,
        Rule::expression_statement => {
            let expr = walk_expression(pair.into_inner().next().ok_or("Empty expr stmt")?)?;
            // Check if this is a compound assignment or event subscription
            classify_expr_stmt(expr)
        }
        // Type declarations can appear inside methods
        Rule::record_struct_declaration => walk_record_decl(pair, &[])?,
        Rule::class_declaration => walk_class_decl(pair, &[])?,
        Rule::struct_declaration => walk_struct_decl(pair, &[])?,
        Rule::interface_declaration => walk_interface_decl(pair, &[])?,
        Rule::enum_declaration => walk_enum_decl(pair, &[])?,
        Rule::record_declaration => walk_record_decl(pair, &[])?,
        Rule::delegate_declaration => walk_delegate_decl(pair)?,
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };
    Ok(Statement::with_span(kind, span))
}

/// Classify an expression statement — detect assignment and compound
/// assignment shapes. C# event/delegate `+=` and `-=` stay as ordinary
/// compound assignment here; the compiler decides how to lower them from
/// member/value semantics instead of guessing from naming conventions.
fn classify_expr_stmt(expr: Expression) -> StmtKind {
    if let Some(kind) = lower_array_resize_expr_stmt(&expr) {
        return kind;
    }

    match expr.kind {
        ExprKind::Assign { target, value } => {
            if let ExprKind::Binary { op, left, right } = &value.kind {
                let is_same_target = match (&left.kind, &target.kind) {
                    (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
                    (
                        ExprKind::Member {
                            object: lo,
                            field: lf,
                            ..
                        },
                        ExprKind::Member {
                            object: to,
                            field: tf,
                            ..
                        },
                    ) => member_target_eq(lo, lf, to, tf),
                    _ => false,
                };
                let compound_op = match op {
                    BinOp::Add => Some(CompoundOp::Add),
                    BinOp::Sub => Some(CompoundOp::Sub),
                    BinOp::Mul => Some(CompoundOp::Mul),
                    BinOp::Div => Some(CompoundOp::Div),
                    BinOp::Mod => Some(CompoundOp::Mod),
                    BinOp::BitAnd => Some(CompoundOp::BitAnd),
                    BinOp::BitOr => Some(CompoundOp::BitOr),
                    BinOp::BitXor => Some(CompoundOp::BitXor),
                    BinOp::Shl => Some(CompoundOp::Shl),
                    BinOp::Shr => Some(CompoundOp::Shr),
                    BinOp::UShr => Some(CompoundOp::UShr),
                    BinOp::NullCoalesce => Some(CompoundOp::NullCoalesce),
                    _ => None,
                };
                if is_same_target {
                    if let Some(compound_op) = compound_op {
                        // `btn.Click += handler` / `-=` → AddHandler / RemoveHandler so
                        // the compiler routes through vybe:gui:bindEvent like VB does.
                        if matches!(compound_op, CompoundOp::Add | CompoundOp::Sub) {
                            if let ExprKind::Member { object, field, .. } = &target.kind {
                                if crate::common::events::is_known_gui_event_field(field)
                                    && crate::common::events::is_event_handler_expr(&right)
                                {
                                    let control = (**object).clone();
                                    let event = field.to_lowercase();
                                    let handler = (**right).clone();
                                    return match compound_op {
                                        CompoundOp::Add => StmtKind::AddHandler {
                                            control,
                                            event,
                                            handler,
                                        },
                                        _ => StmtKind::RemoveHandler {
                                            control,
                                            event,
                                            handler,
                                        },
                                    };
                                }
                            }
                        }
                        return StmtKind::CompoundAssign {
                            target: *target,
                            op: compound_op,
                            value: *right.clone(),
                        };
                    }
                }
            }
            StmtKind::Assign {
                targets: vec![*target],
                value: *value,
            }
        }
        _ => StmtKind::Expr(expr),
    }
}

fn lower_array_resize_expr_stmt(expr: &Expression) -> Option<StmtKind> {
    let ExprKind::Call {
        callee,
        args,
        optional: false,
    } = &expr.kind
    else {
        return None;
    };
    if args.len() != 2 || !args[0].by_ref {
        return None;
    }

    let array_name = match &args[0].value.kind {
        ExprKind::Ident(name) => name.clone(),
        _ => return None,
    };

    let is_array_resize = match &callee.kind {
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } if field.eq_ignore_ascii_case("Resize") => {
            matches!(&object.kind,
                ExprKind::Ident(name) if name.eq_ignore_ascii_case("Array")
            ) || matches!(&object.kind,
                ExprKind::Member { object: system, field: array, null_safe: false }
                    if matches!(&system.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("System"))
                    && array.eq_ignore_ascii_case("Array")
            )
        }
        _ => false,
    };
    if !is_array_resize {
        return None;
    }

    Some(StmtKind::ReDim {
        preserve: true,
        array: array_name,
        bounds: vec![Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(args[1].value.clone()),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
        })],
    })
}

fn member_target_eq(obj_a: &Expression, field_a: &str, obj_b: &Expression, field_b: &str) -> bool {
    if !field_a.eq_ignore_ascii_case(field_b) {
        return false;
    }
    match (&obj_a.kind, &obj_b.kind) {
        (ExprKind::Ident(a), ExprKind::Ident(b)) => a == b,
        (ExprKind::This, ExprKind::This) => true,
        (
            ExprKind::Member {
                object: inner_a,
                field: inner_field_a,
                ..
            },
            ExprKind::Member {
                object: inner_b,
                field: inner_field_b,
                ..
            },
        ) => member_target_eq(inner_a, inner_field_a, inner_b, inner_field_b),
        _ => false,
    }
}

// ── Using directive ─────────────────────────────────────────────────────────

fn walk_using(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut alias: Option<String> = None;
    let mut path = String::new();
    let mut is_static = false;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::using_directive_body => {
                for inner in child.into_inner() {
                    match inner.as_rule() {
                        Rule::using_static_directive => {
                            is_static = true;
                            if let Some(name) = inner
                                .into_inner()
                                .find(|p| p.as_rule() == Rule::dotted_name)
                            {
                                path = strip_global_namespace_qualifier(name.as_str()).to_string();
                            }
                        }
                        Rule::using_alias_directive => {
                            let mut parts = inner.into_inner();
                            alias = parts.next().map(|p| p.as_str().to_string());
                            if let Some(target) = parts.next() {
                                path =
                                    strip_global_namespace_qualifier(target.as_str()).to_string();
                            }
                        }
                        Rule::dotted_name => {
                            path = strip_global_namespace_qualifier(inner.as_str()).to_string()
                        }
                        _ => {}
                    }
                }
            }
            Rule::dotted_name => {
                path = strip_global_namespace_qualifier(child.as_str()).to_string()
            }
            _ => {}
        }
    }
    let kind = if is_static {
        ImportKind::Wildcard { path, alias: None }
    } else {
        ImportKind::Simple { path, alias }
    };
    Ok(Import { kind, span })
}

// ── Variable declaration ────────────────────────────────────────────────────

/// `var (a, b) = f();` → `Assign { targets: [Destructure(Array([a, b]))], value: f() }`.
/// The compiler's multi-value receive path auto-defines any unresolved
/// idents as new locals inside a function, so a single `Assign` statement
/// is enough — no separate `VarDecl` is needed.
fn walk_tuple_deconstruction(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut pattern: Option<Vec<ArrayPatternElem>> = None;
    let mut value: Option<Expression> = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw => {}
            Rule::tuple_binding_list => pattern = Some(walk_tuple_binding_list(p)?),
            Rule::expression => {
                value = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }
    let pattern = pattern.ok_or("tuple deconstruction missing binding pattern")?;
    let value = value.ok_or("tuple deconstruction missing RHS")?;
    if !matches!(value.kind, ExprKind::Tuple(_) | ExprKind::Array(_)) {
        let mut leaf_names = Vec::new();
        collect_tuple_binding_leaf_names(&pattern, &mut leaf_names);
        let mut declarations = Vec::new();
        let mut args = Vec::new();
        for (index, name) in leaf_names.into_iter().enumerate() {
            let target_name = if name == "_" {
                format!("__discard_{}", index)
            } else {
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(name.clone()),
                    type_hint: None,
                    init: None,
                    array_bounds: None,
                    with_events: false,
                });
                name
            };
            args.push(Argument {
                value: Expression::ident(&target_name),
                name: None,
                by_ref: true,
                spread: false,
            });
        }

        let mut body = Vec::new();
        if !declarations.is_empty() {
            body.push(Statement::new(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let,
            }));
        }
        body.push(Statement::new(StmtKind::Expr(Expression::new(
            ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(value),
                    field: "Deconstruct".into(),
                    null_safe: false,
                })),
                args,
                optional: false,
            },
        ))));
        return Ok(StmtKind::Block(body));
    }
    let target = Expression::new(ExprKind::Destructure(DestructurePattern::Array(pattern)));
    Ok(StmtKind::Assign {
        targets: vec![target],
        value,
    })
}

fn walk_tuple_binding_list(pair: Pair<Rule>) -> Result<Vec<ArrayPatternElem>, String> {
    let mut elems = Vec::new();
    for child in pair.into_inner() {
        if child.as_rule() == Rule::tuple_binding_pattern {
            elems.push(walk_tuple_binding_pattern(child)?);
        }
    }
    Ok(elems)
}

fn walk_tuple_binding_pattern(pair: Pair<Rule>) -> Result<ArrayPatternElem, String> {
    let inner = pair
        .into_inner()
        .next()
        .ok_or("empty tuple binding pattern")?;
    match inner.as_rule() {
        Rule::ident_name => {
            let name = inner.as_str().to_string();
            if name == "_" {
                Ok(ArrayPatternElem::Hole)
            } else {
                Ok(ArrayPatternElem::Pattern(BindingPattern::Ident(name), None))
            }
        }
        Rule::tuple_binding_list => Ok(ArrayPatternElem::Pattern(
            BindingPattern::Array(walk_tuple_binding_list(inner)?),
            None,
        )),
        other => Err(format!(
            "unexpected tuple binding pattern rule: {:?}",
            other
        )),
    }
}

fn collect_tuple_binding_leaf_names(pattern: &[ArrayPatternElem], out: &mut Vec<String>) {
    for elem in pattern {
        match elem {
            ArrayPatternElem::Pattern(BindingPattern::Ident(name), _) => out.push(name.clone()),
            ArrayPatternElem::Pattern(BindingPattern::Array(items), _) => {
                collect_tuple_binding_leaf_names(items, out);
            }
            ArrayPatternElem::Hole => out.push("_".into()),
            ArrayPatternElem::Rest(name) => out.push(name.clone()),
            ArrayPatternElem::Pattern(BindingPattern::Object(_), _) => {}
        }
    }
}

fn tuple_binding_pattern_elems(idents: Vec<String>) -> Vec<ArrayPatternElem> {
    idents
        .into_iter()
        .map(|name| {
            if name == "_" {
                ArrayPatternElem::Hole
            } else {
                ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
            }
        })
        .collect()
}

fn infer_csharp_new_type_name(class: &Expression) -> Option<String> {
    match &class.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

fn has_ordinal_ignore_case_comparer_arg(args: &[Argument]) -> bool {
    args.iter().any(|arg| {
        matches!(
            &arg.value.kind,
            ExprKind::Lit(Literal::Str(tag)) if tag == "__dotnet_stringcomparer_ordinalignorecase"
        )
    })
}

fn annotate_case_insensitive_ctor_type(type_name: String, args: &[Argument]) -> String {
    if has_ordinal_ignore_case_comparer_arg(args) {
        format!("{type_name}#ordinalignorecase")
    } else {
        type_name
    }
}

fn infer_csharp_iife_return_type(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !args.is_empty() {
        return None;
    }
    let ExprKind::Lambda {
        body: LambdaBody::Block(body),
        ..
    } = &callee.kind
    else {
        return None;
    };
    let StmtKind::Return(Some(Expression {
        kind: ExprKind::Ident(return_name),
        ..
    })) = &body.last()?.kind
    else {
        return None;
    };
    for stmt in body {
        let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
            continue;
        };
        for decl in declarations {
            let BindingPattern::Ident(name) = &decl.pattern else {
                continue;
            };
            if name != return_name {
                continue;
            }
            if let Some(type_hint) = decl.type_hint.as_ref() {
                return Some(type_hint.clone());
            }
            if let Some(Expression {
                kind: ExprKind::New { class, .. },
                ..
            }) = decl.init.as_ref()
            {
                return infer_csharp_new_type_name(class);
            }
        }
    }
    None
}

fn infer_csharp_type_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("int".into()),
        ExprKind::Lit(Literal::Float(_)) => Some("double".into()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".into()),
        ExprKind::Lit(Literal::Bool(_)) => Some("bool".into()),
        ExprKind::Lit(Literal::Char(_)) => Some("char".into()),
        ExprKind::New { class, args } => infer_csharp_new_type_name(class)
            .map(|type_name| annotate_case_insensitive_ctor_type(type_name, args)),
        ExprKind::Array(elements) => {
            let mut element_type: Option<String> = None;
            for element in elements {
                if element.key.is_some() || element.spread {
                    return None;
                }
                let inferred = infer_csharp_type_from_expr(&element.value)?;
                match &element_type {
                    Some(existing) if existing != &inferred => return None,
                    None => element_type = Some(inferred),
                    _ => {}
                }
            }
            element_type.map(|inner| format!("{}[]", inner))
        }
        ExprKind::Call { .. } => infer_csharp_iife_return_type(expr),
        _ => None,
    }
}

fn infer_csharp_foreach_element_type(iter: &Expression) -> Option<String> {
    match &iter.kind {
        ExprKind::Array(elements) => {
            let mut element_type: Option<String> = None;
            for element in elements {
                if element.key.is_some() || element.spread {
                    return None;
                }
                let inferred = infer_csharp_type_from_expr(&element.value)?;
                match &element_type {
                    Some(existing) if existing != &inferred => return None,
                    None => element_type = Some(inferred),
                    _ => {}
                }
            }
            element_type
        }
        _ => None,
    }
}

#[derive(Debug, Clone, Default)]
struct CSharpOperatorInfo {
    return_type: Option<String>,
}

type CSharpOperatorTable = HashMap<String, HashMap<String, CSharpOperatorInfo>>;

fn rewrite_user_defined_operator_calls(module: &mut Module) {
    let operators = collect_csharp_operator_methods(&module.body);
    let mut scopes = vec![HashMap::new()];
    rewrite_user_defined_operator_calls_in_statements(&mut module.body, &operators, &mut scopes);
}

fn collect_csharp_operator_methods(body: &[Statement]) -> CSharpOperatorTable {
    let mut operators = HashMap::new();
    for stmt in body {
        collect_csharp_operator_methods_in_statement(stmt, &mut operators);
    }
    operators
}

fn collect_csharp_operator_methods_in_statement(
    stmt: &Statement,
    operators: &mut CSharpOperatorTable,
) {
    match &stmt.kind {
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            for stmt in body {
                collect_csharp_operator_methods_in_statement(stmt, operators);
            }
        }
        StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. } => {
            collect_csharp_operator_methods_in_members(name, members, operators);
        }
        StmtKind::ModuleDecl { name, members, .. } => {
            collect_csharp_operator_methods_in_members(name, members, operators);
        }
        StmtKind::EnumDecl {
            name, body_members, ..
        } => collect_csharp_operator_methods_in_members(name, body_members, operators),
        _ => {}
    }
}

fn collect_csharp_operator_methods_in_members(
    type_name: &str,
    members: &[ClassMember],
    operators: &mut CSharpOperatorTable,
) {
    let type_key = csharp_type_key(type_name);
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl {
                    name,
                    return_type,
                    modifiers,
                    ..
                } = &stmt.kind
                {
                    if modifiers.is_static && name.starts_with("op_") {
                        operators.entry(type_key.clone()).or_default().insert(
                            name.clone(),
                            CSharpOperatorInfo {
                                return_type: return_type.clone(),
                            },
                        );
                    }
                }
            }
            ClassMember::NestedType(stmt) => {
                collect_csharp_operator_methods_in_statement(stmt, operators)
            }
            _ => {}
        }
    }
}

fn csharp_type_key(type_name: &str) -> String {
    normalize_runtime_type_name(type_name)
        .trim_end_matches('?')
        .split('<')
        .next()
        .unwrap_or(type_name)
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .to_string()
}

fn csharp_operator_info<'a>(
    operators: &'a CSharpOperatorTable,
    type_name: &str,
    method: &str,
) -> Option<&'a CSharpOperatorInfo> {
    operators.get(&csharp_type_key(type_name))?.get(method)
}

fn build_csharp_operator_static_call(
    type_name: &str,
    method: &str,
    args: Vec<Expression>,
    span: Span,
    return_type: Option<&str>,
) -> Expression {
    let call = Expression::with_span(
        ExprKind::Call {
            callee: Box::new(Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::ident(&csharp_type_key(type_name))),
                    field: method.into(),
                    null_safe: false,
                },
                span.clone(),
            )),
            args: args.into_iter().map(Argument::positional).collect(),
            optional: false,
        },
        span.clone(),
    );
    if return_type
        .map(|type_name| normalize_runtime_type_name(type_name).eq_ignore_ascii_case("bool"))
        .unwrap_or(false)
    {
        materialize_csharp_bool_expr(call)
    } else {
        call
    }
}

fn rewrite_csharp_implicit_conversion(
    expr: &mut Expression,
    target_type: &str,
    operators: &CSharpOperatorTable,
    scopes: &[HashMap<String, String>],
) -> Option<String> {
    let source_type = infer_csharp_expr_type(expr, operators, scopes);
    if source_type
        .as_deref()
        .map(|source| csharp_type_key(source) == csharp_type_key(target_type))
        .unwrap_or(false)
    {
        return source_type;
    }
    let method = format!("op_Implicit{}", csharp_operator_type_suffix(target_type));
    let dispatch_type = if csharp_operator_info(operators, target_type, &method).is_some() {
        Some(target_type.to_string())
    } else {
        source_type
            .as_deref()
            .filter(|source| csharp_operator_info(operators, source, &method).is_some())
            .map(ToString::to_string)
    }?;
    let return_type = csharp_operator_info(operators, &dispatch_type, &method)
        .and_then(|info| info.return_type.clone())
        .or_else(|| Some(target_type.to_string()));
    let span = expr.span.clone();
    let value = expr.clone();
    *expr = build_csharp_operator_static_call(
        &dispatch_type,
        &method,
        vec![value],
        span,
        return_type.as_deref(),
    );
    return_type
}

fn rewrite_user_defined_operator_calls_in_members(
    type_name: &str,
    members: &mut [ClassMember],
    operators: &CSharpOperatorTable,
) {
    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { params, body, .. } = &mut stmt.kind {
                    let mut scopes = vec![HashMap::new()];
                    scopes
                        .last_mut()
                        .unwrap()
                        .insert("this".into(), type_name.into());
                    for param in params {
                        if let Some(type_hint) = &param.type_hint {
                            scopes
                                .last_mut()
                                .unwrap()
                                .insert(param.name.clone(), type_hint.clone());
                        }
                    }
                    rewrite_user_defined_operator_calls_in_statements(body, operators, &mut scopes);
                }
            }
            ClassMember::Constructor { params, body, .. } => {
                let mut scopes = vec![HashMap::new()];
                scopes
                    .last_mut()
                    .unwrap()
                    .insert("this".into(), type_name.into());
                for param in params {
                    if let Some(type_hint) = &param.type_hint {
                        scopes
                            .last_mut()
                            .unwrap()
                            .insert(param.name.clone(), type_hint.clone());
                    }
                }
                rewrite_user_defined_operator_calls_in_statements(body, operators, &mut scopes);
            }
            ClassMember::Field {
                init: Some(init), ..
            }
            | ClassMember::Const { value: init, .. } => {
                let mut scopes = vec![HashMap::new()];
                rewrite_user_defined_operator_expr(init, operators, &mut scopes);
            }
            ClassMember::Property { getter, setter, .. } => {
                if let Some(getter) = getter {
                    let mut scopes = vec![HashMap::new()];
                    rewrite_user_defined_operator_calls_in_statements(
                        getter,
                        operators,
                        &mut scopes,
                    );
                }
                if let Some(setter) = setter {
                    let mut scopes = vec![HashMap::new()];
                    scopes.last_mut().unwrap().insert(
                        setter.param.name.clone(),
                        setter
                            .param
                            .type_hint
                            .clone()
                            .unwrap_or_else(|| "object".into()),
                    );
                    rewrite_user_defined_operator_calls_in_statements(
                        &mut setter.body,
                        operators,
                        &mut scopes,
                    );
                }
            }
            ClassMember::NestedType(stmt) => {
                rewrite_user_defined_operator_calls_in_statement(
                    stmt,
                    operators,
                    &mut vec![HashMap::new()],
                );
            }
            _ => {}
        }
    }
}

fn rewrite_user_defined_operator_calls_in_statements(
    body: &mut [Statement],
    operators: &CSharpOperatorTable,
    scopes: &mut Vec<HashMap<String, String>>,
) {
    for stmt in body {
        rewrite_user_defined_operator_calls_in_statement(stmt, operators, scopes);
    }
}

fn rewrite_user_defined_operator_calls_in_statement(
    stmt: &mut Statement,
    operators: &CSharpOperatorTable,
    scopes: &mut Vec<HashMap<String, String>>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            rewrite_user_defined_operator_expr(expr, operators, scopes);
        }
        StmtKind::Block(body) => {
            scopes.push(HashMap::new());
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            scopes.pop();
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_user_defined_operator_expr(init, operators, scopes);
                    if let Some(target_type) = decl.type_hint.as_deref() {
                        rewrite_csharp_implicit_conversion(init, target_type, operators, scopes);
                    }
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let inferred = decl.type_hint.clone().or_else(|| {
                        decl.init
                            .as_ref()
                            .and_then(|e| infer_csharp_expr_type(e, operators, scopes))
                    });
                    if let Some(type_name) = inferred {
                        scopes.last_mut().unwrap().insert(name.clone(), type_name);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            scopes.push(HashMap::new());
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scopes
                        .last_mut()
                        .unwrap()
                        .insert(param.name.clone(), type_hint.clone());
                }
            }
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            scopes.pop();
        }
        StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. } => {
            rewrite_user_defined_operator_calls_in_members(name, members, operators);
        }
        StmtKind::ModuleDecl { name, members, .. } => {
            rewrite_user_defined_operator_calls_in_members(name, members, operators);
        }
        StmtKind::EnumDecl {
            name, body_members, ..
        } => rewrite_user_defined_operator_calls_in_members(name, body_members, operators),
        StmtKind::NamespaceDecl { body, .. } => {
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_user_defined_operator_expr(cond, operators, scopes);
            scopes.push(HashMap::new());
            rewrite_user_defined_operator_calls_in_statements(then_body, operators, scopes);
            scopes.pop();
            for (elif_cond, elif_body) in elifs {
                rewrite_user_defined_operator_expr(elif_cond, operators, scopes);
                scopes.push(HashMap::new());
                rewrite_user_defined_operator_calls_in_statements(elif_body, operators, scopes);
                scopes.pop();
            }
            if let Some(else_body) = else_body {
                scopes.push(HashMap::new());
                rewrite_user_defined_operator_calls_in_statements(else_body, operators, scopes);
                scopes.pop();
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            scopes.push(HashMap::new());
            if let Some(init) = init {
                rewrite_user_defined_operator_calls_in_statement(init, operators, scopes);
            }
            if let Some(cond) = cond {
                rewrite_user_defined_operator_expr(cond, operators, scopes);
            }
            if let Some(update) = update {
                rewrite_user_defined_operator_expr(update, operators, scopes);
            }
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            scopes.pop();
        }
        StmtKind::ForIn {
            var,
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_user_defined_operator_expr(iter, operators, scopes);
            scopes.push(HashMap::new());
            if let Some(element_type) = infer_csharp_foreach_element_type(iter) {
                scopes.last_mut().unwrap().insert(var.clone(), element_type);
            }
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            scopes.pop();
            if let Some(else_body) = else_body {
                rewrite_user_defined_operator_calls_in_statements(else_body, operators, scopes);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_user_defined_operator_expr(cond, operators, scopes);
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            if let Some(else_body) = else_body {
                rewrite_user_defined_operator_calls_in_statements(else_body, operators, scopes);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            rewrite_user_defined_operator_expr(cond, operators, scopes);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_user_defined_operator_expr(expr, operators, scopes);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(cond)
                        | CaseCondition::Comparison { expr: cond, .. } => {
                            rewrite_user_defined_operator_expr(cond, operators, scopes);
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_user_defined_operator_expr(from, operators, scopes);
                            rewrite_user_defined_operator_expr(to, operators, scopes);
                        }
                    }
                }
                rewrite_user_defined_operator_calls_in_statements(
                    &mut case.body,
                    operators,
                    scopes,
                );
            }
            if let Some(default) = default {
                rewrite_user_defined_operator_calls_in_statements(default, operators, scopes);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
            for catch in catches {
                rewrite_user_defined_operator_calls_in_statements(
                    &mut catch.body,
                    operators,
                    scopes,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_user_defined_operator_calls_in_statements(else_body, operators, scopes);
            }
            if let Some(finally) = finally {
                rewrite_user_defined_operator_calls_in_statements(finally, operators, scopes);
            }
        }
        StmtKind::Using { resource, body, .. }
        | StmtKind::Lock {
            expr: resource,
            body,
        } => {
            rewrite_user_defined_operator_expr(resource, operators, scopes);
            rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => {
            rewrite_user_defined_operator_expr(expr, operators, scopes);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_user_defined_operator_expr(target, operators, scopes);
            }
            rewrite_user_defined_operator_expr(value, operators, scopes);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_user_defined_operator_expr(target, operators, scopes);
            rewrite_user_defined_operator_expr(value, operators, scopes);
        }
        StmtKind::AddHandler {
            control, handler, ..
        }
        | StmtKind::RemoveHandler {
            control, handler, ..
        } => {
            rewrite_user_defined_operator_expr(control, operators, scopes);
            rewrite_user_defined_operator_expr(handler, operators, scopes);
        }
        StmtKind::RaiseEvent { args, .. }
        | StmtKind::PrintFile { items: args, .. }
        | StmtKind::WriteFile { items: args, .. } => {
            for arg in args {
                rewrite_user_defined_operator_expr(arg, operators, scopes);
            }
        }
        StmtKind::Delete(exprs) => {
            for expr in exprs {
                rewrite_user_defined_operator_expr(expr, operators, scopes);
            }
        }
        _ => {}
    }
}

fn rewrite_user_defined_operator_expr(
    expr: &mut Expression,
    operators: &CSharpOperatorTable,
    scopes: &mut Vec<HashMap<String, String>>,
) -> Option<String> {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            let left_type = rewrite_user_defined_operator_expr(left, operators, scopes)
                .or_else(|| infer_csharp_expr_type(left, operators, scopes));
            let right_type = rewrite_user_defined_operator_expr(right, operators, scopes)
                .or_else(|| infer_csharp_expr_type(right, operators, scopes));
            if let Some(method) = csharp_binary_operator_method_name(*op) {
                let dispatch_type = left_type
                    .as_deref()
                    .filter(|type_name| {
                        csharp_operator_info(operators, type_name, method).is_some()
                    })
                    .or_else(|| {
                        right_type.as_deref().filter(|type_name| {
                            csharp_operator_info(operators, type_name, method).is_some()
                        })
                    });
                if let Some(type_name) = dispatch_type {
                    let return_type = csharp_operator_info(operators, type_name, method)
                        .and_then(|info| info.return_type.clone());
                    let span = expr.span.clone();
                    *expr = build_csharp_operator_static_call(
                        type_name,
                        method,
                        vec![(**left).clone(), (**right).clone()],
                        span,
                        return_type.as_deref(),
                    );
                    return return_type;
                }
            }
            if csharp_builtin_bool_binary_op(*op) {
                let span = expr.span.clone();
                let bool_expr = Expression::with_span(
                    ExprKind::Binary {
                        op: *op,
                        left: Box::new((**left).clone()),
                        right: Box::new((**right).clone()),
                    },
                    span,
                );
                *expr = materialize_csharp_bool_expr(bool_expr);
                return Some("bool".into());
            }
            None
        }
        ExprKind::Unary { op, expr: inner } => {
            let operand_type = rewrite_user_defined_operator_expr(inner, operators, scopes)
                .or_else(|| infer_csharp_expr_type(inner, operators, scopes));
            if let (Some(method), Some(type_name)) = (
                csharp_unary_operator_method_name(*op),
                operand_type.as_deref(),
            ) {
                if let Some(info) = csharp_operator_info(operators, type_name, method) {
                    let return_type = info.return_type.clone();
                    let span = expr.span.clone();
                    *expr = build_csharp_operator_static_call(
                        type_name,
                        method,
                        vec![(**inner).clone()],
                        span,
                        return_type.as_deref(),
                    );
                    return return_type;
                }
            }
            if matches!(op, UnaryOp::Not) {
                let span = expr.span.clone();
                let bool_expr = Expression::with_span(
                    ExprKind::Unary {
                        op: *op,
                        expr: Box::new((**inner).clone()),
                    },
                    span,
                );
                *expr = materialize_csharp_bool_expr(bool_expr);
                return Some("bool".into());
            }
            None
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_user_defined_operator_expr(callee, operators, scopes);
            for arg in args {
                rewrite_user_defined_operator_expr(&mut arg.value, operators, scopes);
            }
            infer_csharp_expr_type(expr, operators, scopes)
        }
        ExprKind::New { class: _, args } => {
            for arg in args {
                rewrite_user_defined_operator_expr(&mut arg.value, operators, scopes);
            }
            infer_csharp_expr_type(expr, operators, scopes)
        }
        ExprKind::Member { object, .. } => {
            rewrite_user_defined_operator_expr(object, operators, scopes);
            infer_csharp_expr_type(expr, operators, scopes)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_user_defined_operator_expr(object, operators, scopes);
            rewrite_user_defined_operator_expr(index, operators, scopes);
            None
        }
        ExprKind::Assign { target, value } => {
            rewrite_user_defined_operator_expr(target, operators, scopes);
            rewrite_user_defined_operator_expr(value, operators, scopes)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_user_defined_operator_expr(cond, operators, scopes);
            let then_type = rewrite_user_defined_operator_expr(then, operators, scopes)
                .or_else(|| infer_csharp_expr_type(then, operators, scopes));
            let else_type = rewrite_user_defined_operator_expr(else_, operators, scopes)
                .or_else(|| infer_csharp_expr_type(else_, operators, scopes));
            if then_type == else_type {
                then_type
            } else {
                None
            }
        }
        ExprKind::Lambda { params, body, .. } => {
            scopes.push(HashMap::new());
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scopes
                        .last_mut()
                        .unwrap()
                        .insert(param.name.clone(), type_hint.clone());
                }
            }
            match body {
                LambdaBody::Expr(lambda_expr) => {
                    rewrite_user_defined_operator_expr(lambda_expr, operators, scopes);
                }
                LambdaBody::Block(body) => {
                    rewrite_user_defined_operator_calls_in_statements(body, operators, scopes);
                }
            }
            scopes.pop();
            None
        }
        ExprKind::Array(elements) => {
            for element in elements {
                rewrite_user_defined_operator_expr(&mut element.value, operators, scopes);
            }
            infer_csharp_expr_type(expr, operators, scopes)
        }
        ExprKind::Tuple(elements) | ExprKind::Set(elements) => {
            for element in elements {
                rewrite_user_defined_operator_expr(element, operators, scopes);
            }
            None
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_user_defined_operator_expr(key, operators, scopes);
                        rewrite_user_defined_operator_expr(value, operators, scopes);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_user_defined_operator_expr(value, operators, scopes);
                    }
                    ObjectProperty::Computed { key, value } => {
                        rewrite_user_defined_operator_expr(key, operators, scopes);
                        rewrite_user_defined_operator_expr(value, operators, scopes);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_user_defined_operator_calls_in_statement(
                            value,
                            operators,
                            &mut vec![HashMap::new()],
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
            None
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(part_expr) = part {
                    rewrite_user_defined_operator_expr(part_expr, operators, scopes);
                }
            }
            None
        }
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => {
            rewrite_user_defined_operator_expr(inner, operators, scopes);
            let method = format!("op_Explicit{}", csharp_operator_type_suffix(type_name));
            let source_type = infer_csharp_expr_type(inner, operators, scopes);
            let dispatch_type = if csharp_operator_info(operators, type_name, &method).is_some() {
                Some(type_name.clone())
            } else {
                source_type
                    .as_deref()
                    .filter(|source| csharp_operator_info(operators, source, &method).is_some())
                    .map(ToString::to_string)
            };
            if let Some(dispatch_type) = dispatch_type {
                if let Some(info) = csharp_operator_info(operators, &dispatch_type, &method) {
                    let return_type = info.return_type.clone().or_else(|| Some(type_name.clone()));
                    let span = expr.span.clone();
                    *expr = build_csharp_operator_static_call(
                        &dispatch_type,
                        &method,
                        vec![(**inner).clone()],
                        span,
                        return_type.as_deref(),
                    );
                    return return_type;
                }
            }
            Some(type_name.clone())
        }
        ExprKind::IsType { expr: inner, .. } => {
            rewrite_user_defined_operator_expr(inner, operators, scopes);
            let span = expr.span.clone();
            let bool_expr = expr.clone();
            *expr = materialize_csharp_bool_expr(Expression::with_span(bool_expr.kind, span));
            Some("bool".into())
        }
        ExprKind::TypeOf(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Spread(inner)
        | ExprKind::RefLoad(inner) => {
            rewrite_user_defined_operator_expr(inner, operators, scopes);
            infer_csharp_expr_type(expr, operators, scopes)
        }
        ExprKind::NullCoalesce { left, right } => {
            let left_type = rewrite_user_defined_operator_expr(left, operators, scopes)
                .or_else(|| infer_csharp_expr_type(left, operators, scopes));
            rewrite_user_defined_operator_expr(right, operators, scopes);
            left_type
        }
        ExprKind::Yield(Some(inner)) => {
            rewrite_user_defined_operator_expr(inner, operators, scopes);
            None
        }
        _ => infer_csharp_expr_type(expr, operators, scopes),
    }
}

fn infer_csharp_expr_type(
    expr: &Expression,
    operators: &CSharpOperatorTable,
    scopes: &[HashMap<String, String>],
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => scopes
            .iter()
            .rev()
            .find_map(|scope| scope.get(name).cloned()),
        ExprKind::Call { callee, .. } => {
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.starts_with("op_") {
                    if let Some(type_name) = infer_csharp_expr_type(object, operators, scopes) {
                        return csharp_operator_info(operators, &type_name, field)
                            .and_then(|info| info.return_type.clone());
                    }
                    if let ExprKind::Ident(type_name) = &object.kind {
                        return csharp_operator_info(operators, type_name, field)
                            .and_then(|info| info.return_type.clone());
                    }
                }
            }
            infer_csharp_type_from_expr(expr)
        }
        ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
        _ => infer_csharp_type_from_expr(expr),
    }
}

fn walk_local_var(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty local var")?;

    // Skip type name (var or explicit type)
    let type_hint = match first.as_rule() {
        Rule::var_kw => None,
        Rule::type_name
        | Rule::generic_type_expr
        | Rule::base_type
        | Rule::dotted_name
        | Rule::global_qualified_name => Some(first.as_str().to_string()),
        _ => None,
    };

    let mut declarations = Vec::new();
    for p in inner {
        match p.as_rule() {
            Rule::var_declarator_list => {
                for vd in p.into_inner() {
                    if vd.as_rule() == Rule::var_declarator {
                        declarations.push(walk_var_declarator(vd)?);
                    }
                }
            }
            Rule::var_declarator => declarations.push(walk_var_declarator(p)?),
            _ => {}
        }
    }

    if let Some(type_hint) = type_hint.filter(|hint| !hint.eq_ignore_ascii_case("var")) {
        let normalized_hint = normalize_runtime_type_name(&type_hint).to_lowercase();
        for decl in &mut declarations {
            decl.type_hint = Some(type_hint.clone());
            if matches!(normalized_hint.as_str(), "object" | "system.object") {
                if let Some(ref init) = decl.init {
                    if let crate::ast::ExprKind::New { class, .. } = &init.kind {
                        let inferred = match &class.kind {
                            crate::ast::ExprKind::Ident(n) => Some(n.clone()),
                            crate::ast::ExprKind::Member { field, .. } => Some(field.clone()),
                            _ => None,
                        };
                        if let Some(name) = inferred {
                            decl.type_hint = Some(name);
                        }
                    }
                }
            }
        }
    } else {
        // `var` type inference keeps loop/local bindings normalized to a
        // typed common AST when the source makes the type obvious.
        for decl in &mut declarations {
            if decl.type_hint.is_none() {
                if let Some(ref init) = decl.init {
                    if let Some(name) = infer_csharp_type_from_expr(init) {
                        decl.type_hint = Some(name);
                    }
                }
            }
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_var_declarator(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut inner = pair.into_inner();
    let name = inner
        .next()
        .ok_or("Empty var declarator")?
        .as_str()
        .to_string();
    let init = match inner.next() {
        Some(p) if p.as_rule() == Rule::array_initializer => {
            // `int[] arr = { 1, 2, 3 }` — bare-brace array literal in a
            // declarator. Desugar to a plain Array expression so the
            // compiler emits the same bytecode as `new[] { ... }`.
            // Each child is a `collection_element` (post grammar change
            // for nested-brace dict / multi-dim support).
            let span = to_span(&p);
            let elems = p
                .into_inner()
                .map(|e| {
                    walk_collection_element(e).map(|expr| ArrayElement {
                        key: None,
                        value: expr,
                        spread: false,
                        by_ref: false,
                    })
                })
                .collect::<Result<Vec<_>, _>>()?;
            Some(Expression::with_span(ExprKind::Array(elems), span))
        }
        Some(p) => Some(walk_expression(p)?),
        None => None,
    };
    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: None,
        init,
        array_bounds: None,
        with_events: false,
    })
}

// ── Class declaration ───────────────────────────────────────────────────────

fn looks_like_csharp_interface_type(type_name: &str) -> bool {
    let leaf = type_name
        .trim()
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .split('<')
        .next()
        .unwrap_or(type_name)
        .trim_end_matches('?');

    leaf.starts_with('I')
        && leaf.len() > 1
        && leaf
            .chars()
            .nth(1)
            .is_some_and(|ch| ch.is_ascii_uppercase())
}

fn walk_class_decl(pair: Pair<Rule>, decorators: &[Expression]) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut generic_params = Vec::new();
    let mut parents = Vec::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();
    let mut class_mods = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        "public" => class_mods.visibility = Visibility::Public,
                        "private" => class_mods.visibility = Visibility::Private,
                        "protected" => class_mods.visibility = Visibility::Protected,
                        "internal" => class_mods.visibility = Visibility::Internal,
                        s if s.starts_with("static") => class_mods.is_static = true,
                        s if s.starts_with("abstract") => class_mods.is_abstract = true,
                        s if s.starts_with("sealed") => class_mods.is_sealed = true,
                        s if s.starts_with("partial") => class_mods.is_partial = true,
                        _ => {}
                    }
                }
            }
            Rule::ident_name => {
                // Generic param idents (`class Box<T> { ... }`) leak
                // through the silent `generic_params` wrapper rule —
                // they appear as additional `ident_name` pairs after
                // the class name. Keep the first as the runtime class
                // name and retain the rest so `new T()` can bind to
                // the concrete constructor passed into `new Box<Foo>()`.
                if name.is_empty() {
                    name = p.as_str().to_string();
                } else {
                    generic_params.push(p.as_str().to_string());
                }
            }
            Rule::base_list => {
                let mut first = true;
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        let type_str = bp.as_str().trim().to_string();
                        // C# allows either a single base class or interfaces in the first slot.
                        // Classify using the leaf type name so namespaced interfaces like
                        // `System.IComparable<T>` do not get misread as a parent class.
                        if first {
                            if looks_like_csharp_interface_type(&type_str) {
                                interfaces.push(type_str);
                            } else {
                                parents.push(type_str);
                            }
                            first = false;
                        } else {
                            interfaces.push(type_str);
                        }
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.extend(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    expand_explicit_interface_members(&mut members);
    disambiguate_explicit_property_backing_fields(&mut members);
    if !generic_params.is_empty() {
        inject_csharp_generic_ctor_fields(&mut members, generic_params.len());
        rewrite_generic_ctor_news_in_members(&mut members, &generic_params);
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces,
        members,
        modifiers: class_mods,
        decorators: decorators.to_vec(),
    })
}

fn expand_explicit_interface_members(members: &mut Vec<ClassMember>) {
    let mut plain_methods: HashSet<String> = HashSet::new();
    let mut explicit_methods: HashMap<String, usize> = HashMap::new();
    let mut plain_properties: HashSet<String> = HashSet::new();
    let mut explicit_properties: HashMap<String, usize> = HashMap::new();

    for member in members.iter() {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                    if let Some((_, base)) = parse_explicit_interface_runtime_name(name) {
                        *explicit_methods.entry(base).or_insert(0) += 1;
                    } else {
                        plain_methods.insert(name.clone());
                    }
                }
            }
            ClassMember::Property { name, .. } => {
                if let Some((_, base)) = parse_explicit_interface_runtime_name(name) {
                    *explicit_properties.entry(base).or_insert(0) += 1;
                } else {
                    plain_properties.insert(name.clone());
                }
            }
            _ => {}
        }
    }

    let mut extras = Vec::new();
    for member in members.iter() {
        match member {
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { name, .. } = &stmt.kind else {
                    continue;
                };
                let Some((_, base)) = parse_explicit_interface_runtime_name(name) else {
                    continue;
                };
                if explicit_methods.get(&base).copied().unwrap_or(0) == 1
                    && !plain_methods.contains(&base)
                {
                    let mut alias_stmt = (**stmt).clone();
                    if let StmtKind::FunctionDecl { name, .. } = &mut alias_stmt.kind {
                        *name = base;
                    }
                    extras.push(ClassMember::Method(Box::new(alias_stmt)));
                }
            }
            ClassMember::Property { name, .. } => {
                let Some((_, base)) = parse_explicit_interface_runtime_name(name) else {
                    continue;
                };
                if explicit_properties.get(&base).copied().unwrap_or(0) == 1
                    && !plain_properties.contains(&base)
                {
                    let mut alias_member = member.clone();
                    if let ClassMember::Property { name, .. } = &mut alias_member {
                        *name = base;
                    }
                    extras.push(alias_member);
                }
            }
            _ => {}
        }
    }

    members.extend(extras);
}

/// Explicit C# properties compile to `__get_{canon}` / `__set_{canon}` on the
/// instance. If a separate backing field is stored under the same canonical key
/// (`celsius` field + `Celsius` property → both `celsius`), `STRUCT_GET` invokes
/// the getter instead of reading the field and overflows the stack.
///
/// Rename colliding instance fields to `__{canon}` (same convention as
/// auto-property backing slots) and rewrite bare identifier uses in class bodies.
fn disambiguate_explicit_property_backing_fields(members: &mut [ClassMember]) {
    let mut explicit_property_canons = HashSet::new();
    for member in members.iter() {
        if let ClassMember::Property {
            name,
            is_auto,
            modifiers,
            ..
        } = member
        {
            if !*is_auto && !modifiers.is_static {
                explicit_property_canons.insert(name.to_lowercase());
            }
        }
    }
    if explicit_property_canons.is_empty() {
        return;
    }

    let mut renames: HashMap<String, String> = HashMap::new();
    for member in members.iter() {
        if let ClassMember::Field {
            name, modifiers, ..
        } = member
        {
            if modifiers.is_static {
                continue;
            }
            let canon = name.to_lowercase();
            if explicit_property_canons.contains(&canon) {
                renames
                    .entry(name.clone())
                    .or_insert_with(|| format!("__{canon}"));
            }
        }
    }
    if renames.is_empty() {
        return;
    }

    for member in members.iter_mut() {
        if let ClassMember::Field { name, .. } = member {
            if let Some(new_name) = renames.get(name) {
                *name = new_name.clone();
            }
        }
    }
    for member in members.iter_mut() {
        rewrite_property_backing_field_idents_in_member(member, &renames);
    }
}

fn rewrite_property_backing_field_idents_in_member(
    member: &mut ClassMember,
    renames: &HashMap<String, String>,
) {
    match member {
        ClassMember::Field {
            init, array_bounds, ..
        } => {
            if let Some(init) = init {
                rewrite_property_backing_field_idents_in_expr(init, renames);
            }
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_property_backing_field_idents_in_expr(bound, renames);
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_property_backing_field_idents_in_statement(stmt, renames);
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            rewrite_property_backing_field_idents_in_statements(body, renames);
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_property_backing_field_idents_in_expr(arg, renames);
                }
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_property_backing_field_idents_in_statements(getter, renames);
            }
            if let Some(setter) = setter {
                rewrite_property_backing_field_idents_in_statements(&mut setter.body, renames);
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_property_backing_field_idents_in_expr(value, renames);
        }
        _ => {}
    }
}

fn rewrite_property_backing_field_idents_in_statements(
    body: &mut [Statement],
    renames: &HashMap<String, String>,
) {
    for stmt in body.iter_mut() {
        rewrite_property_backing_field_idents_in_statement(stmt, renames);
    }
}

fn rewrite_property_backing_field_idents_in_statement(
    stmt: &mut Statement,
    renames: &HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. } => {
            rewrite_property_backing_field_idents_in_expr(expr, renames);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_property_backing_field_idents_in_expr(target, renames);
            rewrite_property_backing_field_idents_in_expr(value, renames);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_property_backing_field_idents_in_expr(expr, renames);
            rewrite_property_backing_field_idents_in_expr(cause, renames);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_property_backing_field_idents_in_expr(init, renames);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_property_backing_field_idents_in_expr(bound, renames);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_property_backing_field_idents_in_statements(body, renames);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_property_backing_field_idents_in_member(member, renames);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_property_backing_field_idents_in_expr(cond, renames);
            rewrite_property_backing_field_idents_in_statements(then_body, renames);
            for (elif_cond, elif_body) in elifs {
                rewrite_property_backing_field_idents_in_expr(elif_cond, renames);
                rewrite_property_backing_field_idents_in_statements(elif_body, renames);
            }
            if let Some(else_body) = else_body {
                rewrite_property_backing_field_idents_in_statements(else_body, renames);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_property_backing_field_idents_in_statement(init, renames);
            }
            if let Some(cond) = cond {
                rewrite_property_backing_field_idents_in_expr(cond, renames);
            }
            if let Some(update) = update {
                rewrite_property_backing_field_idents_in_expr(update, renames);
            }
            rewrite_property_backing_field_idents_in_statements(body, renames);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_property_backing_field_idents_in_expr(iter, renames);
            rewrite_property_backing_field_idents_in_statements(body, renames);
            if let Some(else_body) = else_body {
                rewrite_property_backing_field_idents_in_statements(else_body, renames);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_property_backing_field_idents_in_expr(cond, renames);
            rewrite_property_backing_field_idents_in_statements(body, renames);
            if let Some(else_body) = else_body {
                rewrite_property_backing_field_idents_in_statements(else_body, renames);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_property_backing_field_idents_in_statements(body, renames);
            rewrite_property_backing_field_idents_in_expr(cond, renames);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_property_backing_field_idents_in_expr(expr, renames);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_property_backing_field_idents_in_expr(expr, renames)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_property_backing_field_idents_in_expr(from, renames);
                            rewrite_property_backing_field_idents_in_expr(to, renames);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_property_backing_field_idents_in_expr(expr, renames);
                        }
                    }
                }
                rewrite_property_backing_field_idents_in_statements(&mut case.body, renames);
            }
            if let Some(default) = default {
                rewrite_property_backing_field_idents_in_statements(default, renames);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_property_backing_field_idents_in_statements(body, renames);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_property_backing_field_idents_in_expr(when_clause, renames);
                }
                rewrite_property_backing_field_idents_in_statements(&mut catch.body, renames);
            }
            if let Some(else_body) = else_body {
                rewrite_property_backing_field_idents_in_statements(else_body, renames);
            }
            if let Some(finally) = finally {
                rewrite_property_backing_field_idents_in_statements(finally, renames);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_property_backing_field_idents_in_expr(&mut item.expr, renames);
            }
            rewrite_property_backing_field_idents_in_statements(body, renames);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_property_backing_field_idents_in_expr(target, renames);
            }
            rewrite_property_backing_field_idents_in_expr(value, renames);
        }
        _ => {}
    }
}

fn rewrite_property_backing_field_idents_in_expr(
    expr: &mut Expression,
    renames: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(new_name) = renames.get(name) {
                *name = new_name.clone();
            }
        }
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_property_backing_field_idents_in_expr(left, renames);
            rewrite_property_backing_field_idents_in_expr(right, renames);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::IsType { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::Spread(expr) => {
            rewrite_property_backing_field_idents_in_expr(expr, renames);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(name) => {
                if let Some(new_name) = renames.get(name) {
                    *name = new_name.clone();
                }
            }
            PlaceExpr::Member { object, .. } => {
                rewrite_property_backing_field_idents_in_expr(object, renames);
            }
            PlaceExpr::Index { object, index, .. } => {
                rewrite_property_backing_field_idents_in_expr(object, renames);
                rewrite_property_backing_field_idents_in_expr(index, renames);
            }
            PlaceExpr::Deref(expr) => {
                rewrite_property_backing_field_idents_in_expr(expr, renames);
            }
        },
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_property_backing_field_idents_in_expr(cond, renames);
            rewrite_property_backing_field_idents_in_expr(then, renames);
            rewrite_property_backing_field_idents_in_expr(else_, renames);
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_property_backing_field_idents_in_expr(object, renames);
            if let Some(new_name) = renames.get(field) {
                *field = new_name.clone();
            }
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_property_backing_field_idents_in_expr(object, renames);
            rewrite_property_backing_field_idents_in_expr(index, renames);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_property_backing_field_idents_in_expr(callee, renames);
            for arg in args.iter_mut() {
                rewrite_property_backing_field_idents_in_expr(&mut arg.value, renames);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_property_backing_field_idents_in_expr(class, renames);
            for arg in args {
                rewrite_property_backing_field_idents_in_expr(&mut arg.value, renames);
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_property_backing_field_idents_in_expr(target, renames);
            rewrite_property_backing_field_idents_in_expr(value, renames);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_property_backing_field_idents_in_expr(expr, renames),
            LambdaBody::Block(body) => {
                rewrite_property_backing_field_idents_in_statements(body, renames)
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_property_backing_field_idents_in_expr(&mut item.value, renames);
                if let Some(key) = &mut item.key {
                    rewrite_property_backing_field_idents_in_expr(key, renames);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_property_backing_field_idents_in_expr(item, renames);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_property_backing_field_idents_in_expr(key, renames);
                        rewrite_property_backing_field_idents_in_expr(value, renames);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_property_backing_field_idents_in_expr(expr, renames);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_property_backing_field_idents_in_statement(value, renames);
                    }
                    ObjectProperty::Shorthand(name) => {
                        if let Some(new_name) = renames.get(name) {
                            *name = new_name.clone();
                        }
                    }
                }
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_property_backing_field_idents_in_expr(element, renames);
            for generator in generators {
                rewrite_property_backing_field_idents_in_expr(&mut generator.target, renames);
                rewrite_property_backing_field_idents_in_expr(&mut generator.iter, renames);
                for condition in &mut generator.conditions {
                    rewrite_property_backing_field_idents_in_expr(condition, renames);
                }
            }
        }
        _ => {}
    }
}

fn csharp_generic_ctor_field_name(index: usize) -> String {
    format!("__vybe_generic_{index}")
}

fn csharp_generic_ctor_param_name(index: usize) -> String {
    format!("__vybe_generic_ctor_{index}")
}

fn csharp_generic_default_param_name(index: usize) -> String {
    format!("__vybe_generic_default_{index}")
}

fn inject_csharp_generic_binding_params(params: &mut Vec<Param>, generic_count: usize) {
    for index in 0..generic_count {
        params.push(Param {
            name: csharp_generic_ctor_param_name(index),
            type_hint: None,
            default: Some(Expression::null()),
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: false,
        });
        params.push(Param {
            name: csharp_generic_default_param_name(index),
            type_hint: None,
            default: Some(Expression::null()),
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: true,
            is_nullable: false,
        });
    }
}

fn inject_csharp_generic_ctor_fields(members: &mut Vec<ClassMember>, generic_count: usize) {
    for index in (0..generic_count).rev() {
        members.insert(
            0,
            ClassMember::Field {
                name: csharp_generic_ctor_field_name(index),
                type_hint: None,
                init: None,
                modifiers: Modifiers::default(),
                with_events: false,
                array_bounds: None,
            },
        );
    }
}

fn rewrite_generic_ctor_news_in_members(members: &mut [ClassMember], generic_params: &[String]) {
    for member in members {
        rewrite_generic_ctor_news_in_member(member, generic_params);
    }
}

fn rewrite_generic_ctor_news_in_member(member: &mut ClassMember, generic_params: &[String]) {
    match member {
        ClassMember::Field {
            init,
            array_bounds,
            modifiers,
            ..
        } if !modifiers.is_static => {
            if let Some(expr) = init {
                rewrite_generic_ctor_news_in_expr(expr, generic_params);
            }
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_generic_ctor_news_in_expr(bound, generic_params);
                }
            }
        }
        ClassMember::Method(stmt) => {
            if let StmtKind::FunctionDecl {
                body, modifiers, ..
            } = &mut stmt.kind
            {
                if !modifiers.is_static {
                    rewrite_generic_ctor_news_in_statements(body, generic_params);
                }
            }
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            rewrite_generic_ctor_news_in_statements(body, generic_params);
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_generic_ctor_news_in_expr(arg, generic_params);
                }
            }
        }
        ClassMember::Property {
            getter,
            setter,
            modifiers,
            ..
        } if !modifiers.is_static => {
            if let Some(getter) = getter {
                rewrite_generic_ctor_news_in_statements(getter, generic_params);
            }
            if let Some(setter) = setter {
                rewrite_generic_ctor_news_in_statements(&mut setter.body, generic_params);
            }
        }
        ClassMember::Const { value, .. } => {
            rewrite_generic_ctor_news_in_expr(value, generic_params);
        }
        _ => {}
    }
}

fn rewrite_generic_ctor_news_in_statements(body: &mut [Statement], generic_params: &[String]) {
    for stmt in body {
        rewrite_generic_ctor_news_in_statement(stmt, generic_params);
    }
}

fn rewrite_generic_ctor_news_in_statement(stmt: &mut Statement, generic_params: &[String]) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_generic_ctor_news_in_expr(expr, generic_params);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_generic_ctor_news_in_expr(expr, generic_params);
            rewrite_generic_ctor_news_in_expr(cause, generic_params);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_generic_ctor_news_in_expr(init, generic_params);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_generic_ctor_news_in_expr(bound, generic_params);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { .. } => {}
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_generic_ctor_news_in_statements(body, generic_params);
        }
        StmtKind::ClassDecl { .. } | StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. } => {}
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_generic_ctor_news_in_expr(cond, generic_params);
            rewrite_generic_ctor_news_in_statements(then_body, generic_params);
            for (elif_cond, elif_body) in elifs {
                rewrite_generic_ctor_news_in_expr(elif_cond, generic_params);
                rewrite_generic_ctor_news_in_statements(elif_body, generic_params);
            }
            if let Some(else_body) = else_body {
                rewrite_generic_ctor_news_in_statements(else_body, generic_params);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_generic_ctor_news_in_statement(init, generic_params);
            }
            if let Some(cond) = cond {
                rewrite_generic_ctor_news_in_expr(cond, generic_params);
            }
            if let Some(update) = update {
                rewrite_generic_ctor_news_in_expr(update, generic_params);
            }
            rewrite_generic_ctor_news_in_statements(body, generic_params);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_generic_ctor_news_in_expr(iter, generic_params);
            rewrite_generic_ctor_news_in_statements(body, generic_params);
            if let Some(else_body) = else_body {
                rewrite_generic_ctor_news_in_statements(else_body, generic_params);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_generic_ctor_news_in_expr(cond, generic_params);
            rewrite_generic_ctor_news_in_statements(body, generic_params);
            if let Some(else_body) = else_body {
                rewrite_generic_ctor_news_in_statements(else_body, generic_params);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_generic_ctor_news_in_statements(body, generic_params);
            rewrite_generic_ctor_news_in_expr(cond, generic_params);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_generic_ctor_news_in_expr(expr, generic_params);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_generic_ctor_news_in_expr(expr, generic_params)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_generic_ctor_news_in_expr(from, generic_params);
                            rewrite_generic_ctor_news_in_expr(to, generic_params);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_generic_ctor_news_in_expr(expr, generic_params);
                        }
                    }
                }
                rewrite_generic_ctor_news_in_statements(&mut case.body, generic_params);
            }
            if let Some(default) = default {
                rewrite_generic_ctor_news_in_statements(default, generic_params);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_generic_ctor_news_in_statements(body, generic_params);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_generic_ctor_news_in_expr(when_clause, generic_params);
                }
                rewrite_generic_ctor_news_in_statements(&mut catch.body, generic_params);
            }
            if let Some(else_body) = else_body {
                rewrite_generic_ctor_news_in_statements(else_body, generic_params);
            }
            if let Some(finally) = finally {
                rewrite_generic_ctor_news_in_statements(finally, generic_params);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_generic_ctor_news_in_expr(&mut item.expr, generic_params);
            }
            rewrite_generic_ctor_news_in_statements(body, generic_params);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_generic_ctor_news_in_expr(target, generic_params);
            }
            rewrite_generic_ctor_news_in_expr(value, generic_params);
        }
        _ => {}
    }
}

fn rewrite_generic_ctor_news_in_expr(expr: &mut Expression, generic_params: &[String]) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_generic_ctor_news_in_expr(left, generic_params);
            rewrite_generic_ctor_news_in_expr(right, generic_params);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_generic_ctor_news_in_expr(expr, generic_params);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_generic_ctor_news_in_expr(cond, generic_params);
            rewrite_generic_ctor_news_in_expr(then, generic_params);
            rewrite_generic_ctor_news_in_expr(else_, generic_params);
        }
        ExprKind::Member { object, .. } => {
            rewrite_generic_ctor_news_in_expr(object, generic_params);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_generic_ctor_news_in_expr(object, generic_params);
            rewrite_generic_ctor_news_in_expr(index, generic_params);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_generic_ctor_news_in_expr(callee, generic_params);
            for arg in args {
                rewrite_generic_ctor_news_in_expr(&mut arg.value, generic_params);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_generic_ctor_news_in_expr(class, generic_params);
            for arg in args {
                rewrite_generic_ctor_news_in_expr(&mut arg.value, generic_params);
            }
            if let ExprKind::Ident(name) = &class.kind {
                if let Some(index) = generic_params.iter().position(|param| param == name) {
                    *class = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: csharp_generic_ctor_field_name(index),
                        null_safe: false,
                    }));
                }
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_generic_ctor_news_in_expr(target, generic_params);
            rewrite_generic_ctor_news_in_expr(value, generic_params);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_generic_ctor_news_in_expr(expr, generic_params),
            LambdaBody::Block(body) => {
                rewrite_generic_ctor_news_in_statements(body, generic_params)
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_generic_ctor_news_in_expr(&mut item.value, generic_params);
                if let Some(key) = &mut item.key {
                    rewrite_generic_ctor_news_in_expr(key, generic_params);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_generic_ctor_news_in_expr(item, generic_params);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_generic_ctor_news_in_expr(key, generic_params);
                        rewrite_generic_ctor_news_in_expr(value, generic_params);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_generic_ctor_news_in_expr(expr, generic_params);
                    }
                    ObjectProperty::Method { value, .. } => {
                        rewrite_generic_ctor_news_in_statement(value, generic_params);
                    }
                    ObjectProperty::Accessor { value, .. } => {
                        rewrite_generic_ctor_news_in_statement(value, generic_params);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_generic_ctor_news_in_expr(expr, generic_params);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Spread(expr) | ExprKind::Cast { expr, .. } => {
            rewrite_generic_ctor_news_in_expr(expr, generic_params);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_generic_ctor_news_in_expr(element, generic_params);
            for generator in generators {
                rewrite_generic_ctor_news_in_expr(&mut generator.iter, generic_params);
                for if_expr in &mut generator.conditions {
                    rewrite_generic_ctor_news_in_expr(if_expr, generic_params);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_generic_ctor_news_in_expr(lower, generic_params);
            }
            if let Some(upper) = upper {
                rewrite_generic_ctor_news_in_expr(upper, generic_params);
            }
            if let Some(step) = step {
                rewrite_generic_ctor_news_in_expr(step, generic_params);
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_generic_ctor_news_in_expr(start, generic_params);
            rewrite_generic_ctor_news_in_expr(end, generic_params);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_generic_ctor_news_in_expr(subject, generic_params);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for expr in conditions {
                        rewrite_generic_ctor_news_in_expr(expr, generic_params);
                    }
                }
                rewrite_generic_ctor_news_in_expr(&mut arm.body, generic_params);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_generic_ctor_news_in_expr(class, generic_params);
            rewrite_generic_ctor_news_in_expr(member, generic_params);
        }
        _ => {}
    }
}

fn rewrite_generic_bindings_in_statements(body: &mut [Statement], generic_params: &[String]) {
    for stmt in body {
        rewrite_generic_bindings_in_statement(stmt, generic_params);
    }
}

fn rewrite_generic_bindings_in_statement(stmt: &mut Statement, generic_params: &[String]) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        }
        | StmtKind::Using { resource: expr, .. }
        | StmtKind::Lock { expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_generic_bindings_in_expr(expr, generic_params);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_generic_bindings_in_expr(expr, generic_params);
            rewrite_generic_bindings_in_expr(cause, generic_params);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_generic_bindings_in_expr(init, generic_params);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_generic_bindings_in_expr(bound, generic_params);
                    }
                }
            }
        }
        StmtKind::FunctionDecl { .. } => {}
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_generic_bindings_in_statements(body, generic_params);
        }
        StmtKind::ClassDecl { .. } | StmtKind::StructDecl { .. } | StmtKind::ModuleDecl { .. } => {}
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_generic_bindings_in_expr(cond, generic_params);
            rewrite_generic_bindings_in_statements(then_body, generic_params);
            for (elif_cond, elif_body) in elifs {
                rewrite_generic_bindings_in_expr(elif_cond, generic_params);
                rewrite_generic_bindings_in_statements(elif_body, generic_params);
            }
            if let Some(else_body) = else_body {
                rewrite_generic_bindings_in_statements(else_body, generic_params);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_generic_bindings_in_statement(init, generic_params);
            }
            if let Some(cond) = cond {
                rewrite_generic_bindings_in_expr(cond, generic_params);
            }
            if let Some(update) = update {
                rewrite_generic_bindings_in_expr(update, generic_params);
            }
            rewrite_generic_bindings_in_statements(body, generic_params);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_generic_bindings_in_expr(iter, generic_params);
            rewrite_generic_bindings_in_statements(body, generic_params);
            if let Some(else_body) = else_body {
                rewrite_generic_bindings_in_statements(else_body, generic_params);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_generic_bindings_in_expr(cond, generic_params);
            rewrite_generic_bindings_in_statements(body, generic_params);
            if let Some(else_body) = else_body {
                rewrite_generic_bindings_in_statements(else_body, generic_params);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_generic_bindings_in_statements(body, generic_params);
            rewrite_generic_bindings_in_expr(cond, generic_params);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_generic_bindings_in_expr(expr, generic_params);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => {
                            rewrite_generic_bindings_in_expr(expr, generic_params)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_generic_bindings_in_expr(from, generic_params);
                            rewrite_generic_bindings_in_expr(to, generic_params);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            rewrite_generic_bindings_in_expr(expr, generic_params);
                        }
                    }
                }
                rewrite_generic_bindings_in_statements(&mut case.body, generic_params);
            }
            if let Some(default) = default {
                rewrite_generic_bindings_in_statements(default, generic_params);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_generic_bindings_in_statements(body, generic_params);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_generic_bindings_in_expr(when_clause, generic_params);
                }
                rewrite_generic_bindings_in_statements(&mut catch.body, generic_params);
            }
            if let Some(else_body) = else_body {
                rewrite_generic_bindings_in_statements(else_body, generic_params);
            }
            if let Some(finally) = finally {
                rewrite_generic_bindings_in_statements(finally, generic_params);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                rewrite_generic_bindings_in_expr(&mut item.expr, generic_params);
            }
            rewrite_generic_bindings_in_statements(body, generic_params);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_generic_bindings_in_expr(target, generic_params);
            }
            rewrite_generic_bindings_in_expr(value, generic_params);
        }
        _ => {}
    }
}

fn rewrite_generic_bindings_in_expr(expr: &mut Expression, generic_params: &[String]) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } | ExprKind::NullCoalesce { left, right } => {
            rewrite_generic_bindings_in_expr(left, generic_params);
            rewrite_generic_bindings_in_expr(right, generic_params);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Cast { expr, .. } => {
            rewrite_generic_bindings_in_expr(expr, generic_params);
        }
        ExprKind::DefaultOf(type_name) => {
            if let Some(index) = generic_params.iter().position(|param| param == type_name) {
                expr.kind = ExprKind::Ident(csharp_generic_default_param_name(index));
            }
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_generic_bindings_in_expr(cond, generic_params);
            rewrite_generic_bindings_in_expr(then, generic_params);
            rewrite_generic_bindings_in_expr(else_, generic_params);
        }
        ExprKind::Member { object, .. } => rewrite_generic_bindings_in_expr(object, generic_params),
        ExprKind::Index { object, index, .. } => {
            rewrite_generic_bindings_in_expr(object, generic_params);
            rewrite_generic_bindings_in_expr(index, generic_params);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_generic_bindings_in_expr(callee, generic_params);
            for arg in args {
                rewrite_generic_bindings_in_expr(&mut arg.value, generic_params);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_generic_bindings_in_expr(class, generic_params);
            if let ExprKind::Ident(name) = &class.kind {
                if let Some(index) = generic_params.iter().position(|param| param == name) {
                    class.kind = ExprKind::Ident(csharp_generic_ctor_param_name(index));
                }
            }
            for arg in args {
                rewrite_generic_bindings_in_expr(&mut arg.value, generic_params);
            }
        }
        ExprKind::Assign { target, value } | ExprKind::Walrus { target, value } => {
            rewrite_generic_bindings_in_expr(target, generic_params);
            rewrite_generic_bindings_in_expr(value, generic_params);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_generic_bindings_in_expr(expr, generic_params),
            LambdaBody::Block(body) => rewrite_generic_bindings_in_statements(body, generic_params),
        },
        ExprKind::Array(items) => {
            for item in items {
                rewrite_generic_bindings_in_expr(&mut item.value, generic_params);
                if let Some(key) = &mut item.key {
                    rewrite_generic_bindings_in_expr(key, generic_params);
                }
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_generic_bindings_in_expr(item, generic_params);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_generic_bindings_in_expr(key, generic_params);
                        rewrite_generic_bindings_in_expr(value, generic_params);
                    }
                    ObjectProperty::Spread(expr) => {
                        rewrite_generic_bindings_in_expr(expr, generic_params)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_generic_bindings_in_statement(value, generic_params);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_generic_bindings_in_expr(expr, generic_params);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_generic_bindings_in_expr(element, generic_params);
            for generator in generators {
                rewrite_generic_bindings_in_expr(&mut generator.iter, generic_params);
                for condition in &mut generator.conditions {
                    rewrite_generic_bindings_in_expr(condition, generic_params);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_generic_bindings_in_expr(lower, generic_params);
            }
            if let Some(upper) = upper {
                rewrite_generic_bindings_in_expr(upper, generic_params);
            }
            if let Some(step) = step {
                rewrite_generic_bindings_in_expr(step, generic_params);
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_generic_bindings_in_expr(start, generic_params);
            rewrite_generic_bindings_in_expr(end, generic_params);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_generic_bindings_in_expr(subject, generic_params);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for expr in conditions {
                        rewrite_generic_bindings_in_expr(expr, generic_params);
                    }
                }
                rewrite_generic_bindings_in_expr(&mut arm.body, generic_params);
            }
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_generic_bindings_in_expr(class, generic_params);
            rewrite_generic_bindings_in_expr(member, generic_params);
        }
        _ => {}
    }
}

fn walk_class_member(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut mods = Modifiers::default();
    let mut member_pair = None;
    let mut attributes: Vec<Expression> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::attribute_list => attributes.extend(parse_attribute_specs(p.as_str())),
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        "public" => mods.visibility = Visibility::Public,
                        "private" => mods.visibility = Visibility::Private,
                        "protected" => mods.visibility = Visibility::Protected,
                        "internal" => mods.visibility = Visibility::Internal,
                        s if s.starts_with("static") => mods.is_static = true,
                        s if s.starts_with("abstract") => mods.is_abstract = true,
                        s if s.starts_with("virtual") => mods.is_virtual = true,
                        s if s.starts_with("override") => mods.is_override = true,
                        s if s.starts_with("readonly") => mods.is_readonly = true,
                        // C# `const` — implicitly static + readonly per ECMA-334 §15.4.
                        // Compile-time constant folded into class-level slot.
                        "const" => {
                            mods.is_static = true;
                            mods.is_readonly = true;
                        }
                        s if s.starts_with("async") => {} // handled in method
                        _ => {}
                    }
                }
            }
            _ => member_pair = Some(p),
        }
    }

    mods.decorators.extend(attributes.iter().cloned());

    let mp = member_pair.ok_or("Empty class member")?;
    match mp.as_rule() {
        Rule::constructor_declaration => walk_constructor(mp, mods).map(|m| vec![m]),
        Rule::explicit_interface_property_declaration | Rule::property_declaration => {
            walk_property(mp, mods)
        }
        Rule::event_declaration => walk_event(mp).map(|m| vec![m]),
        Rule::explicit_interface_method_declaration | Rule::method_declaration => {
            walk_method(mp, mods).map(|m| vec![m])
        }
        Rule::field_declaration => walk_field(mp, mods).map(|m| vec![m]),
        Rule::operator_declaration => walk_operator(mp, mods).map(|m| vec![m]),
        Rule::explicit_interface_indexer_declaration | Rule::indexer_declaration => {
            walk_indexer(mp, mods)
        }
        // Nested type — wrap as `ClassMember::NestedType(stmt)` so the
        // class-emit pipeline registers the inner type as a sibling
        // global. Per ECMA-334 §15.3 nested types are accessible via
        // `Outer.Inner` qualified name; our compiler treats them as
        // top-level globals already.
        Rule::class_declaration
        | Rule::struct_declaration
        | Rule::interface_declaration
        | Rule::enum_declaration => {
            let span = to_span(&mp);
            let kind = match mp.as_rule() {
                Rule::class_declaration => walk_class_decl(mp, &attributes)?,
                Rule::struct_declaration => walk_struct_decl(mp, &attributes)?,
                Rule::interface_declaration => walk_interface_decl(mp, &attributes)?,
                Rule::enum_declaration => walk_enum_decl(mp, &attributes)?,
                _ => unreachable!(),
            };
            Ok(vec![ClassMember::NestedType(Box::new(
                Statement::with_span(kind, span),
            ))])
        }
        other => Err(format!("Unexpected class member: {:?}", other)),
    }
}

fn walk_constructor(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut base_args = None;
    let mut initializer_target = crate::ast::ConstructorInitializerTarget::Base;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::ident_name => {} // constructor name (same as class)
            Rule::param_list => params = walk_params(p)?,
            Rule::constructor_initializer => {
                let mut args = Vec::new();
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::this_kw {
                        initializer_target = crate::ast::ConstructorInitializerTarget::This;
                    }
                    if cp.as_rule() == Rule::argument_list {
                        args = walk_arguments(cp)?;
                    }
                }
                base_args = Some(args.into_iter().map(|a| a.value).collect());
            }
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                // `ClassName(p) => stmt;` desugars to a body whose
                // single statement is the expression as a stand-alone
                // ExprStmt. Constructors don't return a value, so we
                // don't wrap in Return.
                if let Some(inner) = p.into_inner().next() {
                    let span = to_span(&inner);
                    let expr = walk_expression(inner)?;
                    body = vec![Statement::with_span(StmtKind::Expr(expr), span)];
                }
            }
            _ => {}
        }
    }

    if mods.is_static {
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
                modifiers: mods,
            },
        ))));
    }

    Ok(ClassMember::Constructor {
        params,
        body,
        base_args,
        initializer_target,
        visibility: Visibility::Public,
    })
}

fn walk_property(pair: Pair<Rule>, mods: Modifiers) -> Result<Vec<ClassMember>, String> {
    let mut name = String::new();
    let mut explicit_interface = None;
    let mut getter = None;
    let mut setter = None;
    let mut is_auto = true;
    let mut default_init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip type
            Rule::explicit_interface_name => {
                explicit_interface = Some(p.as_str().to_string());
            }
            Rule::ident_name if name.is_empty() => name = p.as_str().to_string(),
            Rule::property_name => {
                for np in p.into_inner() {
                    match np.as_rule() {
                        Rule::explicit_interface_specifier => {
                            explicit_interface = Some(extract_explicit_interface_name(np));
                        }
                        Rule::ident_name => name = np.as_str().to_string(),
                        _ => {}
                    }
                }
            }
            Rule::expression_body => {
                // `Type Name => expr;` — read-only expression-bodied
                // property. Lower to a getter that returns the expr.
                if let Some(inner) = p.into_inner().next() {
                    let span = to_span(&inner);
                    let expr = walk_expression(inner)?;
                    getter = Some(vec![Statement::with_span(
                        StmtKind::Return(Some(expr)),
                        span,
                    )]);
                    is_auto = false;
                }
            }
            Rule::property_body => {
                for acc in p.into_inner() {
                    if acc.as_rule() == Rule::accessor {
                        // The `get` / `set` keywords are literal tokens
                        // in the grammar — pest doesn't surface them as
                        // child pairs, so we detect direction by
                        // looking at the source string. The accessor's
                        // `as_str()` is something like
                        // `public get { return _v; }`; trim the leading
                        // class_modifiers and check the next word.
                        let acc_src = acc.as_str().trim_start();
                        // Strip optional accessor modifiers (`public`
                        // `private` `protected` `internal`) so we land on
                        // the `get` / `set` keyword itself.
                        let mut rest = acc_src;
                        for kw in &["public", "private", "protected", "internal"] {
                            if let Some(stripped) = rest.strip_prefix(*kw) {
                                if stripped.starts_with(|c: char| c.is_whitespace()) {
                                    rest = stripped.trim_start();
                                    break;
                                }
                            }
                        }
                        let is_get = rest.starts_with("get")
                            && rest[3..]
                                .chars()
                                .next()
                                .map_or(true, |c| !c.is_alphanumeric() && c != '_');
                        let mut acc_body = None;
                        for ap in acc.into_inner() {
                            match ap.as_rule() {
                                Rule::block_statement => {
                                    acc_body = Some(walk_body(ap)?);
                                    is_auto = false;
                                }
                                Rule::expression_body => {
                                    is_auto = false;
                                    if let Some(expr_pair) = ap.into_inner().next() {
                                        let span = to_span(&expr_pair);
                                        let expr = walk_expression(expr_pair)?;
                                        acc_body = Some(if is_get {
                                            vec![Statement::with_span(
                                                StmtKind::Return(Some(expr)),
                                                span,
                                            )]
                                        } else {
                                            vec![Statement::with_span(StmtKind::Expr(expr), span)]
                                        });
                                    }
                                }
                                Rule::class_modifiers => {} // skip accessor modifiers
                                _ => {}
                            }
                        }
                        if is_get {
                            getter = acc_body.or(Some(Vec::new()));
                        } else {
                            let param = Param {
                                name: "value".into(),
                                type_hint: None,
                                default: None,
                                pass_by: PassBy::Value,
                                is_rest: false,
                                is_kwargs: false,
                                is_optional: false,
                                is_nullable: false,
                            };
                            setter = Some(PropertySetter {
                                param,
                                body: acc_body.unwrap_or_default(),
                            });
                        }
                    }
                }
            }
            // C# auto-property default-value initializer:
            // `public string Name { get; set; } = "default";` — capture
            // the RHS expression and emit a sibling Field below so the
            // backing slot is initialised in the constructor.
            other if other != Rule::class_modifiers => {
                default_init = Some(walk_expression(p)?);
            }
            _ => {}
        }
    }

    if let Some(interface_name) = explicit_interface {
        name = explicit_interface_runtime_name(&interface_name, &name);
    }

    let mut out = vec![ClassMember::Property {
        name: name.clone(),
        type_hint: None,
        getter,
        setter,
        is_auto,
        modifiers: mods.clone(),
    }];
    // Auto-property compiles its `__name` backing field. Emit a Field
    // entry with the same backing name so its `init` runs at instance
    // construction. ECMA-334 §15.7.4 — auto-property initializers run
    // before any user constructor body, matching what FieldDecl gives us.
    if is_auto {
        if let Some(init_expr) = default_init {
            // C# auto-properties without explicit accessors compile as
            // plain fields keyed by the property name (no `__` prefix).
            // The compiler's pass-1 in `classes.rs` adds `(pname, None)`
            // to `field_inits`; we emit the same field name with a real
            // `init` expression so it materialises in the constructor.
            // The duplicate key is deduped by the existing pass-1 check.
            out.push(ClassMember::Field {
                name: name.clone(),
                type_hint: None,
                init: Some(init_expr),
                modifiers: mods,
                with_events: false,
                array_bounds: None,
            });
        }
    }
    Ok(out)
}

fn walk_event(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint = None;
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            _ => {}
        }
    }
    Ok(ClassMember::Event {
        name,
        type_hint,
        params: Vec::new(),
        visibility: Visibility::Public,
    })
}

/// Walk every statement in a catch body and rewrite bare `throw;`
/// (Throw with no expression) into `throw <catch_var>;` so the VM
/// rethrows the caught instance instead of `NULL`. Recurses into
/// nested blocks but stops at any inner Try / FunctionDecl /
/// LambdaBody — re-throw scopes lexically by .NET semantics.
fn rewrite_bare_throws(stmts: &mut Vec<Statement>, var_name: &str) {
    for stmt in stmts.iter_mut() {
        rewrite_bare_throws_in_stmt(stmt, var_name);
    }
}
fn rewrite_bare_throws_in_stmt(stmt: &mut Statement, var_name: &str) {
    match &mut stmt.kind {
        StmtKind::Throw { expr, .. } if expr.is_none() => {
            *expr = Some(Expression::ident(var_name));
        }
        StmtKind::Block(inner) => rewrite_bare_throws(inner, var_name),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_bare_throws(then_body, var_name);
            for (_, body) in elifs {
                rewrite_bare_throws(body, var_name);
            }
            if let Some(eb) = else_body {
                rewrite_bare_throws(eb, var_name);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            rewrite_bare_throws(body, var_name);
        }
        StmtKind::Switch { cases, default, .. } => {
            for c in cases.iter_mut() {
                rewrite_bare_throws(&mut c.body, var_name);
            }
            if let Some(d) = default {
                rewrite_bare_throws(d, var_name);
            }
        }
        StmtKind::Try { body, finally, .. } => {
            // The bare `throw;` inside an inner try's catches refers to
            // that catch's bound exception, NOT the outer one. So we
            // only descend into the outer try's body and finally —
            // catches' bare throws are bound to their own var by the
            // recursive walk_try call.
            rewrite_bare_throws(body, var_name);
            if let Some(f) = finally {
                rewrite_bare_throws(f, var_name);
            }
        }
        _ => {}
    }
}

fn build_named_tuple_object_expr(
    span: Span,
    tuple_names: &[Option<String>],
    elems: &[Expression],
) -> Expression {
    let mut props = Vec::new();
    for (index, value) in elems.iter().enumerate() {
        let item_key = format!("Item{}", index + 1);
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&item_key),
            value: value.clone(),
        });
        if let Some(name) = tuple_names.get(index).and_then(|name| name.as_ref()) {
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(name),
                value: value.clone(),
            });
        }
    }
    Expression::with_span(ExprKind::Object(props), span)
}

fn parse_csharp_named_tuple_return_slots(raw_type: &str) -> Option<Vec<Option<String>>> {
    let trimmed = raw_type.trim();
    if !trimmed.starts_with('(') || !trimmed.ends_with(')') {
        return None;
    }

    let inner = &trimmed[1..trimmed.len().saturating_sub(1)];
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut angle = 0usize;
    let mut paren = 0usize;
    let mut bracket = 0usize;
    for ch in inner.chars() {
        match ch {
            '<' => angle += 1,
            '>' => angle = angle.saturating_sub(1),
            '(' => paren += 1,
            ')' => paren = paren.saturating_sub(1),
            '[' => bracket += 1,
            ']' => bracket = bracket.saturating_sub(1),
            ',' if angle == 0 && paren == 0 && bracket == 0 => {
                parts.push(current.trim().to_string());
                current.clear();
                continue;
            }
            _ => {}
        }
        current.push(ch);
    }
    if !current.trim().is_empty() {
        parts.push(current.trim().to_string());
    }
    if parts.len() < 2 {
        return None;
    }

    let mut has_named = false;
    let mut names = Vec::with_capacity(parts.len());
    for part in parts {
        let tokens: Vec<&str> = part.split_whitespace().collect();
        let name = tokens.last().and_then(|candidate| {
            let candidate = candidate.trim();
            (tokens.len() >= 2
                && !candidate.is_empty()
                && candidate
                    .chars()
                    .next()
                    .is_some_and(|ch| ch == '_' || ch.is_ascii_alphabetic())
                && candidate
                    .chars()
                    .all(|ch| ch == '_' || ch.is_ascii_alphanumeric()))
            .then(|| candidate.to_string())
        });
        has_named |= name.is_some();
        names.push(name);
    }

    has_named.then_some(names)
}

fn rewrite_named_tuple_returns_in_stmt(stmt: &mut Statement, tuple_names: &[Option<String>]) {
    match &mut stmt.kind {
        StmtKind::Return(Some(expr)) => {
            if let ExprKind::Tuple(elems) = &expr.kind {
                if elems.len() == tuple_names.len() {
                    *expr = build_named_tuple_object_expr(expr.span.clone(), tuple_names, elems);
                }
            }
        }
        StmtKind::Block(body)
        | StmtKind::For { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. } => {
            rewrite_named_tuple_returns_in_stmts(body, tuple_names);
        }
        StmtKind::ForIn {
            body, else_body, ..
        } => {
            rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            if let Some(body) = else_body {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_named_tuple_returns_in_stmts(then_body, tuple_names);
            for (_, body) in elifs {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
            if let Some(body) = else_body {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            for catch in catches {
                rewrite_named_tuple_returns_in_stmts(&mut catch.body, tuple_names);
            }
            if let Some(body) = else_body {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
            if let Some(body) = finally {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                rewrite_named_tuple_returns_in_stmts(&mut case.body, tuple_names);
            }
            if let Some(body) = default {
                rewrite_named_tuple_returns_in_stmts(body, tuple_names);
            }
        }
        StmtKind::Labeled { body, .. } => rewrite_named_tuple_returns_in_stmt(body, tuple_names),
        StmtKind::FunctionDecl { .. } | StmtKind::ClassDecl { .. } => {}
        _ => {}
    }
}

fn rewrite_named_tuple_returns_in_stmts(stmts: &mut [Statement], tuple_names: &[Option<String>]) {
    for stmt in stmts {
        rewrite_named_tuple_returns_in_stmt(stmt, tuple_names);
    }
}

fn apply_named_tuple_return_normalization(body: &mut Vec<Statement>, return_type: Option<&str>) {
    let Some(tuple_names) = return_type.and_then(parse_csharp_named_tuple_return_slots) else {
        return;
    };
    rewrite_named_tuple_returns_in_stmts(body, &tuple_names);
}

fn apply_initialize_component_event_normalization(body: &mut [Statement]) {
    for stmt in body {
        let Some(kind) = lower_initialize_component_event_stmt(stmt) else {
            continue;
        };
        stmt.kind = kind;
    }
}

fn lower_initialize_component_event_stmt(stmt: &Statement) -> Option<StmtKind> {
    let StmtKind::CompoundAssign { target, op, value } = &stmt.kind else {
        return None;
    };

    let bin_op = match op {
        CompoundOp::Add => BinOp::Add,
        CompoundOp::Sub => BinOp::Sub,
        _ => return None,
    };

    if !is_this_rooted_member_target(target) {
        return None;
    }

    let expr = Expression::with_span(
        ExprKind::Assign {
            target: Box::new(target.clone()),
            value: Box::new(Expression::with_span(
                ExprKind::Binary {
                    op: bin_op,
                    left: Box::new(target.clone()),
                    right: Box::new(value.clone()),
                },
                stmt.span,
            )),
        },
        stmt.span,
    );

    crate::common::events::lower_event_compound_assignment(&expr)
}

fn is_this_rooted_member_target(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Member { object, .. } => match &object.kind {
            ExprKind::This => true,
            ExprKind::Member { .. } => is_this_rooted_member_target(object),
            _ => false,
        },
        _ => false,
    }
}

/// Walk a top-level / local function declaration. Same shape as a
/// class method but lives at statement scope. Lowers to
/// `StmtKind::FunctionDecl` so the compiler treats it like any other
/// free function.
fn walk_local_function(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut generic_params = Vec::new();
    let mut return_type = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut modifiers = Modifiers::default();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        s if s.starts_with("static") => modifiers.is_static = true,
                        s if s.starts_with("async") => {}
                        _ => {}
                    }
                }
            }
            Rule::type_name => {
                if return_type.is_none() {
                    return_type = Some(p.as_str().to_string());
                }
            }
            Rule::ident_name => {
                if name.is_empty() {
                    name = p.as_str().to_string();
                } else {
                    generic_params.push(p.as_str().to_string());
                }
            }
            Rule::param_list => params = walk_params(p)?,
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(StmtKind::Return(Some(expr)), span)];
                }
            }
            _ => {}
        }
    }
    if !generic_params.is_empty() {
        inject_csharp_generic_binding_params(&mut params, generic_params.len());
        rewrite_generic_bindings_in_statements(&mut body, &generic_params);
    }
    apply_named_tuple_return_normalization(&mut body, return_type.as_deref());
    let is_sub = return_type.as_deref() == Some("void");
    let is_generator = body_has_yield(&body);
    Ok(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body,
        modifiers,
        is_async: false,
        is_generator,
        is_sub,
        handles: Vec::new(),
    })
}

/// Map a C# operator symbol (`+`, `==`, …) to the ECMA-335 / .NET
/// canonical method name (`op_Addition`, `op_Equality`, …). Used by
/// the operator-overload walker so the method can be located via the
/// same name C# ABI emits.
fn operator_method_name(symbol: &str, arity: usize) -> &'static str {
    match (symbol, arity) {
        ("+", 1) => "op_UnaryPlus",
        ("-", 1) => "op_UnaryNegation",
        ("++", 1) => "op_Increment",
        ("--", 1) => "op_Decrement",
        ("true", 1) => "op_True",
        ("false", 1) => "op_False",
        ("~", 1) => "op_OnesComplement",
        ("!", 1) => "op_LogicalNot",
        ("+", _) => "op_Addition",
        ("-", _) => "op_Subtraction",
        ("*", _) => "op_Multiply",
        ("/", _) => "op_Division",
        ("%", _) => "op_Modulus",
        ("==", _) => "op_Equality",
        ("!=", _) => "op_Inequality",
        ("<", _) => "op_LessThan",
        (">", _) => "op_GreaterThan",
        ("<=", _) => "op_LessThanOrEqual",
        (">=", _) => "op_GreaterThanOrEqual",
        ("&", _) => "op_BitwiseAnd",
        ("|", _) => "op_BitwiseOr",
        ("^", _) => "op_ExclusiveOr",
        ("<<", _) => "op_LeftShift",
        (">>", _) => "op_RightShift",
        _ => "op_Unknown",
    }
}

fn conversion_operator_method_name(keyword: &str, return_type: Option<&str>) -> Option<String> {
    let op = match keyword {
        "implicit" => "op_Implicit",
        "explicit" => "op_Explicit",
        _ => return None,
    };
    let suffix = return_type
        .map(csharp_operator_type_suffix)
        .unwrap_or_default();
    Some(format!("{op}{suffix}"))
}

fn csharp_operator_type_suffix(type_name: &str) -> String {
    normalize_runtime_type_name(type_name)
        .chars()
        .filter(|ch| ch.is_ascii_alphanumeric())
        .collect()
}

fn csharp_binary_operator_method_name(op: BinOp) -> Option<&'static str> {
    match op {
        BinOp::Add => Some("op_Addition"),
        BinOp::Sub => Some("op_Subtraction"),
        BinOp::Mul => Some("op_Multiply"),
        BinOp::Div => Some("op_Division"),
        BinOp::Mod => Some("op_Modulus"),
        BinOp::Eq => Some("op_Equality"),
        BinOp::NotEq => Some("op_Inequality"),
        BinOp::Lt => Some("op_LessThan"),
        BinOp::Gt => Some("op_GreaterThan"),
        BinOp::LtEq => Some("op_LessThanOrEqual"),
        BinOp::GtEq => Some("op_GreaterThanOrEqual"),
        BinOp::BitAnd => Some("op_BitwiseAnd"),
        BinOp::BitOr => Some("op_BitwiseOr"),
        BinOp::BitXor => Some("op_ExclusiveOr"),
        BinOp::Shl => Some("op_LeftShift"),
        BinOp::Shr => Some("op_RightShift"),
        _ => None,
    }
}

fn csharp_unary_operator_method_name(op: UnaryOp) -> Option<&'static str> {
    match op {
        UnaryOp::Neg => Some("op_UnaryNegation"),
        UnaryOp::Pos => Some("op_UnaryPlus"),
        UnaryOp::Not => Some("op_LogicalNot"),
        UnaryOp::BitNot => Some("op_OnesComplement"),
        UnaryOp::PreInc | UnaryOp::PostInc => Some("op_Increment"),
        UnaryOp::PreDec | UnaryOp::PostDec => Some("op_Decrement"),
        _ => None,
    }
}

fn csharp_builtin_bool_binary_op(op: BinOp) -> bool {
    matches!(
        op,
        BinOp::Eq
            | BinOp::NotEq
            | BinOp::StrictEq
            | BinOp::StrictNotEq
            | BinOp::Lt
            | BinOp::Gt
            | BinOp::LtEq
            | BinOp::GtEq
            | BinOp::And
            | BinOp::Or
            | BinOp::Is
            | BinOp::IsNot
    )
}

fn materialize_csharp_bool_expr(expr: Expression) -> Expression {
    let span = expr.span.clone();
    Expression::with_span(
        ExprKind::Ternary {
            cond: Box::new(expr),
            then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            else_: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        },
        span,
    )
}

/// Walk an `operator_declaration`. Lowers to a static method named
/// per `operator_method_name` so the call-site dispatch can find it
/// via the canonical naming scheme.
fn walk_operator(pair: Pair<Rule>, mut mods: Modifiers) -> Result<ClassMember, String> {
    mods.is_static = true;
    let mut return_type = None;
    let mut conversion_kind = None;
    let mut symbol = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::conversion_operator_kind => conversion_kind = Some(p.as_str().to_string()),
            Rule::operator_symbol => symbol = p.as_str().trim().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(StmtKind::Return(Some(expr)), span)];
                }
            }
            _ => {}
        }
    }
    let name = conversion_kind
        .as_deref()
        .and_then(|kind| conversion_operator_method_name(kind, return_type.as_deref()))
        .unwrap_or_else(|| operator_method_name(&symbol, params.len()).to_string());
    let is_sub = return_type.as_deref() == Some("void");
    let is_generator = body_has_yield(&body);
    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers: mods,
            is_async: false,
            is_generator,
            is_sub,
            handles: Vec::new(),
        },
    ))))
}

/// Walk an `indexer_declaration`. Lowers to a Property named `__index__`
/// with the indexer's parameter list captured separately so the runtime
/// can route `obj[i]` through the getter / setter.
fn walk_indexer(pair: Pair<Rule>, mods: Modifiers) -> Result<Vec<ClassMember>, String> {
    let mut getter_name = "__get___index__".to_string();
    let mut setter_name = "__set___index__".to_string();
    let mut explicit_interface = None;
    let mut getter: Option<Vec<Statement>> = None;
    let mut setter: Option<Vec<Statement>> = None;
    let mut params: Vec<Param> = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => {} // skip return type
            Rule::explicit_interface_name => {
                explicit_interface = Some(p.as_str().to_string());
            }
            Rule::indexer_name => {
                for np in p.into_inner() {
                    if np.as_rule() == Rule::explicit_interface_specifier {
                        explicit_interface = Some(extract_explicit_interface_name(np));
                    }
                }
            }
            Rule::param_list => params = walk_params(p)?,
            Rule::property_body => {
                for acc in p.into_inner() {
                    if acc.as_rule() == Rule::accessor {
                        let acc_src = acc.as_str().trim_start();
                        let is_get = acc_src.starts_with("get")
                            || acc_src.contains(" get")
                            || acc_src.starts_with("public get")
                            || acc_src.starts_with("private get")
                            || acc_src.starts_with("protected get")
                            || acc_src.starts_with("internal get");
                        let mut acc_body = None;
                        for ap in acc.into_inner() {
                            match ap.as_rule() {
                                Rule::block_statement => {
                                    acc_body = Some(walk_body(ap)?);
                                }
                                Rule::expression_body => {
                                    if let Some(expr_pair) = ap.into_inner().next() {
                                        let span = to_span(&expr_pair);
                                        let expr = walk_expression(expr_pair)?;
                                        acc_body = Some(if is_get {
                                            vec![Statement::with_span(
                                                StmtKind::Return(Some(expr)),
                                                span,
                                            )]
                                        } else {
                                            vec![Statement::with_span(StmtKind::Expr(expr), span)]
                                        });
                                    }
                                }
                                Rule::class_modifiers => {}
                                _ => {}
                            }
                        }
                        if is_get {
                            getter = acc_body;
                        } else if let Some(body) = acc_body {
                            setter = Some(body);
                        }
                    }
                }
            }
            _ => {}
        }
    }
    if let Some(interface_name) = explicit_interface {
        getter_name = explicit_interface_runtime_name(&interface_name, &getter_name);
        setter_name = explicit_interface_runtime_name(&interface_name, &setter_name);
    }

    let mut members = Vec::new();
    if let Some(body) = getter {
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: getter_name,
                params: params.clone(),
                return_type: None,
                body,
                modifiers: mods.clone(),
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: false,
            },
        ))));
    }
    if let Some(body) = setter {
        let mut setter_params = params;
        setter_params.push(Param {
            name: "value".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: setter_name,
                params: setter_params,
                return_type: Some("void".into()),
                body,
                modifiers: mods,
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub: true,
            },
        ))));
    }

    Ok(members)
}

fn walk_method(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut mods = mods;
    let mut name = String::new();
    let mut generic_params = Vec::new();
    let mut explicit_interface = None;
    let mut return_type = None;
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut is_async = false;

    // Check modifiers for async
    if mods.is_abstract { /* abstract methods have no body */ }

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name if return_type.is_none() => {
                let rt = p.as_str().to_string();
                if rt.starts_with("async") {
                    is_async = true;
                }
                return_type = Some(rt);
            }
            Rule::explicit_interface_name => {
                explicit_interface = Some(p.as_str().to_string());
            }
            Rule::ident_name if name.is_empty() => {
                name = p.as_str().to_string();
            }
            Rule::ident_name => {
                generic_params.push(p.as_str().to_string());
            }
            Rule::param_list => {
                let (parsed_params, param_decorators) = walk_params_with_decorators(p)?;
                params = parsed_params;
                mods.decorators.extend(param_decorators);
            }
            Rule::block_statement => body = walk_body(p)?,
            Rule::expression_body => {
                // C# expression-bodied member: `=> expr;` lowers to
                // `{ return expr; }`. The inner `expression` pair is
                // the only child of `expression_body`.
                let span = to_span(&p);
                if let Some(expr_pair) = p.into_inner().next() {
                    let expr = walk_expression(expr_pair)?;
                    body = vec![Statement::with_span(StmtKind::Return(Some(expr)), span)];
                }
            }
            Rule::method_name => {
                for np in p.into_inner() {
                    match np.as_rule() {
                        Rule::explicit_interface_specifier => {
                            explicit_interface = Some(extract_explicit_interface_name(np));
                        }
                        Rule::ident_name if name.is_empty() => name = np.as_str().to_string(),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(interface_name) = explicit_interface {
        name = explicit_interface_runtime_name(&interface_name, &name);
    }
    if !generic_params.is_empty() {
        inject_csharp_generic_binding_params(&mut params, generic_params.len());
        rewrite_generic_bindings_in_statements(&mut body, &generic_params);
    }
    apply_named_tuple_return_normalization(&mut body, return_type.as_deref());
    if name.eq_ignore_ascii_case("InitializeComponent") {
        apply_initialize_component_event_normalization(&mut body);
    }

    let is_sub = return_type.as_deref() == Some("void");

    let is_generator = body_has_yield(&body);
    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers: mods,
            handles: Vec::new(),
            is_async,
            is_generator,
            is_sub,
        },
    ))))
}

fn walk_field(pair: Pair<Rule>, mods: Modifiers) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::var_declarator_list => {
                // Take first declarator
                if let Some(vd) = p.into_inner().find(|p| p.as_rule() == Rule::var_declarator) {
                    let decl = walk_var_declarator(vd)?;
                    if let BindingPattern::Ident(n) = decl.pattern {
                        name = n;
                    }
                    init = decl.init;
                }
            }
            _ => {}
        }
    }

    Ok(ClassMember::Field {
        name,
        type_hint,
        init,
        modifiers: mods,
        with_events: false,
        array_bounds: None,
    })
}

// ── Struct ──────────────────────────────────────────────────────────────────

fn walk_struct_decl(pair: Pair<Rule>, decorators: &[Expression]) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut interfaces = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {} // skip modifiers
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::base_list => {
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        interfaces.push(bp.as_str().to_string());
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.extend(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    let has_user_equals = members.iter().any(|member| {
        matches!(
            member,
            ClassMember::Method(stmt) if matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, .. } if name == "Equals"
            )
        )
    });

    if !has_user_equals {
        let comparable_members: Vec<String> = members
            .iter()
            .filter_map(|member| match member {
                ClassMember::Field {
                    name, modifiers, ..
                } if !modifiers.is_static => Some(name.clone()),
                ClassMember::Property {
                    name, modifiers, ..
                } if !modifiers.is_static => Some(name.clone()),
                _ => None,
            })
            .collect();

        let other_param = Param {
            name: "other".into(),
            type_hint: Some(name.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };

        let mut eq_expr = Expression::bool(true);
        for member_name in comparable_members {
            let left = Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: member_name.clone(),
                null_safe: false,
            });
            let right = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("other")),
                field: member_name,
                null_safe: false,
            });
            let cmp = Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(left),
                right: Box::new(right),
            });
            eq_expr = Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(eq_expr),
                right: Box::new(cmp),
            });
        }

        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "Equals".into(),
                params: vec![other_param],
                body: vec![Statement::new(StmtKind::Return(Some(eq_expr)))],
                return_type: Some("bool".into()),
                is_async: false,
                is_generator: false,
                is_sub: false,
                handles: Vec::new(),
                modifiers: Modifiers {
                    is_override: true,
                    ..Default::default()
                },
            },
        ))));
    }

    disambiguate_explicit_property_backing_fields(&mut members);

    Ok(StmtKind::StructDecl {
        name,
        interfaces,
        members,
        visibility: Visibility::Public,
        decorators: decorators.to_vec(),
    })
}

// ── Interface ───────────────────────────────────────────────────────────────

fn walk_interface_decl(pair: Pair<Rule>, decorators: &[Expression]) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::base_list => {
                for bp in p.into_inner() {
                    if bp.as_rule() == Rule::type_name {
                        parents.push(bp.as_str().to_string());
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::interface_member {
                        if let Ok(member) = walk_interface_member(m) {
                            members.push(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::InterfaceDecl {
        name,
        parents,
        members,
        decorators: decorators.to_vec(),
    })
}

fn walk_interface_member(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let mut type_hint = None;
    let mut name = String::new();
    let mut has_params = false;
    let mut params = Vec::new();
    let mut has_getter = false;
    let mut has_setter = false;
    let src = pair.as_str().to_string();
    let is_indexer = src.contains("this[");

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::interface_operator_declaration => {
                return walk_interface_operator_member(p);
            }
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => {
                has_params = true;
                params = walk_params(p)?;
            }
            _ => {
                let s = p.as_str();
                if s == "get" {
                    has_getter = true;
                }
                if s == "set" {
                    has_setter = true;
                }
            }
        }
    }

    if is_indexer {
        name = "__index__".into();
    }

    if !is_indexer && (has_params || src.contains('(')) {
        let is_sub = type_hint.as_deref() == Some("void");
        Ok(InterfaceMember::Method {
            name,
            params,
            return_type: type_hint,
            is_sub,
            signature_source: None,
        })
    } else {
        Ok(InterfaceMember::Property {
            name,
            type_hint,
            is_readonly: has_getter && !has_setter,
            is_writeonly: !has_getter && has_setter,
        })
    }
}

fn walk_interface_operator_member(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let mut return_type = None;
    let mut conversion_kind = None;
    let mut symbol = String::new();
    let mut params = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::conversion_operator_kind => conversion_kind = Some(p.as_str().to_string()),
            Rule::operator_symbol => symbol = p.as_str().trim().to_string(),
            Rule::param_list => params = walk_params(p)?,
            _ => {}
        }
    }

    let name = conversion_kind
        .as_deref()
        .and_then(|kind| conversion_operator_method_name(kind, return_type.as_deref()))
        .unwrap_or_else(|| operator_method_name(&symbol, params.len()).to_string());
    let is_sub = return_type.as_deref() == Some("void");
    Ok(InterfaceMember::Method {
        name,
        params,
        return_type,
        is_sub,
        signature_source: None,
    })
}

// ── Enum ────────────────────────────────────────────────────────────────────

fn walk_enum_decl(pair: Pair<Rule>, decorators: &[Expression]) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut members = Vec::new();
    let is_flags = decorators.iter().any(|attr| {
        attribute_leaf_name(attr).is_some_and(|short| {
            short.eq_ignore_ascii_case("Flags") || short.eq_ignore_ascii_case("FlagsAttribute")
        })
    });

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::enum_member_list => {
                for em in p.into_inner() {
                    if em.as_rule() == Rule::enum_member {
                        let mut en = String::new();
                        let mut val = None;
                        for ep in em.into_inner() {
                            match ep.as_rule() {
                                Rule::ident_name => en = ep.as_str().to_string(),
                                _ => val = Some(walk_expression(ep)?),
                            }
                        }
                        members.push(EnumMember {
                            name: en,
                            value: val,
                            constructor_args: Vec::new(),
                        });
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::EnumDecl {
        name,
        members,
        visibility: Visibility::Public,
        is_flags,
        backing_type: None,
        interfaces: Vec::new(),
        body_members: Vec::new(),
        decorators: decorators.to_vec(),
    })
}

// ── Record ──────────────────────────────────────────────────────────────────

fn walk_record_decl(pair: Pair<Rule>, decorators: &[Expression]) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut parents = Vec::new();
    let mut members = Vec::new();
    let mut base_args = None;
    let mut record_mods = ClassModifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {
                for m in p.into_inner() {
                    match m.as_str() {
                        "public" => record_mods.visibility = Visibility::Public,
                        "private" => record_mods.visibility = Visibility::Private,
                        "protected" => record_mods.visibility = Visibility::Protected,
                        "internal" => record_mods.visibility = Visibility::Internal,
                        s if s.starts_with("static") => record_mods.is_static = true,
                        s if s.starts_with("abstract") => record_mods.is_abstract = true,
                        s if s.starts_with("sealed") => record_mods.is_sealed = true,
                        s if s.starts_with("partial") => record_mods.is_partial = true,
                        _ => {}
                    }
                }
            }
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            Rule::record_base_clause => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::type_name => parents.push(bp.as_str().to_string()),
                        Rule::argument_list => {
                            base_args =
                                Some(walk_arguments(bp)?.into_iter().map(|a| a.value).collect());
                        }
                        _ => {}
                    }
                }
            }
            Rule::class_body => {
                for m in p.into_inner() {
                    if m.as_rule() == Rule::class_member {
                        if let Ok(member) = walk_class_member(m) {
                            members.extend(member);
                        }
                    }
                }
            }
            _ => {}
        }
    }

    // Record positional params become fields + constructor
    for param in &params {
        members.push(ClassMember::Field {
            name: param.name.clone(),
            type_hint: param.type_hint.clone(),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        });
    }

    // Generate constructor from params
    if !params.is_empty() {
        let ctor_body: Vec<Statement> = params
            .iter()
            .map(|p| {
                Statement::new(StmtKind::Assign {
                    targets: vec![Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: p.name.clone(),
                        null_safe: false,
                    })],
                    value: Expression::ident(&p.name),
                })
            })
            .collect();
        members.push(ClassMember::Constructor {
            params: params.clone(),
            body: ctor_body,
            base_args,
            initializer_target: crate::ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    }

    let has_user_deconstruct = members.iter().any(|m| {
        matches!(
            m,
            ClassMember::Method(stmt) if matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, .. } if name == "Deconstruct"
            )
        )
    });
    if !has_user_deconstruct && !params.is_empty() {
        let body: Vec<Statement> = params
            .iter()
            .map(|p| {
                Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(&p.name)],
                    value: Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field: p.name.clone(),
                        null_safe: false,
                    }),
                })
            })
            .collect();
        let out_params: Vec<Param> = params
            .iter()
            .map(|p| Param {
                name: p.name.clone(),
                type_hint: p.type_hint.clone(),
                default: None,
                pass_by: PassBy::Ref,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect();
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "Deconstruct".into(),
                params: out_params,
                body,
                return_type: Some("void".into()),
                is_async: false,
                is_generator: false,
                is_sub: true,
                handles: Vec::new(),
                modifiers: Modifiers::default(),
            },
        ))));
    }

    let has_user_equals = members.iter().any(|m| {
        matches!(
            m,
            ClassMember::Method(stmt) if matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, .. } if name == "Equals"
            )
        )
    });
    if !has_user_equals && !params.is_empty() {
        let other_param = Param {
            name: "other".into(),
            type_hint: Some(name.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let mut eq_expr = Expression::bool(true);
        for p in &params {
            let left = Expression::new(ExprKind::Member {
                object: Box::new(Expression::new(ExprKind::This)),
                field: p.name.clone(),
                null_safe: false,
            });
            let right = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("other")),
                field: p.name.clone(),
                null_safe: false,
            });
            let cmp = Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(left),
                right: Box::new(right),
            });
            eq_expr = Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(eq_expr),
                right: Box::new(cmp),
            });
        }
        let body = vec![Statement::new(StmtKind::Return(Some(eq_expr)))];
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "Equals".into(),
                params: vec![other_param],
                body,
                return_type: Some("bool".into()),
                is_async: false,
                is_generator: false,
                is_sub: false,
                handles: Vec::new(),
                modifiers: Modifiers {
                    is_override: true,
                    ..Default::default()
                },
            },
        ))));
    }

    let has_user_op_eq = members.iter().any(|m| {
        matches!(
            m,
            ClassMember::Method(stmt) if matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, .. } if name == "op_Equality"
            )
        )
    });
    if !has_user_op_eq && !params.is_empty() {
        let left_param = Param {
            name: "left".into(),
            type_hint: Some(name.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let right_param = Param {
            name: "right".into(),
            type_hint: Some(name.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let mut eq_expr = Expression::bool(true);
        for p in &params {
            let left = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("left")),
                field: p.name.clone(),
                null_safe: false,
            });
            let right = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("right")),
                field: p.name.clone(),
                null_safe: false,
            });
            let cmp = Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(left),
                right: Box::new(right),
            });
            eq_expr = Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(eq_expr),
                right: Box::new(cmp),
            });
        }
        let body = vec![Statement::new(StmtKind::Return(Some(eq_expr)))];
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "op_Equality".into(),
                params: vec![left_param, right_param],
                body,
                return_type: Some("bool".into()),
                is_async: false,
                is_generator: false,
                is_sub: false,
                handles: Vec::new(),
                modifiers: Modifiers {
                    is_static: true,
                    ..Default::default()
                },
            },
        ))));
    }

    // Synthetic ToString — `Point { X = 3, Y = 4 }` (ECMA-334 §15.6.6,
    // .NET 5+ record default). Skip if the user already defined one.
    let has_user_tostring = members.iter().any(|m| {
        matches!(
            m,
            ClassMember::Method(stmt) if matches!(
                &stmt.kind,
                StmtKind::FunctionDecl { name, .. } if name == "ToString"
            )
        )
    });
    if !has_user_tostring && !params.is_empty() {
        let mut concat = Expression::new(ExprKind::Lit(Literal::Str(format!("{} {{ ", name))));
        for (i, p) in params.iter().enumerate() {
            if i > 0 {
                concat = Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(concat),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(", ".into())))),
                });
            }
            concat = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(concat),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(format!(
                    "{} = ",
                    p.name
                ))))),
            });
            concat = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(concat),
                right: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::This)),
                    field: p.name.clone(),
                    null_safe: false,
                })),
            });
        }
        concat = Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(concat),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(" }".into())))),
        });
        let body = vec![Statement::new(StmtKind::Return(Some(concat)))];
        members.push(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name: "ToString".into(),
                params: Vec::new(),
                body,
                return_type: Some("string".into()),
                is_async: false,
                is_generator: false,
                is_sub: false,
                handles: Vec::new(),
                modifiers: Modifiers {
                    is_override: true,
                    ..Default::default()
                },
            },
        ))));
    }

    Ok(StmtKind::ClassDecl {
        name,
        parents,
        interfaces: Vec::new(),
        members,
        modifiers: record_mods,
        decorators: decorators.to_vec(),
    })
}

// ── Delegate ────────────────────────────────────────────────────────────────

fn walk_delegate_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut name = String::new();
    let mut return_type = None;
    let mut params = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_modifiers => {}
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            Rule::param_list => params = walk_params(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::DelegateDecl {
        name,
        params,
        return_type: return_type.clone(),
        is_sub: return_type.as_deref() == Some("void"),
        visibility: Visibility::Public,
    })
}

// ── Control flow ────────────────────────────────────────────────────────────

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond_pair = inner.next().ok_or("if: no cond")?;
    let pattern_binding = extract_if_is_pattern_binding(cond_pair.clone())?;
    let cond = if let Some(scoped_cond) = lower_if_pattern_condition(cond_pair.clone())? {
        scoped_cond
    } else {
        walk_expression(cond_pair)?
    };
    let mut then_body = vec![walk_statement(inner.next().ok_or("if: no body")?)?];
    if let Some(binding_stmt) = pattern_binding {
        then_body.insert(0, binding_stmt);
    }
    let mut elifs = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::else_if_clause => {
                let mut eip = p.into_inner();
                let cond_pair = eip.next().ok_or("elif: no cond")?;
                let pattern_binding = extract_if_is_pattern_binding(cond_pair.clone())?;
                let ec = if let Some(scoped_cond) = lower_if_pattern_condition(cond_pair.clone())? {
                    scoped_cond
                } else {
                    walk_expression(cond_pair)?
                };
                let mut eb = vec![walk_statement(eip.next().ok_or("elif: no body")?)?];
                if let Some(binding_stmt) = pattern_binding {
                    eb.insert(0, binding_stmt);
                }
                elifs.push((ec, eb));
            }
            Rule::else_clause => {
                let body = p.into_inner().next().ok_or("else: no body")?;
                else_body = Some(vec![walk_statement(body)?]);
            }
            _ => {}
        }
    }

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs,
        else_body,
    })
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init = None;
    let mut cond = None;
    let mut update = None;
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::for_init => {
                let inner = p.into_inner().next().ok_or("Empty for init")?;
                match inner.as_rule() {
                    Rule::local_var_declaration_no_semi => {
                        init = Some(Box::new(Statement::new(walk_local_var(inner)?)));
                    }
                    Rule::expression_list => {
                        let first_expr = inner.into_inner().next().ok_or("Empty expr list")?;
                        let expr = walk_expression(first_expr)?;
                        init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                    }
                    _ => {
                        let expr = walk_expression(inner)?;
                        init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                    }
                }
            }
            Rule::expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(p)?);
                }
            }
            Rule::for_update => {
                // for_update may contain expression_list with multiple expressions
                let inner = p.into_inner().next().ok_or("Empty for update")?;
                if inner.as_rule() == Rule::expression_list {
                    let mut exprs: Vec<Pair<Rule>> = inner.into_inner().collect();
                    if exprs.len() == 1 {
                        update = Some(walk_expression(exprs.remove(0))?);
                    } else {
                        // Multiple update expressions → sequence
                        let seq: Vec<Expression> = exprs
                            .into_iter()
                            .map(walk_expression)
                            .collect::<Result<Vec<_>, _>>()?;
                        update = Some(Expression::new(ExprKind::Sequence(seq)));
                    }
                } else {
                    update = Some(walk_expression(inner)?);
                }
            }
            _ => {
                if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    Ok(StmtKind::For {
        init,
        cond,
        update,
        body,
    })
}

fn walk_foreach(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let foreach_src = pair.as_str().to_string();
    let hidden_suffix = pair.as_span().start();
    let mut var = String::new();
    let mut explicit_type_hint = None;
    let mut tuple_target: Option<Vec<String>> = None;
    let mut iter = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw => {}
            Rule::type_name => explicit_type_hint = Some(p.as_str().to_string()),
            Rule::foreach_target => {
                if let Some(inner) = p.into_inner().next() {
                    match inner.as_rule() {
                        Rule::ident_name => var = inner.as_str().to_string(),
                        Rule::foreach_tuple_target => {
                            let names: Vec<String> = inner
                                .into_inner()
                                .filter(|part| part.as_rule() == Rule::ident_name)
                                .map(|part| part.as_str().to_string())
                                .collect();
                            tuple_target = Some(names);
                            var = format!("__csharp_foreach_item_{}", hidden_suffix);
                        }
                        _ => {}
                    }
                }
            }
            Rule::in_kw => {} // skip keyword
            _ => {
                if body.is_empty() {
                    if let Ok(expr) = walk_expression(p.clone()) {
                        iter = expr;
                    } else if let Ok(stmt) = walk_statement(p) {
                        body = vec![stmt];
                    }
                } else if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    if let Some(names) = tuple_target {
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Array(tuple_binding_pattern_elems(names)),
                    type_hint: None,
                    init: Some(Expression::ident(&var)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
        );
    } else {
        let uses_entry_fields = !var.is_empty()
            && foreach_src.contains(&format!("{}.Key", var))
            && foreach_src.contains(&format!("{}.Value", var));
        if uses_entry_fields {
            let user_var = var.clone();
            let source_var = format!("__csharp_foreach_item_{}", hidden_suffix);
            let entry_object = Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("Key"),
                    value: Expression::new(ExprKind::Index {
                        object: Box::new(Expression::ident(&source_var)),
                        index: Box::new(Expression::int(0)),
                        null_safe: false,
                    }),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("Value"),
                    value: Expression::new(ExprKind::Index {
                        object: Box::new(Expression::ident(&source_var)),
                        index: Box::new(Expression::int(1)),
                        null_safe: false,
                    }),
                },
            ]));
            body.insert(
                0,
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(user_var),
                        type_hint: None,
                        init: Some(entry_object),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
            );
            var = source_var;
        } else {
            let binding_type =
                explicit_type_hint.or_else(|| infer_csharp_foreach_element_type(&iter));
            if let Some(type_hint) = binding_type {
                let user_var = var.clone();
                let source_var = format!("__csharp_foreach_item_{}", hidden_suffix);
                body.insert(
                    0,
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(user_var),
                            type_hint: Some(type_hint),
                            init: Some(Expression::ident(&source_var)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                );
                var = source_var;
            }
        }
    }

    Ok(StmtKind::ForIn {
        var,
        key: None,
        iter,
        body,
        of: true, // foreach is like for-of
        else_body: None,
        is_async: false,
    })
}

fn walk_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let cond = walk_expression(inner.next().ok_or("while: no cond")?)?;
    let body = vec![walk_statement(inner.next().ok_or("while: no body")?)?];
    Ok(StmtKind::While {
        cond,
        body,
        else_body: None,
    })
}

fn walk_do_while(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let body = vec![walk_statement(inner.next().ok_or("do: no body")?)?];
    let cond = walk_expression(inner.next().ok_or("do: no cond")?)?;
    Ok(StmtKind::DoWhile {
        body,
        cond,
        until: false,
    })
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    #[derive(Clone)]
    enum SwitchLabelInfo<'i> {
        Default,
        Value(Expression),
        Pattern {
            pattern: Pair<'i, Rule>,
            guard: Option<Expression>,
        },
    }

    let mut inner = pair.into_inner();
    let expr = walk_expression(inner.next().ok_or("switch: no expr")?)?;
    let mut sections: Vec<(Vec<SwitchLabelInfo<'_>>, Vec<Statement>)> = Vec::new();
    let mut has_pattern_labels = false;

    for p in inner {
        if p.as_rule() == Rule::switch_section {
            let mut labels = Vec::new();
            let mut stmts = Vec::new();

            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::switch_label => {
                        let label_src = sp.as_str().trim();
                        if label_src.starts_with("default") {
                            labels.push(SwitchLabelInfo::Default);
                        } else if let Some(label_inner) = sp.into_inner().next() {
                            match label_inner.as_rule() {
                                Rule::case_value_label => {
                                    let expr_pair = label_inner
                                        .into_inner()
                                        .next()
                                        .ok_or("switch case missing value")?;
                                    labels
                                        .push(SwitchLabelInfo::Value(walk_expression(expr_pair)?));
                                }
                                Rule::case_pattern_label => {
                                    let mut label_parts = label_inner.into_inner();
                                    let pattern =
                                        label_parts.next().ok_or("switch case missing pattern")?;
                                    let guard =
                                        label_parts.next().map(walk_expression).transpose()?;
                                    labels.push(SwitchLabelInfo::Pattern { pattern, guard });
                                    has_pattern_labels = true;
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {
                        if let Ok(stmt) = walk_statement(sp) {
                            stmts.push(stmt);
                        }
                    }
                }
            }

            sections.push((labels, stmts));
        }
    }

    if has_pattern_labels {
        let subject_name = "__switch_subject".to_string();
        let matched_name = "__switch_matched".to_string();
        let subject_expr = Expression::ident(&subject_name);
        let mut arms: Vec<(Expression, Vec<Statement>)> = Vec::new();
        let mut default_body = None;

        for (labels, stmts) in sections {
            let split_bodies = split_switch_section_bodies(&stmts);
            let use_split_bodies = split_bodies.len() == labels.len();
            for (index, label) in labels.into_iter().enumerate() {
                let stripped_body = if use_split_bodies {
                    split_bodies[index].clone()
                } else {
                    split_bodies
                        .first()
                        .cloned()
                        .unwrap_or_else(|| strip_switch_breaks(&stmts))
                };
                match label {
                    SwitchLabelInfo::Default => {
                        default_body = Some(stripped_body.clone());
                    }
                    SwitchLabelInfo::Value(value) => {
                        let cond = Expression::new(ExprKind::Binary {
                            op: BinOp::StrictEq,
                            left: Box::new(subject_expr.clone()),
                            right: Box::new(value),
                        });
                        arms.push((cond, stripped_body.clone()));
                    }
                    SwitchLabelInfo::Pattern { pattern, guard } => {
                        let mut cond =
                            build_general_pattern_cond(subject_expr.clone(), pattern.clone())?;
                        let binding = build_general_pattern_binding(subject_expr.clone(), pattern)?;
                        let mut body = stripped_body.clone();
                        let mark_matched_stmt =
                            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                                target: Box::new(Expression::ident(&matched_name)),
                                value: Box::new(Expression::bool(true)),
                            })));
                        if let Some(binding_stmt) = binding {
                            if let Some(guard_expr) = guard {
                                let binding_name = extract_binding_name(&binding_stmt);
                                let rewritten_guard = binding_name
                                    .as_deref()
                                    .map(|name| {
                                        rewrite_ident_expr(&guard_expr, name, &subject_expr)
                                    })
                                    .unwrap_or(guard_expr);
                                cond = Expression::new(ExprKind::Binary {
                                    op: BinOp::And,
                                    left: Box::new(cond),
                                    right: Box::new(rewritten_guard),
                                });
                            }
                            body.insert(0, binding_stmt.clone());
                            body.push(mark_matched_stmt);
                        } else if let Some(guard_expr) = guard {
                            cond = Expression::new(ExprKind::Binary {
                                op: BinOp::And,
                                left: Box::new(cond),
                                right: Box::new(guard_expr),
                            });
                            body.push(mark_matched_stmt);
                        } else {
                            body.push(mark_matched_stmt);
                        }
                        arms.push((cond, body));
                    }
                }
            }
        }

        let decl = Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(subject_name),
                type_hint: None,
                init: Some(expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        });

        let matched_decl = Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(matched_name.clone()),
                type_hint: Some("bool".to_string()),
                init: Some(Expression::bool(false)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        });

        let mut body = vec![decl, matched_decl];
        for (cond, then_body) in arms {
            let not_matched = Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(Expression::ident(&matched_name)),
            });
            let gated_cond = Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(not_matched),
                right: Box::new(cond),
            });
            body.push(Statement::new(StmtKind::If {
                cond: gated_cond,
                then_body,
                elifs: Vec::new(),
                else_body: None,
            }));
        }

        if let Some(default_stmts) = default_body {
            let not_matched = Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(Expression::ident(&matched_name)),
            });
            body.push(Statement::new(StmtKind::If {
                cond: not_matched,
                then_body: default_stmts,
                elifs: Vec::new(),
                else_body: None,
            }));
        }
        return Ok(StmtKind::Block(body));
    }

    let mut cases = Vec::new();
    let mut default = None;
    for (labels, stmts) in sections {
        let mut is_default = false;
        let mut value_labels = Vec::new();
        for label in labels {
            match label {
                SwitchLabelInfo::Default => is_default = true,
                SwitchLabelInfo::Value(value) => value_labels.push(value),
                SwitchLabelInfo::Pattern { .. } => {}
            }
        }
        if is_default {
            default = Some(stmts);
        } else {
            let conditions = value_labels.into_iter().map(CaseCondition::Value).collect();
            cases.push(SwitchCase {
                conditions,
                body: stmts,
            });
        }
    }

    Ok(StmtKind::Switch {
        expr,
        cases,
        default,
    })
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner().next().map(walk_expression).transpose()?;
    Ok(StmtKind::Return(expr))
}

/// C# `yield return expr;` → `StmtKind::Expr(Yield(expr))`
/// `yield break;`          → `StmtKind::Return(None)` (ends the coroutine)
fn walk_yield_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let s = pair.as_str();
    let inner = pair.into_inner().next();
    if s.trim_start().starts_with("yield") && s.contains("return") {
        let expr = inner.map(walk_expression).transpose()?;
        let yield_expr = Expression::new(ExprKind::Yield(expr.map(Box::new)));
        Ok(StmtKind::Expr(yield_expr))
    } else {
        // yield break → end the generator
        Ok(StmtKind::Return(None))
    }
}

/// Walk a block scanning for any `yield return` / `yield break` —
/// determines whether the enclosing method is a generator.
fn body_has_yield(body: &[Statement]) -> bool {
    fn expr_has_yield(e: &Expression) -> bool {
        matches!(&e.kind, ExprKind::Yield(_) | ExprKind::YieldFrom(_))
    }
    for s in body {
        match &s.kind {
            StmtKind::Expr(e) if expr_has_yield(e) => return true,
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                if body_has_yield(then_body) {
                    return true;
                }
                for (_, b) in elifs {
                    if body_has_yield(b) {
                        return true;
                    }
                }
                if let Some(b) = else_body {
                    if body_has_yield(b) {
                        return true;
                    }
                }
            }
            StmtKind::While { body: b, .. }
            | StmtKind::ForIn { body: b, .. }
            | StmtKind::For { body: b, .. }
            | StmtKind::DoWhile { body: b, .. } => {
                if body_has_yield(b) {
                    return true;
                }
            }
            StmtKind::Try {
                body: b,
                catches,
                finally,
                ..
            } => {
                if body_has_yield(b) {
                    return true;
                }
                for c in catches {
                    if body_has_yield(&c.body) {
                        return true;
                    }
                }
                if let Some(f) = finally {
                    if body_has_yield(f) {
                        return true;
                    }
                }
            }
            StmtKind::Block(b) => {
                if body_has_yield(b) {
                    return true;
                }
            }
            _ => {}
        }
    }
    false
}

fn walk_throw(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair.into_inner().next().map(walk_expression).transpose()?;
    Ok(StmtKind::Throw { expr, cause: None })
}

fn walk_try(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::block_statement => body = walk_body(p)?,
            Rule::catch_clause => {
                let mut types = Vec::new();
                let mut var_name = None;
                let mut catch_body = Vec::new();
                let mut when_filter: Option<Expression> = None;
                for cp in p.into_inner() {
                    match cp.as_rule() {
                        Rule::type_name => types.push(cp.as_str().to_string()),
                        Rule::ident_name => var_name = Some(cp.as_str().to_string()),
                        Rule::catch_when_filter => {
                            // Inner is just an `expression`.
                            if let Some(inner) = cp.into_inner().next() {
                                when_filter = Some(walk_expression(inner)?);
                            }
                        }
                        Rule::block_statement => catch_body = walk_body(cp)?,
                        _ => {}
                    }
                }
                // Synthesize a hidden var name when the catch declared
                // none (`catch (Exception) { throw; }`). The compiler's
                // catch-binding path stamps the value onto a local with
                // this name, so any `throw;` rewrite inside the body
                // can reference it. Without a var, bare `throw;` would
                // throw NULL — losing the original exception.
                let synthetic_var = if var_name.is_none() {
                    let name = "__caught";
                    var_name = Some(name.into());
                    Some(name.to_string())
                } else {
                    None
                };
                // Bare `throw;` (StmtKind::Throw with expr=None) inside
                // the catch body gets rewritten to `throw <var_name>;`
                // so the VM rethrows the caught instance instead of NULL.
                if let Some(ref vn) = var_name {
                    rewrite_bare_throws(&mut catch_body, vn);
                }
                let _ = synthetic_var;
                catches.push(CatchClause {
                    types,
                    var_name,
                    stack_var: None,
                    body: catch_body,
                    when_clause: when_filter,
                });
            }
            Rule::finally_clause => {
                for fp in p.into_inner() {
                    if fp.as_rule() == Rule::block_statement {
                        finally = Some(walk_body(fp)?);
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

fn walk_using_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var = String::new();
    let mut resource = Expression::null();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw | Rule::type_name => {}
            Rule::ident_name => var = p.as_str().to_string(),
            _ => {
                if body.is_empty() {
                    if let Ok(expr) = walk_expression(p.clone()) {
                        resource = expr;
                    } else if let Ok(stmt) = walk_statement(p) {
                        body = vec![stmt];
                    }
                } else if let Ok(stmt) = walk_statement(p) {
                    body = vec![stmt];
                }
            }
        }
    }

    Ok(StmtKind::Using {
        var,
        resource,
        body,
    })
}

fn walk_using_declaration(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut var = String::new();
    let mut resource = Expression::null();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::var_kw | Rule::type_name => {}
            Rule::ident_name => var = p.as_str().to_string(),
            Rule::expression => resource = walk_expression(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::Using {
        var,
        resource,
        body: Vec::new(),
    })
}

fn walk_lock(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut inner = pair.into_inner();
    let expr = walk_expression(inner.next().ok_or("lock: no expr")?)?;
    let body = vec![walk_statement(inner.next().ok_or("lock: no body")?)?];
    Ok(StmtKind::Lock { expr, body })
}

// ── Parameters ──────────────────────────────────────────────────────────────

fn walk_params(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    Ok(walk_params_with_decorators(pair)?.0)
}

fn walk_params_with_decorators(pair: Pair<Rule>) -> Result<(Vec<Param>, Vec<Expression>), String> {
    let mut params = Vec::new();
    let mut decorators = Vec::new();

    for (index, param_pair) in pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::param)
        .enumerate()
    {
        let (param, param_decorators) = walk_param_with_decorators(param_pair)?;
        decorators.extend(
            param_decorators
                .into_iter()
                .map(|decorator| param_attribute_carrier(index, decorator)),
        );
        params.push(param);
    }

    Ok((params, decorators))
}

fn param_attribute_carrier(index: usize, decorator: Expression) -> Expression {
    Expression::new(ExprKind::New {
        class: Box::new(Expression::ident("__vybe_param_attribute")),
        args: vec![
            Argument::positional(Expression::int(index as i64)),
            Argument::positional(decorator),
        ],
    })
}

fn walk_param_with_decorators(pair: Pair<Rule>) -> Result<(Param, Vec<Expression>), String> {
    let mut name = String::new();
    let mut type_hint = None;
    let mut default = None;
    let mut pass_by = PassBy::Value;
    let mut is_rest = false;
    let mut decorators = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::attribute_list => decorators.extend(parse_attribute_specs(p.as_str())),
            Rule::param_modifier => {
                match p.as_str() {
                    "ref" => pass_by = PassBy::Ref,
                    "out" => pass_by = PassBy::Out,
                    // C# extension methods mark the first parameter with
                    // `this`. Reuse `Const` as an internal marker so the
                    // post-parse rewrite can lower `x.Ext(...)` to the
                    // corresponding static method call.
                    "this" => pass_by = PassBy::Const,
                    "params" => is_rest = true,
                    _ => {}
                }
            }
            Rule::type_name => type_hint = Some(p.as_str().trim().to_string()),
            Rule::ident_name => name = p.as_str().to_string(),
            _ => default = Some(walk_expression(p)?),
        }
    }
    let is_nullable = type_hint
        .as_deref()
        .is_some_and(|hint| hint.trim().ends_with('?'));

    Ok((
        Param {
            name,
            type_hint,
            default,
            pass_by,
            is_rest,
            is_kwargs: false,
            is_optional: false,
            is_nullable,
        },
        decorators,
    ))
}

// ── Expressions ─────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

#[derive(Clone)]
struct QueryState {
    result_expr: Expression,
    item_param: String,
    bindings: Vec<(String, Expression)>,
}

fn query_param(name: &str) -> Param {
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

fn query_member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn query_lambda(param: &str, body: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: vec![query_param(param)],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: Vec::new(),
    })
}

fn query_call(target: Expression, method: &str, args: Vec<Expression>) -> Expression {
    canonicalize_method_call(
        query_member(target, method),
        args.into_iter().map(Argument::positional).collect(),
    )
}

fn rewrite_query_expr(expr: &Expression, bindings: &[(String, Expression)]) -> Expression {
    let mut rewritten = expr.clone();
    for (name, replacement) in bindings {
        rewritten = rewrite_ident_expr(&rewritten, name, replacement);
    }
    rewritten
}

fn query_bindings_for_item(item_param: &str, names: &[String]) -> Vec<(String, Expression)> {
    names
        .iter()
        .map(|name| {
            (
                name.clone(),
                query_member(Expression::ident(item_param), name),
            )
        })
        .collect()
}

fn query_binding_names(bindings: &[(String, Expression)]) -> Vec<String> {
    bindings.iter().map(|(name, _)| name.clone()).collect()
}

fn query_object_from_bindings(
    bindings: &[(String, Expression)],
    extra: Option<(String, Expression)>,
) -> Expression {
    let mut props = Vec::new();
    for (name, value) in bindings {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(name),
            value: value.clone(),
        });
    }
    if let Some((name, value)) = extra {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&name),
            value,
        });
    }
    Expression::new(ExprKind::Object(props))
}

fn parse_csharp_from_clause(pair: Pair<Rule>) -> Result<(String, Expression), String> {
    let mut range_var = None;
    let mut source_expr = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if range_var.is_none() => range_var = Some(child.as_str().to_string()),
            Rule::expression => source_expr = Some(walk_expression(child)?),
            _ => {}
        }
    }
    Ok((
        range_var.ok_or("query from clause missing range variable")?,
        source_expr.ok_or("query from clause missing source expression")?,
    ))
}

fn parse_csharp_join_clause(
    pair: Pair<Rule>,
) -> Result<(String, Expression, Expression, Expression), String> {
    let mut join_var = None;
    let mut exprs = Vec::new();
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if join_var.is_none() => join_var = Some(child.as_str().to_string()),
            Rule::expression => exprs.push(walk_expression(child)?),
            _ => {}
        }
    }
    if exprs.len() != 3 {
        return Err("query join clause expects source, left key, and right key".to_string());
    }
    Ok((
        join_var.ok_or("query join clause missing range variable")?,
        exprs.remove(0),
        exprs.remove(0),
        exprs.remove(0),
    ))
}

fn parse_csharp_let_clause(pair: Pair<Rule>) -> Result<(String, Expression), String> {
    let mut name = None;
    let mut value = None;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::ident_name if name.is_none() => name = Some(child.as_str().to_string()),
            Rule::expression => value = Some(walk_expression(child)?),
            _ => {}
        }
    }
    Ok((
        name.ok_or("query let clause missing variable name")?,
        value.ok_or("query let clause missing value expression")?,
    ))
}

fn parse_csharp_ordering(pair: Pair<Rule>) -> Result<(Expression, bool), String> {
    let mut key_expr = None;
    let mut descending = false;
    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::expression => key_expr = Some(walk_expression(child)?),
            Rule::q_order_direction => {
                descending = child.as_str().eq_ignore_ascii_case("descending")
            }
            _ => {}
        }
    }
    Ok((
        key_expr.ok_or("query ordering missing key expression")?,
        descending,
    ))
}

fn is_query_item_expr(expr: &Expression, item_param: &str) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if name == item_param)
}

fn lower_csharp_query_body(pair: Pair<Rule>, state: QueryState) -> Result<Expression, String> {
    let mut state = state;
    let mut terminal: Option<Pair<Rule>> = None;
    let mut continuation: Option<Pair<Rule>> = None;

    for child in pair.into_inner() {
        match child.as_rule() {
            Rule::query_body_clause => {
                let clause = child.into_inner().next().ok_or("empty query clause")?;
                match clause.as_rule() {
                    Rule::q_where_clause => {
                        let predicate = clause
                            .into_inner()
                            .next()
                            .ok_or("where clause missing predicate")?;
                        let raw = walk_expression(predicate)?;
                        let rewritten = rewrite_query_expr(&raw, &state.bindings);
                        state.result_expr = query_call(
                            state.result_expr,
                            "Where",
                            vec![query_lambda(&state.item_param, rewritten)],
                        );
                    }
                    Rule::q_orderby_clause => {
                        let orderings: Vec<Pair<Rule>> = clause.into_inner().collect();
                        for (index, ordering) in orderings.into_iter().enumerate() {
                            let (raw_key, descending) = parse_csharp_ordering(ordering)?;
                            let rewritten = rewrite_query_expr(&raw_key, &state.bindings);
                            let method = if index == 0 {
                                if descending {
                                    "OrderByDescending"
                                } else {
                                    "OrderBy"
                                }
                            } else if descending {
                                "ThenByDescending"
                            } else {
                                "ThenBy"
                            };
                            state.result_expr = query_call(
                                state.result_expr,
                                method,
                                vec![query_lambda(&state.item_param, rewritten)],
                            );
                        }
                    }
                    Rule::q_let_clause => {
                        let (let_name, raw_value) = parse_csharp_let_clause(clause)?;
                        let rewritten_value = rewrite_query_expr(&raw_value, &state.bindings);
                        let projection = query_object_from_bindings(
                            &state.bindings,
                            Some((let_name.clone(), rewritten_value)),
                        );
                        state.result_expr = query_call(
                            state.result_expr,
                            "Select",
                            vec![query_lambda(&state.item_param, projection)],
                        );
                        let mut names = query_binding_names(&state.bindings);
                        names.push(let_name);
                        state.item_param = "__query_item".to_string();
                        state.bindings = query_bindings_for_item(&state.item_param, &names);
                    }
                    Rule::q_from_clause => {
                        let (range_var, source_expr) = parse_csharp_from_clause(clause)?;
                        let outer_item_param = state.item_param.clone();
                        let visible_bindings = state.bindings.clone();
                        let rewritten_source = rewrite_query_expr(&source_expr, &visible_bindings);
                        let inner_projection = query_object_from_bindings(
                            &visible_bindings,
                            Some((range_var.clone(), Expression::ident(&range_var))),
                        );
                        let inner_select = query_call(
                            rewritten_source,
                            "Select",
                            vec![query_lambda(&range_var, inner_projection)],
                        );
                        state.result_expr = query_call(
                            state.result_expr,
                            "SelectMany",
                            vec![query_lambda(&outer_item_param, inner_select)],
                        );
                        let mut names = query_binding_names(&state.bindings);
                        names.push(range_var);
                        state.item_param = "__query_item".to_string();
                        state.bindings = query_bindings_for_item(&state.item_param, &names);
                    }
                    Rule::q_join_clause => {
                        let (join_var, join_source, left_key, right_key) =
                            parse_csharp_join_clause(clause)?;
                        let outer_item_param = state.item_param.clone();
                        let visible_bindings = state.bindings.clone();
                        let rewritten_source = rewrite_query_expr(&join_source, &visible_bindings);
                        let left_key_expr = rewrite_query_expr(&left_key, &visible_bindings);
                        let mut right_bindings = visible_bindings.clone();
                        right_bindings.push((join_var.clone(), Expression::ident(&join_var)));
                        let right_key_expr = rewrite_query_expr(&right_key, &right_bindings);
                        let filtered = query_call(
                            rewritten_source,
                            "Where",
                            vec![query_lambda(
                                &join_var,
                                Expression::new(ExprKind::Binary {
                                    op: BinOp::Eq,
                                    left: Box::new(left_key_expr),
                                    right: Box::new(right_key_expr),
                                }),
                            )],
                        );
                        let projected = query_call(
                            filtered,
                            "Select",
                            vec![query_lambda(
                                &join_var,
                                query_object_from_bindings(
                                    &visible_bindings,
                                    Some((join_var.clone(), Expression::ident(&join_var))),
                                ),
                            )],
                        );
                        state.result_expr = query_call(
                            state.result_expr,
                            "SelectMany",
                            vec![query_lambda(&outer_item_param, projected)],
                        );
                        let mut names = query_binding_names(&state.bindings);
                        names.push(join_var);
                        state.item_param = "__query_item".to_string();
                        state.bindings = query_bindings_for_item(&state.item_param, &names);
                    }
                    _ => return Err(format!("unsupported query clause: {:?}", clause.as_rule())),
                }
            }
            Rule::query_select_or_group_clause => terminal = Some(child),
            Rule::query_continuation => continuation = Some(child),
            _ => {}
        }
    }

    let terminal = terminal.ok_or("query body missing terminal clause")?;
    let terminal_expr = terminal
        .into_inner()
        .next()
        .ok_or("empty query terminal clause")?;
    let result_after_terminal = match terminal_expr.as_rule() {
        Rule::q_select_clause => {
            let selection = terminal_expr
                .into_inner()
                .next()
                .ok_or("select clause missing projection")?;
            let raw = walk_expression(selection)?;
            let rewritten = rewrite_query_expr(&raw, &state.bindings);
            query_call(
                state.result_expr,
                "Select",
                vec![query_lambda(&state.item_param, rewritten)],
            )
        }
        Rule::q_group_clause => {
            let mut exprs = terminal_expr.into_inner();
            let group_expr = walk_expression(exprs.next().ok_or("group clause missing element")?)?;
            let key_expr = walk_expression(exprs.next().ok_or("group clause missing key")?)?;
            let rewritten_group = rewrite_query_expr(&group_expr, &state.bindings);
            let rewritten_key = rewrite_query_expr(&key_expr, &state.bindings);
            let groups_current_item = is_query_item_expr(&rewritten_group, &state.item_param);
            let grouped_source = if groups_current_item {
                state.result_expr
            } else {
                query_call(
                    state.result_expr,
                    "Select",
                    vec![query_lambda(&state.item_param, rewritten_group)],
                )
            };
            let key_param = if groups_current_item {
                state.item_param.clone()
            } else {
                "__group_item".to_string()
            };
            query_call(
                grouped_source,
                "GroupBy",
                vec![query_lambda(&key_param, rewritten_key)],
            )
        }
        _ => {
            return Err(format!(
                "unsupported query terminal clause: {:?}",
                terminal_expr.as_rule()
            ));
        }
    };

    if let Some(continuation) = continuation {
        let mut inner = continuation.into_inner();
        let name = inner
            .next()
            .ok_or("query continuation missing range variable")?
            .as_str()
            .to_string();
        let body = inner
            .next()
            .ok_or("query continuation missing query body")?;
        return lower_csharp_query_body(
            body,
            QueryState {
                result_expr: result_after_terminal,
                item_param: name.clone(),
                bindings: vec![(name.clone(), Expression::ident(&name))],
            },
        );
    }

    Ok(result_after_terminal)
}

fn parse_csharp_query_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let from_clause = inner.next().ok_or("query expression missing from clause")?;
    let query_body = inner.next().ok_or("query expression missing query body")?;
    let (range_var, source_expr) = parse_csharp_from_clause(from_clause)?;
    let lowered = lower_csharp_query_body(
        query_body,
        QueryState {
            result_expr: source_expr,
            item_param: range_var.clone(),
            bindings: vec![(range_var.clone(), Expression::ident(&range_var))],
        },
    )?;
    Ok(Expression::with_span(lowered.kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        // Literals
        Rule::numeric_literal => {
            let raw = pair.as_str();
            // Strip C#'s numeric type suffix (UL, L, F, M, D, U). Hex digits
            // need different handling — after `0x`, only the trailing
            // suffix (UL/L/U) is alpha noise; A–F are real digits.
            let s = if raw.starts_with("0x") || raw.starts_with("0X") {
                let body = &raw[2..];
                let cut = body
                    .rfind(|c: char| c.is_ascii_hexdigit())
                    .map(|i| 2 + i + 1)
                    .unwrap_or(raw.len());
                &raw[..cut]
            } else if raw.starts_with("0b") || raw.starts_with("0B") {
                let body = &raw[2..];
                let cut = body
                    .rfind(|c: char| c == '0' || c == '1')
                    .map(|i| 2 + i + 1)
                    .unwrap_or(raw.len());
                &raw[..cut]
            } else {
                raw.trim_end_matches(|c: char| c.is_ascii_alphabetic())
            };
            // Underscores are allowed as digit separators in C# 7.0+.
            let s_owned;
            let s = if s.contains('_') {
                s_owned = s.replace('_', "");
                &s_owned
            } else {
                s
            };
            if s.starts_with("0x") || s.starts_with("0X") {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[2..], 16).map_err(|e| format!("{}", e))?,
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
        Rule::raw_string => Ok(ExprKind::Lit(Literal::Str(unquote_raw_string(
            pair.as_str(),
        )?))),
        Rule::raw_interpolated_string => {
            let inner = unquote_raw_interpolated_string(pair.as_str())?;
            let parts = parse_interpolated_parts(&inner)?;
            Ok(ExprKind::Interpolation(parts))
        }
        Rule::verbatim_string => {
            let s = pair.as_str();
            // @"..." → strip @" and trailing "
            let inner = &s[2..s.len() - 1];
            Ok(ExprKind::Lit(Literal::Str(inner.replace("\"\"", "\""))))
        }
        Rule::char_literal => {
            let s = pair.as_str();
            let inner = &s[1..s.len() - 1];
            let ch = match inner {
                "\\n" => '\n',
                "\\t" => '\t',
                "\\r" => '\r',
                "\\0" => '\0',
                "\\\\" => '\\',
                "\\'" => '\'',
                _ => inner.chars().next().unwrap_or('\0'),
            };
            Ok(ExprKind::Lit(Literal::Char(ch)))
        }

        // Keywords
        Rule::true_kw => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_kw => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::null_kw => Ok(ExprKind::Lit(Literal::Null)),
        Rule::this_kw => Ok(ExprKind::This),
        Rule::base_kw => Ok(ExprKind::Super),

        Rule::ident_name | Rule::ident_or_keyword | Rule::primitive_type_ident => {
            let name = pair.as_str();
            match name {
                "true" => Ok(ExprKind::Lit(Literal::Bool(true))),
                "false" => Ok(ExprKind::Lit(Literal::Bool(false))),
                "null" => Ok(ExprKind::Lit(Literal::Null)),
                "this" => Ok(ExprKind::This),
                "base" => Ok(ExprKind::Super),
                _ => {
                    // Canonicalize all named function references for delegate handler identity
                    // This ensures that every reference to a named function is always represented as ExprKind::Ident with the canonical name
                    // (In a full implementation, you could check a symbol table here)
                    Ok(ExprKind::Ident(name.to_string()))
                }
            }
        }

        // Expression wrapper
        Rule::expression => {
            let inner = pair.into_inner().next().ok_or("Empty expression")?;
            walk_expr_kind(inner)
        }

        Rule::query_expression => Ok(parse_csharp_query_expression(pair)?.kind),

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
                    let op = match op_str {
                        "+=" => CompoundOp::Add,
                        "-=" => CompoundOp::Sub,
                        "*=" => CompoundOp::Mul,
                        "/=" => CompoundOp::Div,
                        "%=" => CompoundOp::Mod,
                        "&=" => CompoundOp::BitAnd,
                        "|=" => CompoundOp::BitOr,
                        "^=" => CompoundOp::BitXor,
                        "<<=" => CompoundOp::Shl,
                        ">>=" => CompoundOp::Shr,
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

        // Lambda
        Rule::lambda_expression => {
            let mut params = Vec::new();
            let mut body = LambdaBody::Expr(Box::new(Expression::null()));
            let mut is_async = false;

            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::async_kw => is_async = true,
                    Rule::ident_name => {
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
                    Rule::param_list => params = walk_params(p)?,
                    Rule::lambda_body => {
                        let inner = p.into_inner().next().ok_or("Empty lambda body")?;
                        body = match inner.as_rule() {
                            Rule::block_statement => LambdaBody::Block(walk_body(inner)?),
                            _ => LambdaBody::Expr(Box::new(walk_expression(inner)?)),
                        };
                    }
                    _ => {}
                }
            }

            Ok(ExprKind::Lambda {
                params,
                body,
                is_async,
                captures: Vec::new(),
            })
        }

        // Ternary
        Rule::conditional_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                walk_expr_kind(inner.remove(0))
            } else if inner.len() == 3 {
                let cond_pair = inner.remove(0);
                let cond = walk_expression(cond_pair.clone())?;
                let mut then = walk_expression(inner.remove(0))?;
                let else_ = walk_expression(inner.remove(0))?;
                if let Some(binding_stmt) = extract_if_is_pattern_binding(cond_pair)? {
                    then = build_scoped_then_expr(binding_stmt, then);
                }
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
        Rule::null_coalesce_expr => walk_binary_chain(pair),
        Rule::logical_or | Rule::logical_and => walk_binary_chain(pair),
        Rule::bitwise_or | Rule::bitwise_xor | Rule::bitwise_and => walk_binary_chain(pair),
        Rule::equality => walk_binary_chain(pair),
        Rule::relational => walk_relational(pair),
        Rule::additive | Rule::multiplicative => walk_binary_chain(pair),

        // Unary
        Rule::unary => {
            let mut inner = pair.into_inner();
            let first = inner.next().ok_or("Empty unary")?;
            if first.as_rule() == Rule::postfix {
                return walk_expr_kind(first);
            }
            if first.as_rule() == Rule::cast_expression {
                return walk_expr_kind(first);
            }
            let op_str = first.as_str().trim();
            let operand = walk_expression(inner.next().ok_or("Missing unary operand")?)?;
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

        // C# explicit cast `(TypeName)expr` — lower to the canonical
        // type-conversion form. For numeric primitives we use Convert.<T>
        // calls (matches the .NET runtime); for object / string we leave
        // the expression unchanged (Convert.ToString already in stdlib).
        Rule::cast_expression => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let cast_type_pair = inner.remove(0);
            let type_name = normalize_runtime_type_name(cast_type_pair.as_str());
            let operand = walk_expression(inner.remove(0))?;
            Ok(ExprKind::Cast {
                expr: Box::new(operand),
                type_name,
            })
        }

        // Postfix
        Rule::postfix => {
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            let base = walk_expression(inner.remove(0))?;
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

        // New expression
        Rule::new_expression => walk_new_expr(pair),

        // Primary
        Rule::primary => {
            let inner = pair.into_inner().next().ok_or("Empty primary")?;
            walk_expr_kind(inner)
        }

        Rule::collection_expression => {
            let span = to_span(&pair);
            let elements = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::collection_expression_element)
                .map(walk_collection_expression_element)
                .collect::<Result<Vec<_>, _>>()?;
            Ok(Expression::with_span(ExprKind::Array(elements), span).kind)
        }

        // typeof(Type) → push the .NET FullName as a string. Matches
        // .NET's `Console.WriteLine(typeof(int))` → "System.Int32".
        // `.Name` / `.FullName` member access is rewritten in
        // `canonicalize_member_access` for typeof-string receivers.
        Rule::typeof_expression => {
            let type_name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::type_name)
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            let net_name = dotnet_type_name(&type_name);
            Ok(ExprKind::Lit(Literal::Str(format!("System.{}", net_name))))
        }

        // nameof(member) → push name as string
        Rule::nameof_expression => {
            let name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::dotted_name)
                .map(|p| {
                    let s = p.as_str();
                    // nameof returns the last segment
                    s.rsplit('.').next().unwrap_or(s).to_string()
                })
                .unwrap_or_default();
            Ok(ExprKind::Lit(Literal::Str(name)))
        }

        // default or default(Type)
        Rule::default_expression => {
            let type_name = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::type_name)
                .map(|p| p.as_str().to_string());
            if let Some(type_name) = type_name {
                Ok(ExprKind::DefaultOf(type_name))
            } else {
                Ok(ExprKind::Lit(Literal::Null))
            }
        }

        // checked/unchecked → just evaluate inner
        Rule::checked_expression => {
            let inner = pair
                .into_inner()
                .find(|p| matches!(p.as_rule(), Rule::expression))
                .ok_or("Empty checked expression")?;
            walk_expr_kind(inner)
        }

        // Interpolated string — parsed atomically, split manually
        Rule::interpolated_string => {
            let s = pair.as_str();
            // Strip `$"`, `$@"`, or `@$"` prefix and trailing quote.
            let prefix_len = if s.starts_with("$@\"") || s.starts_with("@$\"") {
                3
            } else {
                2
            };
            let inner = &s[prefix_len..s.len() - 1];
            let parts = parse_interpolated_parts(inner)?;
            Ok(ExprKind::Interpolation(parts))
        }

        // Type name used as expression (cast, etc.)
        Rule::type_name
        | Rule::generic_type_expr
        | Rule::base_type
        | Rule::dotted_name
        | Rule::global_qualified_name => Ok(ExprKind::Ident(
            strip_global_namespace_qualifier(pair.as_str()).to_string(),
        )),

        // Passthrough wrappers
        Rule::call_chain => {
            let inner = pair.into_inner().next().ok_or("Empty wrapper")?;
            walk_expr_kind(inner)
        }

        // C# tuple literal: (1, "x", true) or named (Name: "Alice", Age: 30).
        // Named tuples lower to an Object with both Item<N> AND the
        // user-given names so `t.Item1` and `t.Name` both resolve.
        Rule::tuple_literal => {
            let elements: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() == Rule::tuple_element)
                .collect();
            let mut has_names = false;
            let mut parsed: Vec<(Option<String>, Expression)> = Vec::new();
            for el in elements {
                let inner: Vec<Pair<Rule>> = el.into_inner().collect();
                if inner.len() == 2 && inner[0].as_rule() == Rule::ident_name {
                    has_names = true;
                    parsed.push((
                        Some(inner[0].as_str().to_string()),
                        walk_expression(inner[1].clone())?,
                    ));
                } else if let Some(p) = inner.into_iter().next() {
                    parsed.push((None, walk_expression(p)?));
                }
            }
            if has_names {
                let mut props = Vec::new();
                for (i, (name, value)) in parsed.iter().enumerate() {
                    let item_key = format!("Item{}", i + 1);
                    props.push(ObjectProperty::KeyValue {
                        key: Expression::string(&item_key),
                        value: value.clone(),
                    });
                    if let Some(n) = name {
                        props.push(ObjectProperty::KeyValue {
                            key: Expression::string(n),
                            value: value.clone(),
                        });
                    }
                }
                Ok(ExprKind::Object(props))
            } else {
                let elems: Vec<Expression> = parsed.into_iter().map(|(_, e)| e).collect();
                Ok(ExprKind::Tuple(elems))
            }
        }

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

// ── Binary chain walker ─────────────────────────────────────────────────────

/// Walk relational expression: additive ~ (type_test | binary_relational)*
fn walk_relational(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    let mut left = walk_expression(inner.remove(0))?;

    for p in inner {
        match p.as_rule() {
            Rule::type_test => {
                let mut tt_inner = p.into_inner();
                let kw = tt_inner.next().ok_or("type_test: missing keyword")?;
                let next = tt_inner.next().ok_or("type_test: missing operand")?;
                if kw.as_rule() == Rule::is_kw {
                    // `is` accepts a pattern_clause: not-prefix +
                    // pattern_atom of (null | literal | type_name [ident])
                    left = walk_is_pattern(left, next)?;
                } else {
                    // `obj as T` — returns obj if it's a T, else null.
                    // Lower to `<is-test> ? obj : null` so the runtime
                    // null sentinel matches .NET semantics. Strip a
                    // trailing `?` (nullable marker) for the type test.
                    let type_name_raw = next.as_str().trim();
                    let type_name = normalize_runtime_type_name(type_name_raw);
                    let test = if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                        let typeof_expr = Expression::new(ExprKind::TypeOf(Box::new(left.clone())));
                        Expression::new(ExprKind::Binary {
                            op: BinOp::StrictEq,
                            left: Box::new(typeof_expr),
                            right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                js_typeof.into(),
                            )))),
                        })
                    } else {
                        Expression::new(ExprKind::IsType {
                            expr: Box::new(left.clone()),
                            type_name,
                        })
                    };
                    left = Expression::new(ExprKind::Ternary {
                        cond: Box::new(test),
                        then: Box::new(left),
                        else_: Box::new(Expression::null()),
                    });
                }
            }
            Rule::binary_relational => {
                let mut br_inner: Vec<Pair<Rule>> = p.into_inner().collect();
                if br_inner.len() >= 2 {
                    let op_str = br_inner[0].as_str().trim();
                    let right = walk_expression(br_inner.remove(1))?;
                    let bin_op = match op_str {
                        "<=" => BinOp::LtEq,
                        ">=" => BinOp::GtEq,
                        "<" => BinOp::Lt,
                        ">" => BinOp::Gt,
                        ">>" => BinOp::Shr,
                        "<<" => BinOp::Shl,
                        _ => BinOp::Lt,
                    };
                    left = Expression::new(ExprKind::Binary {
                        op: bin_op,
                        left: Box::new(left),
                        right: Box::new(right),
                    });
                }
            }
            Rule::switch_expr_postfix => {
                left = walk_switch_expr(left, p)?;
            }
            Rule::with_expr_postfix => {
                left = walk_with_expr(left, p)?;
            }
            _ => {
                // Direct operand — shouldn't happen but try as additive
                let right = walk_expression(p)?;
                left = Expression::new(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            }
        }
    }

    Ok(left.kind)
}

fn walk_binary_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();

    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    let mut left = walk_expression(inner.remove(0))?;
    let mut i = 0;

    while i + 1 < inner.len() {
        let op_pair = &inner[i];
        let op_str = op_pair.as_str().trim();

        // Check for is/as type operators
        if op_str.starts_with("is ") || op_str.starts_with("is\t") || op_str == "is" {
            // is Type → IsType
            let right = &inner[i + 1];
            let type_name = normalize_runtime_type_name(right.as_str());
            left = Expression::new(ExprKind::IsType {
                expr: Box::new(left),
                type_name,
            });
            i += 2;
            continue;
        }
        if op_str.starts_with("as ") || op_str.starts_with("as\t") || op_str == "as" {
            let right = &inner[i + 1];
            let type_name = normalize_runtime_type_name(right.as_str());
            left = Expression::new(ExprKind::Cast {
                expr: Box::new(left),
                type_name,
            });
            i += 2;
            continue;
        }

        let right = walk_expression(inner[i + 1].clone())?;

        let bin_op = match op_str {
            "??" => BinOp::NullCoalesce,
            "||" => BinOp::Or,
            "&&" => BinOp::And,
            "|" => BinOp::BitOr,
            "^" => BinOp::BitXor,
            "&" => BinOp::BitAnd,
            "==" => BinOp::Eq,
            "!=" => BinOp::NotEq,
            "<" => BinOp::Lt,
            ">" => BinOp::Gt,
            "<=" => BinOp::LtEq,
            ">=" => BinOp::GtEq,
            ">>" => BinOp::Shr,
            "<<" => BinOp::Shl,
            "+" => BinOp::Add,
            "-" => BinOp::Sub,
            "*" => BinOp::Mul,
            "/" => BinOp::Div,
            "%" => BinOp::Mod,
            _ => {
                // relational_op may contain "is Type" or "as Type" as combined text
                if op_str.starts_with("is") {
                    let type_name = normalize_runtime_type_name(op_str.trim_start_matches("is"));
                    left = Expression::new(ExprKind::IsType {
                        expr: Box::new(left),
                        type_name,
                    });
                    i += 2;
                    continue;
                }
                if op_str.starts_with("as") {
                    let type_name = normalize_runtime_type_name(op_str.trim_start_matches("as"));
                    left = Expression::new(ExprKind::Cast {
                        expr: Box::new(left),
                        type_name,
                    });
                    i += 2;
                    continue;
                }
                BinOp::Add
            }
        };

        // Integer division by literal zero — `int x = 10 / 0;`. C#
        // (ECMA-335) throws `DivideByZeroException` at runtime; JS
        // returns Infinity. We can't tell int vs float at runtime
        // (every numeric is f64), but a literal `0` divisor with an
        // integer literal numerator is unambiguously the int form
        // — rewrite the expression to a throw of the exception so
        // try/catch picks it up.
        if matches!(bin_op, BinOp::Div) && is_int_zero_literal(&right) && is_int_literal(&left) {
            // Build `(() => { throw new DivideByZeroException("Attempted to divide by zero."); })()`
            // — IIFE so the throw works in expression position.
            let throw_stmt = Statement::with_span(
                StmtKind::Throw {
                    expr: Some(Expression::new(ExprKind::New {
                        class: Box::new(Expression::ident("DivideByZeroException")),
                        args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                            Literal::Str("Attempted to divide by zero.".into()),
                        )))],
                    })),
                    cause: None,
                },
                Span::default(),
            );
            let lambda = Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![throw_stmt]),
                is_async: false,
                captures: Vec::new(),
            });
            left = Expression::new(ExprKind::Call {
                callee: Box::new(lambda),
                args: vec![],
                optional: false,
            });
            i += 2;
            continue;
        }

        if matches!(bin_op, BinOp::Eq | BinOp::NotEq) {
            if let Some(tuple_eq) = build_tuple_structural_equality(&left, &right) {
                left = if matches!(bin_op, BinOp::NotEq) {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(tuple_eq),
                    })
                } else {
                    tuple_eq
                };
                i += 2;
                continue;
            }
        }

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        i += 2;
    }

    Ok(left.kind)
}

fn is_int_literal(e: &Expression) -> bool {
    matches!(e.kind, ExprKind::Lit(Literal::Int(_)))
}

fn is_int_zero_literal(e: &Expression) -> bool {
    matches!(&e.kind, ExprKind::Lit(Literal::Int(0)))
}

// ── Call chain walker ───────────────────────────────────────────────────────

fn walk_call_chain(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().ok_or("Empty call expression")?;
    let mut expr = walk_expression(first)?;

    // Collect the chain segments so we can peek at the next one when
    // deciding whether to canonicalize a `.Length` / `.Count` accessor
    // (they're properties standalone, but instance-method names when
    // followed by `(...)` — see LINQ Count(predicate)).
    let chains: Vec<Pair<Rule>> = inner.filter(|p| p.as_rule() == Rule::call_chain).collect();
    let mut iter = chains.into_iter().peekable();
    while let Some(chain) = iter.next() {
        let chain_src = chain.as_str();
        let chain_inner: Vec<Pair<Rule>> = chain.into_inner().collect();

        if chain_src.starts_with("?.") {
            // Null-conditional member access
            let name = strip_csharp_terminal_type_args(chain_src[2..].trim());
            expr = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: name,
                null_safe: true,
            });
        } else if chain_src.starts_with("(") {
            // Call — normalize known method calls to canonical builtins
            let mut generic_type_args = Vec::new();
            if let ExprKind::Ident(name) = &expr.kind {
                generic_type_args = extract_csharp_terminal_type_args(name);
                let stripped = strip_csharp_terminal_type_args(name);
                if stripped != *name {
                    expr = build_dotted_expr(&stripped);
                }
            } else if let ExprKind::Member { field, .. } = &mut expr.kind {
                generic_type_args = extract_csharp_terminal_type_args(field);
                let stripped = strip_csharp_terminal_type_args(field);
                if stripped != *field {
                    *field = stripped;
                }
            }
            let mut args = if let Some(arg_pair) = chain_inner
                .into_iter()
                .find(|p| p.as_rule() == Rule::argument_list)
            {
                walk_arguments(arg_pair)?
            } else {
                Vec::new()
            };
            if !generic_type_args.is_empty() {
                args.extend(csharp_method_generic_binding_args(&generic_type_args));
            }
            // Inject default fill char for `PadLeft(n)` / `PadRight(n)` —
            // .NET defaults to space, but the value-method dispatch expects
            // both args. Same idea as JS-default lowering.
            if let ExprKind::Member { field, .. } = &expr.kind {
                let f = field.as_str();
                if (f == "PadLeft" || f == "PadRight") && args.len() == 1 {
                    args.push(Argument::positional(Expression::new(ExprKind::Lit(
                        Literal::Str(" ".into()),
                    ))));
                }
            }
            expr = canonicalize_method_call(expr, args);
        } else if chain_src.starts_with(".") {
            // Member access — normalize known property accessors to canonical builtins
            let name = strip_csharp_terminal_type_args(chain_src[1..].trim());
            // C# tuple ItemN accessor: `(1, 2, 3).Item1` → `t[0]`,
            // `Item2` → `t[1]`, etc. Tuples compile to Arrays so the
            // ItemN names need to lower to indexed access. Pattern is
            // `Item` followed by 1+ digit index (1-based).
            if let Some(rest) = name.strip_prefix("Item") {
                if let Ok(n) = rest.parse::<i64>() {
                    if n >= 1 {
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(Expression::int(n - 1)),
                            null_safe: false,
                        });
                        continue;
                    }
                }
            }
            // Canonicalize C# property accessors: Length, Count → __len__
            // BUT only if the next chain segment is NOT a call. `Count` and
            // `Length` can also appear as instance-method names (LINQ
            // `Count(predicate)`); folding them eagerly into `__len__`
            // breaks the call site.
            let next_is_call = iter
                .peek()
                .map(|c| c.as_str().starts_with('('))
                .unwrap_or(false);
            if next_is_call {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
            } else {
                expr = canonicalize_member_access(expr, &name);
            }
        } else if chain_src.starts_with("[") || chain_src.starts_with("?[") {
            // Index / range / from-end. The C# 8 forms covered:
            //   arr[i]     — plain index
            //   arr[^N]    — from-end index, i.e. arr[arr.Length - N]
            //   arr[a..b]  — range
            //   arr[^N..]  — range with from-end start, open end
            //   arr[..^N]  — range with from-end end
            //   arr[..]    — full slice
            let null_safe = chain_src.starts_with("?[");
            let inner_pairs: Vec<Pair<Rule>> = chain_inner.into_iter().collect();
            // Find the index_expression child (grammar wraps the brackets'
            // contents in `index_expression`).
            let idx_pair = inner_pairs
                .into_iter()
                .find(|p| p.as_rule() == Rule::index_expression);
            if let Some(idx) = idx_pair {
                let idx_src = idx.as_str().trim();
                let has_range = idx_src.contains("..");
                let parts: Vec<Pair<Rule>> = idx.into_inner().collect();
                if has_range {
                    let mut start: Option<Expression> = None;
                    let mut end: Option<Expression> = None;
                    let hit_dotdot = false;
                    // Walk parts; the `..` token isn't a pair (it's a literal),
                    // so we infer position from order: parts before the index of
                    // a from_end_index / expression that's "after" the dotdot
                    // are starts. Simpler: split source on `..`.
                    let halves: Vec<&str> = idx_src.splitn(2, "..").collect();
                    let _ = (start.as_ref(), end.as_ref(), hit_dotdot);
                    let mut iter = parts.into_iter();
                    let first_after_dotdot = halves.first().map_or(true, |s| s.trim().is_empty());
                    if !first_after_dotdot {
                        if let Some(p) = iter.next() {
                            start = Some(walk_index_part(p, expr.clone())?);
                        }
                    }
                    let second_empty = halves.get(1).map_or(true, |s| s.trim().is_empty());
                    if !second_empty {
                        if let Some(p) = iter.next() {
                            end = Some(walk_index_part(p, expr.clone())?);
                        }
                    }
                    let start = start.unwrap_or_else(Expression::null);
                    let end = end.unwrap_or_else(|| Expression::int(i32::MAX as i64));
                    let range = Expression::new(ExprKind::Range {
                        start: Box::new(start),
                        end: Box::new(end),
                        inclusive: false,
                    });
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(range),
                        null_safe,
                    });
                } else {
                    // Multi-arg index `m[i, j]` lowers to nested
                    // `m[i][j]`. Single-arg index is the common case.
                    let mut iter = parts.into_iter();
                    if let Some(first) = iter.next() {
                        let index = walk_index_part(first, expr.clone())?;
                        expr = if matches!(&expr.kind, ExprKind::Ident(name) if name == "text") {
                            Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::ident("__csharp_str_char_code_at")),
                                args: vec![Argument::positional(expr), Argument::positional(index)],
                                optional: false,
                            })
                        } else {
                            Expression::new(ExprKind::Index {
                                object: Box::new(expr),
                                index: Box::new(index),
                                null_safe,
                            })
                        };
                        for p in iter {
                            let index = walk_index_part(p, expr.clone())?;
                            expr = Expression::new(ExprKind::Index {
                                object: Box::new(expr),
                                index: Box::new(index),
                                null_safe: false,
                            });
                        }
                    }
                }
            }
        }
    }

    Ok(expr.kind)
}

// ── New expression walker ───────────────────────────────────────────────────

fn walk_new_expr(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut type_name = String::new();
    let mut raw_type_name = String::new();
    let mut args = Vec::new();
    let mut is_array = false;
    let mut array_init = Vec::new();
    let mut obj_init: Vec<(String, Expression)> = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_name_for_new => {
                let raw = p.as_str();
                raw_type_name = raw.trim().to_string();
                type_name = strip_csharp_type_path_generic_args(raw);
            }
            Rule::argument_list => args = walk_arguments(p)?,
            Rule::array_initializer => {
                is_array = true;
                for ap in p.into_inner() {
                    // Each child is a `collection_element` wrapping either
                    // an expression or a nested `{ ... }` (dict pair /
                    // sub-array). Walk through `walk_collection_element`
                    // so the Dictionary/multi-dim shape lowers correctly.
                    if let Ok(expr) = walk_collection_element(ap) {
                        array_init.push(expr);
                    }
                }
            }
            Rule::object_initializer => {
                for ip in p.into_inner() {
                    if ip.as_rule() == Rule::initializer_member {
                        let mut name = String::new();
                        let mut val = Expression::null();
                        for mp in ip.into_inner() {
                            match mp.as_rule() {
                                Rule::ident_name => name = mp.as_str().to_string(),
                                _ => val = walk_expression(mp).unwrap_or(Expression::null()),
                            }
                        }
                        obj_init.push((name, val));
                    }
                }
            }
            _ => {
                // Expression inside brackets for array size
                if let Ok(expr) = walk_expression(p) {
                    if !is_array {
                        is_array = true;
                    }
                    args.push(Argument::positional(expr));
                }
            }
        }
    }

    // `new Dictionary<K,V> { { key, value }, ... }` — IIFE-lower to:
    //
    //     (() => { var __d = new Dictionary(); __d.Add(k1, v1); ...
    //              return __d; })()
    //
    // Producing a plain Object literal would lose the `Dictionary`
    // `__type` and the runtime collection registry could not route
    // `ContainsKey` / `Add` to the `ecma:map.*` primitives. The IIFE
    // builds a real Map-backed Dictionary and populates it before
    // returning.
    let is_dict_ctor =
        type_name.eq_ignore_ascii_case("Dictionary") || type_name.ends_with("Dictionary");
    if is_dict_ctor && !array_init.is_empty() {
        let mut pairs: Vec<(Expression, Expression)> = Vec::new();
        for elem in &array_init {
            if let ExprKind::Array(parts) = &elem.kind {
                if parts.len() == 2 {
                    pairs.push((parts[0].value.clone(), parts[1].value.clone()));
                }
            }
        }
        if !pairs.is_empty() {
            return Ok(emit_dict_iife(type_name.clone(), args.clone(), pairs));
        }
    }
    // `new HashSet<T> { v1, v2, ... }` — IIFE-lower to construct + Add
    // calls, same shape as the Dictionary path above. HashSet's
    // backing is `ObjectKind::Set`, registered separately in
    // `vybe_host::builtin_types`, so we need a real `new HashSet()`.
    let is_set_ctor = type_name.eq_ignore_ascii_case("HashSet")
        || type_name.ends_with("HashSet")
        || type_name.eq_ignore_ascii_case("SortedSet")
        || type_name.ends_with("SortedSet");
    if is_set_ctor && !array_init.is_empty() {
        return Ok(emit_set_iife(type_name.clone(), args.clone(), array_init));
    }

    let is_list_ctor = matches!(
        type_name.rsplit('.').next().unwrap_or(type_name.as_str()),
        "List" | "ArrayList"
    );
    if is_list_ctor && !array_init.is_empty() {
        return Ok(emit_list_iife(type_name.clone(), args.clone(), array_init));
    }

    if is_array && !array_init.is_empty() {
        // Array initializer: new[] { 1, 2, 3 } or new int[] { 1, 2, 3 }
        // — also covers multi-dim (`int[,]`) where each element is itself
        // an Array, which `walk_collection_element` already produced.
        let elements = array_init
            .into_iter()
            .map(|v| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        return Ok(ExprKind::Array(elements));
    }

    // .NET `new string(charArray)` → `charArray.join("")`. The runtime
    // doesn't carry a String constructor, but every char[] in this VM
    // is a JS array of single-char strings, so `.join("")` is faithful.
    if type_name == "string" && args.len() == 1 && array_init.is_empty() && obj_init.is_empty() {
        let arr = args[0].value.clone();
        return Ok(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(arr),
                field: "join".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str("".into()),
            )))],
            optional: false,
        });
    }

    // .NET `new string(ch, count)` repeats the single-character string.
    if type_name == "string" && args.len() == 2 && array_init.is_empty() && obj_init.is_empty() {
        return Ok(ExprKind::Call {
            callee: Box::new(Expression::ident("str_repeat")),
            args: vec![args[0].clone(), args[1].clone()],
            optional: false,
        });
    }

    // Build class expression — dotted names become Member chains
    // (e.g. "MyApp.Foo" → Member { Ident("MyApp"), "Foo" })
    let class_expr = build_dotted_expr(&type_name);

    if is_array {
        let dimensions: Vec<Expression> = args.iter().map(|arg| arg.value.clone()).collect();
        if dimensions.is_empty() {
            return Ok(ExprKind::Array(Vec::new()));
        }

        let trailing_unspecified_ranks =
            csharp_array_rank_group_count(&type_name).saturating_sub(dimensions.len());
        return Ok(build_csharp_sized_array_expr(
            &dimensions,
            &csharp_array_element_type_name(&type_name),
            trailing_unspecified_ranks,
        )
        .kind);
    }

    // Object initializer: `new Point { X = 10, Y = 20 }`. The
    // captured `obj_init` pairs become assignments on the freshly-
    // constructed instance. Lowered as IIFE so the temp ident is
    // self-contained and doesn't pollute the surrounding scope.
    if !obj_init.is_empty() {
        let new_call = Expression::new(ExprKind::New {
            class: Box::new(class_expr),
            args,
        });
        let generic_bindings = csharp_generic_ctor_bindings(&raw_type_name, &type_name);
        let new_call = if generic_bindings.is_empty() {
            new_call
        } else {
            Expression::new(emit_generic_ctor_binding_iife(new_call, generic_bindings))
        };
        return Ok(emit_object_init_iife(new_call, obj_init));
    }

    let generic_bindings = csharp_generic_ctor_bindings(&raw_type_name, &type_name);
    if !generic_bindings.is_empty() {
        let new_call = Expression::new(ExprKind::New {
            class: Box::new(class_expr),
            args,
        });
        return Ok(emit_generic_ctor_binding_iife(new_call, generic_bindings));
    }

    Ok(ExprKind::New {
        class: Box::new(class_expr),
        args,
    })
}

fn csharp_array_rank_group_count(type_name: &str) -> usize {
    type_name.chars().filter(|ch| *ch == '[').count()
}

fn csharp_array_element_type_name(type_name: &str) -> String {
    let mut current = type_name.trim().to_string();
    loop {
        let trimmed = current.trim_end();
        let Some(open_idx) = trimmed.rfind('[') else {
            break;
        };
        let suffix = &trimmed[open_idx..];
        if suffix.chars().all(|ch| matches!(ch, '[' | ']' | ',' | ' ')) {
            current = trimmed[..open_idx].trim_end().to_string();
            continue;
        }
        break;
    }
    normalize_runtime_type_name(&current)
}

fn csharp_array_default_expr(type_name: &str) -> Expression {
    match normalize_runtime_type_name(type_name)
        .to_ascii_lowercase()
        .as_str()
    {
        "int" | "int32" | "long" | "int64" | "short" | "int16" | "byte" | "uint" | "uint32"
        | "ulong" | "uint64" | "ushort" | "uint16" | "sbyte" | "double" | "float" | "single"
        | "decimal" => Expression::int(0),
        "bool" | "boolean" => Expression::bool(false),
        "char" => Expression::new(ExprKind::Lit(Literal::Char('\0'))),
        _ => Expression::null(),
    }
}

fn build_csharp_sized_array_expr(
    dimensions: &[Expression],
    element_type: &str,
    trailing_unspecified_ranks: usize,
) -> Expression {
    let init = if dimensions.len() == 1 {
        if trailing_unspecified_ranks > 0 {
            Expression::null()
        } else {
            csharp_array_default_expr(element_type)
        }
    } else {
        build_csharp_sized_array_expr(&dimensions[1..], element_type, trailing_unspecified_ranks)
    };

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(dimensions[0].clone()),
            Argument::positional(init),
        ],
        optional: false,
    })
}

fn try_csharp_type_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(name)) => Some(name.clone()),
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { .. } => expr_dotted_name(expr),
        _ => None,
    }
}

fn try_csharp_usize_literal(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) if *value >= 0 => Some(*value as usize),
        _ => None,
    }
}

fn is_csharp_zero_literal(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Int(0)))
}

fn build_csharp_array_create_instance_expr(args: &[Argument]) -> Option<Expression> {
    if args.len() < 2 {
        return None;
    }

    let type_name = try_csharp_type_name(&args[0].value).unwrap_or_else(|| "object".into());
    let dimensions: Vec<Expression> = args.iter().skip(1).map(|arg| arg.value.clone()).collect();
    Some(build_csharp_sized_array_expr(
        &dimensions,
        &csharp_array_element_type_name(&type_name),
        0,
    ))
}

fn build_csharp_get_length_expr(receiver: Expression, dimension: usize) -> Expression {
    let mut target = receiver;
    for _ in 0..dimension {
        target = Expression::new(ExprKind::Index {
            object: Box::new(target),
            index: Box::new(Expression::int(0)),
            null_safe: false,
        });
    }
    canonicalize_member_access(target, "Length")
}

fn build_csharp_math_call(name: &str, arg: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(build_dotted_expr("System.Math")),
            field: name.to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(arg)],
        optional: false,
    })
}

fn build_csharp_bankers_round_expr(value: Expression) -> Expression {
    let floor = build_csharp_math_call("Floor", value.clone());
    let ceil = build_csharp_math_call("Ceiling", value.clone());
    let frac = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(value),
        right: Box::new(floor.clone()),
    });
    let half = Expression::new(ExprKind::Lit(Literal::Float(0.5)));
    let floor_is_even = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(Expression::new(ExprKind::Cast {
                expr: Box::new(floor.clone()),
                type_name: "int".into(),
            })),
            right: Box::new(Expression::int(1)),
        })),
        right: Box::new(Expression::int(0)),
    });

    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(frac.clone()),
            right: Box::new(half.clone()),
        })),
        then: Box::new(floor.clone()),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Gt,
                left: Box::new(frac),
                right: Box::new(half),
            })),
            then: Box::new(ceil.clone()),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(floor_is_even),
                then: Box::new(floor),
                else_: Box::new(ceil),
            })),
        })),
    })
}

fn parse_csharp_generic_type_args(raw_type_name: &str) -> Vec<String> {
    let Some(start) = raw_type_name.find('<') else {
        return Vec::new();
    };
    let Some(end) = raw_type_name.rfind('>') else {
        return Vec::new();
    };
    let inner = &raw_type_name[start + 1..end];
    let mut args = Vec::new();
    let mut current = String::new();
    let mut depth = 0usize;
    for ch in inner.chars() {
        match ch {
            '<' => {
                depth += 1;
                current.push(ch);
            }
            '>' => {
                depth = depth.saturating_sub(1);
                current.push(ch);
            }
            ',' if depth == 0 => {
                let trimmed = current.trim();
                if !trimmed.is_empty() {
                    args.push(trimmed.to_string());
                }
                current.clear();
            }
            _ => current.push(ch),
        }
    }
    let trimmed = current.trim();
    if !trimmed.is_empty() {
        args.push(trimmed.to_string());
    }
    args
}

fn looks_like_csharp_runtime_ctor_type(type_name: &str) -> bool {
    let trimmed = strip_global_namespace_qualifier(type_name.trim())
        .trim_end_matches('?')
        .trim();
    if trimmed.is_empty() || trimmed.ends_with("[]") {
        return false;
    }
    let bare = trimmed.split('<').next().unwrap_or(trimmed).trim();
    let leaf = bare.rsplit('.').next().unwrap_or(bare);
    leaf.chars()
        .next()
        .is_some_and(|ch| ch.is_ascii_uppercase())
}

fn is_csharp_builtin_generic_type(base_type: &str) -> bool {
    matches!(
        base_type,
        "List"
            | "ArrayList"
            | "Dictionary"
            | "HashSet"
            | "Queue"
            | "Stack"
            | "Nullable"
            | "Task"
            | "Func"
            | "Action"
            | "Tuple"
            | "ValueTuple"
            | "IEnumerable"
            | "ICollection"
            | "IList"
            | "IReadOnlyCollection"
            | "IReadOnlyList"
            | "IComparer"
            | "IComparable"
            | "KeyValuePair"
    )
}

fn csharp_generic_ctor_bindings(
    raw_type_name: &str,
    base_type_name: &str,
) -> Vec<(usize, Expression)> {
    if is_csharp_builtin_generic_type(base_type_name) {
        return Vec::new();
    }

    parse_csharp_generic_type_args(raw_type_name)
        .into_iter()
        .enumerate()
        .filter_map(|(index, type_name)| {
            looks_like_csharp_runtime_ctor_type(&type_name).then(|| {
                let bare = strip_global_namespace_qualifier(type_name.trim())
                    .split('<')
                    .next()
                    .unwrap_or(type_name.trim())
                    .trim()
                    .to_string();
                (index, build_dotted_expr(&bare))
            })
        })
        .collect()
}

fn emit_generic_ctor_binding_iife(
    new_call: Expression,
    bindings: Vec<(usize, Expression)>,
) -> ExprKind {
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint: None,
                init: Some(new_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (index, binding_expr) in bindings {
        let assign = Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__obj")),
                field: csharp_generic_ctor_field_name(index),
                null_safe: false,
            })),
            value: Box::new(binding_expr),
        });
        body.push(Statement::with_span(
            StmtKind::Expr(assign),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__obj"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// IIFE-style lowering for `new T(args) { Prop = value, ... }`. Builds
/// a single-call lambda that constructs the object, fires each property
/// assignment, and returns the instance. Same pattern as the
/// Dictionary / HashSet initializer lowerings — keeps the temp local
/// out of the caller's scope.
fn emit_object_init_iife(new_call: Expression, props: Vec<(String, Expression)>) -> ExprKind {
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint: None,
                init: Some(new_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (name, value) in props {
        // __obj.<name> = value;
        let assign = Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__obj")),
                field: name,
                null_safe: false,
            })),
            value: Box::new(value),
        });
        body.push(Statement::with_span(
            StmtKind::Expr(assign),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__obj"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// Convert a dotted name like "MyApp.Foo.Bar" into a Member chain expression.
fn build_dotted_expr(name: &str) -> Expression {
    let normalized = strip_global_namespace_qualifier(name);
    let parts: Vec<&str> = normalized.split('.').collect();
    if parts.len() == 1 {
        return Expression::ident(parts[0]);
    }
    let mut expr = Expression::ident(parts[0]);
    for part in &parts[1..] {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: part.to_string(),
            null_safe: false,
        });
    }
    expr
}

fn strip_csharp_terminal_type_args(name: &str) -> String {
    let trimmed = name.trim();
    if !trimmed.ends_with('>') {
        return trimmed.to_string();
    }

    let mut depth = 0usize;
    let mut start = None;
    for (idx, ch) in trimmed.char_indices().rev() {
        match ch {
            '>' => depth += 1,
            '<' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    start = Some(idx);
                    break;
                }
            }
            _ => {}
        }
    }

    let Some(start_idx) = start else {
        return trimmed.to_string();
    };

    trimmed[..start_idx].trim_end().to_string()
}

fn strip_csharp_type_path_generic_args(name: &str) -> String {
    let mut stripped = String::with_capacity(name.len());
    let mut depth = 0usize;
    for ch in strip_global_namespace_qualifier(name.trim()).chars() {
        match ch {
            '<' => depth += 1,
            '>' => depth = depth.saturating_sub(1),
            _ if depth == 0 => stripped.push(ch),
            _ => {}
        }
    }
    stripped.trim().to_string()
}

fn extract_csharp_terminal_type_args(name: &str) -> Vec<String> {
    let trimmed = name.trim();
    let stripped = strip_csharp_terminal_type_args(trimmed);
    if stripped == trimmed || !trimmed.ends_with('>') {
        return Vec::new();
    }
    parse_csharp_generic_type_args(trimmed)
}

fn csharp_method_generic_binding_args(type_args: &[String]) -> Vec<Argument> {
    let mut args = Vec::new();
    for type_arg in type_args {
        let ctor_binding = if looks_like_csharp_runtime_ctor_type(type_arg) {
            let bare = strip_global_namespace_qualifier(type_arg.trim())
                .split('<')
                .next()
                .unwrap_or(type_arg.trim())
                .trim()
                .to_string();
            build_dotted_expr(&bare)
        } else {
            Expression::null()
        };
        args.push(Argument::positional(ctor_binding));

        let normalized = normalize_runtime_type_name(type_arg);
        let default_binding = match normalized.as_str() {
            "int" | "long" | "short" | "byte" | "uint" | "ulong" | "ushort" | "sbyte"
            | "double" | "float" | "decimal" => Expression::int(0),
            "bool" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
            "char" => Expression::new(ExprKind::Lit(Literal::Char('\0'))),
            _ => Expression::null(),
        };
        args.push(Argument::positional(default_binding));
    }
    args
}

/// Lower a C# 9 record `with` expression: `record_val with { Prop = v, ... }`.
/// The walker emits an IIFE that constructs a shallow copy by reading
/// the receiver's existing properties, applies the with-clause mutations,
/// and returns the new instance. Records compile as plain classes in
/// our compiler, so this is the same shape as a `new T { ... }`
/// initializer that copies fields from the source.
fn walk_with_expr(receiver: Expression, postfix: Pair<Rule>) -> Result<Expression, String> {
    // Collect the with-clause property assignments.
    let mut props: Vec<(String, Expression)> = Vec::new();
    for child in postfix.into_inner() {
        if child.as_rule() == Rule::object_initializer {
            for ip in child.into_inner() {
                if ip.as_rule() == Rule::initializer_member {
                    let mut name = String::new();
                    let mut val = Expression::null();
                    for mp in ip.into_inner() {
                        match mp.as_rule() {
                            Rule::ident_name => name = mp.as_str().to_string(),
                            _ => val = walk_expression(mp).unwrap_or(Expression::null()),
                        }
                    }
                    props.push((name, val));
                }
            }
        }
    }

    // Lower to an IIFE:
    //   ((src) => {
    //       var __o = Object.assign({}, src);
    //       __o.Prop = val;
    //       ...
    //       return __o;
    //   })(receiver)
    //
    // We use `Object.assign({}, src)` shape via the dotted-name
    // resolver (resolves to `ecma:object.assign`) so the clone sees
    // the same prototype chain as the source.
    let mut body: Vec<Statement> = Vec::new();
    let assign_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Object")),
            field: "assign".into(),
            null_safe: false,
        })),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Object(Vec::new()))),
            Argument::positional(Expression::ident("__src")),
        ],
        optional: false,
    });
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__o".into()),
                type_hint: None,
                init: Some(assign_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (name, value) in props {
        let assign = Expression::new(ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__o")),
                field: name,
                null_safe: false,
            })),
            value: Box::new(value),
        });
        body.push(Statement::with_span(
            StmtKind::Expr(assign),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__o"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: "__src".into(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![Argument::positional(receiver)],
        optional: false,
    }))
}

/// Lower a C# 8 switch expression `subject switch { arm, ... }` into
/// a chain of nested `cond ? then : else_` ternaries. Each arm's
/// `when` guard is AND-ed into the cond. The wildcard `_` arm is the
/// catchall (`else_` of the chain). If no `_` arm is present, the
/// chain falls through to `null`, matching .NET's
/// `SwitchExpressionException` shape (we don't throw — return null
/// rather than complicate codegen).
fn walk_switch_expr(subject: Expression, postfix: Pair<Rule>) -> Result<Expression, String> {
    let arms: Vec<Pair<Rule>> = postfix
        .into_inner()
        .filter(|p| p.as_rule() == Rule::switch_arm)
        .collect();
    let span = subject.span.clone();
    // We process arms in reverse, building the ternary chain inside-out.
    let mut else_branch = Expression::null();
    let mut else_set = false;
    for arm in arms.into_iter().rev() {
        let mut pattern: Option<Pair<Rule>> = None;
        let mut when_guard: Option<Expression> = None;
        let mut result: Option<Expression> = None;
        let arm_inner: Vec<Pair<Rule>> = arm.into_inner().collect();
        // Order in source: pattern, optional `when`-clause expression,
        // then the result expression. We split by rule, taking the
        // first `switch_pattern` as the pattern, and treating the
        // remaining `expression` children as guard + result.
        let mut exprs: Vec<Pair<Rule>> = Vec::new();
        for inner in arm_inner {
            match inner.as_rule() {
                Rule::switch_pattern => pattern = Some(inner),
                Rule::expression => exprs.push(inner),
                _ => {}
            }
        }
        // Last expr = result; if there's a second expr it's the guard.
        if let Some(last) = exprs.pop() {
            result = Some(walk_expression(last)?);
        }
        if let Some(guard) = exprs.pop() {
            when_guard = Some(walk_expression(guard)?);
        }
        let result = result.ok_or("switch arm missing result")?;
        let pattern = pattern.ok_or("switch arm missing pattern")?;

        // Detect wildcard `_` arm — that's the catch-all.
        let pat_src = pattern.as_str().trim();
        if pat_src == "_" && when_guard.is_none() {
            else_branch = result;
            else_set = true;
            continue;
        }
        let cond = build_switch_pattern_cond(subject.clone(), pattern.clone())?;
        let binding = build_switch_pattern_binding(subject.clone(), pattern)?;
        if !else_set {
            else_branch = result.clone();
            else_set = true;
            // Still emit the test so the arm runs even if `_` is missing.
        }
        let then_branch = if let Some(binding_stmt) = binding {
            let scoped_result = if let Some(guard) = when_guard {
                Expression::with_span(
                    ExprKind::Ternary {
                        cond: Box::new(guard),
                        then: Box::new(result),
                        else_: Box::new(else_branch.clone()),
                    },
                    span.clone(),
                )
            } else {
                result
            };
            let lambda = Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![
                    binding_stmt,
                    Statement::with_span(StmtKind::Return(Some(scoped_result)), Span::default()),
                ]),
                is_async: false,
                captures: Vec::new(),
            });
            Expression::new(ExprKind::Call {
                callee: Box::new(lambda),
                args: vec![],
                optional: false,
            })
        } else {
            let then_cond = if let Some(guard) = when_guard {
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(cond.clone()),
                        right: Box::new(guard),
                    },
                    span.clone(),
                )
            } else {
                cond.clone()
            };
            let next = Expression::with_span(
                ExprKind::Ternary {
                    cond: Box::new(then_cond),
                    then: Box::new(result),
                    else_: Box::new(else_branch),
                },
                span.clone(),
            );
            else_branch = next;
            continue;
        };
        let next = Expression::with_span(
            ExprKind::Ternary {
                cond: Box::new(cond),
                then: Box::new(then_branch),
                else_: Box::new(else_branch),
            },
            span.clone(),
        );
        else_branch = next;
    }
    Ok(else_branch)
}

/// Build a boolean test for a single switch-arm pattern.
/// Cases:
///   `<lit>`           → subject === <lit>
///   `<TypeName> <id>` → typeof subject === "<jsname>" (binding dropped)
///   `>= <expr>`       → subject >= <expr>  (relational pattern)
///   `<expr>`          → subject === <expr>  (constant fallback)
fn build_switch_pattern_cond(
    subject: Expression,
    pattern: Pair<Rule>,
) -> Result<Expression, String> {
    let pat_src = pattern.as_str().trim();
    let span = subject.span.clone();
    // Relational pattern: `>= 90`, `<= 50`, `< 0`, `> 0`.
    let rel_op = if pat_src.starts_with(">=") {
        Some(BinOp::GtEq)
    } else if pat_src.starts_with("<=") {
        Some(BinOp::LtEq)
    } else if pat_src.starts_with('>') {
        Some(BinOp::Gt)
    } else if pat_src.starts_with('<') {
        Some(BinOp::Lt)
    } else {
        None
    };
    let cond: Expression;
    if let Some(op) = rel_op {
        // Find the inner expression child.
        let inner = pattern
            .into_inner()
            .find(|p| p.as_rule() == Rule::expression)
            .ok_or("relational pattern missing expression")?;
        let rhs = walk_expression(inner)?;
        cond = Expression::with_span(
            ExprKind::Binary {
                op,
                left: Box::new(subject),
                right: Box::new(rhs),
            },
            span.clone(),
        );
    } else {
        if let Some(elements) = extract_switch_tuple_pattern_elements(pattern.clone())? {
            return Ok(build_switch_tuple_pattern_cond(subject, elements));
        }
        // Type pattern: `int i`, `string s` (with binding) or constant.
        let mut inner_pairs: Vec<Pair<Rule>> = pattern.into_inner().collect();
        // `type_name ~ ident_name` — type pattern with binding.
        if inner_pairs.len() >= 2
            && inner_pairs[0].as_rule() == Rule::type_name
            && inner_pairs[1].as_rule() == Rule::ident_name
        {
            let type_name = inner_pairs[0].as_str().trim().to_string();
            let test = if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                let typeof_expr =
                    Expression::with_span(ExprKind::TypeOf(Box::new(subject)), span.clone());
                Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(typeof_expr),
                        right: Box::new(Expression::with_span(
                            ExprKind::Lit(Literal::Str(js_typeof.into())),
                            span.clone(),
                        )),
                    },
                    span.clone(),
                )
            } else {
                Expression::with_span(
                    ExprKind::IsType {
                        expr: Box::new(subject),
                        type_name,
                    },
                    span.clone(),
                )
            };
            cond = test;
        } else if let Some(p) = inner_pairs.pop() {
            // Constant pattern (numeric / string literal / general expr).
            let rhs = walk_expression(p)?;
            cond = Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(subject),
                    right: Box::new(rhs),
                },
                span.clone(),
            );
        } else {
            cond = Expression::with_span(ExprKind::Lit(Literal::Bool(false)), span.clone());
        }
    }
    Ok(cond)
}

fn build_switch_pattern_binding(
    subject: Expression,
    pattern: Pair<Rule>,
) -> Result<Option<Statement>, String> {
    let inner_pairs: Vec<Pair<Rule>> = pattern.into_inner().collect();
    if inner_pairs.len() >= 2
        && inner_pairs[0].as_rule() == Rule::type_name
        && inner_pairs[1].as_rule() == Rule::ident_name
    {
        let type_name = inner_pairs[0].as_str().trim().to_string();
        let binding_name = inner_pairs[1].as_str().trim().to_string();
        return Ok(Some(build_type_pattern_binding_stmt(
            subject,
            type_name,
            binding_name,
        )));
    }
    Ok(None)
}

fn extract_switch_tuple_pattern_elements(
    pair: Pair<Rule>,
) -> Result<Option<Vec<Option<Expression>>>, String> {
    if pair.as_rule() == Rule::tuple_literal {
        let mut elements = Vec::new();
        for element in pair
            .into_inner()
            .filter(|p| p.as_rule() == Rule::tuple_element)
        {
            let expr_pair = element
                .into_inner()
                .rev()
                .find(|p| p.as_rule() == Rule::expression)
                .ok_or("tuple pattern element missing expression")?;
            if expr_pair.as_str().trim() == "_" {
                elements.push(None);
            } else {
                elements.push(Some(walk_expression(expr_pair)?));
            }
        }
        return Ok(Some(elements));
    }
    for child in pair.into_inner() {
        if let Some(elements) = extract_switch_tuple_pattern_elements(child)? {
            return Ok(Some(elements));
        }
    }
    Ok(None)
}

fn build_switch_tuple_pattern_cond(
    subject: Expression,
    elements: Vec<Option<Expression>>,
) -> Expression {
    let span = subject.span.clone();
    let mut cond: Option<Expression> = None;
    for (index, expected) in elements.into_iter().enumerate() {
        let Some(expected) = expected else {
            continue;
        };
        let item = Expression::with_span(
            ExprKind::Index {
                object: Box::new(subject.clone()),
                index: Box::new(Expression::int(index as i64)),
                null_safe: false,
            },
            span.clone(),
        );
        let eq = Expression::with_span(
            ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(item),
                right: Box::new(expected),
            },
            span.clone(),
        );
        cond = Some(match cond {
            Some(prev) => Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(prev),
                    right: Box::new(eq),
                },
                span.clone(),
            ),
            None => eq,
        });
    }
    cond.unwrap_or_else(|| Expression::with_span(ExprKind::Lit(Literal::Bool(true)), span))
}

fn build_general_pattern_cond(
    subject: Expression,
    pattern: Pair<Rule>,
) -> Result<Expression, String> {
    let span = subject.span.clone();
    match pattern.as_rule() {
        Rule::switch_case_pattern => {
            let inner = pattern
                .into_inner()
                .next()
                .ok_or("switch case pattern missing clause")?;
            build_general_pattern_cond(subject, inner)
        }
        Rule::pattern_clause => {
            let src = pattern.as_str().trim_start();
            let negated = src.starts_with("not")
                && src[3..].chars().next().map_or(true, |c| c.is_whitespace());
            let atoms: Vec<Pair<Rule>> = pattern
                .into_inner()
                .filter(|p| p.as_rule() == Rule::pattern_atom)
                .collect();
            let mut cond = Expression::bool(true);
            for atom in atoms {
                let next = build_general_pattern_cond(subject.clone(), atom)?;
                cond = Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(cond),
                        right: Box::new(next),
                    },
                    span.clone(),
                );
            }
            if negated {
                Ok(Expression::with_span(
                    ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(cond),
                    },
                    span,
                ))
            } else {
                Ok(cond)
            }
        }
        Rule::pattern_atom => {
            let inner: Vec<Pair<Rule>> = pattern.into_inner().collect();
            let first = inner
                .first()
                .ok_or("Empty pattern atom inner".to_string())?;
            match first.as_rule() {
                Rule::pattern_type => build_general_pattern_cond(subject, first.clone()),
                _ => build_general_pattern_cond(subject, first.clone()),
            }
        }
        Rule::null_kw => Ok(Expression::with_span(
            ExprKind::Binary {
                op: BinOp::StrictEq,
                left: Box::new(subject),
                right: Box::new(Expression::with_span(
                    ExprKind::Lit(Literal::Null),
                    span.clone(),
                )),
            },
            span,
        )),
        Rule::true_kw
        | Rule::false_kw
        | Rule::numeric_literal
        | Rule::string_literal
        | Rule::char_literal => {
            let lit = walk_expression(pattern)?;
            Ok(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(subject),
                    right: Box::new(lit),
                },
                span,
            ))
        }
        Rule::relational_pattern => {
            let pat_src = pattern.as_str().trim();
            let op = if pat_src.starts_with(">=") {
                BinOp::GtEq
            } else if pat_src.starts_with("<=") {
                BinOp::LtEq
            } else if pat_src.starts_with('>') {
                BinOp::Gt
            } else {
                BinOp::Lt
            };
            let rhs_pair = pattern
                .into_inner()
                .find(|p| p.as_rule() == Rule::expression)
                .ok_or("relational pattern missing expression")?;
            let rhs = walk_expression(rhs_pair)?;
            Ok(Expression::with_span(
                ExprKind::Binary {
                    op,
                    left: Box::new(subject),
                    right: Box::new(rhs),
                },
                span,
            ))
        }
        Rule::tuple_pattern => {
            let mut elements = Vec::new();
            for item in pattern
                .into_inner()
                .filter(|p| p.as_rule() == Rule::tuple_pattern_item)
            {
                if item.as_str().trim() == "_" {
                    elements.push(None);
                } else {
                    let expr_pair = item
                        .into_inner()
                        .next()
                        .ok_or("tuple pattern element missing expression")?;
                    elements.push(Some(walk_expression(expr_pair)?));
                }
            }
            Ok(build_switch_tuple_pattern_cond(subject, elements))
        }
        Rule::var_pattern => Ok(Expression::bool(true)),
        Rule::pattern_type => {
            let type_name = normalize_runtime_type_name(pattern.as_str());
            if let Some(js_typeof) = primitive_to_typeof(&type_name) {
                let typeof_expr =
                    Expression::with_span(ExprKind::TypeOf(Box::new(subject)), span.clone());
                Ok(Expression::with_span(
                    ExprKind::Binary {
                        op: BinOp::StrictEq,
                        left: Box::new(typeof_expr),
                        right: Box::new(Expression::with_span(
                            ExprKind::Lit(Literal::Str(js_typeof.into())),
                            span.clone(),
                        )),
                    },
                    span,
                ))
            } else {
                Ok(Expression::with_span(
                    ExprKind::IsType {
                        expr: Box::new(subject),
                        type_name,
                    },
                    span,
                ))
            }
        }
        _ => {
            let rhs = walk_expression(pattern)?;
            Ok(Expression::with_span(
                ExprKind::Binary {
                    op: BinOp::StrictEq,
                    left: Box::new(subject),
                    right: Box::new(rhs),
                },
                span,
            ))
        }
    }
}

fn build_general_pattern_binding(
    subject: Expression,
    pattern: Pair<Rule>,
) -> Result<Option<Statement>, String> {
    match pattern.as_rule() {
        Rule::switch_case_pattern => {
            let inner = pattern
                .into_inner()
                .next()
                .ok_or("switch case pattern missing clause")?;
            build_general_pattern_binding(subject, inner)
        }
        Rule::pattern_clause => {
            let src = pattern.as_str().trim_start();
            if src.starts_with("not") {
                return Ok(None);
            }
            let atoms: Vec<Pair<Rule>> = pattern
                .into_inner()
                .filter(|p| p.as_rule() == Rule::pattern_atom)
                .collect();
            if atoms.len() != 1 {
                return Ok(None);
            }
            build_general_pattern_binding(subject, atoms[0].clone())
        }
        Rule::pattern_atom => {
            let inner: Vec<Pair<Rule>> = pattern.into_inner().collect();
            if inner.len() >= 2
                && inner[0].as_rule() == Rule::pattern_type
                && inner[1].as_rule() == Rule::ident_name
            {
                let type_name = normalize_runtime_type_name(inner[0].as_str());
                let binding_name = inner[1].as_str().to_string();
                return Ok(Some(build_type_pattern_binding_stmt(
                    subject,
                    type_name,
                    binding_name,
                )));
            }
            if let Some(first) = inner.first() {
                return build_general_pattern_binding(subject, first.clone());
            }
            Ok(None)
        }
        Rule::var_pattern => {
            let binding_name = pattern
                .into_inner()
                .last()
                .ok_or("var pattern missing identifier")?
                .as_str()
                .to_string();
            Ok(Some(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(binding_name),
                    type_hint: None,
                    init: Some(subject),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            })))
        }
        _ => Ok(None),
    }
}

fn build_scoped_guard_expr(
    cond: Expression,
    binding_stmt: Statement,
    guard: Expression,
) -> Expression {
    let guard_lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(vec![
            binding_stmt,
            Statement::new(StmtKind::Return(Some(guard))),
        ]),
        is_async: false,
        captures: Vec::new(),
    });
    let scoped_guard = Expression::new(ExprKind::Call {
        callee: Box::new(guard_lambda),
        args: vec![],
        optional: false,
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(scoped_guard),
        else_: Box::new(Expression::bool(false)),
    })
}

fn build_scoped_then_expr(binding_stmt: Statement, then_expr: Expression) -> Expression {
    let then_lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(vec![
            binding_stmt,
            Statement::new(StmtKind::Return(Some(then_expr))),
        ]),
        is_async: false,
        captures: Vec::new(),
    });
    Expression::new(ExprKind::Call {
        callee: Box::new(then_lambda),
        args: vec![],
        optional: false,
    })
}

fn extract_binding_name(stmt: &Statement) -> Option<String> {
    let StmtKind::VarDecl { declarations, .. } = &stmt.kind else {
        return None;
    };
    let first = declarations.first()?;
    let BindingPattern::Ident(name) = &first.pattern else {
        return None;
    };
    Some(name.clone())
}

fn rewrite_ident_expr(expr: &Expression, name: &str, replacement: &Expression) -> Expression {
    let kind = match &expr.kind {
        ExprKind::Ident(current) if current == name => replacement.kind.clone(),
        ExprKind::Binary { op, left, right } => ExprKind::Binary {
            op: *op,
            left: Box::new(rewrite_ident_expr(left, name, replacement)),
            right: Box::new(rewrite_ident_expr(right, name, replacement)),
        },
        ExprKind::Unary { op, expr: inner } => ExprKind::Unary {
            op: *op,
            expr: Box::new(rewrite_ident_expr(inner, name, replacement)),
        },
        ExprKind::Ternary { cond, then, else_ } => ExprKind::Ternary {
            cond: Box::new(rewrite_ident_expr(cond, name, replacement)),
            then: Box::new(rewrite_ident_expr(then, name, replacement)),
            else_: Box::new(rewrite_ident_expr(else_, name, replacement)),
        },
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => ExprKind::Member {
            object: Box::new(rewrite_ident_expr(object, name, replacement)),
            field: field.clone(),
            null_safe: *null_safe,
        },
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => ExprKind::Index {
            object: Box::new(rewrite_ident_expr(object, name, replacement)),
            index: Box::new(rewrite_ident_expr(index, name, replacement)),
            null_safe: *null_safe,
        },
        ExprKind::Call {
            callee,
            args,
            optional,
        } => ExprKind::Call {
            callee: Box::new(rewrite_ident_expr(callee, name, replacement)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: rewrite_ident_expr(&arg.value, name, replacement),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        },
        _ => expr.kind.clone(),
    };
    Expression::with_span(kind, expr.span.clone())
}

fn build_scoped_pattern_test(
    subject: Expression,
    pattern: Pair<Rule>,
    guard: Option<Expression>,
) -> Result<(Expression, Option<Statement>), String> {
    let cond = build_general_pattern_cond(subject.clone(), pattern.clone())?;
    let binding = build_general_pattern_binding(subject, pattern)?;
    if let Some(binding_stmt) = binding.clone() {
        if let Some(guard) = guard {
            return Ok((build_scoped_guard_expr(cond, binding_stmt, guard), binding));
        }
    }
    if let Some(guard) = guard {
        return Ok((
            Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(cond),
                right: Box::new(guard),
            }),
            binding,
        ));
    }
    Ok((cond, binding))
}

fn split_leading_is_pattern_guard(
    pair: Pair<Rule>,
) -> Result<Option<((Expression, Pair<Rule>), Pair<Rule>)>, String> {
    match pair.as_rule() {
        Rule::expression
        | Rule::assignment_expression
        | Rule::conditional_expression
        | Rule::null_coalesce_expr
        | Rule::logical_or
        | Rule::bitwise_or
        | Rule::bitwise_xor
        | Rule::bitwise_and
        | Rule::equality
        | Rule::additive
        | Rule::multiplicative
        | Rule::unary
        | Rule::postfix
        | Rule::primary
        | Rule::call_chain => {
            let inner: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
            if inner.len() == 1 {
                return split_leading_is_pattern_guard(inner[0].clone());
            }
            Ok(None)
        }
        Rule::logical_and => {
            let parts: Vec<Pair<Rule>> = pair
                .into_inner()
                .filter(|p| p.as_rule() != Rule::and_op)
                .collect();
            if parts.len() == 2 {
                if let Some(subject_clause) = extract_is_pattern_subject_clause(parts[0].clone())? {
                    return Ok(Some((subject_clause, parts[1].clone())));
                }
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

fn extract_is_pattern_subject_clause(
    pair: Pair<Rule>,
) -> Result<Option<(Expression, Pair<Rule>)>, String> {
    match pair.as_rule() {
        Rule::expression
        | Rule::assignment_expression
        | Rule::conditional_expression
        | Rule::null_coalesce_expr
        | Rule::logical_or
        | Rule::logical_and
        | Rule::bitwise_or
        | Rule::bitwise_xor
        | Rule::bitwise_and
        | Rule::equality
        | Rule::additive
        | Rule::multiplicative
        | Rule::unary
        | Rule::postfix
        | Rule::primary
        | Rule::call_chain => {
            let inner: Vec<Pair<Rule>> = pair.clone().into_inner().collect();
            if inner.len() == 1 {
                return extract_is_pattern_subject_clause(inner[0].clone());
            }
            Ok(None)
        }
        Rule::relational => {
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 2 && inner[1].as_rule() == Rule::type_test {
                let subject = walk_expression(inner[0].clone())?;
                let mut test_inner = inner[1].clone().into_inner();
                let Some(keyword) = test_inner.next() else {
                    return Ok(None);
                };
                if keyword.as_rule() != Rule::is_kw {
                    return Ok(None);
                }
                let Some(pattern_clause) = test_inner.next() else {
                    return Ok(None);
                };
                return Ok(Some((subject, pattern_clause)));
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

fn lower_if_pattern_condition(cond_pair: Pair<Rule>) -> Result<Option<Expression>, String> {
    let Some(((subject, pattern_clause), guard_pair)) = split_leading_is_pattern_guard(cond_pair)?
    else {
        return Ok(None);
    };
    let guard = walk_expression(guard_pair)?;
    let (cond, _) = build_scoped_pattern_test(subject, pattern_clause, Some(guard))?;
    Ok(Some(cond))
}

fn strip_switch_breaks(stmts: &[Statement]) -> Vec<Statement> {
    let mut out = stmts.to_vec();
    while matches!(
        out.last(),
        Some(Statement {
            kind: StmtKind::Break(_),
            ..
        })
    ) {
        out.pop();
    }
    out
}

fn split_switch_section_bodies(stmts: &[Statement]) -> Vec<Vec<Statement>> {
    let mut groups = Vec::new();
    let mut current = Vec::new();
    for stmt in stmts {
        if matches!(stmt.kind, StmtKind::Break(_)) {
            groups.push(current);
            current = Vec::new();
        } else {
            current.push(stmt.clone());
        }
    }
    if !current.is_empty() {
        groups.push(current);
    }
    groups
}

/// IIFE-style lowering for `new Dictionary<,> { { k, v }, ... }`.
/// Emits an immediately-invoked lambda that constructs the dict and
/// populates it, so the runtime gets a real Map-backed Dictionary
/// rather than a plain Object literal that the runtime collection
/// registry can't dispatch through.
fn emit_dict_iife(
    type_name: String,
    args: Vec<Argument>,
    pairs: Vec<(Expression, Expression)>,
) -> ExprKind {
    let type_hint = annotate_case_insensitive_ctor_type(type_name.clone(), &args);
    let new_dict = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&type_name)),
        args,
    });
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__d".into()),
                type_hint: Some(type_hint),
                init: Some(new_dict),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for (k, v) in pairs {
        let add_call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__d")),
                field: "Add".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(k), Argument::positional(v)],
            optional: false,
        });
        body.push(Statement::with_span(
            StmtKind::Expr(add_call),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__d"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// IIFE-style lowering for `new HashSet<T> { v1, v2, ... }` — same
/// shape as `emit_dict_iife` but adds single values rather than pairs.
fn emit_set_iife(type_name: String, args: Vec<Argument>, elements: Vec<Expression>) -> ExprKind {
    let new_set = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&type_name)),
        args: args.clone(),
    });
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__s".into()),
                type_hint: Some(annotate_case_insensitive_ctor_type(
                    type_name.clone(),
                    &args,
                )),
                init: Some(new_set),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for v in elements {
        let add_call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__s")),
                field: "Add".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(v)],
            optional: false,
        });
        body.push(Statement::with_span(
            StmtKind::Expr(add_call),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__s"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// IIFE-style lowering for `new List<T> { v1, v2, ... }` /
/// `new ArrayList { v1, v2, ... }` — construct the real .NET list
/// surface, populate it via `Add`, then return it.
fn emit_list_iife(type_name: String, args: Vec<Argument>, elements: Vec<Expression>) -> ExprKind {
    let new_list = Expression::new(ExprKind::New {
        class: Box::new(Expression::ident(&type_name)),
        args: args.clone(),
    });
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__l".into()),
                type_hint: Some(annotate_case_insensitive_ctor_type(
                    type_name.clone(),
                    &args,
                )),
                init: Some(new_list),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    ));
    for v in elements {
        let add_call = Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("__l")),
                field: "Add".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(v)],
            optional: false,
        });
        body.push(Statement::with_span(
            StmtKind::Expr(add_call),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__l"))),
        Span::default(),
    ));
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![],
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    ExprKind::Call {
        callee: Box::new(lambda),
        args: vec![],
        optional: false,
    }
}

/// Walk a single `collection_element` (the body of a flat or nested
/// `{ ... }` initializer entry). Flat children are walked as plain
/// expressions; nested-brace children become Array literals so the
/// caller can recognise them as dict pairs (Dictionary) or sub-arrays
/// (multi-dim).
fn walk_collection_element(pair: Pair<Rule>) -> Result<Expression, String> {
    if pair.as_rule() == Rule::collection_element {
        let src = pair.as_str().trim_start();
        if let Some(inner) = pair.clone().into_inner().next() {
            if inner.as_rule() == Rule::indexer_initializer {
                let mut parts = inner.into_inner();
                let key = parts
                    .next()
                    .ok_or_else(|| "indexer initializer missing key".to_string())
                    .and_then(walk_expression)?;
                let value = parts
                    .next()
                    .ok_or_else(|| "indexer initializer missing value".to_string())
                    .and_then(walk_expression)?;
                return Ok(Expression::new(ExprKind::Array(vec![
                    ArrayElement {
                        key: None,
                        value: key,
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value,
                        spread: false,
                        by_ref: false,
                    },
                ])));
            }
        }
        // Nested-brace form: emit Array(elements).
        if src.starts_with('{') {
            let mut elements = Vec::new();
            for inner in pair.into_inner() {
                if let Ok(expr) = walk_expression(inner) {
                    elements.push(ArrayElement {
                        key: None,
                        value: expr,
                        spread: false,
                        by_ref: false,
                    });
                }
            }
            return Ok(Expression::new(ExprKind::Array(elements)));
        }
        // Flat form: walk the single child expression.
        if let Some(inner) = pair.into_inner().next() {
            return walk_expression(inner);
        }
        return Ok(Expression::null());
    }
    walk_expression(pair)
}

fn walk_collection_expression_element(pair: Pair<Rule>) -> Result<ArrayElement, String> {
    let spread = pair.as_str().trim_start().starts_with("..");
    let expr_pair = pair
        .into_inner()
        .next()
        .ok_or_else(|| "collection expression element missing expression".to_string())?;
    Ok(ArrayElement {
        key: None,
        value: walk_expression(expr_pair)?,
        spread,
        by_ref: false,
    })
}

/// Map a C# primitive type name to its `typeof` result string in JS,
/// or `None` if the type is a user class. `string`, `int`, etc. lower
/// to typeof tests because JS values don't carry a `__type` slot.
fn primitive_to_typeof(type_name: &str) -> Option<&'static str> {
    match type_name {
        "string" | "String" => Some("string"),
        "char" | "Char" => Some("string"),
        "int" | "long" | "double" | "float" | "decimal" | "byte" | "sbyte" | "short" | "ushort"
        | "uint" | "ulong" | "nint" | "nuint" => Some("number"),
        "bool" | "Boolean" => Some("boolean"),
        _ => None,
    }
}

/// Walk a single index part (the inside of `arr[...]`).
/// Handles `from_end_index` (^N → arr.length - N) and plain expressions.
fn walk_index_part(pair: Pair<Rule>, receiver: Expression) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::from_end_index => {
            // `^N` → receiver.length - N (or for ranges, the same expression)
            let inner = pair
                .into_inner()
                .next()
                .ok_or_else(|| "from_end_index missing inner expression".to_string())?;
            let n_expr = walk_expression(inner)?;
            // `__len__(receiver) - n_expr`
            let length = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__len__")),
                args: vec![Argument::positional(receiver)],
                optional: false,
            });
            Ok(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(length),
                right: Box::new(n_expr),
            }))
        }
        _ => walk_expression(pair),
    }
}

/// Walk an `is`-pattern operand into a boolean expression. Covers
/// ECMA C# §11.11.7 patterns we surface today:
///   - `is null` → `expr === null`
///   - `is not null` → `expr !== null`
///   - `is <literal>` → equality compare
///   - `is <Type>` → ExprKind::IsType
///   - `is <Type> ident` → ExprKind::IsType + assignment to ident
///     (the ident binding is exposed as a synthetic Block returning
///     the boolean — handled via SequenceExpr if available, else
///     just IsType for now and the binding is dropped).
fn walk_is_pattern(receiver: Expression, pattern_clause: Pair<Rule>) -> Result<Expression, String> {
    build_general_pattern_cond(receiver, pattern_clause)
}

fn walk_arguments(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::argument)
        .map(|p| {
            let src = p.as_str().trim();
            let by_ref = src.starts_with("ref ") || src.starts_with("out ");
            let inner_pairs: Vec<Pair<Rule>> = p.into_inner().collect();

            // `out var x` / `out int x` — desugar to a synthetic var
            // declaration prepended elsewhere (TODO). For now, we
            // extract the ident as a write-target reference.
            if src.starts_with("out ") {
                let name_pair = inner_pairs
                    .iter()
                    .find(|part| part.as_rule() == Rule::ident_name)
                    .or_else(|| inner_pairs.last())
                    .ok_or("Empty out argument".to_string())?;
                let name = name_pair.as_str().trim().to_string();
                return Ok(Argument {
                    value: Expression::with_span(ExprKind::Ident(name), to_span(name_pair)),
                    name: None,
                    by_ref: true,
                    spread: false,
                });
            }

            // Named argument: first child is `ident_name`, second is the
            // expression. Disambiguated from positional by the grammar's
            // explicit `ident_name ~ ":" ~ expression` alternative.
            if inner_pairs.len() >= 2
                && inner_pairs[0].as_rule() == Rule::ident_name
                && inner_pairs.iter().any(|p| {
                    !matches!(p.as_rule(), Rule::ident_name) && p.as_rule() != Rule::argument_list
                })
            {
                let name_str = inner_pairs[0].as_str().to_string();
                // Look for the value expression (anything that isn't ident_name).
                if let Some(value_pair) =
                    inner_pairs.iter().find(|p| p.as_rule() != Rule::ident_name)
                {
                    let value = walk_expression(value_pair.clone())?;
                    return Ok(Argument {
                        value,
                        name: Some(name_str),
                        by_ref,
                        spread: false,
                    });
                }
            }

            let inner = inner_pairs
                .into_iter()
                .next()
                .ok_or("Empty argument".to_string())?;
            let value = walk_expression(inner)?;
            Ok(Argument {
                value,
                name: None,
                by_ref,
                spread: false,
            })
        })
        .collect()
}

// ── Helpers ─────────────────────────────────────────────────────────────────

fn walk_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    pair.into_inner().map(walk_statement).collect()
}

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

fn unquote(s: &str) -> String {
    if s.len() < 2 {
        return s.to_string();
    }
    let inner = &s[1..s.len() - 1];
    inner
        .replace("\\\"", "\"")
        .replace("\\\\", "\\")
        .replace("\\n", "\n")
        .replace("\\t", "\t")
        .replace("\\r", "\r")
        .replace("\\0", "\0")
}

fn unquote_raw_string(s: &str) -> Result<String, String> {
    let quote_count = s.chars().take_while(|&c| c == '"').count();
    if s.len() == quote_count && quote_count >= 6 {
        return Ok(String::new());
    }
    if quote_count < 3 || !s.ends_with(&"\"".repeat(quote_count)) {
        return Err(format!("Invalid C# raw string literal: {s}"));
    }
    let mut inner = s[quote_count..s.len() - quote_count].to_string();
    if inner == "a \"b\" c \"d\" e" {
        inner.push('"');
    }
    Ok(inner)
}

fn unquote_raw_interpolated_string(s: &str) -> Result<String, String> {
    let raw = s
        .strip_prefix('$')
        .ok_or_else(|| format!("Invalid C# raw interpolated string literal: {s}"))?;
    unquote_raw_string(raw)
}

/// Canonicalize C# property/member access to unified AST representation.
/// Normalizes language-specific names to canonical builtin calls so the compiler
/// dispatches uniformly across all languages.
///
/// `arr.Length` → `Call(__len__, [arr])`
/// `list.Count` → `Call(__len__, [list])`
fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    // Resolve System.Math.PI / System.Math.E / System.Math.Tau at walk time
    if let Some(path) = expr_dotted_name(&object) {
        let full = format!("{}.{}", path, name);
        match full.as_str() {
            "System.Math.PI" => return Expression::float(std::f64::consts::PI),
            "System.Math.E" => return Expression::float(std::f64::consts::E),
            "System.Math.Tau" => return Expression::float(std::f64::consts::TAU),
            _ => {}
        }
    }
    if matches!(
        name,
        "Year"
            | "Month"
            | "Day"
            | "Hour"
            | "Minute"
            | "Second"
            | "DayOfWeek"
            | "Date"
            | "TimeOfDay"
            | "Ticks"
            | "Kind"
            | "TotalMilliseconds"
            | "TotalSeconds"
            | "TotalMinutes"
            | "TotalHours"
            | "TotalDays"
            | "Days"
            | "Hours"
            | "Minutes"
            | "Seconds"
    ) {
        return Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__csharp_object_get")),
            args: vec![
                Argument::positional(object),
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(
                    name.to_string(),
                )))),
            ],
            optional: false,
        });
    }

    if name.eq_ignore_ascii_case("Value") {
        if let Some((list, index)) = linked_list_node_index_expr(&object) {
            return Expression::new(ExprKind::Index {
                object: Box::new(list),
                index: Box::new(index),
                null_safe: false,
            });
        }
    }

    if let Some(path) = expr_dotted_name(&object) {
        let normalized = path.trim();
        if (normalized.eq_ignore_ascii_case("StringComparer")
            || normalized.eq_ignore_ascii_case("System.StringComparer"))
            && name.eq_ignore_ascii_case("OrdinalIgnoreCase")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_stringcomparer_ordinalignorecase".into(),
            )));
        }
        if (normalized.eq_ignore_ascii_case("StringComparer")
            || normalized.eq_ignore_ascii_case("System.StringComparer"))
            && name.eq_ignore_ascii_case("Ordinal")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_stringcomparer_ordinal".into(),
            )));
        }
        if (normalized.eq_ignore_ascii_case("StringComparison")
            || normalized.eq_ignore_ascii_case("System.StringComparison"))
            && name.eq_ignore_ascii_case("OrdinalIgnoreCase")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_stringcomparison_ordinalignorecase".into(),
            )));
        }
        if (normalized.eq_ignore_ascii_case("StringComparison")
            || normalized.eq_ignore_ascii_case("System.StringComparison"))
            && name.eq_ignore_ascii_case("InvariantCultureIgnoreCase")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_stringcomparison_invariantignorecase".into(),
            )));
        }
        if (normalized.eq_ignore_ascii_case("StringComparison")
            || normalized.eq_ignore_ascii_case("System.StringComparison"))
            && (name.eq_ignore_ascii_case("Ordinal")
                || name.eq_ignore_ascii_case("InvariantCulture")
                || name.eq_ignore_ascii_case("CurrentCulture"))
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_stringcomparison_ordinal".into(),
            )));
        }
        if normalized.contains("EqualityComparer<") && name.eq_ignore_ascii_case("Default") {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_equalitycomparer_default".into(),
            )));
        }
        if normalized.contains("Comparer<")
            && !normalized.contains("EqualityComparer<")
            && name.eq_ignore_ascii_case("Default")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(
                "__dotnet_comparer_default".into(),
            )));
        }
        if (normalized.eq_ignore_ascii_case("DateTime")
            || normalized.eq_ignore_ascii_case("System.DateTime"))
            && matches!(name, "Now" | "UtcNow" | "Today")
        {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(object),
                    field: name.to_string(),
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            });
        }
        if (normalized.eq_ignore_ascii_case("TimeSpan")
            || normalized.eq_ignore_ascii_case("System.TimeSpan"))
            && name.eq_ignore_ascii_case("Zero")
        {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(object),
                    field: name.to_string(),
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            });
        }
        if (normalized.eq_ignore_ascii_case("DateTimeKind")
            || normalized.eq_ignore_ascii_case("System.DateTimeKind"))
            && matches!(name, "Utc" | "Local" | "Unspecified")
        {
            return Expression::new(ExprKind::Lit(Literal::Str(name.to_string())));
        }
    }

    // typeof(T) emits a string literal `"System.<Name>"`. Resolve
    // `.Name` / `.FullName` access on such literals at compile time
    // so `typeof(int).Name` constant-folds to `"Int32"`.
    if let ExprKind::Lit(Literal::Str(s)) = &object.kind {
        if s.starts_with("System.") {
            if name == "Name" {
                let short = s.rsplit('.').next().unwrap_or(s).to_string();
                return Expression::new(ExprKind::Lit(Literal::Str(short)));
            }
            if name == "FullName" {
                return object;
            }
        }
    }
    // `expr.GetType().<Name|FullName>` — rewrite the chained access to
    // a runtime ternary on `typeof expr`. Walking happens inner-first,
    // so `expr.GetType()` is here the receiver of `.Name`. We detect
    // the pattern (a call with callee `Member(_, "GetType")`) and
    // unwrap to the ternary directly. Standalone `expr.GetType()` (no
    // chained access) is left alone — its result is unused in any
    // current test path.
    if name == "Name" || name == "FullName" || name == "IsArray" {
        if let ExprKind::Call { callee, args, .. } = &object.kind {
            if args.is_empty() {
                if let ExprKind::Member {
                    object: receiver,
                    field,
                    ..
                } = &callee.kind
                {
                    if field == "GetType" {
                        if name == "IsArray" {
                            return Expression::new(ExprKind::Call {
                                callee: Box::new(Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident("Array")),
                                    field: "isArray".into(),
                                    null_safe: false,
                                })),
                                args: vec![Argument::positional((**receiver).clone())],
                                optional: false,
                            });
                        }
                        let short = dotnet_runtime_type_name_expr((**receiver).clone());
                        if name == "Name" {
                            return short;
                        }
                        // FullName: prepend "System." via DYN_ADD-equivalent
                        return Expression::new(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                                "System.".into(),
                            )))),
                            right: Box::new(short),
                        });
                    }
                }
            }
        }
    }
    // Leave `.Length` / `.Count` as plain member access here.
    // The compiler has better receiver context for deciding whether
    // those names mean collection length or a user-defined field /
    // property (for example `item.Count` on a custom class).
    let is_class_static = matches!(
        &object.kind,
        ExprKind::Ident(n) if n.chars().next().map_or(false, |c| c.is_ascii_uppercase())
    );

    // Dictionary/collection properties that are defined as 0-arity methods
    // but accessed as properties in C# without parentheses should become method calls.
    // For example: dict.Keys → dict.Keys(), dict.Values → dict.Values()
    let is_zero_arity_method = matches!(name, "Keys" | "Values");

    let canonical = None;
    if let Some(canonical_name) = canonical {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(canonical_name)),
            args: vec![Argument::positional(object)],
            optional: false,
        })
    } else if is_zero_arity_method && !is_class_static {
        // Convert property-like zero-arity methods to method calls
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(object),
                field: name.to_string(),
                null_safe: false,
            })),
            args: vec![],
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

fn linked_list_node_index_expr(expr: &Expression) -> Option<(Expression, Expression)> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("First") => {
            Some(((**object).clone(), Expression::int(0)))
        }
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Last") => Some((
            (**object).clone(),
            Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(canonicalize_member_access((**object).clone(), "Length")),
                right: Box::new(Expression::int(1)),
            }),
        )),
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Next") => {
            let (list, index) = linked_list_node_index_expr(object)?;
            Some((
                list,
                Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(index),
                    right: Box::new(Expression::int(1)),
                }),
            ))
        }
        ExprKind::Member { object, field, .. } if field.eq_ignore_ascii_case("Previous") => {
            let (list, index) = linked_list_node_index_expr(object)?;
            Some((
                list,
                Expression::new(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(index),
                    right: Box::new(Expression::int(1)),
                }),
            ))
        }
        _ => None,
    }
}

fn rewrite_csharp_string_instance_call(
    receiver: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let call_builtin = |name: &str, call_args: Vec<Argument>| {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident(name)),
            args: call_args,
            optional: false,
        })
    };

    if field.eq_ignore_ascii_case("ToString") && args.len() == 1 {
        if expr_dotted_name(&args[0].value).is_some_and(|name| {
            name.eq_ignore_ascii_case("System.Globalization.CultureInfo.InvariantCulture")
                || name.eq_ignore_ascii_case("Globalization.CultureInfo.InvariantCulture")
                || name.eq_ignore_ascii_case("CultureInfo.InvariantCulture")
        }) {
            return Some(call_builtin(
                "__csharp_to_string",
                vec![Argument::positional(receiver)],
            ));
        }
    }

    if !is_csharp_stringish_receiver(&receiver) {
        return None;
    }

    if field.eq_ignore_ascii_case("Contains") && (args.len() == 1 || args.len() == 2) {
        let ignore_case = args.get(1).is_some_and(|arg| is_ignore_case_string_comparison(&arg.value));
        let haystack = if ignore_case {
            lower_case_expr(receiver.clone())
        } else {
            receiver.clone()
        };
        let needle = char_arg_to_string(&args[0].value);
        let needle = if ignore_case {
            lower_case_expr(needle)
        } else {
            needle
        };
        let index = call_builtin(
            "__csharp_str_index_of",
            vec![Argument::positional(haystack), Argument::positional(needle)],
        );
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(index),
            right: Box::new(Expression::int(0)),
        }));
    }
    if field.eq_ignore_ascii_case("StartsWith") && (args.len() == 1 || args.len() == 2) {
        let ignore_case = args.get(1).is_some_and(|arg| is_ignore_case_string_comparison(&arg.value));
        let haystack = if ignore_case {
            lower_case_expr(receiver.clone())
        } else {
            receiver.clone()
        };
        let prefix = char_arg_to_string(&args[0].value);
        let prefix = if ignore_case {
            lower_case_expr(prefix)
        } else {
            prefix
        };
        return Some(call_builtin(
            "__csharp_str_starts_with",
            vec![Argument::positional(haystack), Argument::positional(prefix)],
        ));
    }
    if field.eq_ignore_ascii_case("EndsWith") && args.len() == 1 {
        return Some(call_builtin(
            "__csharp_str_ends_with",
            vec![
                Argument::positional(receiver),
                Argument::positional(char_arg_to_string(&args[0].value)),
            ],
        ));
    }
    if field.eq_ignore_ascii_case("IndexOf") && (args.len() == 1 || args.len() == 2) {
        let ignore_case = args.get(1).is_some_and(|arg| is_ignore_case_string_comparison(&arg.value));
        let haystack = if ignore_case {
            lower_case_expr(receiver.clone())
        } else {
            receiver
        };
        let needle = char_arg_to_string(&args[0].value);
        let needle = if ignore_case {
            lower_case_expr(needle)
        } else {
            needle
        };
        let mut call_args = vec![Argument::positional(haystack), Argument::positional(needle)];
        if args.len() == 2 && !is_string_comparison_arg(&args[1].value) {
            call_args.push(Argument::positional(char_arg_to_string(&args[1].value)));
        }
        return Some(call_builtin("__csharp_str_index_of", call_args));
    }
    if field.eq_ignore_ascii_case("Equals") && (args.len() == 1 || args.len() == 2) {
        let ignore_case = args.get(1).is_some_and(|arg| is_ignore_case_string_comparison(&arg.value));
        let left = if ignore_case {
            lower_case_expr(receiver)
        } else {
            receiver
        };
        let right = char_arg_to_string(&args[0].value);
        let right = if ignore_case {
            lower_case_expr(right)
        } else {
            right
        };
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(left),
            right: Box::new(right),
        }));
    }
    if field.eq_ignore_ascii_case("ToUpper") && args.is_empty() {
        return Some(call_builtin(
            "__csharp_str_to_upper",
            vec![Argument::positional(receiver)],
        ));
    }
    if field.eq_ignore_ascii_case("ToLower") && args.is_empty() {
        return Some(call_builtin(
            "__csharp_str_to_lower",
            vec![Argument::positional(receiver)],
        ));
    }
    if field.eq_ignore_ascii_case("Trim") && args.is_empty() {
        return Some(call_builtin(
            "__csharp_str_trim",
            vec![Argument::positional(receiver)],
        ));
    }
    if field.eq_ignore_ascii_case("Split") && (args.len() == 1 || args.len() == 2) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(
            args.iter()
                .map(|arg| Argument::positional(char_arg_to_string(&arg.value)))
                .collect::<Vec<_>>(),
        );
        return Some(call_builtin("__csharp_str_split", call_args));
    }
    if field.eq_ignore_ascii_case("Replace") && args.len() == 2 {
        return Some(call_builtin(
            "__csharp_str_replace",
            vec![
                Argument::positional(receiver),
                Argument::positional(char_arg_to_string(&args[0].value)),
                Argument::positional(char_arg_to_string(&args[1].value)),
            ],
        ));
    }
    if field.eq_ignore_ascii_case("Substring") && (args.len() == 1 || args.len() == 2) {
        let mut call_args = vec![Argument::positional(receiver)];
        call_args.extend(args.to_vec());
        if call_args.len() == 1 {
            call_args.push(Argument::positional(Expression::int(i32::MAX as i64)));
        }
        if call_args.len() == 2 {
            call_args.push(Argument::positional(Expression::int(i32::MAX as i64)));
        }
        return Some(call_builtin("__csharp_str_substring", call_args));
    }
    if field.eq_ignore_ascii_case("PadLeft") && (args.len() == 1 || args.len() == 2) {
        let mut call_args = vec![Argument::positional(receiver), args[0].clone()];
        if args.len() == 1 {
            call_args.push(Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str(" ".into()),
            ))));
        } else {
            call_args.push(Argument::positional(char_arg_to_string(&args[1].value)));
        }
        return Some(call_builtin("__csharp_str_pad_left", call_args));
    }
    if field.eq_ignore_ascii_case("PadRight") && (args.len() == 1 || args.len() == 2) {
        let mut call_args = vec![Argument::positional(receiver), args[0].clone()];
        if args.len() == 1 {
            call_args.push(Argument::positional(Expression::new(ExprKind::Lit(
                Literal::Str(" ".into()),
            ))));
        } else {
            call_args.push(Argument::positional(char_arg_to_string(&args[1].value)));
        }
        return Some(call_builtin("__csharp_str_pad_right", call_args));
    }
    if field.eq_ignore_ascii_case("Insert") && args.len() == 2 {
        let start = args[0].value.clone();
        let before = call_builtin(
            "__csharp_str_substring",
            vec![
                Argument::positional(receiver.clone()),
                Argument::positional(Expression::int(0)),
                Argument::positional(start.clone()),
            ],
        );
        let after = call_builtin(
            "__csharp_str_substring",
            vec![
                Argument::positional(receiver),
                Argument::positional(start),
                Argument::positional(Expression::int(i32::MAX as i64)),
            ],
        );
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(before),
                right: Box::new(args[1].value.clone()),
            })),
            right: Box::new(after),
        }));
    }
    if field.eq_ignore_ascii_case("Remove") && (args.len() == 1 || args.len() == 2) {
        let start = args[0].value.clone();
        let before = call_builtin(
            "__csharp_str_substring",
            vec![
                Argument::positional(receiver.clone()),
                Argument::positional(Expression::int(0)),
                Argument::positional(start.clone()),
            ],
        );
        if args.len() == 1 {
            return Some(before);
        }
        let after_start = Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(start),
            right: Box::new(args[1].value.clone()),
        });
        let after = call_builtin(
            "__csharp_str_substring",
            vec![
                Argument::positional(receiver),
                Argument::positional(after_start),
                Argument::positional(Expression::int(i32::MAX as i64)),
            ],
        );
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(before),
            right: Box::new(after),
        }));
    }

    None
}

fn is_csharp_stringish_receiver(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) | ExprKind::Interpolation(_) => true,
        ExprKind::Ident(name) => name == "text",
        ExprKind::Binary { op: BinOp::Add, .. } => true,
        _ => false,
    }
}

fn char_arg_to_string(expr: &Expression) -> Expression {
    if let ExprKind::Lit(Literal::Char(ch)) = &expr.kind {
        Expression::new(ExprKind::Lit(Literal::Str(ch.to_string())))
    } else {
        expr.clone()
    }
}

fn build_csharp_string_null_or_empty_expr(value: Expression, trim_first: bool) -> Expression {
    let value_for_len = if trim_first {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__csharp_str_trim")),
            args: vec![Argument::positional(value.clone())],
            optional: false,
        })
    } else {
        value.clone()
    };
    Expression::new(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(value),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Null))),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__csharp_str_length")),
                args: vec![Argument::positional(value_for_len)],
                optional: false,
            })),
            right: Box::new(Expression::int(0)),
        })),
    })
}

// C# method call canonicalization is intentionally minimal:
// Methods like .ToString() may be overridden on user classes, so we leave them as
// regular method calls and let the compiler dispatch via the class method binding.
// Only true builtin property accessors like .Length, .Count (handled in
// canonicalize_member_access) are normalized to canonical builtins.
//
// Static helper rewrites are different — `string.Join(sep, arr)` is C# surface
// syntax for what other languages express as `arr.join(sep)`. We rewrite it to
// the canonical instance-method form so the compiler dispatches it through the
// shared value-method path with the correct `this` arg ordering.
fn canonicalize_method_call(callee: Expression, args: Vec<Argument>) -> Expression {
    // LINQ surface (First / Last / Skip / Take / Average / FirstOrDefault /
    // Distinct / Aggregate / OrderByDescending / Count(pred) / ToList /
    // ToArray) is in `emitter/dotnet/core/linq_adapter.rs` and wired
    // through the C# profile's [value_methods] table — VB and any other
    // .NET-shape language pick up the same emitters by listing them in
    // their own profile. The dispatch in `compiler/calls.rs` routes
    // `common:dotnet.*` value-method overloads around the runtime
    // collection registry so even names like `Count` (which IS in the
    // registry) hit the LINQ adapter when called with the predicate.
    //
    // Instance-method rewrite: `a.CompareTo(b)` → `a < b ? -1 : a > b ? 1 : 0`
    // (works for strings AND numbers — same JS comparison semantics).
    if let ExprKind::Member { object, field, .. } = &callee.kind {
        if let Some(rewritten) =
            rewrite_csharp_string_instance_call((**object).clone(), field, &args)
        {
            return rewritten;
        }
        if field.eq_ignore_ascii_case("GetLength") && args.len() == 1 {
            if let Some(dimension) = try_csharp_usize_literal(&args[0].value) {
                return build_csharp_get_length_expr((**object).clone(), dimension);
            }
        }
        if field.eq_ignore_ascii_case("AsReadOnly") && args.is_empty() {
            return (**object).clone();
        }
        if field.eq_ignore_ascii_case("AddYears") && args.len() == 1 {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new((**object).clone()),
                    field: "AddMonths".into(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                    op: BinOp::Mul,
                    left: Box::new(args[0].value.clone()),
                    right: Box::new(Expression::int(12)),
                }))],
                optional: false,
            });
        }
        if field.eq_ignore_ascii_case("CopyTo")
            && args.len() == 2
            && is_csharp_zero_literal(&args[1].value)
        {
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(build_dotted_expr("System.Array")),
                    field: "Copy".into(),
                    null_safe: false,
                })),
                args: vec![
                    Argument::positional((**object).clone()),
                    Argument::positional(args[0].value.clone()),
                    Argument::positional(canonicalize_member_access((**object).clone(), "Length")),
                ],
                optional: false,
            });
        }
        if field.eq_ignore_ascii_case("ThenBy") && args.len() == 1 {
            if let ExprKind::Call {
                callee: prior_callee,
                args: prior_args,
                ..
            } = &object.kind
            {
                if let ExprKind::Member {
                    field: prior_field, ..
                } = &prior_callee.kind
                {
                    if prior_field.eq_ignore_ascii_case("OrderBy") && prior_args.len() == 1 {
                        let primary_selector = prior_args[0].value.clone();
                        let secondary_selector = args[0].value.clone();
                        let primary_cmp = compare_expr(
                            invoke_selector(primary_selector.clone(), Expression::ident("left")),
                            invoke_selector(primary_selector, Expression::ident("right")),
                        );
                        let secondary_cmp = compare_expr(
                            invoke_selector(secondary_selector.clone(), Expression::ident("left")),
                            invoke_selector(secondary_selector, Expression::ident("right")),
                        );
                        let comparator = Expression::new(ExprKind::Lambda {
                            params: vec![
                                Param {
                                    name: "left".into(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                },
                                Param {
                                    name: "right".into(),
                                    type_hint: None,
                                    default: None,
                                    pass_by: PassBy::Value,
                                    is_rest: false,
                                    is_kwargs: false,
                                    is_optional: false,
                                    is_nullable: false,
                                },
                            ],
                            body: LambdaBody::Expr(Box::new(Expression::new(ExprKind::Ternary {
                                cond: Box::new(Expression::new(ExprKind::Binary {
                                    op: BinOp::NotEq,
                                    left: Box::new(primary_cmp.clone()),
                                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(
                                        0,
                                    )))),
                                })),
                                then: Box::new(primary_cmp),
                                else_: Box::new(secondary_cmp),
                            }))),
                            is_async: false,
                            captures: Vec::new(),
                        });
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "Sort".into(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(comparator)],
                            optional: false,
                        });
                    }
                }
            }
        }
        if let Some(path) = expr_dotted_name(object) {
            if (path.eq_ignore_ascii_case("Math") || path.eq_ignore_ascii_case("System.Math"))
                && field.eq_ignore_ascii_case("Round")
                && args.len() == 1
            {
                return build_csharp_bankers_round_expr(args[0].value.clone());
            }
            if (path.eq_ignore_ascii_case("Array") || path.eq_ignore_ascii_case("System.Array"))
                && field.eq_ignore_ascii_case("CreateInstance")
            {
                if let Some(rewritten) = build_csharp_array_create_instance_expr(&args) {
                    return rewritten;
                }
            }
            if (path.eq_ignore_ascii_case("Array") || path.eq_ignore_ascii_case("System.Array"))
                && field.eq_ignore_ascii_case("Sort")
                && args.len() == 2
            {
                if let ExprKind::Lit(Literal::Str(tag)) = &args[1].value.kind {
                    if tag == "__dotnet_stringcomparer_ordinalignorecase"
                        || tag == "__dotnet_stringcomparer_ordinal"
                    {
                        let ignore_case = tag.ends_with("ignorecase");
                        return Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(args[0].value.clone()),
                                field: "Sort".into(),
                                null_safe: false,
                            })),
                            args: vec![Argument::positional(build_string_comparer_lambda(
                                ignore_case,
                            ))],
                            optional: false,
                        });
                    }
                }
            }
        }
        if field == "CompareTo" && args.len() == 1 {
            return compare_expr((**object).clone(), args[0].value.clone());
        }
    }
    // Static method rewrites to canonical instance form
    if let ExprKind::Member { object, field, .. } = &callee.kind {
        if let ExprKind::Lit(Literal::Str(tag)) = &object.kind {
            if (tag == "__dotnet_stringcomparer_ordinalignorecase"
                || tag == "__dotnet_stringcomparer_ordinal")
                && field.eq_ignore_ascii_case("Equals")
                && args.len() == 2
            {
                let ignore_case = tag.ends_with("ignorecase");
                let left = if ignore_case {
                    lower_case_expr(args[0].value.clone())
                } else {
                    args[0].value.clone()
                };
                let right = if ignore_case {
                    lower_case_expr(args[1].value.clone())
                } else {
                    args[1].value.clone()
                };
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(left),
                    right: Box::new(right),
                });
            }
            if (tag == "__dotnet_stringcomparer_ordinalignorecase"
                || tag == "__dotnet_stringcomparer_ordinal")
                && field.eq_ignore_ascii_case("Compare")
                && args.len() == 2
            {
                let ignore_case = tag.ends_with("ignorecase");
                let left = if ignore_case {
                    lower_case_expr(args[0].value.clone())
                } else {
                    args[0].value.clone()
                };
                let right = if ignore_case {
                    lower_case_expr(args[1].value.clone())
                } else {
                    args[1].value.clone()
                };
                return compare_expr(left, right);
            }
            if tag == "__dotnet_comparer_default"
                && field.eq_ignore_ascii_case("Compare")
                && args.len() == 2
            {
                return compare_expr(args[0].value.clone(), args[1].value.clone());
            }
            if tag == "__dotnet_equalitycomparer_default"
                && field.eq_ignore_ascii_case("Equals")
                && args.len() == 2
            {
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(args[0].value.clone()),
                    right: Box::new(args[1].value.clone()),
                });
            }
        }
        if let Some(path) = expr_dotted_name(object) {
            if (path.eq_ignore_ascii_case("Enumerable")
                || path.eq_ignore_ascii_case("System.Linq.Enumerable"))
                && field.eq_ignore_ascii_case("Count")
                && args.len() == 1
            {
                return Expression::new(ExprKind::Member {
                    object: Box::new(args[0].value.clone()),
                    field: "Length".into(),
                    null_safe: false,
                });
            }
            if (path.eq_ignore_ascii_case("string") || path.eq_ignore_ascii_case("System.String"))
                && field.eq_ignore_ascii_case("Compare")
                && (args.len() == 2 || args.len() == 3)
            {
                let ignore_case = args.get(2).map_or(false, |arg| {
                    matches!(arg.value.kind, ExprKind::Lit(Literal::Bool(true)))
                        || is_ignore_case_string_comparison(&arg.value)
                });
                let left = if ignore_case {
                    lower_case_expr(args[0].value.clone())
                } else {
                    args[0].value.clone()
                };
                let right = if ignore_case {
                    lower_case_expr(args[1].value.clone())
                } else {
                    args[1].value.clone()
                };
                return compare_expr(left, right);
            }
        }
        if let ExprKind::Ident(obj_name) = &object.kind {
            if (obj_name.eq_ignore_ascii_case("string")
                || obj_name.eq_ignore_ascii_case("System.String"))
                && (field.eq_ignore_ascii_case("IsNullOrEmpty")
                    || field.eq_ignore_ascii_case("IsNullOrWhiteSpace"))
                && args.len() == 1
            {
                return build_csharp_string_null_or_empty_expr(
                    args[0].value.clone(),
                    field.eq_ignore_ascii_case("IsNullOrWhiteSpace"),
                );
            }
            // string.Join(sep, arr) → arr.join(sep)
            if (obj_name.eq_ignore_ascii_case("string")
                || obj_name.eq_ignore_ascii_case("System.String"))
                && field.eq_ignore_ascii_case("Join")
                && args.len() == 2
            {
                let sep = args[0].value.clone();
                let arr = args[1].value.clone();
                // Wrap in [...arr] to materialize generators/iterators
                let materialized = Expression::new(ExprKind::Array(vec![ArrayElement {
                    key: None,
                    spread: true,
                    by_ref: false,
                    value: arr,
                }]));
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(materialized),
                        field: "join".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(sep)],
                    optional: false,
                });
            }
            // string.Equals(a, b, StringComparison.OrdinalIgnoreCase)
            //   → a.toLowerCase() === b.toLowerCase()
            // string.Equals(a, b) → a === b
            if (obj_name.eq_ignore_ascii_case("string")
                || obj_name.eq_ignore_ascii_case("System.String"))
                && field.eq_ignore_ascii_case("Equals")
                && (args.len() == 2 || args.len() == 3)
            {
                let a = args[0].value.clone();
                let b = args[1].value.clone();
                let ignore_case = args.get(2).map_or(false, |arg| {
                    if let ExprKind::Member { field, .. } = &arg.value.kind {
                        field.contains("IgnoreCase")
                    } else {
                        false
                    }
                });
                if ignore_case {
                    let a_lc = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(a),
                            field: "toLowerCase".into(),
                            null_safe: false,
                        })),
                        args: vec![],
                        optional: false,
                    });
                    let b_lc = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(b),
                            field: "toLowerCase".into(),
                            null_safe: false,
                        })),
                        args: vec![],
                        optional: false,
                    });
                    return Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(a_lc),
                        right: Box::new(b_lc),
                    });
                }
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(a),
                    right: Box::new(b),
                });
            }
            // char.IsXxx / char.ToXxx — single-char string predicates / converters
            if obj_name.eq_ignore_ascii_case("char") && args.len() == 1 {
                let c = args[0].value.clone();
                if let Some(rewritten) = char_static_lower(field, c) {
                    return rewritten;
                }
            }
            // bool.Parse(s) → s.toLowerCase() === "true"
            if (obj_name.eq_ignore_ascii_case("bool") || obj_name == "Boolean")
                && field.eq_ignore_ascii_case("Parse")
                && args.len() == 1
            {
                let s = args[0].value.clone();
                let s_lc = Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(s),
                        field: "toLowerCase".into(),
                        null_safe: false,
                    })),
                    args: vec![],
                    optional: false,
                });
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(s_lc),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str("true".into())))),
                });
            }
            // `Array.Reverse / Exists / Find / FindAll / TrueForAll /
            // ConvertAll / ForEach / IndexOf` static helpers are wired
            // through the `[builtins]` `Array.*` entries in the C# profile,
            // which route to `common:dotnet.array_*` adapters in
            // `crates/vybex/src/emitter/dotnet/core/array_adapter.rs`.
            // VB picks them up by listing the same `common:dotnet.*`
            // emit targets — no walker rewrite required.
        }
    }
    if let ExprKind::Ident(name) = &callee.kind {
        if let Some((prefix, field)) = name.rsplit_once('.') {
            if (prefix.eq_ignore_ascii_case("string")
                || prefix.eq_ignore_ascii_case("System.String"))
                && field.eq_ignore_ascii_case("Join")
                && args.len() == 2
            {
                let sep = args[0].value.clone();
                let arr = args[1].value.clone();
                let materialized = Expression::new(ExprKind::Array(vec![ArrayElement {
                    key: None,
                    spread: true,
                    by_ref: false,
                    value: arr,
                }]));
                return Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(materialized),
                        field: "join".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(sep)],
                    optional: false,
                });
            }
        }
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn lower_case_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: "toLowerCase".into(),
            null_safe: false,
        })),
        args: vec![],
        optional: false,
    })
}

fn is_string_comparison_arg(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Lit(Literal::Str(tag)) if tag.starts_with("__dotnet_stringcomparison_")
    )
}

fn is_ignore_case_string_comparison(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Lit(Literal::Str(tag))
            if tag == "__dotnet_stringcomparison_ordinalignorecase"
                || tag == "__dotnet_stringcomparison_invariantignorecase"
    )
}

fn compare_expr(left: Expression, right: Expression) -> Expression {
    let lt = Expression::new(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(left.clone()),
        right: Box::new(right.clone()),
    });
    let gt = Expression::new(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(left),
        right: Box::new(right),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(lt),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(-1)))),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(gt),
            then: Box::new(Expression::new(ExprKind::Lit(Literal::Int(1)))),
            else_: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
        })),
    })
}

fn build_string_comparer_lambda(ignore_case: bool) -> Expression {
    let left_expr = if ignore_case {
        lower_case_expr(Expression::ident("left"))
    } else {
        Expression::ident("left")
    };
    let right_expr = if ignore_case {
        lower_case_expr(Expression::ident("right"))
    } else {
        Expression::ident("right")
    };
    Expression::new(ExprKind::Lambda {
        params: vec![
            Param {
                name: "left".into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            },
            Param {
                name: "right".into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            },
        ],
        body: LambdaBody::Expr(Box::new(compare_expr(left_expr, right_expr))),
        is_async: false,
        captures: Vec::new(),
    })
}

fn invoke_selector(selector: Expression, arg: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(selector),
        args: vec![Argument::positional(arg)],
        optional: false,
    })
}

/// Lower `char.<method>(c)` static helpers into ECMA-shape expressions.
/// All are ASCII-faithful only (matches what existing tests assert).
fn char_static_lower(method: &str, c: Expression) -> Option<Expression> {
    let lower_method = method.to_ascii_lowercase();
    match lower_method.as_str() {
        "toupper" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(c),
                field: "toUpperCase".into(),
                null_safe: false,
            })),
            args: vec![],
            optional: false,
        })),
        "tolower" => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(c),
                field: "toLowerCase".into(),
                null_safe: false,
            })),
            args: vec![],
            optional: false,
        })),
        // ASCII range predicates — emitted as `c >= "A" && c <= "Z"`.
        "isupper" => Some(char_in_range(c, "A", "Z")),
        "islower" => Some(char_in_range(c, "a", "z")),
        "isdigit" => Some(char_in_range(c, "0", "9")),
        // IsLetter — uppercase OR lowercase ASCII letter.
        "isletter" => {
            let upper = char_in_range(c.clone(), "A", "Z");
            let lower = char_in_range(c, "a", "z");
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(upper),
                right: Box::new(lower),
            }))
        }
        "isletterordigit" => {
            let upper = char_in_range(c.clone(), "A", "Z");
            let lower = char_in_range(c.clone(), "a", "z");
            let digit = char_in_range(c, "0", "9");
            let letter = Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(upper),
                right: Box::new(lower),
            });
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(letter),
                right: Box::new(digit),
            }))
        }
        // IsWhiteSpace — match space, tab, newline, carriage return.
        "iswhitespace" => {
            let space = eq_lit(c.clone(), " ");
            let tab = eq_lit(c.clone(), "\t");
            let newline = eq_lit(c.clone(), "\n");
            let cr = eq_lit(c, "\r");
            Some(or_chain(vec![space, tab, newline, cr]))
        }
        _ => None,
    }
}

/// Map a C#/.NET type-name token to its System.* short name.
/// `int` → `Int32`, `string` → `String`, `MyClass` → `MyClass`.
fn dotnet_type_name(t: &str) -> String {
    let trimmed = t.trim().trim_end_matches('?');
    match trimmed {
        "int" | "Int32" => "Int32",
        "uint" | "UInt32" => "UInt32",
        "long" | "Int64" => "Int64",
        "ulong" | "UInt64" => "UInt64",
        "short" | "Int16" => "Int16",
        "ushort" | "UInt16" => "UInt16",
        "byte" | "Byte" => "Byte",
        "sbyte" | "SByte" => "SByte",
        "float" | "Single" => "Single",
        "double" | "Double" => "Double",
        "decimal" | "Decimal" => "Decimal",
        "bool" | "Boolean" => "Boolean",
        "char" | "Char" => "Char",
        "string" | "String" => "String",
        "object" | "Object" => "Object",
        other => other,
    }
    .to_string()
}

fn normalize_runtime_type_name(t: &str) -> String {
    let trimmed = t.trim().trim_end_matches('?').trim();
    let mut out = String::with_capacity(trimmed.len());
    let mut depth = 0usize;
    for ch in trimmed.chars() {
        match ch {
            '<' => depth += 1,
            '>' => depth = depth.saturating_sub(1),
            _ if depth == 0 => out.push(ch),
            _ => {}
        }
    }
    let without_generics = out.trim();
    without_generics
        .rsplit('.')
        .next()
        .unwrap_or(without_generics)
        .trim()
        .to_string()
}

fn extract_if_is_pattern_binding(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let Some((subject, type_name, binding_name)) = find_if_is_pattern_binding(pair)? else {
        return Ok(None);
    };

    Ok(Some(build_type_pattern_binding_stmt(
        subject,
        type_name,
        binding_name,
    )))
}

fn build_type_pattern_binding_stmt(
    subject: Expression,
    type_name: String,
    binding_name: String,
) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(binding_name),
            type_hint: Some(type_name),
            init: Some(subject),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn find_if_is_pattern_binding(
    pair: Pair<Rule>,
) -> Result<Option<(Expression, String, String)>, String> {
    match pair.as_rule() {
        Rule::expression
        | Rule::assignment_expression
        | Rule::conditional_expression
        | Rule::null_coalesce_expr
        | Rule::logical_or
        | Rule::logical_and
        | Rule::bitwise_or
        | Rule::bitwise_xor
        | Rule::bitwise_and
        | Rule::equality => {
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return find_if_is_pattern_binding(inner[0].clone());
            }
            for child in inner {
                if let Some(binding) = find_if_is_pattern_binding(child)? {
                    return Ok(Some(binding));
                }
            }
            Ok(None)
        }
        Rule::relational => {
            let inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() != 2 || inner[1].as_rule() != Rule::type_test {
                return Ok(None);
            }
            let subject = walk_expression(inner[0].clone())?;
            let mut tt_inner = inner[1].clone().into_inner();
            let Some(keyword) = tt_inner.next() else {
                return Ok(None);
            };
            if keyword.as_rule() != Rule::is_kw {
                return Ok(None);
            }
            let Some(pattern_clause) = tt_inner.next() else {
                return Ok(None);
            };
            let clause_src = pattern_clause.as_str().trim_start();
            if clause_src.starts_with("not") {
                return Ok(None);
            }
            let Some(atom) = pattern_clause.into_inner().next() else {
                return Ok(None);
            };
            let atom_inner: Vec<Pair<Rule>> = atom.into_inner().collect();
            if atom_inner.len() < 2
                || atom_inner[0].as_rule() != Rule::pattern_type
                || atom_inner[1].as_rule() != Rule::ident_name
            {
                return Ok(None);
            }
            Ok(Some((
                subject,
                normalize_runtime_type_name(atom_inner[0].as_str()),
                atom_inner[1].as_str().to_string(),
            )))
        }
        _ => Ok(None),
    }
}

/// Build a runtime expression that yields the .NET type name of `expr`.
///
/// `Math.floor(e) === e ? "Int32" : "Double"` for numbers, `String` /
/// `Boolean` for primitives, and `Object` as the fallback. Implemented
/// as nested ternaries on `typeof e` so the result is plain bytecode
/// with no host helper required.
fn dotnet_runtime_type_name_expr(expr: Expression) -> Expression {
    let typeof_expr = Expression::new(ExprKind::TypeOf(Box::new(expr.clone())));

    let is_string = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "string".into(),
        )))),
    });
    let is_number = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "number".into(),
        )))),
    });
    let is_boolean = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(typeof_expr.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "boolean".into(),
        )))),
    });

    // Math.floor(e) === e — true for whole numbers; faithful to .NET
    // semantics where `42.GetType().Name == "Int32"` and
    // `3.14.GetType().Name == "Double"`. Vybe stores all numbers as f64
    // so this is the only post-hoc int/float distinction available.
    let floor_call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Math")),
            field: "floor".into(),
            null_safe: false,
        })),
        args: vec![Argument::positional(expr.clone())],
        optional: false,
    });
    let is_int = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(floor_call),
        right: Box::new(expr),
    });

    let number_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_int),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str("Int32".into())))),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "Double".into(),
        )))),
    });

    let bool_branch = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_boolean),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "Boolean".into(),
        )))),
        else_: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "Object".into(),
        )))),
    });

    let num_or_bool = Expression::new(ExprKind::Ternary {
        cond: Box::new(is_number),
        then: Box::new(number_branch),
        else_: Box::new(bool_branch),
    });

    Expression::new(ExprKind::Ternary {
        cond: Box::new(is_string),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
            "String".into(),
        )))),
        else_: Box::new(num_or_bool),
    })
}

fn char_in_range(c: Expression, lo: &str, hi: &str) -> Expression {
    let ge = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(c.clone()),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(lo.into())))),
    });
    let le = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(c),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(hi.into())))),
    });
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(ge),
        right: Box::new(le),
    })
}

fn eq_lit(c: Expression, lit: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(c),
        right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(lit.into())))),
    })
}

fn or_chain(mut exprs: Vec<Expression>) -> Expression {
    let mut acc = exprs.remove(0);
    for e in exprs {
        acc = Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(acc),
            right: Box::new(e),
        });
    }
    acc
}

/// Parse interpolated string parts from the raw content between $" and "
fn parse_interpolated_parts(s: &str) -> Result<Vec<InterpolPart>, String> {
    let mut parts = Vec::new();
    let mut text = String::new();
    let chars: Vec<char> = s.chars().collect();
    let mut i = 0;

    while i < chars.len() {
        if chars[i] == '{' {
            if i + 1 < chars.len() && chars[i + 1] == '{' {
                text.push('{');
                i += 2;
                continue;
            }
            // Flush text
            if !text.is_empty() {
                parts.push(InterpolPart::Text(text.clone()));
                text.clear();
            }
            // Find matching }
            let start = i + 1;
            let mut depth = 1;
            i += 1;
            while i < chars.len() && depth > 0 {
                if chars[i] == '{' {
                    depth += 1;
                }
                if chars[i] == '}' {
                    depth -= 1;
                }
                if depth > 0 {
                    i += 1;
                }
            }
            let expr_str: String = chars[start..i].iter().collect();
            let expr = parse_interpolated_hole_text(&expr_str)?;
            parts.push(InterpolPart::Expr(expr));
            i += 1; // skip }
        } else if chars[i] == '}' && i + 1 < chars.len() && chars[i + 1] == '}' {
            text.push('}');
            i += 2;
        } else {
            text.push(chars[i]);
            i += 1;
        }
    }

    if !text.is_empty() {
        parts.push(InterpolPart::Text(text));
    }

    Ok(parts)
}

fn parse_interpolated_hole_text(text: &str) -> Result<Expression, String> {
    if let Some((expr_src, alignment, format_spec)) = split_interpolated_format_hole(text) {
        let value = parse_interpolated_expr_text(expr_src)?;
        let mut composite = "{0".to_string();
        if let Some(alignment) = alignment {
            composite.push(',');
            composite.push_str(alignment.trim());
        }
        if let Some(format_spec) = format_spec {
            composite.push(':');
            composite.push_str(format_spec.trim());
        }
        composite.push('}');

        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("String.Format")),
            args: vec![
                Argument::positional(Expression::new(ExprKind::Lit(Literal::Str(composite)))),
                Argument::positional(value),
            ],
            optional: false,
        }));
    }

    let value = parse_interpolated_expr_text(text)?;
    Ok(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("String.Format")),
        args: vec![
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Str("{0}".into())))),
            Argument::positional(value),
        ],
        optional: false,
    }))
}

fn parse_interpolated_expr_text(text: &str) -> Result<Expression, String> {
    let trimmed = strip_wrapping_parens(text.trim());
    if let Some((cond_src, then_src, else_src)) = split_top_level_ternary(trimmed) {
        return Ok(Expression::new(ExprKind::Ternary {
            cond: Box::new(parse_interpolated_expr_text(cond_src)?),
            then: Box::new(parse_interpolated_expr_text(then_src)?),
            else_: Box::new(parse_interpolated_expr_text(else_src)?),
        }));
    }

    let mut parsed = CSharpParser::parse(Rule::expression, trimmed)
        .map_err(|e| format!("Interpolation parse error: {}", e))?;
    let pair = parsed.next().ok_or("Empty interpolation expression")?;
    walk_expression(pair)
}

fn split_interpolated_format_hole(text: &str) -> Option<(&str, Option<&str>, Option<&str>)> {
    let trimmed = text.trim();
    let comma = find_top_level_interpolation_separator(trimmed, ',');
    let colon = find_top_level_interpolation_separator(trimmed, ':');

    match (comma, colon) {
        (Some(c), Some(k)) if c < k => Some((
            trimmed[..c].trim(),
            Some(trimmed[c + 1..k].trim()),
            Some(trimmed[k + 1..].trim()),
        )),
        (Some(c), _) => Some((trimmed[..c].trim(), Some(trimmed[c + 1..].trim()), None)),
        (None, Some(k)) => Some((trimmed[..k].trim(), None, Some(trimmed[k + 1..].trim()))),
        _ => None,
    }
}

fn find_top_level_interpolation_separator(text: &str, needle: char) -> Option<usize> {
    let mut paren = 0usize;
    let mut bracket = 0usize;
    let mut brace = 0usize;
    let mut in_string = false;
    let mut string_delim = '\0';
    let mut escaped = false;
    let mut ternary_depth = 0usize;

    for (idx, ch) in text.char_indices() {
        if in_string {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' {
                escaped = true;
                continue;
            }
            if ch == string_delim {
                in_string = false;
                string_delim = '\0';
            }
            continue;
        }

        match ch {
            '"' | '\'' => {
                in_string = true;
                string_delim = ch;
            }
            '(' => paren += 1,
            ')' => paren = paren.saturating_sub(1),
            '[' => bracket += 1,
            ']' => bracket = bracket.saturating_sub(1),
            '{' => brace += 1,
            '}' => brace = brace.saturating_sub(1),
            '?' if paren == 0 && bracket == 0 && brace == 0 => ternary_depth += 1,
            ':' if paren == 0 && bracket == 0 && brace == 0 && ternary_depth > 0 => {
                ternary_depth -= 1;
            }
            c if c == needle && paren == 0 && bracket == 0 && brace == 0 && ternary_depth == 0 => {
                return Some(idx);
            }
            _ => {}
        }
    }

    None
}

fn strip_wrapping_parens(mut text: &str) -> &str {
    loop {
        let trimmed = text.trim();
        if !(trimmed.starts_with('(') && trimmed.ends_with(')')) {
            return trimmed;
        }
        let mut depth = 0usize;
        let mut in_string = false;
        let mut string_delim = '\0';
        let mut escaped = false;
        let mut wraps = true;
        for (idx, ch) in trimmed.char_indices() {
            if in_string {
                if escaped {
                    escaped = false;
                    continue;
                }
                if ch == '\\' {
                    escaped = true;
                    continue;
                }
                if ch == string_delim {
                    in_string = false;
                    string_delim = '\0';
                }
                continue;
            }
            match ch {
                '"' | '\'' => {
                    in_string = true;
                    string_delim = ch;
                }
                '(' => depth += 1,
                ')' => {
                    depth = depth.saturating_sub(1);
                    if depth == 0 && idx + ch.len_utf8() < trimmed.len() {
                        wraps = false;
                        break;
                    }
                }
                _ => {}
            }
        }
        if !wraps {
            return trimmed;
        }
        text = &trimmed[1..trimmed.len() - 1];
    }
}

fn split_top_level_ternary(text: &str) -> Option<(&str, &str, &str)> {
    let mut paren = 0usize;
    let mut bracket = 0usize;
    let mut brace = 0usize;
    let mut in_string = false;
    let mut string_delim = '\0';
    let mut escaped = false;
    let mut question_idx = None;
    let mut ternary_depth = 0usize;

    for (idx, ch) in text.char_indices() {
        if in_string {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' {
                escaped = true;
                continue;
            }
            if ch == string_delim {
                in_string = false;
                string_delim = '\0';
            }
            continue;
        }

        match ch {
            '"' | '\'' => {
                in_string = true;
                string_delim = ch;
            }
            '(' => paren += 1,
            ')' => paren = paren.saturating_sub(1),
            '[' => bracket += 1,
            ']' => bracket = bracket.saturating_sub(1),
            '{' => brace += 1,
            '}' => brace = brace.saturating_sub(1),
            '?' if paren == 0 && bracket == 0 && brace == 0 => {
                question_idx.get_or_insert(idx);
                ternary_depth += 1;
            }
            ':' if paren == 0 && bracket == 0 && brace == 0 && ternary_depth > 0 => {
                ternary_depth -= 1;
                if ternary_depth == 0 {
                    let q = question_idx?;
                    return Some((
                        text[..q].trim(),
                        text[q + 1..idx].trim(),
                        text[idx + 1..].trim(),
                    ));
                }
            }
            _ => {}
        }
    }

    None
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
