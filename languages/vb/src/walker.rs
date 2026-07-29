use chrono::{Datelike, Duration, Local, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};

use super::{Rule, VbParser};
use vybe_ast::*;
use vybe_compiler::compiler::generics as common_generics;
use vybe_platform_dotnet::emitter::core::exceptions as dotnet_exceptions;
use vybe_platform_dotnet::emitter::core::lowering as dotnet_vb;

const VB_PARTIAL_METHOD_MARKER: &str = "__vb_partial_method_decl";

fn vb_decl_starts_with_partial(raw: &str) -> bool {
    let trimmed = raw.trim_start();
    let Some(prefix) = trimmed.get(..7) else {
        return false;
    };
    prefix.eq_ignore_ascii_case("partial")
        && trimmed[7..].chars().next().is_some_and(char::is_whitespace)
}

thread_local! {
    static VB_CUSTOM_EVENTS: std::cell::RefCell<HashMap<String, String>> =
        std::cell::RefCell::new(HashMap::new());
}

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let option_compare_text = vb_source_uses_option_compare_text(source);
    let normalized_source = normalize_vb_multi_field_lines(&normalize_vb_inline_query_clauses(
        &normalize_vb_option_directive_lines(source),
    ));
    let xml_namespaces = parse_vb_xml_namespace_imports(&normalized_source);
    VB_CUSTOM_EVENTS.with(|events| events.borrow_mut().clear());
    let pairs = VbParser::parse(Rule::program, &normalized_source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut pending_decorators: Vec<Expression> = Vec::new();

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::imports_statement => imports.push(parse_imports_statement(inner)?),
                Rule::option_directive => {}
                Rule::attribute_line => {
                    pending_decorators.extend(parse_vb_attribute_specs(inner.as_str()));
                }
                Rule::statement_line => {
                    for stmt_pair in inner.into_inner() {
                        if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI
                        {
                            continue;
                        }
                        let mut stmt = if stmt_pair.as_rule() == Rule::module_decl {
                            parse_module_decl(stmt_pair)?
                        } else if stmt_pair.as_rule() == Rule::namespace_decl {
                            parse_namespace_decl(stmt_pair)?
                        } else if stmt_pair.as_rule() == Rule::dim_statement {
                            parse_statement(stmt_pair)?
                        } else if let Some(decl_stmt) = try_parse_declaration(stmt_pair.clone())? {
                            decl_stmt
                        } else {
                            parse_statement(stmt_pair)?
                        };
                        apply_vb_pending_decorators(&mut stmt, &mut pending_decorators);
                        body.push(stmt);
                    }
                }
                Rule::NEWLINE | Rule::EOI => {}
                _ => {}
            }
        }
    }

    normalize_vb_partial_classes(&mut body);
    normalize_vb_implicit_method_self_classes(&mut body);

    let mut synthesized = dotnet_exceptions::synthesize_exception_classes();
    synthesized.extend(body);

    let mut module = Module {
        name: "main".into(),
        language: Lang::VB,
        body: synthesized,
        imports,
    };
    normalize_vb_type_hint_whitespace(&mut module);
    rewrite_vb_import_aliases(&mut module);
    normalize_vb_xml_surface(&mut module, xml_namespaces);
    normalize_vb_legacy_file_io(&mut module);
    normalize_vb_visualbasic_strings_calls(&mut module);
    normalize_vb_callbyname_calls(&mut module);
    normalize_vb_extension_method_calls(&mut module);
    normalize_vb_custom_event_calls(&mut module);
    normalize_vb_environment_properties(&mut module);
    normalize_vb_uri_instance_calls(&mut module);
    normalize_vb_trycast_known_locals(&mut module);
    normalize_vb_bitwise_logic(&mut module);
    normalize_vb_flags_enum_ops(&mut module);
    normalize_vb_anonymous_equals(&mut module);
    normalize_vb_generic_new_factory_calls(&mut module);
    normalize_vb_dotnet_collection_calls(&mut module);
    normalize_vb_nested_member_arg_calls(&mut module);
    normalize_vb_for_existing_loop_vars(&mut module);
    normalize_vb_array_paren_indexes(&mut module);
    normalize_vb_custom_collection_for_each(&mut module);
    normalize_vb_default_indexer_calls(&mut module);
    normalize_vb_stringbuilder_member_access(&mut module);
    normalize_vb_operator_calls(&mut module);
    normalize_vb_bitwise_logic(&mut module);
    if option_compare_text {
        normalize_vb_option_compare_text(&mut module);
    }
    normalize_vb_interface_dispatch_type_hints(&mut module);
    normalize_vb_char_storage_types(&mut module);
    normalize_vb_anonymous_equals(&mut module);
    Ok(module)
}

fn normalize_vb_type_hint_whitespace(module: &mut Module) {
    normalize_vb_type_hint_whitespace_statements(&mut module.body);
}

fn normalize_vb_for_existing_loop_vars(module: &mut Module) {
    normalize_vb_for_existing_loop_var_statements(&mut module.body, &mut HashSet::new());
}

fn normalize_vb_for_existing_loop_var_statements(
    body: &mut [Statement],
    declared: &mut HashSet<String>,
) {
    for stmt in body {
        normalize_vb_for_existing_loop_var_statement(stmt, declared);
    }
}

fn normalize_vb_for_existing_loop_var_statement(
    stmt: &mut Statement,
    declared: &mut HashSet<String>,
) {
    if let Some(rewritten) = lower_vb_for_existing_loop_var(stmt, declared) {
        *stmt = rewritten;
        normalize_vb_for_existing_loop_var_statement(stmt, declared);
        return;
    }
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_for_existing_loop_var_expr(init);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    declared.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(cond) = cond {
                normalize_vb_for_existing_loop_var_expr(cond);
            }
            if let Some(update) = update {
                normalize_vb_for_existing_loop_var_expr(update);
            }
            if let Some(init_stmt) = init {
                rewrite_vb_for_existing_loop_var_init(init_stmt, declared);
            }
            let mut scoped = declared.clone();
            if let Some(init_stmt) = init.as_deref() {
                collect_vb_for_declared_names(init_stmt, &mut scoped);
            }
            normalize_vb_for_existing_loop_var_statements(body, &mut scoped);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = HashSet::new();
            for param in params {
                scoped.insert(param.name.to_ascii_lowercase());
            }
            normalize_vb_for_existing_loop_var_statements(body, &mut scoped);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_for_existing_loop_var_member(member, declared);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_for_existing_loop_var_expr(cond);
            normalize_vb_for_existing_loop_var_statements(then_body, &mut declared.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_for_existing_loop_var_expr(elif_cond);
                normalize_vb_for_existing_loop_var_statements(elif_body, &mut declared.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_for_existing_loop_var_statements(else_body, &mut declared.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_for_existing_loop_var_expr(cond);
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
            if let Some(else_body) = else_body {
                normalize_vb_for_existing_loop_var_statements(else_body, &mut declared.clone());
            }
        }
        StmtKind::DoWhile { cond, body, .. } => {
            normalize_vb_for_existing_loop_var_expr(cond);
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_for_existing_loop_var_expr(iter);
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
            if let Some(else_body) = else_body {
                normalize_vb_for_existing_loop_var_statements(else_body, &mut declared.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_for_existing_loop_var_expr(when_clause);
                }
                normalize_vb_for_existing_loop_var_statements(
                    &mut catch.body,
                    &mut declared.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_for_existing_loop_var_statements(else_body, &mut declared.clone());
            }
            if let Some(finally) = finally {
                normalize_vb_for_existing_loop_var_statements(finally, &mut declared.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_for_existing_loop_var_expr(expr);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_for_existing_loop_var_expr(target);
            }
            normalize_vb_for_existing_loop_var_expr(value);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_for_existing_loop_var_expr(target);
            normalize_vb_for_existing_loop_var_expr(value);
        }
        _ => {}
    }
}

fn lower_vb_for_existing_loop_var(
    stmt: &Statement,
    declared: &HashSet<String>,
) -> Option<Statement> {
    let StmtKind::For {
        init: Some(init),
        cond,
        update,
        body,
    } = &stmt.kind
    else {
        return None;
    };
    if vb_body_has_for_control_transfer(body) {
        return None;
    }
    let StmtKind::VarDecl { declarations, kind } = &init.kind else {
        return None;
    };
    let first = declarations.first()?;
    let BindingPattern::Ident(name) = &first.pattern else {
        return None;
    };
    if first.type_hint.is_some() || !declared.contains(&name.to_ascii_lowercase()) {
        return None;
    }

    let mut block = Vec::new();
    block.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(name)],
        value: first.init.clone().unwrap_or_else(Expression::null),
    }));

    let remaining = declarations[1..].to_vec();
    if !remaining.is_empty() {
        block.push(Statement::new(StmtKind::VarDecl {
            declarations: remaining,
            kind: kind.clone(),
        }));
    }

    let mut while_body = body.clone();
    if let Some(update) = update.clone() {
        while_body.push(Statement::new(StmtKind::Expr(update)));
    }
    block.push(Statement::new(StmtKind::While {
        cond: cond.clone().unwrap_or_else(|| Expression::bool(true)),
        body: while_body,
        else_body: None,
    }));

    Some(Statement::new(StmtKind::Block(block)))
}

fn vb_body_has_for_control_transfer(body: &[Statement]) -> bool {
    body.iter().any(vb_stmt_has_for_control_transfer)
}

fn vb_stmt_has_for_control_transfer(stmt: &Statement) -> bool {
    match &stmt.kind {
        StmtKind::Break(BreakTarget::Kind(ExitKind::For))
        | StmtKind::Continue(ContinueTarget::Kind(ContinueKind::For)) => true,
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. } => vb_body_has_for_control_transfer(body),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            vb_body_has_for_control_transfer(then_body)
                || elifs
                    .iter()
                    .any(|(_, body)| vb_body_has_for_control_transfer(body))
                || else_body
                    .as_deref()
                    .is_some_and(vb_body_has_for_control_transfer)
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            vb_body_has_for_control_transfer(body)
                || catches
                    .iter()
                    .any(|catch| vb_body_has_for_control_transfer(&catch.body))
                || else_body
                    .as_deref()
                    .is_some_and(vb_body_has_for_control_transfer)
                || finally
                    .as_deref()
                    .is_some_and(vb_body_has_for_control_transfer)
        }
        _ => false,
    }
}

fn rewrite_vb_for_existing_loop_var_init(
    init_stmt: &mut Box<Statement>,
    declared: &HashSet<String>,
) {
    let StmtKind::VarDecl { declarations, kind } = &mut init_stmt.kind else {
        normalize_vb_for_existing_loop_var_statement(init_stmt, &mut declared.clone());
        return;
    };
    let Some(first) = declarations.first() else {
        return;
    };
    let BindingPattern::Ident(name) = &first.pattern else {
        return;
    };
    if first.type_hint.is_some() || !declared.contains(&name.to_ascii_lowercase()) {
        for decl in declarations {
            if let Some(init) = &mut decl.init {
                normalize_vb_for_existing_loop_var_expr(init);
            }
        }
        return;
    }
    let start = first.init.clone().unwrap_or_else(Expression::null);
    let target_name = name.clone();
    let remaining = declarations[1..].to_vec();
    let mut block = vec![Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&target_name)],
        value: start,
    })];
    if !remaining.is_empty() {
        block.push(Statement::new(StmtKind::VarDecl {
            declarations: remaining,
            kind: kind.clone(),
        }));
    }
    *init_stmt = Box::new(Statement::new(StmtKind::Block(block)));
}

fn collect_vb_for_declared_names(stmt: &Statement, declared: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let BindingPattern::Ident(name) = &decl.pattern {
                    declared.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::Block(stmts) => {
            for stmt in stmts {
                collect_vb_for_declared_names(stmt, declared);
            }
        }
        _ => {}
    }
}

fn normalize_vb_for_existing_loop_var_member(member: &mut ClassMember, declared: &HashSet<String>) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_for_existing_loop_var_statement(stmt, &mut declared.clone());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_for_existing_loop_var_statements(body, &mut declared.clone());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_for_existing_loop_var_statements(getter, &mut declared.clone());
            }
            if let Some(setter) = setter {
                normalize_vb_for_existing_loop_var_statements(
                    &mut setter.body,
                    &mut declared.clone(),
                );
            }
        }
        _ => {}
    }
}

fn normalize_vb_for_existing_loop_var_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_for_existing_loop_var_expr(callee);
            for arg in args {
                normalize_vb_for_existing_loop_var_expr(&mut arg.value);
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_for_existing_loop_var_expr(object),
        ExprKind::Index { object, index, .. } => {
            normalize_vb_for_existing_loop_var_expr(object);
            normalize_vb_for_existing_loop_var_expr(index);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_for_existing_loop_var_expr(left);
            normalize_vb_for_existing_loop_var_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_for_existing_loop_var_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_for_existing_loop_var_expr(cond);
            normalize_vb_for_existing_loop_var_expr(then);
            normalize_vb_for_existing_loop_var_expr(else_);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_for_existing_loop_var_expr(&mut item.value);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_vb_for_existing_loop_var_expr(item);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_for_existing_loop_var_expr(class);
            for arg in args {
                normalize_vb_for_existing_loop_var_expr(&mut arg.value);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_for_existing_loop_var_expr(target);
            normalize_vb_for_existing_loop_var_expr(value);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_for_existing_loop_var_expr(expr),
            LambdaBody::Block(body) => {
                normalize_vb_for_existing_loop_var_statements(body, &mut HashSet::new());
            }
        },
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) = part {
                    normalize_vb_for_existing_loop_var_expr(expr);
                }
            }
        }
        _ => {}
    }
}

fn normalize_vb_type_hint_whitespace_statements(body: &mut [Statement]) {
    for stmt in body {
        normalize_vb_type_hint_whitespace_statement(stmt);
    }
}

fn trim_vb_type_hint(type_hint: &mut Option<String>) {
    if let Some(value) = type_hint {
        let trimmed = value.trim();
        if trimmed.len() != value.len() {
            *value = trimmed.to_string();
        }
    }
}

fn normalize_vb_type_hint_whitespace_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                trim_vb_type_hint(&mut decl.type_hint);
            }
        }
        StmtKind::FunctionDecl {
            params,
            return_type,
            body,
            ..
        } => {
            for param in params {
                trim_vb_type_hint(&mut param.type_hint);
            }
            trim_vb_type_hint(return_type);
            normalize_vb_type_hint_whitespace_statements(body);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_type_hint_whitespace_member(member);
            }
        }
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_type_hint_whitespace_statements(body);
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_type_hint_whitespace_statements(then_body);
            for (_, body) in elifs {
                normalize_vb_type_hint_whitespace_statements(body);
            }
            if let Some(body) = else_body {
                normalize_vb_type_hint_whitespace_statements(body);
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            normalize_vb_type_hint_whitespace_statements(body);
            if let Some(body) = else_body {
                normalize_vb_type_hint_whitespace_statements(body);
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                normalize_vb_type_hint_whitespace_statement(init);
            }
            normalize_vb_type_hint_whitespace_statements(body);
        }
        StmtKind::ForIn {
            body, else_body, ..
        } => {
            normalize_vb_type_hint_whitespace_statements(body);
            if let Some(body) = else_body {
                normalize_vb_type_hint_whitespace_statements(body);
            }
        }
        _ => {}
    }
}

fn normalize_vb_type_hint_whitespace_member(member: &mut ClassMember) {
    match member {
        ClassMember::Field { type_hint, .. } => trim_vb_type_hint(type_hint),
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_type_hint_whitespace_statement(stmt);
        }
        ClassMember::Constructor { params, body, .. } => {
            for param in params {
                trim_vb_type_hint(&mut param.type_hint);
            }
            normalize_vb_type_hint_whitespace_statements(body);
        }
        ClassMember::Property {
            type_hint,
            getter,
            setter,
            ..
        } => {
            trim_vb_type_hint(type_hint);
            if let Some(getter) = getter {
                normalize_vb_type_hint_whitespace_statements(getter);
            }
            if let Some(setter) = setter {
                trim_vb_type_hint(&mut setter.param.type_hint);
                normalize_vb_type_hint_whitespace_statements(&mut setter.body);
            }
        }
        ClassMember::Event { type_hint, .. } => trim_vb_type_hint(type_hint),
        ClassMember::Const { type_hint, .. } => trim_vb_type_hint(type_hint),
        // VB declares no augmentations; carries no type hint to normalize.
        ClassMember::Augment(_) => {}
    }
}

fn normalize_vb_uri_instance_calls(module: &mut Module) {
    let mut locals = HashMap::new();
    normalize_vb_uri_instance_statements(&mut module.body, &mut locals);
}

fn normalize_vb_uri_instance_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_uri_instance_statement(stmt, locals);
    }
}

fn normalize_vb_uri_instance_statement(stmt: &mut Statement, locals: &mut HashMap<String, String>) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                let is_uri_relative = decl
                    .init
                    .as_ref()
                    .is_some_and(|init| vb_is_uri_make_relative_call(init, locals));
                if let Some(init) = &mut decl.init {
                    normalize_vb_uri_instance_expr(init, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let type_name = decl
                        .type_hint
                        .as_deref()
                        .map(vb_canonical_type_name)
                        .or_else(|| {
                            decl.init
                                .as_ref()
                                .and_then(|init| vb_infer_expr_type(init, locals))
                        });
                    if let Some(type_name) = type_name.filter(|type_name| {
                        matches!(
                            type_name.as_str(),
                            "Uri" | "Version" | "Stopwatch" | "TimeSpan" | "Task"
                        )
                    }) {
                        locals.insert(name.to_ascii_lowercase(), type_name);
                    }
                    if is_uri_relative {
                        locals.insert(name.to_ascii_lowercase(), "UriRelativeString".into());
                    }
                }
            }
        }
        StmtKind::Expr(expr) => normalize_vb_uri_instance_expr(expr, locals),
        StmtKind::Assign { targets, value } => {
            normalize_vb_uri_instance_expr(value, locals);
            for target in targets {
                normalize_vb_uri_instance_expr(target, locals);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_uri_instance_expr(cond, locals);
            let mut then_locals = locals.clone();
            normalize_vb_uri_instance_statements(then_body, &mut then_locals);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_uri_instance_expr(elif_cond, locals);
                let mut elif_locals = locals.clone();
                normalize_vb_uri_instance_statements(elif_body, &mut elif_locals);
            }
            if let Some(else_body) = else_body {
                let mut else_locals = locals.clone();
                normalize_vb_uri_instance_statements(else_body, &mut else_locals);
            }
        }
        StmtKind::While { cond, body, .. } => {
            normalize_vb_uri_instance_expr(cond, locals);
            let mut inner = locals.clone();
            normalize_vb_uri_instance_statements(body, &mut inner);
        }
        StmtKind::For { body, .. } | StmtKind::ForIn { body, .. } => {
            let mut inner = locals.clone();
            normalize_vb_uri_instance_statements(body, &mut inner);
        }
        StmtKind::FunctionDecl { body, .. } => {
            let mut inner = HashMap::new();
            normalize_vb_uri_instance_statements(body, &mut inner);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        let mut inner = HashMap::new();
                        normalize_vb_uri_instance_statement(stmt, &mut inner);
                    }
                    ClassMember::Constructor { body, .. } => {
                        let mut inner = HashMap::new();
                        normalize_vb_uri_instance_statements(body, &mut inner);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn normalize_vb_uri_instance_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if args.is_empty()
                    && matches!(
                        field.to_ascii_lowercase().as_str(),
                        "result" | "iscompleted" | "iscanceled"
                    )
                    && vb_infer_expr_type(object, locals).as_deref() == Some("Task")
                {
                    normalize_vb_uri_instance_expr(object, locals);
                    return;
                }
                if args.len() == 1
                    && field.eq_ignore_ascii_case("Result")
                    && vb_infer_expr_type(object, locals).as_deref() == Some("Task")
                {
                    let result_call = call_expr(
                        Expression::new(ExprKind::Member {
                            object: Box::new((**object).clone()),
                            field: field.clone(),
                            null_safe: false,
                        }),
                        vec![],
                    );
                    *expr = Expression::new(ExprKind::Index {
                        object: Box::new(result_call),
                        index: Box::new(args[0].value.clone()),
                        null_safe: false,
                    });
                    normalize_vb_uri_instance_expr(expr, locals);
                    return;
                }
            }
            normalize_vb_uri_instance_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_uri_instance_expr(&mut arg.value, locals);
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if args.is_empty()
                    && field.eq_ignore_ascii_case("ToString")
                    && vb_infer_expr_type(object, locals).as_deref() == Some("UriRelativeString")
                {
                    *expr = (**object).clone();
                    return;
                }
                if args.is_empty()
                    && matches!(
                        field.to_ascii_lowercase().as_str(),
                        "result" | "iscompleted" | "iscanceled"
                    )
                    && vb_infer_expr_type(object, locals).as_deref() == Some("Task")
                {
                    return;
                }
                if field.eq_ignore_ascii_case("StartNew")
                    && dotted_expr_name(object).is_some_and(|path| {
                        path.eq_ignore_ascii_case("Task.Factory")
                            || path.eq_ignore_ascii_case("System.Threading.Tasks.Task.Factory")
                    })
                {
                    *callee = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Task")),
                        field: "Run".into(),
                        null_safe: false,
                    }));
                    return;
                }
                if field.eq_ignore_ascii_case("WhenAny")
                    && !args.is_empty()
                    && dotted_expr_name(object).is_some_and(|path| {
                        path.eq_ignore_ascii_case("Task")
                            || path.eq_ignore_ascii_case("System.Threading.Tasks.Task")
                    })
                {
                    *expr = call_expr(
                        Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident("Task")),
                            field: "FromResult".into(),
                            null_safe: false,
                        }),
                        vec![Argument::positional(args[0].value.clone())],
                    );
                    normalize_vb_uri_instance_expr(expr, locals);
                    return;
                }
                if args.len() == 1 {
                    let receiver_type = vb_infer_expr_type(object, locals);
                    if receiver_type.as_deref() == Some("Uri")
                        && field.eq_ignore_ascii_case("IsBaseOf")
                    {
                        let child_href = vb_uri_href_expr(args[0].value.clone());
                        let base_href = vb_uri_href_expr((**object).clone());
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(child_href),
                                field: "StartsWith".into(),
                                null_safe: false,
                            }),
                            vec![Argument::positional(base_href)],
                        );
                        return;
                    }
                    if receiver_type.as_deref() == Some("Uri")
                        && field.eq_ignore_ascii_case("MakeRelativeUri")
                    {
                        let target_href = vb_uri_href_expr(args[0].value.clone());
                        let base_href = vb_uri_href_expr((**object).clone());
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(target_href),
                                field: "Replace".into(),
                                null_safe: false,
                            }),
                            vec![
                                Argument::positional(base_href),
                                Argument::positional(Expression::string("")),
                            ],
                        );
                        return;
                    }
                    if receiver_type.as_deref() == Some("Version")
                        && field.eq_ignore_ascii_case("CompareTo")
                    {
                        let receiver = (**object).clone();
                        let first = args[0].value.clone();
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident("Version")),
                                field: "CompareTo".into(),
                                null_safe: false,
                            }),
                            vec![Argument::positional(first), Argument::positional(receiver)],
                        );
                        return;
                    }
                    let static_owner = match (
                        receiver_type.as_deref(),
                        field.to_ascii_lowercase().as_str(),
                    ) {
                        (Some("Uri"), "isbaseof" | "makerelativeuri") => Some("Uri"),
                        _ => None,
                    };
                    if let Some(static_owner) = static_owner {
                        let method = field.clone();
                        let receiver = (**object).clone();
                        let first = args[0].value.clone();
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident(static_owner)),
                                field: method,
                                null_safe: false,
                            }),
                            vec![Argument::positional(receiver), Argument::positional(first)],
                        );
                    }
                }
            }
        }
        ExprKind::Member { object, field, .. } => {
            if field.eq_ignore_ascii_case("Result") {
                if let ExprKind::Call { callee, args, .. } = &object.kind {
                    if let ExprKind::Member {
                        object: task_object,
                        field: method,
                        ..
                    } = &callee.kind
                    {
                        if method.eq_ignore_ascii_case("WhenAny")
                            && !args.is_empty()
                            && dotted_expr_name(task_object).is_some_and(|path| {
                                path.eq_ignore_ascii_case("Task")
                                    || path.eq_ignore_ascii_case("System.Threading.Tasks.Task")
                            })
                        {
                            *expr = args[0].value.clone();
                            normalize_vb_uri_instance_expr(expr, locals);
                            return;
                        }
                    }
                }
            }
            normalize_vb_uri_instance_expr(object, locals);
            if matches!(
                field.to_ascii_lowercase().as_str(),
                "result" | "iscompleted" | "iscanceled"
            ) && vb_infer_expr_type(object, locals).as_deref() == Some("Task")
            {
                *expr = call_expr(
                    Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: field.clone(),
                        null_safe: false,
                    }),
                    vec![],
                );
            }
        }
        ExprKind::Binary { op, left, right } => {
            if let Some(rewritten) = rewrite_vb_version_compareto_binary(op, left, right, locals) {
                *expr = rewritten;
                normalize_vb_uri_instance_expr(expr, locals);
                return;
            }
            if let Some(rewritten) = rewrite_vb_version_binary(*op, left, right, locals) {
                *expr = rewritten;
                normalize_vb_uri_instance_expr(expr, locals);
                return;
            }
            normalize_vb_uri_instance_expr(left, locals);
            normalize_vb_uri_instance_expr(right, locals);
            if let ExprKind::Binary { op, left, right } = &expr.kind {
                if let Some(rewritten) =
                    rewrite_vb_runtime_timespan_binary(*op, left, right, locals)
                {
                    *expr = rewritten;
                    normalize_vb_uri_instance_expr(expr, locals);
                }
            }
        }
        ExprKind::Unary { expr: inner, .. } => normalize_vb_uri_instance_expr(inner, locals),
        ExprKind::Index { object, index, .. } => {
            normalize_vb_uri_instance_expr(object, locals);
            normalize_vb_uri_instance_expr(index, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_uri_instance_expr(cond, locals);
            normalize_vb_uri_instance_expr(then, locals);
            normalize_vb_uri_instance_expr(else_, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    normalize_vb_uri_instance_expr(key, locals);
                }
                normalize_vb_uri_instance_expr(&mut item.value, locals);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_uri_instance_expr(item, locals);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_version_binary(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    if vb_infer_expr_type(left, locals).as_deref() != Some("Version")
        || vb_infer_expr_type(right, locals).as_deref() != Some("Version")
    {
        return None;
    }
    match op {
        BinOp::Lt => Some(vb_version_lex_compare_expr(
            left.clone(),
            right.clone(),
            BinOp::Lt,
        )),
        BinOp::Gt => Some(vb_version_lex_compare_expr(
            left.clone(),
            right.clone(),
            BinOp::Gt,
        )),
        BinOp::LtEq => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(vb_version_lex_compare_expr(
                left.clone(),
                right.clone(),
                BinOp::Gt,
            )),
        })),
        BinOp::GtEq => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(vb_version_lex_compare_expr(
                left.clone(),
                right.clone(),
                BinOp::Lt,
            )),
        })),
        BinOp::Eq | BinOp::StrictEq => Some(vb_version_all_fields_equal_expr(
            left.clone(),
            right.clone(),
        )),
        BinOp::NotEq | BinOp::StrictNotEq => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(vb_version_all_fields_equal_expr(
                left.clone(),
                right.clone(),
            )),
        })),
        _ => None,
    }
}

fn vb_uri_href_expr(object: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: "href".into(),
        null_safe: false,
    })
}

fn vb_is_uri_make_relative_call(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    if args.len() != 1 {
        return false;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return false;
    };
    field.eq_ignore_ascii_case("MakeRelativeUri")
        && (vb_infer_expr_type(object, locals).as_deref() == Some("Uri")
            || dotted_expr_name(object)
                .as_deref()
                .is_some_and(|name| vb_canonical_type_name(name) == "Uri"))
}

fn rewrite_vb_version_compareto_binary(
    op: &BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    if literal_i64(right) != Some(0) {
        return None;
    }
    let ExprKind::Call { callee, args, .. } = &left.kind else {
        return None;
    };
    if args.len() != 1 {
        return None;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !field.eq_ignore_ascii_case("CompareTo")
        || vb_infer_expr_type(object, locals).as_deref() != Some("Version")
    {
        return None;
    }
    let lhs = (**object).clone();
    let rhs = args[0].value.clone();
    match op {
        BinOp::Lt => Some(vb_version_lex_compare_expr(lhs, rhs, BinOp::Lt)),
        BinOp::Gt => Some(vb_version_lex_compare_expr(lhs, rhs, BinOp::Gt)),
        BinOp::Eq => Some(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(lhs),
                field: "Equals".into(),
                null_safe: false,
            })),
            args: vec![Argument::positional(rhs)],
            optional: false,
        })),
        BinOp::NotEq => Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(lhs),
                    field: "Equals".into(),
                    null_safe: false,
                })),
                args: vec![Argument::positional(rhs)],
                optional: false,
            })),
        })),
        _ => None,
    }
}

fn vb_version_lex_compare_expr(lhs: Expression, rhs: Expression, op: BinOp) -> Expression {
    let mut result = None;
    for field in ["Revision", "Build", "Minor", "Major"] {
        let left_field = Expression::new(ExprKind::Member {
            object: Box::new(lhs.clone()),
            field: field.into(),
            null_safe: false,
        });
        let right_field = Expression::new(ExprKind::Member {
            object: Box::new(rhs.clone()),
            field: field.into(),
            null_safe: false,
        });
        let cmp = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left_field.clone()),
            right: Box::new(right_field.clone()),
        });
        let eq = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(left_field),
            right: Box::new(right_field),
        });
        result = Some(match result {
            Some(next) => Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(cmp),
                right: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(eq),
                    right: Box::new(next),
                })),
            }),
            None => cmp,
        });
    }
    result.unwrap_or_else(|| Expression::bool(false))
}

fn vb_version_all_fields_equal_expr(lhs: Expression, rhs: Expression) -> Expression {
    let mut result = None;
    for field in ["Revision", "Build", "Minor", "Major"] {
        let eq = Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(lhs.clone()),
                field: field.into(),
                null_safe: false,
            })),
            right: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(rhs.clone()),
                field: field.into(),
                null_safe: false,
            })),
        });
        result = Some(match result {
            Some(prev) => Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(eq),
                right: Box::new(prev),
            }),
            None => eq,
        });
    }
    result.unwrap_or_else(|| Expression::bool(true))
}

fn normalize_vb_environment_properties(module: &mut Module) {
    normalize_vb_environment_property_statements(&mut module.body);
}

fn normalize_vb_environment_property_statements(body: &mut [Statement]) {
    for stmt in body {
        normalize_vb_environment_property_statement(stmt);
    }
}

fn normalize_vb_environment_property_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => normalize_vb_environment_property_expr(expr),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_environment_property_expr(init);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            normalize_vb_environment_property_expr(value);
            if targets.len() == 1 && is_vb_environment_member(&targets[0], "ExitCode") {
                let call = call_expr(
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Environment")),
                        field: "SetExitCode".into(),
                        null_safe: false,
                    }),
                    vec![Argument::positional(value.clone())],
                );
                stmt.kind = StmtKind::Expr(call);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_environment_property_expr(target);
            normalize_vb_environment_property_expr(value);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_environment_property_expr(cond);
            normalize_vb_environment_property_statements(then_body);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_environment_property_expr(elif_cond);
                normalize_vb_environment_property_statements(elif_body);
            }
            if let Some(else_body) = else_body {
                normalize_vb_environment_property_statements(else_body);
            }
        }
        StmtKind::While { cond, body, .. } => {
            normalize_vb_environment_property_expr(cond);
            normalize_vb_environment_property_statements(body);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_vb_environment_property_statement(init);
            }
            if let Some(cond) = cond {
                normalize_vb_environment_property_expr(cond);
            }
            if let Some(update) = update {
                normalize_vb_environment_property_expr(update);
            }
            normalize_vb_environment_property_statements(body);
        }
        StmtKind::ForIn { iter, body, .. } => {
            normalize_vb_environment_property_expr(iter);
            normalize_vb_environment_property_statements(body);
        }
        StmtKind::Return(Some(expr)) => {
            normalize_vb_environment_property_expr(expr);
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                normalize_vb_environment_property_expr(expr);
            }
            if let Some(cause) = cause {
                normalize_vb_environment_property_expr(cause);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_environment_property_statements(body);
        }
        StmtKind::FunctionDecl { body, .. } => {
            normalize_vb_environment_property_statements(body);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        normalize_vb_environment_property_statement(stmt);
                    }
                    ClassMember::Constructor { body, .. } => {
                        normalize_vb_environment_property_statements(body);
                    }
                    ClassMember::Field {
                        init: Some(init), ..
                    } => {
                        normalize_vb_environment_property_expr(init);
                    }
                    ClassMember::Const { value, .. } => {
                        normalize_vb_environment_property_expr(value);
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            normalize_vb_environment_property_statements(getter);
                        }
                        if let Some(setter) = setter {
                            normalize_vb_environment_property_statements(&mut setter.body);
                        }
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn is_vb_environment_member(expr: &Expression, field_name: &str) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, field, .. }
            if field.eq_ignore_ascii_case(field_name)
                && dotted_expr_name(object).is_some_and(|name| {
                    name.eq_ignore_ascii_case("Environment")
                        || name.eq_ignore_ascii_case("System.Environment")
                })
    )
}

fn normalize_vb_environment_property_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            normalize_vb_environment_property_expr(object);
            if field.eq_ignore_ascii_case("Version") || field.eq_ignore_ascii_case("ExitCode") {
                if dotted_expr_name(object).is_some_and(|name| {
                    name.eq_ignore_ascii_case("Environment")
                        || name.eq_ignore_ascii_case("System.Environment")
                }) {
                    let method = field.clone();
                    *expr = call_expr(
                        Expression::new(ExprKind::Member {
                            object: Box::new((**object).clone()),
                            field: method,
                            null_safe: false,
                        }),
                        Vec::new(),
                    );
                }
            }
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_environment_property_expr(callee);
            for arg in args {
                normalize_vb_environment_property_expr(&mut arg.value);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_vb_environment_property_expr(left);
            normalize_vb_environment_property_expr(right);
        }
        ExprKind::Unary { expr: inner, .. } => normalize_vb_environment_property_expr(inner),
        ExprKind::Assign { target, value } => {
            normalize_vb_environment_property_expr(value);
            if is_vb_environment_member(target, "ExitCode") {
                *expr = call_expr(
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Environment")),
                        field: "SetExitCode".into(),
                        null_safe: false,
                    }),
                    vec![Argument::positional((**value).clone())],
                );
            }
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_environment_property_expr(object);
            normalize_vb_environment_property_expr(index);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_environment_property_expr(cond);
            normalize_vb_environment_property_expr(then);
            normalize_vb_environment_property_expr(else_);
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    normalize_vb_environment_property_expr(key);
                }
                normalize_vb_environment_property_expr(&mut item.value);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_environment_property_expr(item);
            }
        }
        _ => {}
    }
}

fn normalize_vb_char_storage_types(module: &mut Module) {
    normalize_vb_char_storage_statements(&mut module.body);
}

fn normalize_vb_char_type_hint(type_hint: &mut Option<String>) {
    if type_hint
        .as_deref()
        .is_some_and(|hint| vb_canonical_type_name(hint) == "Char")
    {
        *type_hint = Some("String".into());
    }
}

fn normalize_vb_char_storage_statements(body: &mut [Statement]) {
    for stmt in body {
        normalize_vb_char_storage_statement(stmt);
    }
}

fn normalize_vb_char_storage_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                normalize_vb_char_type_hint(&mut decl.type_hint);
            }
        }
        StmtKind::FunctionDecl {
            params,
            return_type,
            body,
            ..
        } => {
            for param in params {
                normalize_vb_char_type_hint(&mut param.type_hint);
            }
            normalize_vb_char_type_hint(return_type);
            normalize_vb_char_storage_statements(body);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_char_storage_member(member);
            }
        }
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_char_storage_statements(body);
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_char_storage_statements(then_body);
            for (_, body) in elifs {
                normalize_vb_char_storage_statements(body);
            }
            if let Some(body) = else_body {
                normalize_vb_char_storage_statements(body);
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                normalize_vb_char_storage_statement(init);
            }
            normalize_vb_char_storage_statements(body);
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            normalize_vb_char_storage_statements(body);
            if let Some(body) = else_body {
                normalize_vb_char_storage_statements(body);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_char_storage_statements(body);
            for catch in catches {
                for ty in &mut catch.types {
                    if vb_canonical_type_name(ty) == "Char" {
                        *ty = "String".into();
                    }
                }
                normalize_vb_char_storage_statements(&mut catch.body);
            }
            if let Some(body) = else_body {
                normalize_vb_char_storage_statements(body);
            }
            if let Some(body) = finally {
                normalize_vb_char_storage_statements(body);
            }
        }
        _ => {}
    }
}

fn normalize_vb_char_storage_member(member: &mut ClassMember) {
    match member {
        ClassMember::Field { type_hint, .. } | ClassMember::Const { type_hint, .. } => {
            normalize_vb_char_type_hint(type_hint);
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_char_storage_statement(stmt);
        }
        ClassMember::Constructor { params, body, .. } => {
            for param in params {
                normalize_vb_char_type_hint(&mut param.type_hint);
            }
            normalize_vb_char_storage_statements(body);
        }
        ClassMember::Property {
            type_hint,
            getter,
            setter,
            ..
        } => {
            normalize_vb_char_type_hint(type_hint);
            if let Some(getter) = getter {
                normalize_vb_char_storage_statements(getter);
            }
            if let Some(setter) = setter {
                normalize_vb_char_type_hint(&mut setter.param.type_hint);
                normalize_vb_char_storage_statements(&mut setter.body);
            }
        }
        ClassMember::Event { .. } | ClassMember::Augment(_) => {}
    }
}

fn normalize_vb_multi_field_lines(source: &str) -> String {
    let mut out = String::with_capacity(source.len());
    for line in source.lines() {
        let trimmed = line.trim_start();
        let indent = &line[..line.len() - trimmed.len()];
        let lower = trimmed.to_ascii_lowercase();
        let is_field = ["public ", "private ", "friend ", "protected "]
            .iter()
            .any(|prefix| lower.starts_with(prefix));
        if is_field
            && trimmed.contains(',')
            && lower.contains(" as ")
            && !lower.contains(" sub ")
            && !lower.contains(" function ")
            && !lower.contains(" property ")
            && !lower.contains(" operator ")
        {
            let Some(as_idx) = lower.rfind(" as ") else {
                out.push_str(line);
                out.push('\n');
                continue;
            };
            let (left, right) = trimmed.split_at(as_idx);
            let type_suffix = right.trim_start();
            let mut parts = left.split_whitespace();
            let Some(modifier) = parts.next() else {
                out.push_str(line);
                out.push('\n');
                continue;
            };
            let names_text = left[modifier.len()..].trim();
            for name in names_text
                .split(',')
                .map(str::trim)
                .filter(|name| !name.is_empty())
            {
                out.push_str(indent);
                out.push_str(modifier);
                out.push(' ');
                out.push_str(name);
                out.push(' ');
                out.push_str(type_suffix);
                out.push('\n');
            }
        } else {
            out.push_str(line);
            out.push('\n');
        }
    }
    out
}

fn normalize_vb_option_directive_lines(source: &str) -> String {
    let mut normalized = source.to_string();
    for directive in [
        "Option Strict On:",
        "Option Strict Off:",
        "Option Explicit On:",
        "Option Explicit Off:",
        "Option Infer On:",
        "Option Infer Off:",
        "Option Compare Binary:",
        "Option Compare Text:",
    ] {
        normalized =
            normalized.replace(directive, &format!("{}\n", directive.trim_end_matches(':')));
    }
    normalized
}

fn vb_source_uses_option_compare_text(source: &str) -> bool {
    for line in source.lines() {
        let trimmed = line.trim_start();
        if trimmed.starts_with('\'') {
            continue;
        }
        let lower = trimmed.to_ascii_lowercase();
        if lower.starts_with("option compare text") {
            return true;
        }
        if lower.starts_with("option compare binary") {
            return false;
        }
    }
    false
}

fn normalize_vb_inline_query_clauses(source: &str) -> String {
    let mut out = String::with_capacity(source.len());
    for line in source.lines() {
        let lower = line.to_ascii_lowercase();
        let query_start = lower.find(" from ").map(|idx| idx + 1).or_else(|| {
            lower
                .trim_start()
                .starts_with("from ")
                .then_some(line.len() - line.trim_start().len())
        });
        let Some(start) = query_start else {
            out.push_str(line);
            out.push('\n');
            continue;
        };

        let indent_len = line.len() - line.trim_start().len();
        let indent = &line[..indent_len];
        let patterns = [
            " order by ",
            " skip while ",
            " take while ",
            " group join ",
            " group by ",
            " where ",
            " let ",
            " join ",
            " skip ",
            " take ",
            " distinct",
            " select ",
        ];
        let mut cursor = 0usize;
        let mut scan = start + "from ".len();
        let mut changed = false;
        while scan < lower.len() {
            let mut next: Option<(usize, &str)> = None;
            for pattern in patterns {
                if let Some(rel) = lower[scan..].find(pattern) {
                    let idx = scan + rel;
                    if next.map_or(true, |(best, _)| idx < best) {
                        next = Some((idx, pattern));
                    }
                }
            }
            let Some((idx, pattern)) = next else {
                break;
            };
            out.push_str(&line[cursor..idx]);
            out.push('\n');
            out.push_str(indent);
            cursor = idx + 1;
            scan = cursor + pattern.len().saturating_sub(1);
            changed = true;
        }

        if changed {
            out.push_str(&line[cursor..]);
            out.push('\n');
            if !lower[start..].contains(" select ")
                && !lower[start..].contains(" group ")
                && let Some(range_var) = vb_inline_query_range_var(line, start)
            {
                out.push_str(indent);
                out.push_str("Select ");
                out.push_str(&range_var);
                out.push('\n');
            }
        } else {
            out.push_str(line);
            out.push('\n');
        }
    }
    out
}

fn vb_inline_query_range_var(line: &str, from_start: usize) -> Option<String> {
    let rest = line.get(from_start + "from".len()..)?.trim_start();
    let name = rest
        .split(|ch: char| ch.is_whitespace() || ch == ',')
        .next()?
        .trim();
    (!name.is_empty()).then(|| name.to_string())
}

pub fn parse_expression_str(source: &str) -> Result<Expression, String> {
    let mut pairs =
        VbParser::parse(Rule::expression, source).map_err(|e| format!("Parse error: {}", e))?;
    let pair = pairs
        .next()
        .ok_or_else(|| "Missing VB expression".to_string())?;
    parse_expression(pair)
}

fn normalize_vb_identifier(name: &str) -> String {
    name.strip_prefix('[')
        .and_then(|s| s.strip_suffix(']'))
        .unwrap_or(name)
        .to_string()
}

#[derive(Clone, Debug, Default)]
struct VbXmlNormalizeState {
    element_locals: HashSet<String>,
    document_locals: HashSet<String>,
    sequence_locals: HashSet<String>,
    name_infos: HashMap<String, VbXmlNameInfo>,
    namespace_imports: HashMap<String, String>,
}

#[derive(Clone, Debug)]
struct VbXmlNameInfo {
    namespace_uri: String,
    local_name: String,
    prefix: String,
}

fn normalize_vb_xml_surface(module: &mut Module, namespace_imports: HashMap<String, String>) {
    let mut state = VbXmlNormalizeState {
        namespace_imports,
        ..VbXmlNormalizeState::default()
    };
    normalize_vb_xml_statements(&mut module.body, &mut state);
}

fn normalize_vb_visualbasic_strings_calls(module: &mut Module) {
    normalize_vb_visualbasic_strings_statements(&mut module.body);
}

fn normalize_vb_callbyname_calls(module: &mut Module) {
    normalize_vb_callbyname_statements(&mut module.body);
}

fn normalize_vb_callbyname_statements(body: &mut Vec<Statement>) {
    for stmt in body {
        normalize_vb_callbyname_statement(stmt);
    }
}

fn normalize_vb_callbyname_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            if let Some(replacement) = lower_vb_callbyname_statement(expr) {
                stmt.kind = replacement;
            } else {
                normalize_vb_callbyname_expr(expr);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Assign { value: expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => normalize_vb_callbyname_expr(expr),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_callbyname_expr(init);
                }
            }
        }
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::Try { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => normalize_vb_callbyname_statements(body),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_callbyname_statements(then_body);
            for (_, body) in elifs {
                normalize_vb_callbyname_statements(body);
            }
            if let Some(body) = else_body {
                normalize_vb_callbyname_statements(body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                normalize_vb_callbyname_statements(&mut case.body);
            }
            if let Some(body) = default {
                normalize_vb_callbyname_statements(body);
            }
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_callbyname_member(member);
            }
        }
        _ => {}
    }
}

fn normalize_vb_callbyname_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_callbyname_statement(stmt)
        }
        ClassMember::Constructor { body, .. } => normalize_vb_callbyname_statements(body),
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_callbyname_statements(getter);
            }
            if let Some(setter) = setter {
                normalize_vb_callbyname_statements(&mut setter.body);
            }
        }
        _ => {}
    }
}

fn lower_vb_callbyname_statement(expr: &mut Expression) -> Option<StmtKind> {
    let (object, member, call_type, mut rest) = parse_vb_callbyname_parts(expr)?;
    if call_type != "set" {
        return None;
    }
    let value = rest.pop()?.value;
    Some(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: member,
            null_safe: false,
        })],
        value,
    })
}

fn normalize_vb_callbyname_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_callbyname_expr(callee);
            for arg in args.iter_mut() {
                normalize_vb_callbyname_expr(&mut arg.value);
            }
            if let Some((object, member, call_type, rest)) = parse_vb_callbyname_parts(expr) {
                match call_type.as_str() {
                    "get" => {
                        *expr = Expression::new(ExprKind::Member {
                            object: Box::new(object),
                            field: member,
                            null_safe: false,
                        });
                    }
                    "method" => {
                        *expr = Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(object),
                                field: member,
                                null_safe: false,
                            })),
                            args: rest,
                            optional: false,
                        });
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_callbyname_expr(object),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_callbyname_expr(left);
            normalize_vb_callbyname_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => normalize_vb_callbyname_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_callbyname_expr(cond);
            normalize_vb_callbyname_expr(then);
            normalize_vb_callbyname_expr(else_);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_callbyname_expr(object);
            normalize_vb_callbyname_expr(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_callbyname_expr(&mut item.value);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_vb_callbyname_expr(item);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_callbyname_expr(class);
            for arg in args {
                normalize_vb_callbyname_expr(&mut arg.value);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_callbyname_expr(target);
            normalize_vb_callbyname_expr(value);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_callbyname_expr(expr),
            LambdaBody::Block(body) => normalize_vb_callbyname_statements(body),
        },
        ExprKind::ClassExpr { members, .. } => {
            for member in members {
                normalize_vb_callbyname_member(member);
            }
        }
        ExprKind::FunctionExpr(stmt) => normalize_vb_callbyname_statement(stmt),
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { value, .. }
                    | ObjectProperty::Computed { value, .. } => normalize_vb_callbyname_expr(value),
                    ObjectProperty::Spread(value) => normalize_vb_callbyname_expr(value),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_callbyname_statement(value);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn parse_vb_callbyname_parts(
    expr: &Expression,
) -> Option<(Expression, String, String, Vec<Argument>)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !dotted_expr_name(callee)
        .as_deref()
        .map(|name| name.eq_ignore_ascii_case("CallByName"))
        .unwrap_or(false)
        || args.len() < 3
    {
        return None;
    }
    let member = literal_string(&args.get(1)?.value)?;
    let call_type = vb_call_type_name(&args.get(2)?.value)?;
    Some((
        args.first()?.value.clone(),
        member,
        call_type,
        args.iter().skip(3).cloned().collect(),
    ))
}

fn vb_call_type_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("CallType")) => {
            Some(field.to_ascii_lowercase())
        }
        ExprKind::Ident(name) => Some(name.to_ascii_lowercase()),
        ExprKind::Lit(Literal::Str(value)) => Some(value.to_ascii_lowercase()),
        _ => None,
    }
}

fn normalize_vb_legacy_file_io(module: &mut Module) {
    normalize_vb_legacy_file_io_statements(&mut module.body);
}

fn normalize_vb_legacy_file_io_statements(body: &mut Vec<Statement>) {
    let mut out = Vec::with_capacity(body.len());
    for mut stmt in body.drain(..) {
        if let Some(mut replacement) = lower_vb_legacy_file_io_statement(stmt.clone()) {
            normalize_vb_legacy_file_io_statements(&mut replacement);
            out.extend(replacement);
            continue;
        }
        normalize_vb_legacy_file_io_statement(&mut stmt);
        out.push(stmt);
    }
    *body = out;
}

fn lower_vb_legacy_file_io_statement(stmt: Statement) -> Option<Vec<Statement>> {
    let span = stmt.span.clone();
    match stmt.kind {
        StmtKind::OpenFile {
            path,
            mode,
            file_number,
        } => Some(vec![Statement::with_span(
            StmtKind::Expr(vb_filesystem_call(
                "FileOpen",
                vec![
                    Argument::positional(file_number),
                    Argument::positional(path),
                    Argument::positional(Expression::string(vb_file_mode_name(mode))),
                ],
            )),
            span,
        )]),
        StmtKind::CloseFile(file_number) => Some(vec![Statement::with_span(
            StmtKind::Expr(vb_filesystem_call(
                "FileClose",
                file_number.into_iter().map(Argument::positional).collect(),
            )),
            span,
        )]),
        StmtKind::PrintFile { file_number, items } => Some(vec![Statement::with_span(
            StmtKind::Expr(vb_filesystem_call(
                "PrintLine",
                std::iter::once(Argument::positional(file_number))
                    .chain(items.into_iter().map(Argument::positional))
                    .collect(),
            )),
            span,
        )]),
        StmtKind::WriteFile { file_number, items } => Some(vec![Statement::with_span(
            StmtKind::Expr(vb_filesystem_call(
                "WriteLine",
                std::iter::once(Argument::positional(file_number))
                    .chain(items.into_iter().map(Argument::positional))
                    .collect(),
            )),
            span,
        )]),
        StmtKind::LineInput {
            file_number,
            variable,
        } => Some(vec![Statement::with_span(
            StmtKind::Assign {
                targets: vec![Expression::ident(&variable)],
                value: vb_filesystem_call("LineInput", vec![Argument::positional(file_number)]),
            },
            span,
        )]),
        StmtKind::InputFile {
            file_number,
            variables,
        } => Some(
            variables
                .into_iter()
                .map(|target| {
                    Statement::with_span(
                        StmtKind::Assign {
                            targets: vec![target],
                            value: vb_filesystem_call(
                                "Input",
                                vec![Argument::positional(file_number.clone())],
                            ),
                        },
                        span.clone(),
                    )
                })
                .collect(),
        ),
        _ => None,
    }
}

fn normalize_vb_legacy_file_io_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::FunctionDecl { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::Try { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => normalize_vb_legacy_file_io_statements(body),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_legacy_file_io_statements(then_body);
            for (_, body) in elifs {
                normalize_vb_legacy_file_io_statements(body);
            }
            if let Some(body) = else_body {
                normalize_vb_legacy_file_io_statements(body);
            }
        }
        StmtKind::Switch { cases, default, .. } => {
            for case in cases {
                normalize_vb_legacy_file_io_statements(&mut case.body);
            }
            if let Some(body) = default {
                normalize_vb_legacy_file_io_statements(body);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_legacy_file_io_member(member);
            }
        }
        StmtKind::Expr(expr) => {
            if let Some(replacement) = lower_vb_input_call_statement(expr) {
                stmt.kind = replacement;
            } else {
                normalize_vb_legacy_file_io_expr(expr);
            }
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Assign { value: expr, .. }
        | StmtKind::CompoundAssign { value: expr, .. } => normalize_vb_legacy_file_io_expr(expr),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_legacy_file_io_expr(init);
                }
            }
        }
        _ => {}
    }
}

fn normalize_vb_legacy_file_io_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(method) => {
            normalize_vb_legacy_file_io_statement(method);
        }
        ClassMember::Constructor { body, .. } => normalize_vb_legacy_file_io_statements(body),
        ClassMember::Property { getter, setter, .. } => {
            if let Some(body) = getter {
                normalize_vb_legacy_file_io_statements(body);
            }
            if let Some(setter) = setter {
                normalize_vb_legacy_file_io_statements(&mut setter.body);
            }
        }
        _ => {}
    }
}

fn lower_vb_input_call_statement(expr: &mut Expression) -> Option<StmtKind> {
    let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
        return None;
    };
    if !callee_is_vb_filesystem_name(callee, "Input") || args.len() < 2 {
        return None;
    }
    let file_number = args.first()?.value.clone();
    let target = args.get(1)?.value.clone();
    Some(StmtKind::Assign {
        targets: vec![target],
        value: vb_filesystem_call("Input", vec![Argument::positional(file_number)]),
    })
}

fn normalize_vb_legacy_file_io_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_legacy_file_io_expr(callee);
            for arg in args.iter_mut() {
                normalize_vb_legacy_file_io_expr(&mut arg.value);
            }
            if let Some(name) = vb_filesystem_leaf_name(callee) {
                *callee = Box::new(build_dotted_expr(&format!(
                    "dotnet.Microsoft.VisualBasic.FileSystem.{}",
                    name
                )));
                if name.eq_ignore_ascii_case("FileOpen") && args.len() >= 3 {
                    if let Some(mode) = vb_open_mode_arg_to_string(&args[2].value) {
                        args[2].value = Expression::string(mode);
                    }
                }
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_legacy_file_io_expr(object),
        ExprKind::Binary { left, right, .. } => {
            normalize_vb_legacy_file_io_expr(left);
            normalize_vb_legacy_file_io_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr)) => normalize_vb_legacy_file_io_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_legacy_file_io_expr(cond);
            normalize_vb_legacy_file_io_expr(then);
            normalize_vb_legacy_file_io_expr(else_);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_legacy_file_io_expr(object);
            normalize_vb_legacy_file_io_expr(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_legacy_file_io_expr(&mut item.value);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_vb_legacy_file_io_expr(item);
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                normalize_vb_legacy_file_io_expr(&mut arg.value);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_legacy_file_io_expr(target);
            normalize_vb_legacy_file_io_expr(value);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_legacy_file_io_expr(expr),
            LambdaBody::Block(body) => normalize_vb_legacy_file_io_statements(body),
        },
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { value, .. }
                    | ObjectProperty::Computed { value, .. } => {
                        normalize_vb_legacy_file_io_expr(value);
                    }
                    ObjectProperty::Spread(value) => normalize_vb_legacy_file_io_expr(value),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_legacy_file_io_statement(value);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn vb_filesystem_call(name: &str, args: Vec<Argument>) -> Expression {
    call_expr(
        build_dotted_expr(&format!("dotnet.Microsoft.VisualBasic.FileSystem.{}", name)),
        args,
    )
}

fn vb_file_mode_name(mode: FileMode) -> &'static str {
    match mode {
        FileMode::Input => "Input",
        FileMode::Output => "Output",
        FileMode::Append => "Append",
        FileMode::Binary => "Binary",
        FileMode::Random => "Random",
    }
}

fn vb_filesystem_leaf_name(callee: &Expression) -> Option<&'static str> {
    const NAMES: &[&str] = &[
        "FreeFile",
        "FileOpen",
        "FileClose",
        "PrintLine",
        "WriteLine",
        "LineInput",
        "Input",
        "EOF",
        "LOF",
        "Loc",
        "FileAttr",
        "GetAttr",
        "SetAttr",
        "Seek",
        "Dir",
        "FileCopy",
        "Kill",
        "FileLen",
        "FileDateTime",
        "CurDir",
        "ChDir",
        "MkDir",
        "RmDir",
        "Name",
        "Rename",
        "Get",
        "Put",
    ];
    let name = dotted_expr_name(callee)?;
    let lower = name.to_ascii_lowercase();
    if lower.contains('.') {
        let qualified = lower.starts_with("microsoft.visualbasic.filesystem.")
            || lower.starts_with("dotnet.microsoft.visualbasic.filesystem.");
        if !qualified {
            return None;
        }
    }
    let leaf = lower.rsplit('.').next().unwrap_or(&lower);
    NAMES.iter().copied().find(|n| n.eq_ignore_ascii_case(leaf))
}

fn callee_is_vb_filesystem_name(callee: &Expression, expected: &str) -> bool {
    vb_filesystem_leaf_name(callee)
        .map(|name| name.eq_ignore_ascii_case(expected))
        .unwrap_or(false)
}

fn vb_open_mode_arg_to_string(expr: &Expression) -> Option<&'static str> {
    let name = dotted_expr_name(expr)?;
    match name.to_ascii_lowercase().as_str() {
        "openmode.input" | "microsoft.visualbasic.openmode.input" => Some("Input"),
        "openmode.output" | "microsoft.visualbasic.openmode.output" => Some("Output"),
        "openmode.append" | "microsoft.visualbasic.openmode.append" => Some("Append"),
        "openmode.binary" | "microsoft.visualbasic.openmode.binary" => Some("Binary"),
        "openmode.random" | "microsoft.visualbasic.openmode.random" => Some("Random"),
        _ => None,
    }
}

fn normalize_vb_visualbasic_strings_statements(body: &mut [Statement]) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
                normalize_vb_visualbasic_strings_expr(expr)
            }
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations {
                    if let Some(init) = &mut decl.init {
                        normalize_vb_visualbasic_strings_expr(init);
                    }
                }
            }
            StmtKind::Assign { targets, value } => {
                for target in targets {
                    normalize_vb_visualbasic_strings_expr(target);
                }
                normalize_vb_visualbasic_strings_expr(value);
            }
            StmtKind::CompoundAssign { target, value, .. } => {
                normalize_vb_visualbasic_strings_expr(target);
                normalize_vb_visualbasic_strings_expr(value);
            }
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body,
            } => {
                normalize_vb_visualbasic_strings_expr(cond);
                normalize_vb_visualbasic_strings_statements(then_body);
                for (elif_cond, elif_body) in elifs {
                    normalize_vb_visualbasic_strings_expr(elif_cond);
                    normalize_vb_visualbasic_strings_statements(elif_body);
                }
                if let Some(else_body) = else_body {
                    normalize_vb_visualbasic_strings_statements(else_body);
                }
            }
            StmtKind::For {
                init,
                cond,
                update,
                body,
            } => {
                if let Some(init) = init {
                    normalize_vb_visualbasic_strings_statements(std::slice::from_mut(init));
                }
                if let Some(cond) = cond {
                    normalize_vb_visualbasic_strings_expr(cond);
                }
                if let Some(update) = update {
                    normalize_vb_visualbasic_strings_expr(update);
                }
                normalize_vb_visualbasic_strings_statements(body);
            }
            StmtKind::ForIn {
                iter,
                body,
                else_body,
                ..
            } => {
                normalize_vb_visualbasic_strings_expr(iter);
                normalize_vb_visualbasic_strings_statements(body);
                if let Some(else_body) = else_body {
                    normalize_vb_visualbasic_strings_statements(else_body);
                }
            }
            StmtKind::While {
                cond,
                body,
                else_body,
            } => {
                normalize_vb_visualbasic_strings_expr(cond);
                normalize_vb_visualbasic_strings_statements(body);
                if let Some(else_body) = else_body {
                    normalize_vb_visualbasic_strings_statements(else_body);
                }
            }
            StmtKind::FunctionDecl { body, .. }
            | StmtKind::Block(body)
            | StmtKind::NamespaceDecl { body, .. } => {
                normalize_vb_visualbasic_strings_statements(body)
            }
            StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    normalize_vb_visualbasic_strings_member(member);
                }
            }
            _ => {}
        }
    }
}

fn normalize_vb_visualbasic_strings_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_visualbasic_strings_statements(std::slice::from_mut(stmt));
        }
        ClassMember::Constructor { body, .. } => normalize_vb_visualbasic_strings_statements(body),
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_visualbasic_strings_statements(getter);
            }
            if let Some(setter) = setter {
                normalize_vb_visualbasic_strings_statements(&mut setter.body);
            }
        }
        _ => {}
    }
}

fn normalize_vb_visualbasic_strings_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_visualbasic_strings_expr(callee);
            for arg in args {
                normalize_vb_visualbasic_strings_expr(&mut arg.value);
            }
        }
        ExprKind::Member { object, field, .. } => {
            normalize_vb_visualbasic_strings_expr(object);
            if matches!(
                field.to_ascii_lowercase().as_str(),
                "left" | "right" | "mid"
            ) && matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Strings"))
            {
                *object = Box::new(build_dotted_expr("dotnet.Microsoft.VisualBasic.Strings"));
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_visualbasic_strings_expr(left);
            normalize_vb_visualbasic_strings_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_visualbasic_strings_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_visualbasic_strings_expr(cond);
            normalize_vb_visualbasic_strings_expr(then);
            normalize_vb_visualbasic_strings_expr(else_);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_visualbasic_strings_expr(object);
            normalize_vb_visualbasic_strings_expr(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_visualbasic_strings_expr(&mut item.value);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_visualbasic_strings_expr(class);
            for arg in args {
                normalize_vb_visualbasic_strings_expr(&mut arg.value);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_visualbasic_strings_expr(expr),
            LambdaBody::Block(body) => normalize_vb_visualbasic_strings_statements(body),
        },
        ExprKind::ClassExpr { members, .. } => {
            for member in members {
                normalize_vb_visualbasic_strings_member(member);
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            normalize_vb_visualbasic_strings_statements(std::slice::from_mut(stmt));
        }
        ExprKind::Yield(Some(expr)) => normalize_vb_visualbasic_strings_expr(expr),
        ExprKind::Sequence(exprs) => {
            for expr in exprs {
                normalize_vb_visualbasic_strings_expr(expr);
            }
        }
        _ => {}
    }
}

fn normalize_vb_extension_method_calls(module: &mut Module) {
    let mut methods = HashMap::new();
    collect_vb_extension_methods(&module.body, &mut methods);
    if methods.is_empty() {
        return;
    }
    rewrite_vb_extension_method_statements(&mut module.body, &methods);
}

fn collect_vb_extension_methods(body: &[Statement], methods: &mut HashMap<String, String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name, modifiers, ..
            } if modifiers.is_extension => {
                methods.insert(name.to_ascii_lowercase(), name.clone());
            }
            StmtKind::FunctionDecl { body, .. }
            | StmtKind::Block(body)
            | StmtKind::NamespaceDecl { body, .. } => collect_vb_extension_methods(body, methods),
            StmtKind::ModuleDecl { name, members, .. } => {
                for member in members {
                    collect_vb_extension_methods_member(member, methods, Some(name));
                }
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                for member in members {
                    collect_vb_extension_methods_member(member, methods, None);
                }
            }
            _ => {}
        }
    }
}

fn collect_vb_extension_methods_member(
    member: &ClassMember,
    methods: &mut HashMap<String, String>,
    owner: Option<&str>,
) {
    match member {
        ClassMember::Method(stmt) => {
            if let StmtKind::FunctionDecl {
                name, modifiers, ..
            } = &stmt.kind
            {
                if modifiers.is_extension {
                    let target = owner
                        .map(|owner| format!("{owner}.{name}"))
                        .unwrap_or_else(|| name.clone());
                    methods.insert(name.to_ascii_lowercase(), target);
                }
            }
        }
        ClassMember::NestedType(stmt) => {
            collect_vb_extension_methods(std::slice::from_ref(stmt), methods);
        }
        ClassMember::Constructor { body, .. } => collect_vb_extension_methods(body, methods),
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                collect_vb_extension_methods(getter, methods);
            }
            if let Some(setter) = setter {
                collect_vb_extension_methods(&setter.body, methods);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_extension_method_statements(
    body: &mut [Statement],
    methods: &HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_extension_method_statement(stmt, methods);
    }
}

fn rewrite_vb_extension_method_statement(stmt: &mut Statement, methods: &HashMap<String, String>) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_extension_method_expr(init, methods);
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_extension_method_expr(expr, methods);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_extension_method_expr(target, methods);
            }
            rewrite_vb_extension_method_expr(value, methods);
        }
        StmtKind::If {
            cond,
            then_body,
            else_body,
            ..
        } => {
            rewrite_vb_extension_method_expr(cond, methods);
            rewrite_vb_extension_method_statements(then_body, methods);
            if let Some(else_body) = else_body {
                rewrite_vb_extension_method_statements(else_body, methods);
            }
        }
        StmtKind::While { cond, body, .. } => {
            rewrite_vb_extension_method_expr(cond, methods);
            rewrite_vb_extension_method_statements(body, methods);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_vb_extension_method_statement(init, methods);
            }
            if let Some(cond) = cond {
                rewrite_vb_extension_method_expr(cond, methods);
            }
            if let Some(update) = update {
                rewrite_vb_extension_method_expr(update, methods);
            }
            rewrite_vb_extension_method_statements(body, methods);
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_vb_extension_method_expr(iter, methods);
            rewrite_vb_extension_method_statements(body, methods);
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_extension_method_statements(body, methods)
        }
        StmtKind::ModuleDecl { members, .. }
        | StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. } => {
            for member in members {
                rewrite_vb_extension_method_member(member, methods);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_extension_method_member(member: &mut ClassMember, methods: &HashMap<String, String>) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_extension_method_statement(stmt, methods);
        }
        ClassMember::Constructor { body, .. } => {
            rewrite_vb_extension_method_statements(body, methods)
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_extension_method_statements(getter, methods);
            }
            if let Some(setter) = setter {
                rewrite_vb_extension_method_statements(&mut setter.body, methods);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_extension_method_expr(expr: &mut Expression, methods: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_extension_method_expr(callee, methods);
            for arg in &mut *args {
                rewrite_vb_extension_method_expr(&mut arg.value, methods);
            }
            if let ExprKind::Member {
                object,
                field,
                null_safe: false,
            } = &callee.kind
            {
                if let Some(target) = methods.get(&field.to_ascii_lowercase()) {
                    let receiver = (**object).clone();
                    args.insert(0, Argument::positional(receiver));
                    *callee = Box::new(build_dotted_expr(target));
                }
            }
        }
        ExprKind::Member { object, .. } => rewrite_vb_extension_method_expr(object, methods),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_extension_method_expr(left, methods);
            rewrite_vb_extension_method_expr(right, methods);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => rewrite_vb_extension_method_expr(expr, methods),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_extension_method_expr(cond, methods);
            rewrite_vb_extension_method_expr(then, methods);
            rewrite_vb_extension_method_expr(else_, methods);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_extension_method_expr(object, methods);
            rewrite_vb_extension_method_expr(index, methods);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_extension_method_expr(&mut item.value, methods);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_vb_extension_method_expr(key, methods);
                        rewrite_vb_extension_method_expr(value, methods);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_vb_extension_method_expr(value, methods)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_extension_method_statement(value, methods);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_extension_method_expr(class, methods);
            for arg in args {
                rewrite_vb_extension_method_expr(&mut arg.value, methods);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_vb_extension_method_expr(expr, methods),
            LambdaBody::Block(body) => rewrite_vb_extension_method_statements(body, methods),
        },
        _ => {}
    }
}

fn normalize_vb_xml_statements(body: &mut [Statement], state: &mut VbXmlNormalizeState) {
    for stmt in body {
        normalize_vb_xml_statement(stmt, state);
    }
}

fn normalize_vb_xml_statement(stmt: &mut Statement, state: &mut VbXmlNormalizeState) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_xml_expr(init, state);
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if vb_expr_is_xml_document_value(init) {
                            if decl.type_hint.is_none() {
                                decl.type_hint = Some("XDocument".to_string());
                            }
                            state.document_locals.insert(name.to_ascii_lowercase());
                        } else if vb_expr_is_xml_value(init) {
                            if decl.type_hint.is_none() {
                                decl.type_hint = Some("XElement".to_string());
                            }
                            let key = name.to_ascii_lowercase();
                            state.element_locals.insert(key.clone());
                            if vb_expr_is_xml_axis_result(init) {
                                state.sequence_locals.insert(key.clone());
                            }
                            if let Some(info) = vb_xml_name_info_from_expr(init, state) {
                                state.name_infos.insert(key, info);
                            }
                        }
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            normalize_vb_xml_expr(value, state);
            for target in targets {
                normalize_vb_xml_expr(target, state);
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => normalize_vb_xml_expr(expr, state),
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = state.clone();
            for param in params {
                if param
                    .type_hint
                    .as_deref()
                    .is_some_and(|hint| hint.to_ascii_lowercase().contains("xelement"))
                {
                    let lower = param.name.to_ascii_lowercase();
                    let hint = param
                        .type_hint
                        .as_deref()
                        .unwrap_or_default()
                        .to_ascii_lowercase();
                    if hint.contains("ienumerable") || hint.ends_with("()") {
                        scoped.sequence_locals.insert(lower);
                    } else {
                        scoped.element_locals.insert(lower);
                    }
                }
            }
            normalize_vb_xml_statements(body, &mut scoped);
        }
        StmtKind::Block(body) => {
            normalize_vb_xml_statements(body, &mut state.clone());
        }
        StmtKind::If {
            then_body,
            else_body,
            ..
        } => {
            normalize_vb_xml_statements(then_body, &mut state.clone());
            if let Some(else_body) = else_body {
                normalize_vb_xml_statements(else_body, &mut state.clone());
            }
        }
        StmtKind::While { body, .. } | StmtKind::For { body, .. } => {
            normalize_vb_xml_statements(body, &mut state.clone());
        }
        StmtKind::ForIn {
            var, iter, body, ..
        } => {
            normalize_vb_xml_expr(iter, state);
            let mut loop_state = state.clone();
            if vb_expr_is_known_xml_value(iter, state) {
                loop_state.element_locals.insert(var.to_ascii_lowercase());
            }
            normalize_vb_xml_statements(body, &mut loop_state);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_xml_member(member, state);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_xml_statements(body, &mut state.clone())
        }
        _ => {}
    }
}

fn normalize_vb_xml_member(member: &mut ClassMember, state: &VbXmlNormalizeState) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_xml_statement(stmt, &mut state.clone());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_xml_statements(body, &mut state.clone());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_xml_statements(getter, &mut state.clone());
            }
            if let Some(setter) = setter {
                normalize_vb_xml_statements(&mut setter.body, &mut state.clone());
            }
        }
        _ => {}
    }
}

fn normalize_vb_xml_expr(expr: &mut Expression, state: &mut VbXmlNormalizeState) {
    match &mut expr.kind {
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } => {
            normalize_vb_xml_expr(object, state);
            if vb_expr_is_xml_name_value(object) {
                if field.eq_ignore_ascii_case("LocalName") || field.eq_ignore_ascii_case("Local") {
                    *expr = call_expr(
                        build_dotted_expr("xml.local"),
                        vec![Argument::positional((**object).clone())],
                    );
                } else if field.eq_ignore_ascii_case("NamespaceName")
                    || field.eq_ignore_ascii_case("NamespaceURI")
                {
                    *expr = call_expr(
                        build_dotted_expr("xml.namespace"),
                        vec![Argument::positional((**object).clone())],
                    );
                }
                return;
            }
            if field.eq_ignore_ascii_case("Root") && vb_expr_is_known_xml_document(object, state) {
                *expr = Expression::new(ExprKind::Member {
                    object: Box::new((**object).clone()),
                    field: "documentElement".to_string(),
                    null_safe: false,
                });
                return;
            }
            if field.eq_ignore_ascii_case("Count") && vb_expr_is_xml_sequence_value(object, state) {
                *expr = Expression::new(ExprKind::Member {
                    object: Box::new((**object).clone()),
                    field: "length".to_string(),
                    null_safe: false,
                });
            } else if vb_expr_is_xml_axis_result(object) && field.eq_ignore_ascii_case("Value") {
                *expr = Expression::new(ExprKind::Member {
                    object: Box::new(Expression::new(ExprKind::Index {
                        object: Box::new((**object).clone()),
                        index: Box::new(Expression::int(0)),
                        null_safe: false,
                    })),
                    field: "textContent".to_string(),
                    null_safe: false,
                });
            } else if vb_expr_is_known_xml_value(object, state) {
                if field.eq_ignore_ascii_case("Value") {
                    *expr = Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: "textContent".to_string(),
                        null_safe: false,
                    });
                } else if field.eq_ignore_ascii_case("Name") {
                    *expr = vb_xml_name_expr_for_object(object, state).unwrap_or_else(|| {
                        call_expr(
                            build_dotted_expr("xml.node_name"),
                            vec![Argument::positional((**object).clone())],
                        )
                    });
                } else if field.eq_ignore_ascii_case("IsEmpty") {
                    *expr = Expression::new(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "childNodes".to_string(),
                                null_safe: false,
                            })),
                            field: "length".to_string(),
                            null_safe: false,
                        })),
                        right: Box::new(Expression::int(0)),
                    });
                } else if field.eq_ignore_ascii_case("FirstNode") {
                    *expr = Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: "firstChild".to_string(),
                        null_safe: false,
                    });
                }
            }
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_xml_expr(callee, state);
            for arg in &mut *args {
                normalize_vb_xml_expr(&mut arg.value, state);
            }
            if args.is_empty() {
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if field == "length" && vb_expr_is_xml_sequence_value(object, state) {
                        *expr = (**callee).clone();
                        return;
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("Select")
                    && vb_expr_is_xml_sequence_value(object, state)
                {
                    if let Some(first) = args.first_mut() {
                        normalize_vb_xml_lambda_element_body(&mut first.value, state);
                    }
                }
                if field.eq_ignore_ascii_case("Count")
                    && vb_expr_is_xml_sequence_value(object, state)
                {
                    *expr = Expression::new(ExprKind::Member {
                        object: Box::new((**object).clone()),
                        field: "length".to_string(),
                        null_safe: false,
                    });
                } else if field.eq_ignore_ascii_case("ToString")
                    && args.is_empty()
                    && vb_expr_is_known_xml_node(object, state)
                {
                    *expr = call_expr(
                        build_dotted_expr("xml.save"),
                        vec![Argument::positional((**object).clone())],
                    );
                }
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_xml_expr(left, state);
            normalize_vb_xml_expr(right, state);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_xml_expr(expr, state),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_xml_expr(cond, state);
            normalize_vb_xml_expr(then, state);
            normalize_vb_xml_expr(else_, state);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_xml_expr(object, state);
            normalize_vb_xml_expr(index, state);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_xml_expr(&mut item.value, state);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_xml_expr(key, state);
                        normalize_vb_xml_expr(value, state);
                    }
                    ObjectProperty::Spread(value) => normalize_vb_xml_expr(value, state),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_xml_statement(value, &mut state.clone());
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                normalize_vb_xml_expr(&mut arg.value, state);
            }
        }
        _ => {}
    }
}

fn normalize_vb_xml_lambda_element_body(expr: &mut Expression, state: &VbXmlNormalizeState) {
    let ExprKind::Lambda { params, body, .. } = &mut expr.kind else {
        return;
    };
    let mut scoped = state.clone();
    for param in params {
        scoped
            .element_locals
            .insert(param.name.to_ascii_lowercase());
    }
    match body {
        LambdaBody::Expr(expr) => normalize_vb_xml_expr(expr, &mut scoped),
        LambdaBody::Block(body) => normalize_vb_xml_statements(body, &mut scoped),
    }
}

fn vb_expr_is_known_xml_value(expr: &Expression, state: &VbXmlNormalizeState) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => state.element_locals.contains(&name.to_ascii_lowercase()),
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } if field == "documentElement" => vb_expr_is_known_xml_document(object, state),
        _ => vb_expr_is_xml_value(expr),
    }
}

fn vb_expr_is_known_xml_document(expr: &Expression, state: &VbXmlNormalizeState) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => state.document_locals.contains(&name.to_ascii_lowercase()),
        _ => vb_expr_is_xml_document_value(expr),
    }
}

fn vb_expr_is_known_xml_node(expr: &Expression, state: &VbXmlNormalizeState) -> bool {
    vb_expr_is_known_xml_value(expr, state)
        || matches!(
            &expr.kind,
            ExprKind::Member {
                field,
                null_safe: false,
                ..
            } if field == "firstChild" || field == "lastChild"
        )
}

fn vb_expr_is_xml_sequence_value(expr: &Expression, state: &VbXmlNormalizeState) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => state.sequence_locals.contains(&name.to_ascii_lowercase()),
        _ => vb_expr_is_xml_axis_result(expr),
    }
}

fn vb_expr_is_xml_value(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if field == "documentElement" => {
            matches!(
                &object.kind,
                ExprKind::Call { callee, .. }
                    if dotted_expr_name(callee).is_some_and(|name| name.eq_ignore_ascii_case("xml.parse"))
            )
        }
        ExprKind::Call {
            callee, args: _, ..
        } => {
            dotted_expr_name(callee).is_some_and(|name| {
                matches!(
                    name.to_ascii_lowercase().as_str(),
                    "xml.parse" | "xdocument.parse" | "xml.load" | "xdocument.load"
                )
            }) || matches!(
                &callee.kind,
                ExprKind::Member { object, field, .. }
                    if field.eq_ignore_ascii_case("Root")
                        || (field.eq_ignore_ascii_case("First") && vb_expr_is_xml_elements_sequence(object))
            ) || vb_expr_is_xml_axis_result(expr)
        }
        ExprKind::New { class, .. } => dotted_expr_name(class)
            .is_some_and(|name| name.ends_with("XElement") || name.ends_with("XDocument")),
        ExprKind::Index { object, .. } if vb_expr_is_xml_axis_result(object) => true,
        _ => false,
    }
}

fn vb_expr_is_xml_document_value(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if dotted_expr_name(callee).is_some_and(|name| {
                matches!(
                    name.to_ascii_lowercase().as_str(),
                    "xml.parse" | "xdocument.parse" | "xml.load" | "xdocument.load"
                )
            })
    ) || matches!(
        &expr.kind,
        ExprKind::New { class, .. }
            if dotted_expr_name(class).is_some_and(|name| name.ends_with("XDocument"))
    )
}

fn vb_expr_is_xml_name_value(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind, ExprKind::Member { object, field, .. }
                    if matches!(&object.kind, ExprKind::Ident(root) if root == "xml")
                        && (field == "node_name" || field == "name"))
    )
}

fn parse_vb_xml_namespace_imports(source: &str) -> HashMap<String, String> {
    let mut namespaces = HashMap::new();
    for line in source.lines() {
        let trimmed = line.trim();
        if !trimmed
            .get(..7)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("Imports"))
            || !trimmed.contains("<xmlns")
        {
            continue;
        }
        collect_vb_xml_namespaces_from_text(trimmed, &mut namespaces);
    }
    namespaces
}

fn collect_vb_xml_namespaces_from_text(text: &str, namespaces: &mut HashMap<String, String>) {
    let mut rest = text;
    while let Some(idx) = rest.find("xmlns") {
        rest = &rest[idx + "xmlns".len()..];
        let after_xmlns = rest.trim_start();
        let (prefix, after_prefix) = if let Some(after_colon) = after_xmlns.strip_prefix(':') {
            let end = after_colon
                .find(|ch: char| {
                    !(ch.is_ascii_alphanumeric() || ch == '_' || ch == '-' || ch == '.')
                })
                .unwrap_or(after_colon.len());
            (after_colon[..end].to_ascii_lowercase(), &after_colon[end..])
        } else {
            ("".to_string(), after_xmlns)
        };
        let Some(eq_idx) = after_prefix.find('=') else {
            break;
        };
        let after_eq = after_prefix[eq_idx + 1..].trim_start();
        let Some(quote) = after_eq
            .chars()
            .next()
            .filter(|ch| *ch == '"' || *ch == '\'')
        else {
            rest = after_eq;
            continue;
        };
        let value_start = quote.len_utf8();
        let Some(value_end) = after_eq[value_start..].find(quote) else {
            break;
        };
        namespaces.insert(
            prefix,
            after_eq[value_start..value_start + value_end].to_string(),
        );
        rest = &after_eq[value_start + value_end + quote.len_utf8()..];
    }
}

fn vb_xml_name_expr_for_object(
    object: &Expression,
    state: &VbXmlNormalizeState,
) -> Option<Expression> {
    vb_xml_name_info_from_expr(object, state)
        .or_else(|| match &object.kind {
            ExprKind::Ident(name) => state.name_infos.get(&name.to_ascii_lowercase()).cloned(),
            _ => None,
        })
        .map(|info| {
            call_expr(
                build_dotted_expr("xml.name"),
                vec![
                    Argument::positional(Expression::string(&info.namespace_uri)),
                    Argument::positional(Expression::string(&info.local_name)),
                    Argument::positional(Expression::string(&info.prefix)),
                ],
            )
        })
}

fn vb_xml_name_info_from_expr(
    expr: &Expression,
    state: &VbXmlNormalizeState,
) -> Option<VbXmlNameInfo> {
    let source = vb_xml_literal_source_from_expr(expr)?;
    let mut literal_namespaces = state.namespace_imports.clone();
    collect_vb_xml_namespaces_from_text(source, &mut literal_namespaces);
    let raw = vb_xml_root_name(source)?;
    let (prefix, local_name) = raw
        .split_once(':')
        .map(|(prefix, local)| (prefix.to_string(), local.to_string()))
        .unwrap_or_else(|| ("".to_string(), raw));
    let namespace_uri = literal_namespaces
        .get(&prefix.to_ascii_lowercase())
        .cloned()
        .unwrap_or_default();
    Some(VbXmlNameInfo {
        namespace_uri,
        local_name,
        prefix,
    })
}

fn vb_xml_literal_source_from_expr(expr: &Expression) -> Option<&str> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } if field == "documentElement" => {
            vb_xml_literal_source_from_expr(object)
        }
        ExprKind::Call { callee, args, .. }
            if dotted_expr_name(callee)
                .is_some_and(|name| name.eq_ignore_ascii_case("xml.parse")) =>
        {
            match args.first().map(|arg| &arg.value.kind) {
                Some(ExprKind::Lit(Literal::Str(text))) => Some(text.as_str()),
                _ => None,
            }
        }
        _ => None,
    }
}

fn vb_xml_root_name(source: &str) -> Option<String> {
    let mut text = source.trim_start();
    if text.starts_with("<?xml") {
        let end = text.find("?>")?;
        text = text[end + 2..].trim_start();
    }
    while text.starts_with("<!--") {
        let end = text.find("-->")?;
        text = text[end + 3..].trim_start();
    }
    let after_lt = text.strip_prefix('<')?;
    if after_lt.starts_with('/') || after_lt.starts_with('!') || after_lt.starts_with('?') {
        return None;
    }
    let end = after_lt
        .find(|ch: char| ch.is_whitespace() || ch == '/' || ch == '>')
        .unwrap_or(after_lt.len());
    Some(after_lt[..end].to_string())
}

fn strip_vb_generic_suffix(name: &str) -> String {
    common_generics::generic_base_name(name).to_string()
}

fn strip_vb_generic_suffixes_preserve_path(name: &str) -> String {
    let mut out = String::new();
    let mut cursor = 0usize;
    let bytes = name.as_bytes();
    while cursor < bytes.len() {
        let rest = &name[cursor..];
        if rest
            .get(..3)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("(of"))
        {
            if let Some(end) = matching_vb_paren_end(name, cursor) {
                cursor = end + 1;
                continue;
            }
        }
        let ch = rest.chars().next().unwrap();
        out.push(ch);
        cursor += ch.len_utf8();
    }
    out.trim().to_string()
}

fn vb_generic_type_marker(raw: &str) -> Option<String> {
    let trimmed = raw.trim();
    let lower = trimmed.to_ascii_lowercase();
    let start = lower.find("(of")?;
    let end = matching_vb_paren_end(trimmed, start)?;
    if end + 1 != trimmed.len() {
        return None;
    }
    let base = strip_vb_generic_suffix(trimmed);
    if base.is_empty() {
        return None;
    }
    let args = vb_generic_suffix_types(&trimmed[start..=end]);
    if args.is_empty() {
        return None;
    }
    let mut key = format!("__vb_generic_type_{}", sanitize_vb_static_key(&base));
    for arg in args {
        key.push('_');
        key.push_str(&sanitize_vb_static_key(arg));
    }
    Some(key)
}

fn vb_generic_static_name(marker: &str, member: &str) -> Option<String> {
    marker.strip_prefix("__vb_generic_type_").map(|rest| {
        format!(
            "__vb_generic_static_{}_{}",
            rest,
            sanitize_vb_static_key(member)
        )
    })
}

fn vb_generic_type_marker_parts(marker: &str) -> Option<(String, String)> {
    let rest = marker.strip_prefix("__vb_generic_type_")?;
    let (base, type_name) = rest.rsplit_once('_')?;
    if base.is_empty() || type_name.is_empty() {
        return None;
    }
    Some((base.to_string(), type_name.to_string()))
}

fn vb_generic_call_marker(name: &str, type_name: &str) -> String {
    format!(
        "__vb_generic_call_{}__of__{}",
        sanitize_vb_static_key(name),
        sanitize_vb_static_key(type_name)
    )
}

fn vb_generic_call_marker_parts(marker: &str) -> Option<(String, String)> {
    let rest = marker.strip_prefix("__vb_generic_call_")?;
    let (name, type_name) = rest.split_once("__of__")?;
    if name.is_empty() || type_name.is_empty() {
        return None;
    }
    Some((name.to_string(), type_name.to_string()))
}

fn sanitize_vb_static_key(value: &str) -> String {
    let mut out = String::new();
    for ch in value.chars() {
        if ch.is_ascii_alphanumeric() {
            out.push(ch);
        } else if ch == '_' {
            out.push('_');
        } else if !out.ends_with('_') {
            out.push('_');
        }
    }
    out.trim_matches('_').to_string()
}

fn matching_vb_paren_end(text: &str, start: usize) -> Option<usize> {
    let mut depth = 0usize;
    for (offset, ch) in text[start..].char_indices() {
        match ch {
            '(' => depth += 1,
            ')' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Some(start + offset);
                }
            }
            _ => {}
        }
    }
    None
}

fn vb_generic_suffix_types(text: &str) -> Vec<&str> {
    let trimmed = text.trim();
    let Some(inner) = trimmed.strip_prefix('(').and_then(|s| s.strip_suffix(')')) else {
        return Vec::new();
    };
    let Some(rest) = inner
        .trim()
        .strip_prefix("Of")
        .or_else(|| inner.trim().strip_prefix("of"))
    else {
        return Vec::new();
    };
    split_vb_top_level_commas(rest.trim())
        .into_iter()
        .map(str::trim)
        .filter(|part| !part.is_empty())
        .collect()
}

fn vb_declared_base_type_name(raw: &str) -> String {
    let stripped = strip_vb_generic_suffix(raw);
    let leaf = stripped.rsplit('.').next().unwrap_or(&stripped);
    if stripped.contains('.')
        && (stripped.to_ascii_lowercase().starts_with("system.")
            || vybe_platform_dotnet::emitter::component_descriptor_class_interface(&stripped)
                .is_some())
    {
        return leaf.to_string();
    }
    stripped
}

fn vb_generic_suffix_first_type(text: &str) -> Option<String> {
    common_generics::generic_argument_display_names(text)
        .into_iter()
        .next()
}

fn vb_call_generic_first_type(text: &str) -> Option<String> {
    common_generics::first_generic_argument_leaf_name(text)
}

fn consume_vb_generic_suffix(text: &str) {
    let _ = common_generics::parse_generic_params_hint(text);
}

fn normalize_vb_attribute_type_name(name: &str) -> String {
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

fn normalize_vb_interface_dispatch_type_hints(module: &mut Module) {
    let mut interfaces = std::collections::HashSet::new();
    let mut shadowing_interfaces = std::collections::HashSet::new();
    collect_vb_interface_names(&module.body, &mut interfaces, &mut shadowing_interfaces);
    let mut interface_members = HashMap::new();
    collect_vb_interface_member_owners(&module.body, &mut interface_members);
    for owners in interface_members.values() {
        if owners.len() > 1 {
            for owner in owners {
                shadowing_interfaces.insert(owner.clone());
            }
        }
    }
    rewrite_vb_interface_dispatch_type_hint_statements(
        &mut module.body,
        &interfaces,
        &shadowing_interfaces,
    );
    rewrite_vb_interface_qualified_call_statements(
        &mut module.body,
        &interfaces,
        &mut HashMap::new(),
    );
}

fn collect_vb_interface_member_owners(
    body: &[Statement],
    members: &mut HashMap<String, Vec<String>>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::InterfaceDecl {
                name,
                members: interface_members,
                ..
            } => {
                let interface = vb_canonical_type_name(name);
                let interface_key = interface.to_ascii_lowercase();
                for member in interface_members {
                    let member_name = match member {
                        InterfaceMember::Method { name, .. }
                        | InterfaceMember::Property { name, .. }
                        | InterfaceMember::Event { name, .. } => name,
                    };
                    let entry = members.entry(member_name.to_ascii_lowercase()).or_default();
                    if !entry.iter().any(|item| item == &interface_key) {
                        entry.push(interface_key.clone());
                    }
                }
            }
            StmtKind::ClassDecl {
                members: class_members,
                ..
            }
            | StmtKind::StructDecl {
                members: class_members,
                ..
            }
            | StmtKind::ModuleDecl {
                members: class_members,
                ..
            } => {
                for member in class_members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_interface_member_owners(std::slice::from_ref(nested), members);
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_vb_interface_member_owners(body, members)
            }
            _ => {}
        }
    }
}

fn rewrite_vb_interface_qualified_call_statements(
    body: &mut [Statement],
    interfaces: &std::collections::HashSet<String>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_interface_qualified_call_statement(stmt, interfaces, locals);
    }
}

fn rewrite_vb_interface_qualified_call_statement(
    stmt: &mut Statement,
    interfaces: &std::collections::HashSet<String>,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_interface_qualified_call_expr(init, interfaces, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(type_hint) = &decl.type_hint {
                        let hint = vb_interface_type_key(type_hint);
                        if interfaces.contains(&hint.to_ascii_lowercase()) {
                            locals.insert(name.to_ascii_lowercase(), hint);
                        }
                    }
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_interface_qualified_call_expr(expr, interfaces, locals);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_interface_qualified_call_expr(target, interfaces, locals);
            }
            rewrite_vb_interface_qualified_call_expr(value, interfaces, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_vb_interface_qualified_call_expr(target, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(value, interfaces, locals);
        }
        StmtKind::FunctionDecl { body, params, .. } => {
            let mut fn_locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    let hint = vb_interface_type_key(type_hint);
                    if interfaces.contains(&hint.to_ascii_lowercase()) {
                        fn_locals.insert(param.name.to_ascii_lowercase(), hint);
                    }
                }
            }
            rewrite_vb_interface_qualified_call_statements(body, interfaces, &mut fn_locals);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_interface_qualified_call_member(member, interfaces);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_interface_qualified_call_expr(cond, interfaces, locals);
            rewrite_vb_interface_qualified_call_statements(
                then_body,
                interfaces,
                &mut locals.clone(),
            );
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_interface_qualified_call_expr(elif_cond, interfaces, locals);
                rewrite_vb_interface_qualified_call_statements(
                    elif_body,
                    interfaces,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                rewrite_vb_interface_qualified_call_statements(
                    else_body,
                    interfaces,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_interface_qualified_call_statements(body, interfaces, &mut locals.clone());
        }
        _ => {}
    }
}

fn rewrite_vb_interface_qualified_call_member(
    member: &mut ClassMember,
    interfaces: &std::collections::HashSet<String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_interface_qualified_call_statement(stmt, interfaces, &mut HashMap::new());
        }
        ClassMember::Constructor { body, params, .. } => {
            let mut locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    let hint = vb_interface_type_key(type_hint);
                    if interfaces.contains(&hint.to_ascii_lowercase()) {
                        locals.insert(param.name.to_ascii_lowercase(), hint);
                    }
                }
            }
            rewrite_vb_interface_qualified_call_statements(body, interfaces, &mut locals);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_interface_qualified_call_statements(
                    getter,
                    interfaces,
                    &mut HashMap::new(),
                );
            }
            if let Some(setter) = setter {
                rewrite_vb_interface_qualified_call_statements(
                    &mut setter.body,
                    interfaces,
                    &mut HashMap::new(),
                );
            }
        }
        _ => {}
    }
}

fn rewrite_vb_interface_qualified_call_expr(
    expr: &mut Expression,
    interfaces: &std::collections::HashSet<String>,
    locals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, field, .. } = &mut callee.kind {
                if let Some(interface) = vb_interface_receiver_type(object, locals) {
                    if interfaces.contains(&interface.to_ascii_lowercase()) {
                        let forwarder = vb_interface_forwarder_name(&interface, field);
                        let arg_types: Option<Vec<String>> = args
                            .iter()
                            .map(|arg| vb_interface_call_arg_type(&arg.value))
                            .collect();
                        *field = arg_types
                            .filter(|types| !types.is_empty())
                            .map(|types| {
                                format!(
                                    "{}$sig{}",
                                    forwarder.to_ascii_lowercase(),
                                    types
                                        .into_iter()
                                        .map(|ty| ty.to_ascii_lowercase())
                                        .collect::<Vec<_>>()
                                        .join("$")
                                )
                            })
                            .unwrap_or(forwarder);
                    }
                }
                rewrite_vb_interface_qualified_call_expr(object, interfaces, locals);
            } else {
                rewrite_vb_interface_qualified_call_expr(callee, interfaces, locals);
            }
            for arg in args {
                rewrite_vb_interface_qualified_call_expr(&mut arg.value, interfaces, locals);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_vb_interface_qualified_call_expr(object, interfaces, locals);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_interface_qualified_call_expr(left, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(right, interfaces, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::RefLoad(expr) => {
            rewrite_vb_interface_qualified_call_expr(expr, interfaces, locals)
        }
        ExprKind::Assign { target, value } => {
            rewrite_vb_interface_qualified_call_expr(target, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(value, interfaces, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_interface_qualified_call_expr(cond, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(then, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(else_, interfaces, locals);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_interface_qualified_call_expr(object, interfaces, locals);
            rewrite_vb_interface_qualified_call_expr(index, interfaces, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_interface_qualified_call_expr(&mut item.value, interfaces, locals);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                rewrite_vb_interface_qualified_call_expr(item, interfaces, locals);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_interface_qualified_call_expr(class, interfaces, locals);
            for arg in args {
                rewrite_vb_interface_qualified_call_expr(&mut arg.value, interfaces, locals);
            }
        }
        _ => {}
    }
}

fn vb_interface_receiver_type(
    expr: &Expression,
    locals: &HashMap<String, String>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => locals.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Index { object, .. } => vb_interface_receiver_type(object, locals),
        ExprKind::Call { callee, .. } => vb_interface_receiver_type(callee, locals),
        _ => None,
    }
}

fn vb_interface_type_key(type_hint: &str) -> String {
    let hint = vb_canonical_type_name(type_hint);
    hint.trim_end_matches("()")
        .trim_end_matches("[]")
        .trim()
        .to_string()
}

fn vb_interface_call_arg_type(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) => Some("integer".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("double".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("boolean".to_string()),
        ExprKind::Cast { type_name, .. } => Some(vb_canonical_type_name(type_name)),
        _ => None,
    }
}

fn collect_vb_interface_names(
    body: &[Statement],
    interfaces: &mut std::collections::HashSet<String>,
    shadowing_interfaces: &mut std::collections::HashSet<String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::InterfaceDecl { name, members, .. } => {
                let key = vb_canonical_type_name(name).to_ascii_lowercase();
                interfaces.insert(key.clone());
                let mut seen = std::collections::HashSet::new();
                for member in members {
                    if matches!(
                        member,
                        InterfaceMember::Method {
                            signature_source: Some(source),
                            ..
                        } if source.eq_ignore_ascii_case("shadows")
                    ) {
                        shadowing_interfaces.insert(key.clone());
                    }
                    let member_name = match member {
                        InterfaceMember::Method { name, .. }
                        | InterfaceMember::Property { name, .. }
                        | InterfaceMember::Event { name, .. } => name,
                    }
                    .to_ascii_lowercase();
                    if !seen.insert(member_name) {
                        shadowing_interfaces.insert(key.clone());
                    }
                }
            }
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_interface_names(
                            std::slice::from_ref(nested),
                            interfaces,
                            shadowing_interfaces,
                        );
                    }
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_interface_names(
                            std::slice::from_ref(nested),
                            interfaces,
                            shadowing_interfaces,
                        );
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_vb_interface_names(body, interfaces, shadowing_interfaces)
            }
            _ => {}
        }
    }
}

fn rewrite_vb_interface_dispatch_type_hint_statements(
    body: &mut [Statement],
    interfaces: &std::collections::HashSet<String>,
    shadowing_interfaces: &std::collections::HashSet<String>,
) {
    for stmt in body {
        rewrite_vb_interface_dispatch_type_hint_statement(stmt, interfaces, shadowing_interfaces);
    }
}

fn rewrite_vb_interface_dispatch_type_hint_statement(
    stmt: &mut Statement,
    interfaces: &std::collections::HashSet<String>,
    shadowing_interfaces: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                let Some(type_hint) = decl.type_hint.as_ref() else {
                    continue;
                };
                let hint_key = vb_canonical_type_name(type_hint).to_ascii_lowercase();
                if !interfaces.contains(&hint_key) || shadowing_interfaces.contains(&hint_key) {
                    continue;
                }
                if let Some(init) = &decl.init {
                    if let Some(inferred) = vb_infer_expr_type(init, &HashMap::new()) {
                        if !interfaces.contains(&inferred.to_ascii_lowercase())
                            && !matches!(inferred.as_str(), "Object" | "Array")
                        {
                            decl.type_hint = Some(inferred);
                        }
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            rewrite_vb_interface_dispatch_type_hint_statements(
                body,
                interfaces,
                shadowing_interfaces,
            );
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_interface_dispatch_type_hint_member(
                    member,
                    interfaces,
                    shadowing_interfaces,
                );
            }
        }
        StmtKind::InterfaceDecl { .. } => {}
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_vb_interface_dispatch_type_hint_statements(
                then_body,
                interfaces,
                shadowing_interfaces,
            );
            for (_, elif_body) in elifs {
                rewrite_vb_interface_dispatch_type_hint_statements(
                    elif_body,
                    interfaces,
                    shadowing_interfaces,
                );
            }
            if let Some(else_body) = else_body {
                rewrite_vb_interface_dispatch_type_hint_statements(
                    else_body,
                    interfaces,
                    shadowing_interfaces,
                );
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                rewrite_vb_interface_dispatch_type_hint_statement(
                    init,
                    interfaces,
                    shadowing_interfaces,
                );
            }
            rewrite_vb_interface_dispatch_type_hint_statements(
                body,
                interfaces,
                shadowing_interfaces,
            );
        }
        StmtKind::ForIn {
            body, else_body, ..
        }
        | StmtKind::While {
            body, else_body, ..
        } => {
            rewrite_vb_interface_dispatch_type_hint_statements(
                body,
                interfaces,
                shadowing_interfaces,
            );
            if let Some(else_body) = else_body {
                rewrite_vb_interface_dispatch_type_hint_statements(
                    else_body,
                    interfaces,
                    shadowing_interfaces,
                );
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_interface_dispatch_type_hint_statements(
                body,
                interfaces,
                shadowing_interfaces,
            );
        }
        _ => {}
    }
}

fn rewrite_vb_interface_dispatch_type_hint_member(
    member: &mut ClassMember,
    interfaces: &std::collections::HashSet<String>,
    shadowing_interfaces: &std::collections::HashSet<String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_interface_dispatch_type_hint_statement(
                stmt,
                interfaces,
                shadowing_interfaces,
            );
        }
        ClassMember::Constructor { body, .. } => {
            rewrite_vb_interface_dispatch_type_hint_statements(
                body,
                interfaces,
                shadowing_interfaces,
            );
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_interface_dispatch_type_hint_statements(
                    getter,
                    interfaces,
                    shadowing_interfaces,
                );
            }
            if let Some(setter) = setter {
                rewrite_vb_interface_dispatch_type_hint_statements(
                    &mut setter.body,
                    interfaces,
                    shadowing_interfaces,
                );
            }
        }
        _ => {}
    }
}

fn normalize_vb_partial_classes(body: &mut Vec<Statement>) {
    let mut seen: HashMap<String, usize> = HashMap::new();
    let mut merged = Vec::new();
    for mut stmt in body.drain(..) {
        if let StmtKind::NamespaceDecl { body, .. } = &mut stmt.kind {
            normalize_vb_partial_classes(body);
        }
        let merge_key = match &stmt.kind {
            StmtKind::ClassDecl {
                name, modifiers, ..
            } if modifiers.is_partial => Some(name.to_ascii_lowercase()),
            _ => None,
        };
        let Some(key) = merge_key else {
            merged.push(stmt);
            continue;
        };
        if let Some(&idx) = seen.get(&key) {
            let StmtKind::ClassDecl {
                parents,
                interfaces,
                members,
                decorators,
                ..
            } = stmt.kind
            else {
                continue;
            };
            if let StmtKind::ClassDecl {
                parents: target_parents,
                interfaces: target_interfaces,
                members: target_members,
                decorators: target_decorators,
                modifiers,
                ..
            } = &mut merged[idx].kind
            {
                if target_parents.is_empty() {
                    *target_parents = parents;
                }
                for interface in interfaces {
                    if !target_interfaces
                        .iter()
                        .any(|known| known.eq_ignore_ascii_case(&interface))
                    {
                        target_interfaces.push(interface);
                    }
                }
                target_members.extend(members);
                target_decorators.extend(decorators);
                modifiers.is_partial = false;
            }
        } else {
            if let StmtKind::ClassDecl { modifiers, .. } = &mut stmt.kind {
                modifiers.is_partial = false;
            }
            seen.insert(key, merged.len());
            merged.push(stmt);
        }
    }
    for stmt in &mut merged {
        if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
            normalize_vb_partial_methods(members);
        }
    }
    *body = merged;
}

fn normalize_vb_partial_methods(members: &mut Vec<ClassMember>) {
    let implemented: HashSet<String> = members
        .iter()
        .filter(|member| !vb_class_member_is_partial_method_decl(member))
        .filter_map(vb_class_member_method_name)
        .map(|name| name.to_ascii_lowercase())
        .collect();
    if !implemented.is_empty() {
        members.retain(|member| {
            let Some(name) = vb_class_member_method_name(member) else {
                return true;
            };
            !(vb_class_member_is_partial_method_decl(member)
                && implemented.contains(&name.to_ascii_lowercase()))
        });
    }
    for member in members {
        if let ClassMember::NestedType(stmt) = member {
            if let StmtKind::ClassDecl { members, .. }
            | StmtKind::StructDecl { members, .. }
            | StmtKind::ModuleDecl { members, .. } = &mut stmt.kind
            {
                normalize_vb_partial_methods(members);
            }
        }
    }
}

fn vb_class_member_is_partial_method_decl(member: &ClassMember) -> bool {
    let ClassMember::Method(stmt) = member else {
        return false;
    };
    let StmtKind::FunctionDecl { modifiers, .. } = &stmt.kind else {
        return false;
    };
    modifiers.decorators.iter().any(|decorator| {
        matches!(
            &decorator.kind,
            ExprKind::Lit(Literal::Str(value)) if value == VB_PARTIAL_METHOD_MARKER
        )
    })
}

fn normalize_vb_implicit_method_self_classes(body: &mut [Statement]) {
    for stmt in body {
        match &mut stmt.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                normalize_vb_implicit_method_self_members(members);
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(stmt) = member {
                        normalize_vb_implicit_method_self_classes(std::slice::from_mut(stmt));
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                normalize_vb_implicit_method_self_classes(body);
            }
            _ => {}
        }
    }
}

fn normalize_vb_implicit_method_self_members(members: &mut [ClassMember]) {
    let methods: HashSet<String> = members
        .iter()
        .filter_map(|member| {
            let name = vb_class_member_method_name(member)?;
            (!name.starts_with("__vb_myclass_")).then(|| name.to_ascii_lowercase())
        })
        .collect();
    let by_ref_params: HashMap<String, Vec<bool>> = members
        .iter()
        .filter_map(|member| {
            let ClassMember::Method(stmt) = member else {
                return None;
            };
            let StmtKind::FunctionDecl { name, params, .. } = &stmt.kind else {
                return None;
            };
            Some((
                name.to_ascii_lowercase(),
                params
                    .iter()
                    .map(|param| matches!(param.pass_by, PassBy::Ref | PassBy::Out))
                    .collect(),
            ))
        })
        .collect();

    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { params, body, .. } = &mut stmt.kind {
                    let mut locals = params
                        .iter()
                        .map(|param| param.name.to_ascii_lowercase())
                        .collect();
                    normalize_vb_implicit_method_self_statements(
                        body,
                        &methods,
                        &by_ref_params,
                        &mut locals,
                    );
                }
            }
            ClassMember::Constructor { params, body, .. } => {
                let mut locals = params
                    .iter()
                    .map(|param| param.name.to_ascii_lowercase())
                    .collect();
                normalize_vb_implicit_method_self_statements(
                    body,
                    &methods,
                    &by_ref_params,
                    &mut locals,
                );
            }
            ClassMember::Property { getter, setter, .. } => {
                if let Some(getter) = getter {
                    normalize_vb_implicit_method_self_statements(
                        getter,
                        &methods,
                        &by_ref_params,
                        &mut HashSet::new(),
                    );
                }
                if let Some(setter) = setter {
                    let mut locals = HashSet::from([setter.param.name.to_ascii_lowercase()]);
                    normalize_vb_implicit_method_self_statements(
                        &mut setter.body,
                        &methods,
                        &by_ref_params,
                        &mut locals,
                    );
                }
            }
            ClassMember::NestedType(stmt) => {
                normalize_vb_implicit_method_self_classes(std::slice::from_mut(stmt));
            }
            _ => {}
        }
    }
}

fn normalize_vb_implicit_method_self_statements(
    body: &mut [Statement],
    methods: &HashSet<String>,
    by_ref_params: &HashMap<String, Vec<bool>>,
    locals: &mut HashSet<String>,
) {
    for stmt in body {
        normalize_vb_implicit_method_self_statement(stmt, methods, by_ref_params, locals);
    }
}

fn normalize_vb_implicit_method_self_statement(
    stmt: &mut Statement,
    methods: &HashSet<String>,
    by_ref_params: &HashMap<String, Vec<bool>>,
    locals: &mut HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_implicit_method_self_expr(init, methods, by_ref_params, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    locals.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_implicit_method_self_expr(target, methods, by_ref_params, locals);
            }
            normalize_vb_implicit_method_self_expr(value, methods, by_ref_params, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_implicit_method_self_expr(target, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(value, methods, by_ref_params, locals);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_implicit_method_self_expr(expr, methods, by_ref_params, locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_implicit_method_self_expr(cond, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_statements(
                then_body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
            for (elif_cond, elif_body) in elifs {
                normalize_vb_implicit_method_self_expr(elif_cond, methods, by_ref_params, locals);
                normalize_vb_implicit_method_self_statements(
                    elif_body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_implicit_method_self_statements(
                    else_body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_implicit_method_self_expr(cond, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
            if let Some(else_body) = else_body {
                normalize_vb_implicit_method_self_statements(
                    else_body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::DoWhile { cond, body, .. } => {
            normalize_vb_implicit_method_self_expr(cond, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_implicit_method_self_statement(
                    init,
                    methods,
                    by_ref_params,
                    &mut loop_locals,
                );
            }
            if let Some(cond) = cond {
                normalize_vb_implicit_method_self_expr(cond, methods, by_ref_params, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_implicit_method_self_expr(
                    update,
                    methods,
                    by_ref_params,
                    &loop_locals,
                );
            }
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut loop_locals,
            );
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_implicit_method_self_expr(iter, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
            if let Some(else_body) = else_body {
                normalize_vb_implicit_method_self_statements(
                    else_body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_implicit_method_self_expr(
                        when_clause,
                        methods,
                        by_ref_params,
                        locals,
                    );
                }
                normalize_vb_implicit_method_self_statements(
                    &mut catch.body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_implicit_method_self_statements(
                    else_body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
            if let Some(finally) = finally {
                normalize_vb_implicit_method_self_statements(
                    finally,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Block(body) => {
            normalize_vb_implicit_method_self_statements(
                body,
                methods,
                by_ref_params,
                &mut locals.clone(),
            );
        }
        _ => {}
    }
}

fn normalize_vb_implicit_method_self_expr(
    expr: &mut Expression,
    methods: &HashSet<String>,
    by_ref_params: &HashMap<String, Vec<bool>>,
    locals: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            for arg in &mut *args {
                normalize_vb_implicit_method_self_expr(
                    &mut arg.value,
                    methods,
                    by_ref_params,
                    locals,
                );
            }
            if let ExprKind::Ident(name) = &callee.kind {
                let key = name.to_ascii_lowercase();
                if methods.contains(&key) && !locals.contains(&key) {
                    if let Some(by_refs) = by_ref_params.get(&key) {
                        for (idx, by_ref) in by_refs.iter().enumerate() {
                            if let Some(arg) = args.get_mut(idx) {
                                arg.by_ref = *by_ref;
                            }
                        }
                    }
                    let field = name.clone();
                    *callee = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(Expression::new(ExprKind::This)),
                        field,
                        null_safe: false,
                    }));
                    return;
                }
            }
            normalize_vb_implicit_method_self_expr(callee, methods, by_ref_params, locals);
        }
        ExprKind::Member { object, .. } => {
            normalize_vb_implicit_method_self_expr(object, methods, by_ref_params, locals);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_implicit_method_self_expr(left, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(right, methods, by_ref_params, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::RefLoad(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr) => {
            normalize_vb_implicit_method_self_expr(expr, methods, by_ref_params, locals);
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_implicit_method_self_expr(target, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(value, methods, by_ref_params, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_implicit_method_self_expr(cond, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(then, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(else_, methods, by_ref_params, locals);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_implicit_method_self_expr(object, methods, by_ref_params, locals);
            normalize_vb_implicit_method_self_expr(index, methods, by_ref_params, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_implicit_method_self_expr(
                    &mut item.value,
                    methods,
                    by_ref_params,
                    locals,
                );
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_implicit_method_self_expr(item, methods, by_ref_params, locals);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_implicit_method_self_expr(class, methods, by_ref_params, locals);
            for arg in args {
                normalize_vb_implicit_method_self_expr(
                    &mut arg.value,
                    methods,
                    by_ref_params,
                    locals,
                );
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_implicit_method_self_expr(key, methods, by_ref_params, locals);
                        normalize_vb_implicit_method_self_expr(
                            value,
                            methods,
                            by_ref_params,
                            locals,
                        );
                    }
                    ObjectProperty::Spread(value) => {
                        normalize_vb_implicit_method_self_expr(
                            value,
                            methods,
                            by_ref_params,
                            locals,
                        );
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_implicit_method_self_statement(
                            value,
                            methods,
                            by_ref_params,
                            &mut locals.clone(),
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => {
                normalize_vb_implicit_method_self_expr(expr, methods, by_ref_params, locals);
            }
            LambdaBody::Block(body) => {
                normalize_vb_implicit_method_self_statements(
                    body,
                    methods,
                    by_ref_params,
                    &mut locals.clone(),
                );
            }
        },
        _ => {}
    }
}

fn normalize_vb_bitwise_logic(module: &mut Module) {
    rewrite_vb_bitwise_logic_statements(&mut module.body, &mut HashMap::new());
}

fn normalize_vb_trycast_known_locals(module: &mut Module) {
    let mut parents = HashMap::new();
    collect_vb_class_parent_map(&module.body, &mut parents);
    rewrite_vb_trycast_statements(&mut module.body, &parents, &mut HashMap::new());
}

fn collect_vb_class_parent_map(body: &[Statement], parents: &mut HashMap<String, Vec<String>>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents: class_parents,
                members,
                ..
            } => {
                parents.insert(
                    vb_canonical_type_name(name).to_ascii_lowercase(),
                    class_parents
                        .iter()
                        .map(|parent| vb_canonical_type_name(parent).to_ascii_lowercase())
                        .collect(),
                );
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_class_parent_map(std::slice::from_ref(nested.as_ref()), parents);
                    }
                }
            }
            StmtKind::StructDecl { name, members, .. } => {
                parents
                    .entry(vb_canonical_type_name(name).to_ascii_lowercase())
                    .or_default();
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_class_parent_map(std::slice::from_ref(nested), parents);
                    }
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_class_parent_map(std::slice::from_ref(nested), parents);
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => collect_vb_class_parent_map(body, parents),
            _ => {}
        }
    }
}

fn rewrite_vb_trycast_statements(
    body: &mut [Statement],
    parents: &HashMap<String, Vec<String>>,
    actuals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_trycast_statement(stmt, parents, actuals);
    }
}

fn rewrite_vb_trycast_statement(
    stmt: &mut Statement,
    parents: &HashMap<String, Vec<String>>,
    actuals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    if let Some(rewritten) = rewrite_vb_trycast_decl_init(
                        init,
                        decl.type_hint.as_deref(),
                        parents,
                        actuals,
                    ) {
                        *init = rewritten;
                    } else {
                        rewrite_vb_trycast_expr(init, parents, actuals);
                    }
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(actual) = decl.init.as_ref().and_then(vb_new_expr_type_name) {
                        actuals.insert(name.to_ascii_lowercase(), vb_type_key(&actual));
                    }
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_trycast_expr(expr, parents, actuals)
        }
        StmtKind::Assign { targets, value } => {
            for target in &mut *targets {
                rewrite_vb_trycast_expr(target, parents, actuals);
            }
            rewrite_vb_trycast_expr(value, parents, actuals);
            if let Some(Expression {
                kind: ExprKind::Ident(name),
                ..
            }) = targets.first()
            {
                if let Some(actual) = vb_new_expr_type_name(value) {
                    actuals.insert(name.to_ascii_lowercase(), vb_type_key(&actual));
                }
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_trycast_expr(cond, parents, actuals);
            rewrite_vb_trycast_statements(then_body, parents, &mut actuals.clone());
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_trycast_expr(elif_cond, parents, actuals);
                rewrite_vb_trycast_statements(elif_body, parents, &mut actuals.clone());
            }
            if let Some(else_body) = else_body {
                rewrite_vb_trycast_statements(else_body, parents, &mut actuals.clone());
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            rewrite_vb_trycast_statements(body, parents, &mut HashMap::new());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_trycast_member(member, parents);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_trycast_statements(body, parents, &mut HashMap::new())
        }
        _ => {}
    }
}

fn rewrite_vb_trycast_decl_init(
    init: &Expression,
    type_hint: Option<&str>,
    parents: &HashMap<String, Vec<String>>,
    actuals: &HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Cast { expr, type_name } = &init.kind else {
        return None;
    };
    if !type_name.to_ascii_lowercase().starts_with("trycast:") {
        return None;
    }
    let ExprKind::Ident(source) = &expr.kind else {
        return None;
    };
    let actual = actuals.get(&source.to_ascii_lowercase())?;
    let target = type_hint
        .map(vb_type_key)
        .unwrap_or_else(|| vb_type_key(type_name));
    if vb_type_assignable_to(actual, &target, parents)
        || vb_type_token_contains(&target, actual)
        || vb_type_token_contains(type_name, actual)
    {
        Some((**expr).clone())
    } else {
        Some(Expression::null())
    }
}

fn rewrite_vb_trycast_member(member: &mut ClassMember, parents: &HashMap<String, Vec<String>>) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_trycast_statement(stmt, parents, &mut HashMap::new());
        }
        ClassMember::Constructor { body, .. } => {
            rewrite_vb_trycast_statements(body, parents, &mut HashMap::new())
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_trycast_statements(getter, parents, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                rewrite_vb_trycast_statements(&mut setter.body, parents, &mut HashMap::new());
            }
        }
        _ => {}
    }
}

fn rewrite_vb_trycast_expr(
    expr: &mut Expression,
    parents: &HashMap<String, Vec<String>>,
    actuals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => {
            rewrite_vb_trycast_expr(inner, parents, actuals);
            if let Some(target) = type_name.strip_prefix("TryCast:") {
                if let ExprKind::Ident(name) = &inner.kind {
                    if let Some(actual) = actuals.get(&name.to_ascii_lowercase()) {
                        let target = vb_type_key(target);
                        if vb_type_assignable_to(actual, &target, parents)
                            || vb_type_token_contains(&target, actual)
                            || vb_type_token_contains(type_name, actual)
                        {
                            *expr = (**inner).clone();
                        } else {
                            *expr = Expression::null();
                        }
                    }
                }
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_trycast_expr(left, parents, actuals);
            rewrite_vb_trycast_expr(right, parents, actuals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_vb_trycast_expr(expr, parents, actuals),
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_trycast_expr(callee, parents, actuals);
            for arg in args {
                rewrite_vb_trycast_expr(&mut arg.value, parents, actuals);
            }
        }
        ExprKind::Member { object, .. } => rewrite_vb_trycast_expr(object, parents, actuals),
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_trycast_expr(object, parents, actuals);
            rewrite_vb_trycast_expr(index, parents, actuals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_trycast_expr(cond, parents, actuals);
            rewrite_vb_trycast_expr(then, parents, actuals);
            rewrite_vb_trycast_expr(else_, parents, actuals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_trycast_expr(&mut item.value, parents, actuals);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            for item in items {
                rewrite_vb_trycast_expr(item, parents, actuals);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_trycast_expr(class, parents, actuals);
            for arg in args {
                rewrite_vb_trycast_expr(&mut arg.value, parents, actuals);
            }
        }
        _ => {}
    }
}

fn vb_new_expr_type_name(expr: &Expression) -> Option<String> {
    let ExprKind::New { class, .. } = &expr.kind else {
        return None;
    };
    dotted_expr_name(class).map(|name| vb_canonical_type_name(&name))
}

fn vb_type_assignable_to(
    actual: &str,
    target: &str,
    parents: &HashMap<String, Vec<String>>,
) -> bool {
    let actual = vb_type_key(actual);
    let target = vb_type_key(target);
    if actual == target
        || actual.ends_with(&format!(".{target}"))
        || target.ends_with(&format!(".{actual}"))
    {
        return true;
    }
    parents.get(&actual).is_some_and(|items| {
        items
            .iter()
            .any(|parent| vb_type_assignable_to(parent, &target, parents))
    })
}

fn vb_type_key(raw: &str) -> String {
    let raw = raw.split_once(':').map(|(_, ty)| ty).unwrap_or(raw);
    let raw = raw.trim().trim_end_matches('?').trim();
    let raw = raw
        .trim_end_matches(|ch: char| !(ch.is_ascii_alphanumeric() || ch == '_' || ch == '.'))
        .trim();
    vb_canonical_type_name(raw).to_ascii_lowercase()
}

fn vb_type_token_contains(haystack: &str, needle: &str) -> bool {
    let clean = |text: &str| {
        text.chars()
            .filter(|ch| ch.is_ascii_alphanumeric() || *ch == '_')
            .collect::<String>()
    };
    let haystack = clean(haystack);
    let needle = clean(needle);
    !needle.is_empty() && haystack.contains(&needle)
}

fn rewrite_vb_bitwise_logic_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_bitwise_logic_statement(stmt, locals);
    }
}

fn rewrite_vb_bitwise_logic_statement(stmt: &mut Statement, locals: &mut HashMap<String, String>) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            let mut throw_message = None;
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_bitwise_logic_expr(init, locals);
                    if expr_contains_vb_zero_idiv(init, locals) {
                        throw_message =
                            Some("DivideByZeroException:Attempted to divide by zero".to_string());
                    } else if expr_contains_vb_int_overflow(init, locals) {
                        throw_message = Some(
                            "OverflowException:Arithmetic operation resulted in an overflow"
                                .to_string(),
                        );
                    }
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let local_name = name.to_ascii_lowercase();
                    let inferred = decl
                        .type_hint
                        .as_ref()
                        .map(|hint| vb_canonical_type_name(hint))
                        .or_else(|| {
                            decl.init
                                .as_ref()
                                .and_then(|init| vb_infer_expr_type(init, locals))
                        });
                    if let Some(mut ty) = inferred {
                        if ty == "Object"
                            && decl
                                .init
                                .as_ref()
                                .is_some_and(|init| matches!(init.kind, ExprKind::New { .. }))
                        {
                            ty = "ObjectRef".to_string();
                        }
                        locals.insert(local_name.clone(), ty);
                    }
                    if let Some(value) = decl
                        .init
                        .as_ref()
                        .and_then(|init| eval_vb_int_const_expr(init, locals))
                    {
                        locals.insert(format!("{local_name}#const"), value.to_string());
                    }
                }
            }
            if let Some(message) = throw_message {
                *stmt = vb_throw_statement(&message);
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_bitwise_logic_expr(expr, locals);
            if let Some(message) = find_vb_err_raise_message(expr) {
                *stmt = vb_throw_statement(&message);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_bitwise_logic_expr(target, locals);
            }
            rewrite_vb_bitwise_logic_expr(value, locals);
            if expr_contains_vb_zero_idiv(value, locals) {
                *stmt = vb_throw_statement("DivideByZeroException:Attempted to divide by zero");
            }
        }
        StmtKind::CompoundAssign { target, op, value } => {
            rewrite_vb_bitwise_logic_expr(target, locals);
            rewrite_vb_bitwise_logic_expr(value, locals);
            if (matches!(op, CompoundOp::IDiv | CompoundOp::Mod)
                && expr_is_vb_int_zero(value, locals))
                || expr_contains_vb_zero_idiv(value, locals)
            {
                *stmt = vb_throw_statement("DivideByZeroException:Attempted to divide by zero");
            }
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(
                        param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
            }
            rewrite_vb_bitwise_logic_statements(body, &mut scoped);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_bitwise_logic_member(member);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_bitwise_logic_expr(cond, locals);
            rewrite_vb_bitwise_logic_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_bitwise_logic_expr(elif_cond, locals);
                rewrite_vb_bitwise_logic_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                rewrite_vb_bitwise_logic_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = locals.clone();
            if let Some(init) = init {
                rewrite_vb_bitwise_logic_statement(init, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_vb_bitwise_logic_expr(cond, &scoped);
            }
            if let Some(update) = update {
                rewrite_vb_bitwise_logic_expr(update, &scoped);
            }
            rewrite_vb_bitwise_logic_statements(body, &mut scoped);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_bitwise_logic_expr(iter, locals);
            rewrite_vb_bitwise_logic_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_bitwise_logic_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_bitwise_logic_expr(cond, locals);
            rewrite_vb_bitwise_logic_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_bitwise_logic_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            rewrite_vb_bitwise_logic_statements(body, &mut locals.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_vb_bitwise_logic_expr(when_clause, locals);
                }
                rewrite_vb_bitwise_logic_statements(&mut catch.body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                rewrite_vb_bitwise_logic_statements(else_body, &mut locals.clone());
            }
            if let Some(finally) = finally {
                rewrite_vb_bitwise_logic_statements(finally, &mut locals.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_bitwise_logic_statements(body, &mut locals.clone());
        }
        _ => {}
    }
}

fn vb_throw_statement(message: &str) -> Statement {
    let (exception_type, message) = message.split_once(':').unwrap_or(("Exception", message));
    Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(build_dotted_expr(&format!("System.{exception_type}"))),
            args: vec![Argument::positional(Expression::string(message))],
        })),
        cause: None,
    })
}

fn find_vb_err_raise_message(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_err_raise") {
                return args
                    .first()
                    .and_then(|arg| literal_string(&arg.value))
                    .or_else(|| Some("VB runtime error".to_string()));
            }
            find_vb_err_raise_message(callee).or_else(|| {
                args.iter()
                    .find_map(|arg| find_vb_err_raise_message(&arg.value))
            })
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => find_vb_err_raise_message(left).or_else(|| find_vb_err_raise_message(right)),
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => find_vb_err_raise_message(expr),
        ExprKind::Member { object, .. } => find_vb_err_raise_message(object),
        ExprKind::Index { object, index, .. } => {
            find_vb_err_raise_message(object).or_else(|| find_vb_err_raise_message(index))
        }
        ExprKind::Ternary { cond, then, else_ } => find_vb_err_raise_message(cond)
            .or_else(|| find_vb_err_raise_message(then))
            .or_else(|| find_vb_err_raise_message(else_)),
        ExprKind::Array(items) => items
            .iter()
            .find_map(|item| find_vb_err_raise_message(&item.value)),
        ExprKind::Tuple(items) | ExprKind::Set(items) => {
            items.iter().find_map(find_vb_err_raise_message)
        }
        ExprKind::New { class, args } => find_vb_err_raise_message(class).or_else(|| {
            args.iter()
                .find_map(|arg| find_vb_err_raise_message(&arg.value))
        }),
        _ => None,
    }
}

fn expr_contains_vb_zero_idiv(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Binary { op, left, right } => {
            (matches!(op, BinOp::IDiv | BinOp::Mod) && expr_is_vb_int_zero(right, locals))
                || expr_contains_vb_zero_idiv(left, locals)
                || expr_contains_vb_zero_idiv(right, locals)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => expr_contains_vb_zero_idiv(expr, locals),
        ExprKind::Call { callee, args, .. } => {
            expr_contains_vb_zero_idiv(callee, locals)
                || args
                    .iter()
                    .any(|arg| expr_contains_vb_zero_idiv(&arg.value, locals))
        }
        ExprKind::Member { object, .. } => expr_contains_vb_zero_idiv(object, locals),
        ExprKind::Index { object, index, .. } => {
            expr_contains_vb_zero_idiv(object, locals) || expr_contains_vb_zero_idiv(index, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_vb_zero_idiv(cond, locals)
                || expr_contains_vb_zero_idiv(then, locals)
                || expr_contains_vb_zero_idiv(else_, locals)
        }
        ExprKind::Array(items) => items
            .iter()
            .any(|item| expr_contains_vb_zero_idiv(&item.value, locals)),
        ExprKind::Tuple(items) | ExprKind::Set(items) => items
            .iter()
            .any(|item| expr_contains_vb_zero_idiv(item, locals)),
        ExprKind::New { class, args } => {
            expr_contains_vb_zero_idiv(class, locals)
                || args
                    .iter()
                    .any(|arg| expr_contains_vb_zero_idiv(&arg.value, locals))
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => expr_contains_vb_zero_idiv(left, locals) || expr_contains_vb_zero_idiv(right, locals),
        _ => false,
    }
}

fn expr_is_vb_int_zero(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    eval_vb_int_const_expr(expr, locals) == Some(0)
}

fn expr_contains_vb_int_overflow(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Binary { left, right, .. } => {
            eval_vb_int_const_expr(expr, locals)
                .is_some_and(|value| value < i64::from(i32::MIN) || value > i64::from(i32::MAX))
                || expr_contains_vb_int_overflow(left, locals)
                || expr_contains_vb_int_overflow(right, locals)
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => expr_contains_vb_int_overflow(expr, locals),
        ExprKind::Call { callee, args, .. } => {
            expr_contains_vb_int_overflow(callee, locals)
                || args
                    .iter()
                    .any(|arg| expr_contains_vb_int_overflow(&arg.value, locals))
        }
        ExprKind::Member { object, .. } => expr_contains_vb_int_overflow(object, locals),
        ExprKind::Index { object, index, .. } => {
            expr_contains_vb_int_overflow(object, locals)
                || expr_contains_vb_int_overflow(index, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            expr_contains_vb_int_overflow(cond, locals)
                || expr_contains_vb_int_overflow(then, locals)
                || expr_contains_vb_int_overflow(else_, locals)
        }
        ExprKind::Array(items) => items
            .iter()
            .any(|item| expr_contains_vb_int_overflow(&item.value, locals)),
        ExprKind::Tuple(items) | ExprKind::Set(items) => items
            .iter()
            .any(|item| expr_contains_vb_int_overflow(item, locals)),
        ExprKind::New { class, args } => {
            expr_contains_vb_int_overflow(class, locals)
                || args
                    .iter()
                    .any(|arg| expr_contains_vb_int_overflow(&arg.value, locals))
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            expr_contains_vb_int_overflow(left, locals)
                || expr_contains_vb_int_overflow(right, locals)
        }
        _ => false,
    }
}

fn eval_vb_int_const_expr(expr: &Expression, locals: &HashMap<String, String>) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Ident(name) => locals
            .get(&format!("{}#const", name.to_ascii_lowercase()))
            .and_then(|value| value.parse().ok()),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => eval_vb_int_const_expr(expr, locals).and_then(|value| value.checked_neg()),
        ExprKind::Binary { op, left, right } => {
            let left = eval_vb_int_const_expr(left, locals)?;
            let right = eval_vb_int_const_expr(right, locals)?;
            match op {
                BinOp::Add => left.checked_add(right),
                BinOp::Sub => left.checked_sub(right),
                BinOp::Mul => left.checked_mul(right),
                BinOp::IDiv if right != 0 => Some(left / right),
                BinOp::Mod if right != 0 => Some(left % right),
                _ => None,
            }
        }
        _ => None,
    }
}

fn rewrite_vb_bitwise_logic_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_bitwise_logic_statement(stmt, &mut HashMap::new());
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    locals.insert(
                        param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
            }
            rewrite_vb_bitwise_logic_statements(body, &mut locals);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_bitwise_logic_statements(getter, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                let mut locals = HashMap::new();
                if let Some(type_hint) = &setter.param.type_hint {
                    locals.insert(
                        setter.param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
                rewrite_vb_bitwise_logic_statements(&mut setter.body, &mut locals);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_bitwise_logic_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_vb_bitwise_logic_expr(left, locals);
            rewrite_vb_bitwise_logic_expr(right, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_vb_bitwise_logic_expr(expr, locals),
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_bitwise_logic_expr(callee, locals);
            for arg in args {
                rewrite_vb_bitwise_logic_expr(&mut arg.value, locals);
            }
        }
        ExprKind::Member { object, .. } => rewrite_vb_bitwise_logic_expr(object, locals),
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_bitwise_logic_expr(object, locals);
            rewrite_vb_bitwise_logic_expr(index, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_bitwise_logic_expr(cond, locals);
            rewrite_vb_bitwise_logic_expr(then, locals);
            rewrite_vb_bitwise_logic_expr(else_, locals);
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_bitwise_logic_expr(left, locals);
            rewrite_vb_bitwise_logic_expr(right, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_bitwise_logic_expr(&mut item.value, locals);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_bitwise_logic_expr(class, locals);
            for arg in args {
                rewrite_vb_bitwise_logic_expr(&mut arg.value, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_vb_bitwise_logic_expr(key, locals);
                        rewrite_vb_bitwise_logic_expr(value, locals);
                    }
                    ObjectProperty::Spread(value) => rewrite_vb_bitwise_logic_expr(value, locals),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_bitwise_logic_statement(value, &mut locals.clone());
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => rewrite_vb_bitwise_logic_expr(expr, locals),
            LambdaBody::Block(body) => {
                rewrite_vb_bitwise_logic_statements(body, &mut locals.clone())
            }
        },
        _ => {}
    }

    match &expr.kind {
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: inner,
        } if vb_infer_expr_type(inner, locals)
            .as_deref()
            .is_some_and(vb_type_is_numeric) =>
        {
            *expr = Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new((**inner).clone()),
            });
        }
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr: inner,
        } if !vb_expr_is_boolish(inner, locals) => {
            *expr = Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new((**inner).clone()),
            });
        }
        ExprKind::Binary {
            op: BinOp::BitAnd | BinOp::BitOr | BinOp::BitXor,
            left,
            right,
        } if vb_expr_is_boolish(left, locals) && vb_expr_is_boolish(right, locals) => {
            let mut bool_bit_expr = expr.clone();
            if let ExprKind::Binary { left, right, .. } = &mut bool_bit_expr.kind {
                coerce_vb_boolish_literal(left);
                coerce_vb_boolish_literal(right);
            }
            *expr = Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(bool_bit_expr),
                right: Box::new(Expression::int(0)),
            });
        }
        ExprKind::Binary {
            op: BinOp::Eqv,
            left,
            right,
        } if vb_infer_expr_type(left, locals).as_deref() == Some("Boolean")
            && vb_infer_expr_type(right, locals).as_deref() == Some("Boolean") =>
        {
            *expr = Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new((**left).clone()),
                right: Box::new((**right).clone()),
            });
        }
        ExprKind::Binary {
            op: BinOp::Imp,
            left,
            right,
        } if vb_infer_expr_type(left, locals).as_deref() == Some("Boolean")
            && vb_infer_expr_type(right, locals).as_deref() == Some("Boolean") =>
        {
            let not_left = Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new((**left).clone()),
            });
            *expr = Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(not_left),
                right: Box::new((**right).clone()),
            });
        }
        ExprKind::Binary {
            op: op @ (BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq),
            left,
            right,
        } if vb_infer_expr_type(left, locals).as_deref() == Some("DateTime")
            && vb_infer_expr_type(right, locals).as_deref() == Some("DateTime") =>
        {
            let cmp = call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("DateTime")),
                    field: "Compare".into(),
                    null_safe: false,
                }),
                vec![
                    Argument::positional((**left).clone()),
                    Argument::positional((**right).clone()),
                ],
            );
            *expr = Expression::new(ExprKind::Binary {
                op: *op,
                left: Box::new(cmp),
                right: Box::new(Expression::int(0)),
            });
        }
        ExprKind::Binary {
            op: op @ (BinOp::Eq | BinOp::NotEq),
            left,
            right,
        } if vb_infer_expr_type(left, locals).as_deref() == Some("Guid")
            && vb_infer_expr_type(right, locals).as_deref() == Some("Guid") =>
        {
            let left_text = call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new((**left).clone()),
                    field: "ToString".into(),
                    null_safe: false,
                }),
                vec![Argument::positional(Expression::string("D"))],
            );
            let right_text = call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new((**right).clone()),
                    field: "ToString".into(),
                    null_safe: false,
                }),
                vec![Argument::positional(Expression::string("D"))],
            );
            *expr = Expression::new(ExprKind::Binary {
                op: *op,
                left: Box::new(left_text),
                right: Box::new(right_text),
            });
        }
        ExprKind::Binary {
            op: BinOp::Eq | BinOp::NotEq,
            left,
            right,
        } if vb_infer_expr_type(left, locals).as_deref() == Some("ObjectRef")
            && vb_infer_expr_type(right, locals).as_deref() == Some("ObjectRef") =>
        {
            *expr = call_expr(
                Expression::ident("__vb_err_raise"),
                vec![Argument::positional(Expression::string(
                    "Operator '=' is not defined for reference Object values",
                ))],
            );
        }
        _ => {}
    }
}

fn vb_type_is_numeric(type_name: &str) -> bool {
    matches!(
        type_name,
        "Byte"
            | "SByte"
            | "Int16"
            | "UInt16"
            | "Int32"
            | "UInt32"
            | "Int64"
            | "UInt64"
            | "Single"
            | "Double"
            | "Decimal"
    )
}

fn vb_expr_is_boolish(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    if vb_infer_expr_type(expr, locals).as_deref() == Some("Boolean") {
        return true;
    }
    matches!(
        literal_string(expr).as_deref(),
        Some("True" | "true" | "False" | "false")
    )
}

fn vb_call_returns_bool(callee: &Expression) -> bool {
    let name = match &callee.kind {
        ExprKind::Ident(name) => name.as_str(),
        ExprKind::Member { field, .. } => field.as_str(),
        _ => return false,
    };
    matches!(
        name.to_ascii_lowercase().as_str(),
        "isnothing"
            | "isnumeric"
            | "isdate"
            | "isnull"
            | "isempty"
            | "isarray"
            | "isobject"
            | "isnullorempty"
            | "isnullorwhitespace"
            | "contains"
            | "containskey"
            | "containsvalue"
            | "exists"
            | "equals"
            | "startswith"
            | "endswith"
    )
}

fn coerce_vb_boolish_literal(expr: &mut Expression) {
    if let Some(value) = literal_string(expr) {
        if value.eq_ignore_ascii_case("true") {
            *expr = Expression::bool(true);
        } else if value.eq_ignore_ascii_case("false") {
            *expr = Expression::bool(false);
        }
    }
}

fn normalize_vb_operator_calls(module: &mut Module) {
    let mut operators: HashMap<String, Vec<(BinOp, &'static str)>> = HashMap::new();
    let mut conversions: HashMap<(String, String), (String, String)> = HashMap::new();
    collect_vb_operator_classes(&module.body, &mut operators, &mut conversions);
    if operators.is_empty() && conversions.is_empty() {
        return;
    }
    rewrite_vb_operator_call_statements(
        &mut module.body,
        &operators,
        &conversions,
        &mut HashMap::new(),
    );
}

fn collect_vb_operator_classes(
    body: &[Statement],
    operators: &mut HashMap<String, Vec<(BinOp, &'static str)>>,
    conversions: &mut HashMap<(String, String), (String, String)>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                let owner = vb_canonical_type_name(name);
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            let StmtKind::FunctionDecl {
                                name: method_name,
                                params,
                                return_type,
                                ..
                            } = &method.kind
                            else {
                                continue;
                            };
                            if method_name.starts_with("__ctype_") {
                                if let (Some(param), Some(to_type)) =
                                    (params.first(), return_type.as_ref())
                                {
                                    if let Some(from_type) = param.type_hint.as_ref() {
                                        conversions.insert(
                                            (
                                                vb_canonical_type_name(from_type)
                                                    .to_ascii_lowercase(),
                                                vb_canonical_type_name(to_type)
                                                    .to_ascii_lowercase(),
                                            ),
                                            (owner.clone(), method_name.clone()),
                                        );
                                    }
                                }
                                continue;
                            }
                            if let Some((op, method)) = vb_binop_for_dunder_method(method_name) {
                                operators
                                    .entry(owner.to_ascii_lowercase())
                                    .or_default()
                                    .push((op, method));
                            } else if method_name == "__neg__" {
                                operators
                                    .entry(owner.to_ascii_lowercase())
                                    .or_default()
                                    .push((BinOp::MatMul, "__neg__"));
                            } else if method_name == "__bitnot__" {
                                operators
                                    .entry(owner.to_ascii_lowercase())
                                    .or_default()
                                    .push((BinOp::UShr, "__bitnot__"));
                            } else if method_name == "__istrue__" {
                                operators
                                    .entry(owner.to_ascii_lowercase())
                                    .or_default()
                                    .push((BinOp::In, "__istrue__"));
                            } else if method_name == "__isfalse__" {
                                operators
                                    .entry(owner.to_ascii_lowercase())
                                    .or_default()
                                    .push((BinOp::NotIn, "__isfalse__"));
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_vb_operator_classes(
                                std::slice::from_ref(nested),
                                operators,
                                conversions,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_vb_operator_classes(body, operators, conversions)
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_operator_classes(
                            std::slice::from_ref(nested),
                            operators,
                            conversions,
                        );
                    }
                }
            }
            _ => {}
        }
    }
}

fn vb_binop_for_dunder_method(name: &str) -> Option<(BinOp, &'static str)> {
    match name {
        "__add__" => Some((BinOp::Add, "__add__")),
        "__sub__" => Some((BinOp::Sub, "__sub__")),
        "__mul__" => Some((BinOp::Mul, "__mul__")),
        "__truediv__" => Some((BinOp::Div, "__truediv__")),
        "__mod__" => Some((BinOp::Mod, "__mod__")),
        "__eq__" => Some((BinOp::Eq, "__eq__")),
        "__lt__" => Some((BinOp::Lt, "__lt__")),
        "__le__" => Some((BinOp::LtEq, "__le__")),
        "__gt__" => Some((BinOp::Gt, "__gt__")),
        "__ge__" => Some((BinOp::GtEq, "__ge__")),
        "__like__" => Some((BinOp::Like, "__like__")),
        _ => None,
    }
}

fn rewrite_vb_operator_call_statements(
    body: &mut [Statement],
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
    conversions: &HashMap<(String, String), (String, String)>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_operator_call_statement(stmt, operators, conversions, locals);
    }
}

fn rewrite_vb_operator_call_statement(
    stmt: &mut Statement,
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
    conversions: &HashMap<(String, String), (String, String)>,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_operator_call_expr(init, operators, conversions, locals);
                    if let Some(type_hint) = &decl.type_hint {
                        rewrite_vb_conversion_expr(init, type_hint, conversions, locals);
                    }
                }
                let BindingPattern::Ident(name) = &decl.pattern else {
                    continue;
                };
                let ty = decl
                    .type_hint
                    .as_ref()
                    .map(|hint| vb_canonical_type_name(hint))
                    .or_else(|| {
                        decl.init
                            .as_ref()
                            .and_then(|expr| vb_infer_operator_result_type(expr, locals, operators))
                    })
                    .or_else(|| {
                        decl.init
                            .as_ref()
                            .and_then(|expr| vb_infer_expr_type(expr, locals))
                    });
                if let Some(ty) = ty {
                    locals.insert(name.to_ascii_lowercase(), ty);
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_operator_call_expr(expr, operators, conversions, locals);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_operator_call_expr(target, operators, conversions, locals);
            }
            rewrite_vb_operator_call_expr(value, operators, conversions, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_vb_operator_call_expr(target, operators, conversions, locals);
            rewrite_vb_operator_call_expr(value, operators, conversions, locals);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = locals.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(
                        param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
            }
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut scoped);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_operator_call_member(member, operators, conversions);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_operator_call_expr(cond, operators, conversions, locals);
            rewrite_vb_operator_truth_expr(cond, operators, locals, false);
            rewrite_vb_operator_call_statements(
                then_body,
                operators,
                conversions,
                &mut locals.clone(),
            );
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_operator_call_expr(elif_cond, operators, conversions, locals);
                rewrite_vb_operator_truth_expr(elif_cond, operators, locals, false);
                rewrite_vb_operator_call_statements(
                    elif_body,
                    operators,
                    conversions,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                rewrite_vb_operator_call_statements(
                    else_body,
                    operators,
                    conversions,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut scoped = locals.clone();
            if let Some(init) = init {
                rewrite_vb_operator_call_statement(init, operators, conversions, &mut scoped);
            }
            if let Some(cond) = cond {
                rewrite_vb_operator_call_expr(cond, operators, conversions, &mut scoped);
                rewrite_vb_operator_truth_expr(cond, operators, &scoped, false);
            }
            if let Some(update) = update {
                rewrite_vb_operator_call_expr(update, operators, conversions, &mut scoped);
            }
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut scoped);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_operator_call_expr(iter, operators, conversions, locals);
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_operator_call_statements(
                    else_body,
                    operators,
                    conversions,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_operator_call_expr(cond, operators, conversions, locals);
            rewrite_vb_operator_truth_expr(cond, operators, locals, false);
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_operator_call_statements(
                    else_body,
                    operators,
                    conversions,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut locals.clone());
        }
        _ => {}
    }
}

fn rewrite_vb_operator_call_member(
    member: &mut ClassMember,
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
    conversions: &HashMap<(String, String), (String, String)>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_operator_call_statement(stmt, operators, conversions, &mut HashMap::new());
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    locals.insert(
                        param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
            }
            rewrite_vb_operator_call_statements(body, operators, conversions, &mut locals);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_operator_call_statements(
                    getter,
                    operators,
                    conversions,
                    &mut HashMap::new(),
                );
            }
            if let Some(setter) = setter {
                let mut locals = HashMap::new();
                locals.insert(
                    setter.param.name.to_ascii_lowercase(),
                    setter
                        .param
                        .type_hint
                        .as_deref()
                        .map(vb_canonical_type_name)
                        .unwrap_or_else(|| "Object".to_string()),
                );
                rewrite_vb_operator_call_statements(
                    &mut setter.body,
                    operators,
                    conversions,
                    &mut locals,
                );
            }
        }
        _ => {}
    }
}

fn vb_infer_operator_result_type(
    expr: &Expression,
    locals: &HashMap<String, String>,
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
) -> Option<String> {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = vb_infer_expr_type(object, locals)?;
    let has_operator = operators
        .get(&receiver_type.to_ascii_lowercase())
        .is_some_and(|entries| entries.iter().any(|(_, method)| *method == field));
    if !has_operator {
        return None;
    }
    match field.as_str() {
        "__eq__" | "__lt__" | "__le__" | "__gt__" | "__ge__" | "__like__" => {
            Some("Boolean".to_string())
        }
        _ => Some(receiver_type),
    }
}

fn rewrite_vb_conversion_expr(
    expr: &mut Expression,
    target_type: &str,
    conversions: &HashMap<(String, String), (String, String)>,
    locals: &HashMap<String, String>,
) {
    let target = vb_canonical_type_name(target_type).to_ascii_lowercase();
    let (source_expr, source_type) = match &expr.kind {
        ExprKind::Cast { expr: inner, .. } => {
            (inner.as_ref().clone(), vb_infer_expr_type(inner, locals))
        }
        _ => (expr.clone(), vb_infer_expr_type(expr, locals)),
    };
    let Some(source_type) = source_type.map(|ty| vb_canonical_type_name(&ty).to_ascii_lowercase())
    else {
        return;
    };
    if source_type == target {
        return;
    }
    let Some((owner, method)) = conversions.get(&(source_type, target)) else {
        return;
    };
    let callee = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(owner)),
        field: method.clone(),
        null_safe: false,
    });
    *expr = call_expr(callee, vec![Argument::positional(source_expr)]);
}

fn rewrite_vb_operator_truth_expr(
    expr: &mut Expression,
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
    locals: &HashMap<String, String>,
    negated: bool,
) {
    if let ExprKind::Unary {
        op: UnaryOp::Not,
        expr: inner,
    } = &mut expr.kind
    {
        rewrite_vb_operator_truth_expr(inner, operators, locals, !negated);
        return;
    }
    let Some(expr_type) = vb_infer_expr_type(expr, locals) else {
        return;
    };
    let wanted = if negated { "__isfalse__" } else { "__istrue__" };
    let has_operator = operators
        .get(&expr_type.to_ascii_lowercase())
        .is_some_and(|entries| entries.iter().any(|(_, method)| *method == wanted));
    if !has_operator {
        return;
    }
    let callee = Expression::new(ExprKind::Member {
        object: Box::new(expr.clone()),
        field: wanted.to_string(),
        null_safe: false,
    });
    *expr = call_expr(callee, Vec::new());
}

fn rewrite_vb_operator_call_expr(
    expr: &mut Expression,
    operators: &HashMap<String, Vec<(BinOp, &'static str)>>,
    conversions: &HashMap<(String, String), (String, String)>,
    locals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_vb_operator_call_expr(left, operators, conversions, locals);
            rewrite_vb_operator_call_expr(right, operators, conversions, locals);
        }
        ExprKind::Member { object, .. } => {
            rewrite_vb_operator_call_expr(object, operators, conversions, locals)
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_operator_call_expr(callee, operators, conversions, locals);
            for arg in args {
                rewrite_vb_operator_call_expr(&mut arg.value, operators, conversions, locals);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_vb_operator_call_expr(target, operators, conversions, locals);
            rewrite_vb_operator_call_expr(value, operators, conversions, locals);
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_operator_call_expr(left, operators, conversions, locals);
            rewrite_vb_operator_call_expr(right, operators, conversions, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr)
        | ExprKind::Yield(Some(expr)) => {
            rewrite_vb_operator_call_expr(expr, operators, conversions, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_operator_call_expr(cond, operators, conversions, locals);
            rewrite_vb_operator_call_expr(then, operators, conversions, locals);
            rewrite_vb_operator_call_expr(else_, operators, conversions, locals);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_operator_call_expr(object, operators, conversions, locals);
            rewrite_vb_operator_call_expr(index, operators, conversions, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_operator_call_expr(&mut item.value, operators, conversions, locals);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_operator_call_expr(class, operators, conversions, locals);
            for arg in args {
                rewrite_vb_operator_call_expr(&mut arg.value, operators, conversions, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_vb_operator_call_expr(key, operators, conversions, locals);
                        rewrite_vb_operator_call_expr(value, operators, conversions, locals);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_vb_operator_call_expr(value, operators, conversions, locals)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_operator_call_statement(
                            value,
                            operators,
                            conversions,
                            &mut locals.clone(),
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => {
                rewrite_vb_operator_call_expr(expr, operators, conversions, locals)
            }
            LambdaBody::Block(body) => {
                rewrite_vb_operator_call_statements(
                    body,
                    operators,
                    conversions,
                    &mut locals.clone(),
                );
            }
        },
        _ => {}
    }

    if let ExprKind::Cast { type_name, .. } = &expr.kind {
        let target_type = type_name.clone();
        rewrite_vb_conversion_expr(expr, &target_type, conversions, locals);
    }

    let ExprKind::Binary { op, left, right } = &expr.kind else {
        if let ExprKind::Unary { op, expr: inner } = &expr.kind {
            let Some(left_type) = vb_infer_expr_type(inner, locals) else {
                return;
            };
            let wanted = match op {
                UnaryOp::Neg => "__neg__",
                UnaryOp::BitNot | UnaryOp::Not => "__bitnot__",
                _ => return,
            };
            let Some(method_name) =
                operators
                    .get(&left_type.to_ascii_lowercase())
                    .and_then(|entries| {
                        entries
                            .iter()
                            .find_map(|(_, method)| (*method == wanted).then_some(*method))
                    })
            else {
                return;
            };
            let callee = Expression::new(ExprKind::Member {
                object: Box::new((**inner).clone()),
                field: method_name.to_string(),
                null_safe: false,
            });
            *expr = call_expr(callee, Vec::new());
        }
        return;
    };
    if *op == BinOp::Like {
        let left_has_like = vb_infer_expr_type(left, locals)
            .and_then(|left_type| operators.get(&left_type.to_ascii_lowercase()))
            .is_some_and(|entries| {
                entries
                    .iter()
                    .any(|(known_op, method)| *known_op == BinOp::Like && *method == "__like__")
            });
        let right_has_like = vb_infer_expr_type(right, locals)
            .and_then(|right_type| operators.get(&right_type.to_ascii_lowercase()))
            .is_some_and(|entries| {
                entries
                    .iter()
                    .any(|(known_op, method)| *known_op == BinOp::Like && *method == "__like__")
            });
        if !left_has_like && !right_has_like {
            let pattern = if let Some(lit) = literal_string(right) {
                Expression::string(&vb_like_pattern_to_regex(&lit))
            } else {
                Expression::new(ExprKind::Cast {
                    expr: Box::new((**right).clone()),
                    type_name: "__vb_like_pattern".to_string(),
                })
            };
            *expr = vb_bool_expr(call_expr(
                build_dotted_expr("System.Text.RegularExpressions.Regex.IsMatch"),
                vec![
                    Argument::positional((**left).clone()),
                    Argument::positional(pattern),
                ],
            ));
            return;
        }
    }
    let Some(left_type) = vb_infer_expr_type(left, locals) else {
        return;
    };
    let Some(method_name) =
        operators
            .get(&left_type.to_ascii_lowercase())
            .and_then(|entries| {
                entries.iter().find_map(
                    |(known_op, method)| if known_op == op { Some(*method) } else { None },
                )
            })
            .or_else(|| {
                vb_infer_expr_type(right, locals).and_then(|right_type| {
                    operators
                        .get(&right_type.to_ascii_lowercase())
                        .and_then(|entries| {
                            entries.iter().find_map(|(known_op, method)| {
                                if known_op == op { Some(*method) } else { None }
                            })
                        })
                })
            })
    else {
        return;
    };
    let receiver_is_right = vb_infer_expr_type(right, locals)
        .and_then(|right_type| {
            operators
                .get(&right_type.to_ascii_lowercase())
                .map(|entries| {
                    entries
                        .iter()
                        .any(|(known_op, method)| known_op == op && *method == method_name)
                })
        })
        .unwrap_or(false)
        && !vb_infer_expr_type(left, locals)
            .and_then(|left_type| {
                operators
                    .get(&left_type.to_ascii_lowercase())
                    .map(|entries| {
                        entries
                            .iter()
                            .any(|(known_op, method)| known_op == op && *method == method_name)
                    })
            })
            .unwrap_or(false);
    let (receiver, argument) = if receiver_is_right {
        ((**right).clone(), (**left).clone())
    } else {
        ((**left).clone(), (**right).clone())
    };
    let mut callee = Expression::new(ExprKind::Member {
        object: Box::new(receiver),
        field: method_name.to_string(),
        null_safe: false,
    });
    callee.span = expr.span;
    *expr = call_expr(callee, vec![Argument::positional(argument)]);
}

fn normalize_vb_flags_enum_ops(module: &mut Module) {
    let mut enums = HashMap::new();
    collect_vb_enum_infos(&module.body, &mut enums);
    if enums.is_empty() {
        return;
    }
    rewrite_vb_flags_enum_statements(&mut module.body, &enums, &mut HashMap::new());
}

#[derive(Clone, Default)]
struct VbEnumInfo {
    display_name: String,
    members_by_name: HashMap<String, i64>,
    names_by_value: HashMap<i64, String>,
}

fn collect_vb_enum_infos(body: &[Statement], enums: &mut HashMap<String, VbEnumInfo>) {
    collect_vb_enum_infos_scoped(body, None, enums);
}

fn collect_vb_enum_infos_scoped(
    body: &[Statement],
    scope: Option<&str>,
    enums: &mut HashMap<String, VbEnumInfo>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::EnumDecl {
                name,
                members,
                is_flags: _,
                ..
            } => {
                let mut info = VbEnumInfo {
                    display_name: name.clone(),
                    ..VbEnumInfo::default()
                };
                for member in members {
                    if let Some(value) = member.value.as_ref().and_then(literal_int) {
                        info.members_by_name
                            .insert(member.name.to_ascii_lowercase(), value);
                        info.names_by_value
                            .entry(value)
                            .or_insert_with(|| member.name.clone());
                    }
                }
                let key = name.to_ascii_lowercase();
                if let Some(scope) = scope {
                    enums.insert(
                        format!("{}.{}", scope.to_ascii_lowercase(), key),
                        info.clone(),
                    );
                    enums.entry(key).or_insert(info);
                } else {
                    enums.insert(key, info);
                }
            }
            StmtKind::NamespaceDecl { name, body, .. } => {
                let next_scope = scoped_vb_name(scope, name);
                collect_vb_enum_infos_scoped(body, Some(&next_scope), enums);
            }
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                let next_scope = scoped_vb_name(scope, name);
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_enum_infos_scoped(
                            std::slice::from_ref(nested),
                            Some(&next_scope),
                            enums,
                        );
                    }
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_enum_infos_scoped(std::slice::from_ref(nested), scope, enums);
                    }
                }
            }
            _ => {}
        }
    }
}

fn scoped_vb_name(scope: Option<&str>, name: &str) -> String {
    match scope {
        Some(scope) if !scope.is_empty() => format!("{}.{}", scope, name),
        _ => name.to_string(),
    }
}

fn literal_int(expr: &Expression) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => literal_int(expr).map(|value| -value),
        ExprKind::Cast { expr, .. } => literal_int(expr),
        _ => None,
    }
}

fn vb_enum_member_value(
    expr: &Expression,
    enums: &HashMap<String, VbEnumInfo>,
) -> Option<(String, i64)> {
    match &expr.kind {
        ExprKind::Member { object, .. } => {
            let enum_name = dotted_expr_name(object)?.to_ascii_lowercase();
            let field = match &expr.kind {
                ExprKind::Member { field, .. } => field.to_ascii_lowercase(),
                _ => return None,
            };
            let value = *enums.get(&enum_name)?.members_by_name.get(&field)?;
            Some((enum_name, value))
        }
        _ => None,
    }
}

fn rewrite_vb_flags_enum_statements(
    body: &mut [Statement],
    enums: &HashMap<String, VbEnumInfo>,
    locals: &mut HashMap<String, (String, Option<i64>)>,
) {
    for stmt in body {
        rewrite_vb_flags_enum_statement(stmt, enums, locals);
    }
}

fn rewrite_vb_flags_enum_statement(
    stmt: &mut Statement,
    enums: &HashMap<String, VbEnumInfo>,
    locals: &mut HashMap<String, (String, Option<i64>)>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_flags_enum_expr(expr, enums, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_flags_enum_expr(init, enums, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let enum_type = decl
                        .type_hint
                        .as_ref()
                        .map(|hint| vb_canonical_type_name(hint).to_ascii_lowercase())
                        .filter(|ty| enums.contains_key(ty));
                    if let Some(enum_type) = enum_type {
                        let value = decl
                            .init
                            .as_ref()
                            .and_then(|init| eval_vb_flags_enum_int(init, locals));
                        locals.insert(name.to_ascii_lowercase(), (enum_type, value));
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in &mut *targets {
                rewrite_vb_flags_enum_expr(target, enums, locals);
            }
            rewrite_vb_flags_enum_expr(value, enums, locals);
            if let Some(Expression {
                kind: ExprKind::Ident(name),
                ..
            }) = targets.first()
            {
                let new_value = eval_vb_flags_enum_int(value, locals);
                if let Some((enum_type, known)) = locals.get_mut(&name.to_ascii_lowercase()) {
                    if enums.contains_key(enum_type) {
                        *known = new_value;
                    }
                }
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_vb_flags_enum_expr(target, enums, locals);
            rewrite_vb_flags_enum_expr(value, enums, locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_flags_enum_expr(cond, enums, locals);
            rewrite_vb_flags_enum_statements(then_body, enums, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_flags_enum_expr(elif_cond, enums, locals);
                rewrite_vb_flags_enum_statements(elif_body, enums, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                rewrite_vb_flags_enum_statements(else_body, enums, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_vb_flags_enum_statement(init, enums, &mut locals.clone());
            }
            if let Some(cond) = cond {
                rewrite_vb_flags_enum_expr(cond, enums, locals);
            }
            if let Some(update) = update {
                rewrite_vb_flags_enum_expr(update, enums, locals);
            }
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone());
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_flags_enum_expr(iter, enums, locals);
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_flags_enum_statements(else_body, enums, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_flags_enum_expr(cond, enums, locals);
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_flags_enum_statements(else_body, enums, &mut locals.clone());
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone());
            rewrite_vb_flags_enum_expr(cond, enums, locals);
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_flags_enum_member(member, enums);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_flags_enum_statements(body, enums, &mut locals.clone())
        }
        _ => {}
    }
}

fn rewrite_vb_flags_enum_member(member: &mut ClassMember, enums: &HashMap<String, VbEnumInfo>) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_flags_enum_statement(stmt, enums, &mut HashMap::new());
        }
        ClassMember::Constructor { body, .. } => {
            rewrite_vb_flags_enum_statements(body, enums, &mut HashMap::new())
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_flags_enum_statements(getter, enums, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                rewrite_vb_flags_enum_statements(&mut setter.body, enums, &mut HashMap::new());
            }
        }
        _ => {}
    }
}

fn rewrite_vb_flags_enum_expr(
    expr: &mut Expression,
    enums: &HashMap<String, VbEnumInfo>,
    locals: &HashMap<String, (String, Option<i64>)>,
) {
    match &mut expr.kind {
        ExprKind::Binary { op: _, left, right } => {
            rewrite_vb_flags_enum_expr(left, enums, locals);
            rewrite_vb_flags_enum_expr(right, enums, locals);
        }
        ExprKind::Call { callee, args, .. } => {
            let is_writeline = is_vb_console_writeline(callee);
            if is_writeline {
                for arg in &mut *args {
                    if let ExprKind::Ident(local) = &arg.value.kind {
                        if let Some((enum_type, Some(value))) =
                            locals.get(&local.to_ascii_lowercase())
                        {
                            if let Some(name) = enums
                                .get(enum_type)
                                .and_then(|info| info.names_by_value.get(value))
                            {
                                arg.value = Expression::string(name);
                            }
                        }
                    }
                    if let Some(text) = fold_vb_writeline_int64_boundary_cast(&arg.value) {
                        arg.value = Expression::string(&text);
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("HasFlag") {
                    if let ExprKind::Ident(local) = &object.kind {
                        if let Some((enum_type, _)) = locals.get(&local.to_ascii_lowercase()) {
                            if enums.contains_key(enum_type) {
                                if let Some(arg) = args.first() {
                                    let flag = arg.value.clone();
                                    *expr = Expression::new(ExprKind::Binary {
                                        op: BinOp::Eq,
                                        left: Box::new(Expression::new(ExprKind::Binary {
                                            op: BinOp::BitAnd,
                                            left: Box::new((**object).clone()),
                                            right: Box::new(flag.clone()),
                                        })),
                                        right: Box::new(flag),
                                    });
                                    return;
                                }
                            }
                        }
                    }
                } else if field.eq_ignore_ascii_case("ToString") {
                    if let Some((enum_type, value)) = vb_enum_member_value(object, enums) {
                        if let Some(name) = enums
                            .get(&enum_type)
                            .and_then(|info| info.names_by_value.get(&value))
                        {
                            *expr = Expression::string(name);
                            return;
                        }
                    } else if let ExprKind::Ident(local) = &object.kind {
                        if let Some((enum_type, Some(value))) =
                            locals.get(&local.to_ascii_lowercase())
                        {
                            if let Some(name) = enums
                                .get(enum_type)
                                .and_then(|info| info.names_by_value.get(value))
                            {
                                *expr = Expression::string(name);
                                return;
                            }
                        }
                    }
                } else if field.eq_ignore_ascii_case("GetType") {
                    if let Some((enum_type, _)) = vb_enum_member_value(object, enums) {
                        if let Some(info) = enums.get(&enum_type) {
                            *expr = Expression::string(&info.display_name);
                            return;
                        }
                        return;
                    }
                }
            }
            rewrite_vb_flags_enum_expr(callee, enums, locals);
            for arg in &mut *args {
                rewrite_vb_flags_enum_expr(&mut arg.value, enums, locals);
                if is_writeline {
                    if let Some(text) = fold_vb_writeline_int64_boundary_cast(&arg.value) {
                        arg.value = Expression::string(&text);
                    }
                }
            }
        }
        ExprKind::Member { object, field, .. } => {
            if field.eq_ignore_ascii_case("Name") {
                if let ExprKind::Call { callee, .. } = &object.kind {
                    if let ExprKind::Member {
                        object: enum_member,
                        field: get_type,
                        ..
                    } = &callee.kind
                    {
                        if get_type.eq_ignore_ascii_case("GetType") {
                            if let Some((enum_type, _)) = vb_enum_member_value(enum_member, enums) {
                                if let Some(info) = enums.get(&enum_type) {
                                    *expr = Expression::string(&info.display_name);
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            rewrite_vb_flags_enum_expr(object, enums, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_vb_flags_enum_expr(expr, enums, locals),
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_flags_enum_expr(object, enums, locals);
            rewrite_vb_flags_enum_expr(index, enums, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_flags_enum_expr(&mut item.value, enums, locals);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_flags_enum_expr(class, enums, locals);
            for arg in args {
                rewrite_vb_flags_enum_expr(&mut arg.value, enums, locals);
            }
        }
        _ => {}
    }

    if let Some((_, value)) = vb_enum_member_value(expr, enums) {
        *expr = Expression::int(value);
    }
}

fn is_vb_console_writeline(callee: &Expression) -> bool {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return false;
    };
    field.eq_ignore_ascii_case("WriteLine")
        && dotted_expr_name(object).is_some_and(|name| {
            name.eq_ignore_ascii_case("Console") || name.eq_ignore_ascii_case("System.Console")
        })
}

fn fold_vb_writeline_int64_boundary_cast(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !(name.eq_ignore_ascii_case("CLng") || name.eq_ignore_ascii_case("Long")) || args.len() != 1
    {
        return None;
    }
    let value = literal_int(&args[0].value)?;
    (value == i64::MAX).then(|| value.to_string())
}

fn eval_vb_flags_enum_int(
    expr: &Expression,
    locals: &HashMap<String, (String, Option<i64>)>,
) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Ident(name) => locals
            .get(&name.to_ascii_lowercase())
            .and_then(|(_, value)| *value),
        ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr,
        } => eval_vb_flags_enum_int(expr, locals).map(|value| !value),
        ExprKind::Binary { op, left, right } => {
            let left = eval_vb_flags_enum_int(left, locals)?;
            let right = eval_vb_flags_enum_int(right, locals)?;
            match op {
                BinOp::BitOr | BinOp::Or => Some(left | right),
                BinOp::BitAnd | BinOp::And => Some(left & right),
                BinOp::BitXor | BinOp::Xor => Some(left ^ right),
                _ => None,
            }
        }
        _ => None,
    }
}

fn parse_vb_attribute_argument(text: &str) -> Option<Argument> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }
    if let Some((lhs, rhs)) = trimmed.split_once(":=") {
        let name = lhs.trim().rsplit('.').next()?.to_string();
        return Some(Argument {
            value: parse_expression_str(rhs.trim()).ok()?,
            name: Some(name),
            by_ref: false,
            spread: false,
        });
    }
    Some(Argument::positional(parse_expression_str(trimmed).ok()?))
}

fn vb_attribute_line_is_extension(raw: &str) -> bool {
    let normalized = raw
        .trim_start()
        .trim_start_matches('<')
        .trim_start()
        .to_ascii_lowercase();
    normalized.starts_with("extension")
        || normalized.starts_with("runtime.compilerservices.extension")
}

fn parse_vb_attribute_specs(raw: &str) -> Vec<Expression> {
    let mut attrs = Vec::new();
    let mut rest = raw.trim_start();
    while let Some(after_lt) = rest.strip_prefix('<') {
        let Some(close_idx) = after_lt.find('>') else {
            break;
        };
        let attr_src = &after_lt[..close_idx];
        for part in split_vb_top_level_list(attr_src) {
            let trimmed = part.trim();
            if trimmed.is_empty() {
                continue;
            }
            let (name, args) = if let Some(open_idx) = trimmed.find('(') {
                let close_idx = trimmed
                    .rfind(')')
                    .unwrap_or(trimmed.len().saturating_sub(1));
                let name = trimmed[..open_idx].trim();
                let args = split_vb_top_level_list(&trimmed[open_idx + 1..close_idx])
                    .into_iter()
                    .filter_map(|arg| parse_vb_attribute_argument(&arg))
                    .collect();
                (name, args)
            } else {
                (trimmed, Vec::new())
            };
            attrs.push(Expression::new(ExprKind::New {
                class: Box::new(build_dotted_expr(&normalize_vb_attribute_type_name(name))),
                args,
            }));
        }
        rest = after_lt[close_idx + 1..].trim_start();
        if rest.starts_with('_') {
            rest = rest[1..].trim_start();
        }
        if !rest.starts_with('<') {
            break;
        }
    }
    attrs
}

fn apply_vb_pending_decorators(stmt: &mut Statement, pending: &mut Vec<Expression>) {
    if pending.is_empty() {
        return;
    }
    match &mut stmt.kind {
        StmtKind::ClassDecl { decorators, .. }
        | StmtKind::StructDecl { decorators, .. }
        | StmtKind::InterfaceDecl { decorators, .. }
        | StmtKind::EnumDecl { decorators, .. } => {
            decorators.splice(0..0, pending.drain(..));
        }
        StmtKind::FunctionDecl { modifiers, .. } => {
            modifiers.decorators.splice(0..0, pending.drain(..));
        }
        _ => {}
    }
}

fn vb_attribute_leaf_name(expr: &Expression) -> Option<String> {
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

#[derive(Clone)]
enum VbDateValue {
    Date(NaiveDate),
    Time(NaiveTime),
    DateTime(NaiveDateTime),
}

#[derive(Clone, Copy)]
struct VbTimeSpanValue {
    days: i64,
    total_seconds: i64,
}

fn zero_arg_call(name: &str) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: vec![],
        optional: false,
    })
}

fn call_expr(callee: Expression, args: Vec<Argument>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn dotted_expr_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } => Some(format!("{}.{}", dotted_expr_name(object)?, field)),
        _ => None,
    }
}

fn build_vb_math_call(name: &str, arg: Expression) -> Expression {
    call_expr(
        build_dotted_expr(&format!("System.Math.{}", name)),
        vec![Argument::positional(arg)],
    )
}

fn build_vb_bankers_round_expr(value: Expression) -> Expression {
    let floor = build_vb_math_call("Floor", value.clone());
    let ceil = build_vb_math_call("Ceiling", value.clone());
    let frac = Expression::new(ExprKind::Binary {
        op: BinOp::Sub,
        left: Box::new(value),
        right: Box::new(floor.clone()),
    });
    let half = Expression::new(ExprKind::Lit(Literal::Float(0.5)));
    let floor_is_even = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Mod,
            left: Box::new(floor.clone()),
            right: Box::new(Expression::int(2)),
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

fn round_decimal_text_to_even(text: &str, digits: usize) -> Option<f64> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }

    let (negative, body) = if let Some(rest) = trimmed.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = trimmed.strip_prefix('+') {
        (false, rest)
    } else {
        (false, trimmed)
    };

    let mut parts = body.splitn(2, '.');
    let mut whole = parts.next().unwrap_or("0").to_string();
    let frac = parts.next().unwrap_or("");
    if digits >= frac.len() {
        return trimmed.parse().ok();
    }

    let mut kept: Vec<u8> = frac.as_bytes()[..digits].to_vec();
    let next = frac.as_bytes()[digits];
    let rest_nonzero = frac.as_bytes()[digits + 1..]
        .iter()
        .any(|digit| *digit != b'0');
    let last_kept = kept
        .last()
        .copied()
        .or_else(|| whole.as_bytes().last().copied())
        .unwrap_or(b'0');
    let round_up = match next.cmp(&b'5') {
        std::cmp::Ordering::Greater => true,
        std::cmp::Ordering::Less => false,
        std::cmp::Ordering::Equal => rest_nonzero || ((last_kept - b'0') % 2 == 1),
    };

    if round_up {
        let mut carry = true;
        for digit in kept.iter_mut().rev() {
            if *digit == b'9' {
                *digit = b'0';
            } else {
                *digit += 1;
                carry = false;
                break;
            }
        }
        if carry {
            let mut whole_digits: Vec<u8> = whole.into_bytes();
            for digit in whole_digits.iter_mut().rev() {
                if *digit == b'9' {
                    *digit = b'0';
                } else if digit.is_ascii_digit() {
                    *digit += 1;
                    carry = false;
                    break;
                }
            }
            if carry {
                whole_digits.insert(0, b'1');
            }
            whole = String::from_utf8(whole_digits).ok()?;
        }
    }

    let frac_text = String::from_utf8(kept).ok()?;
    let rounded = if digits == 0 {
        whole
    } else {
        format!("{}.{}", whole, frac_text)
    };
    let signed = if negative {
        format!("-{}", rounded)
    } else {
        rounded
    };
    signed.parse().ok()
}

fn try_fold_vb_decimal_round(value: &Expression, digits: &Expression) -> Option<Expression> {
    let digits = literal_number(digits)?;
    if digits < 0.0 || digits.fract() != 0.0 {
        return None;
    }
    let digits = digits as usize;
    let ExprKind::Call { callee, args, .. } = &value.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if !name.eq_ignore_ascii_case("cdec") || args.len() != 1 {
        return None;
    }
    let text = literal_string(&args[0].value)?;
    let rounded = round_decimal_text_to_even(&text, digits)?;
    Some(Expression::new(ExprKind::Lit(Literal::Float(rounded))))
}

fn try_fold_vb_double_round(value: &Expression, digits: &Expression) -> Option<Expression> {
    let digits = literal_number(digits)?;
    if digits < 0.0 || digits.fract() != 0.0 {
        return None;
    }
    let value = literal_number(value)?;
    let rounded = format!("{:.*}", digits as usize, value).parse().ok()?;
    Some(Expression::new(ExprKind::Lit(Literal::Float(rounded))))
}

fn build_vb_precision_round_expr(value: Expression, digits: Expression) -> Expression {
    if let Some(folded) = try_fold_vb_decimal_round(&value, &digits) {
        return folded;
    }

    let pow_left = call_expr(
        build_dotted_expr("Math.Pow"),
        vec![
            Argument::positional(Expression::float(10.0)),
            Argument::positional(digits.clone()),
        ],
    );
    let pow_right = call_expr(
        build_dotted_expr("Math.Pow"),
        vec![
            Argument::positional(Expression::float(10.0)),
            Argument::positional(digits),
        ],
    );
    let scaled = Expression::new(ExprKind::Binary {
        op: BinOp::Mul,
        left: Box::new(value),
        right: Box::new(pow_left),
    });
    let rounded = build_vb_bankers_round_expr(scaled);
    Expression::new(ExprKind::Binary {
        op: BinOp::Div,
        left: Box::new(rounded),
        right: Box::new(pow_right),
    })
}

fn vb_like_pattern_to_regex(pattern: &str) -> String {
    let mut regex = String::from("^");
    let mut chars = pattern.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '*' => regex.push_str(".*"),
            '?' => regex.push('.'),
            '#' => regex.push_str("\\d"),
            '[' => {
                let mut content = String::new();
                if matches!(chars.peek(), Some('!')) {
                    chars.next();
                    content.push('^');
                }
                let mut closed = false;
                for item in chars.by_ref() {
                    match item {
                        ']' => {
                            closed = true;
                            break;
                        }
                        '\\' | '^' => {
                            content.push('\\');
                            content.push(item);
                        }
                        _ => content.push(item),
                    }
                }
                if closed {
                    regex.push('[');
                    regex.push_str(&content);
                    regex.push(']');
                } else {
                    regex.push_str("\\[");
                    for ch in content.chars() {
                        match ch {
                            '.' | '+' | '(' | ')' | '{' | '}' | ']' | '^' | '$' | '|' | '\\'
                            | '*' | '?' | '[' => {
                                regex.push('\\');
                                regex.push(ch);
                            }
                            _ => regex.push(ch),
                        }
                    }
                }
            }
            '.' | '+' | '(' | ')' | '{' | '}' | ']' | '^' | '$' | '|' | '\\' => {
                regex.push('\\');
                regex.push(ch);
            }
            _ => regex.push(ch),
        }
    }
    regex.push('$');
    regex
}

fn maybe_rewrite_vb_binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    let (left, right) = if op == BinOp::Concat {
        (
            vb_stringify_bool_for_concat(left, &HashMap::new()),
            vb_stringify_bool_for_concat(right, &HashMap::new()),
        )
    } else {
        (left, right)
    };

    if matches!(
        op,
        BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq
    ) {
        if let Some(expr) = maybe_rewrite_vb_loose_comparison(op, &left, &right) {
            return expr;
        }
    }

    let binary = Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    });

    if matches!(op, BinOp::Is | BinOp::IsNot) {
        return vb_bool_expr(binary);
    }

    binary
}

fn vb_stringify_bool_for_concat(expr: Expression, locals: &HashMap<String, String>) -> Expression {
    match expr.kind {
        ExprKind::Lit(Literal::Bool(value)) => {
            Expression::string(if value { "True" } else { "False" })
        }
        ExprKind::Member {
            object,
            field,
            null_safe,
        } if field.eq_ignore_ascii_case("Success") => Expression::new(ExprKind::Ternary {
            cond: Box::new(Expression::new(ExprKind::Member {
                object,
                field,
                null_safe,
            })),
            then: Box::new(Expression::string("True")),
            else_: Box::new(Expression::string("False")),
        }),
        ExprKind::Ternary { cond, then, else_ }
            if matches!(then.kind, ExprKind::Lit(Literal::Bool(true)))
                && matches!(else_.kind, ExprKind::Lit(Literal::Bool(false))) =>
        {
            Expression::new(ExprKind::Ternary {
                cond,
                then: Box::new(Expression::string("True")),
                else_: Box::new(Expression::string("False")),
            })
        }
        ExprKind::Ternary { cond, then, else_ }
            if literal_string(&then).is_some_and(|value| value.eq_ignore_ascii_case("true"))
                && literal_string(&else_)
                    .is_some_and(|value| value.eq_ignore_ascii_case("false")) =>
        {
            Expression::new(ExprKind::Ternary { cond, then, else_ })
        }
        ExprKind::Binary { op, left, right }
            if matches!(
                op,
                BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq
            ) =>
        {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary { op, left, right })),
                then: Box::new(Expression::string("True")),
                else_: Box::new(Expression::string("False")),
            })
        }
        kind if vb_expr_is_boolish(&Expression::new(kind.clone()), locals) => {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(vb_bool_condition_expr(Expression::new(kind))),
                then: Box::new(Expression::string("True")),
                else_: Box::new(Expression::string("False")),
            })
        }
        kind => vb_to_string_for_concat(Expression::new(kind), locals),
    }
}

fn vb_to_string_for_concat(expr: Expression, locals: &HashMap<String, String>) -> Expression {
    if vb_infer_expr_type(&expr, locals).as_deref() == Some("String")
        || matches!(expr.kind, ExprKind::Lit(Literal::Str(_)))
    {
        return expr;
    }
    if let ExprKind::Call { callee, .. } = &expr.kind {
        if dotted_expr_name(callee).as_deref().is_some_and(|name| {
            matches!(
                name.to_ascii_lowercase().as_str(),
                "str" | "str$" | "strconv"
            )
        }) {
            return expr;
        }
    }
    call_expr(
        Expression::ident("StrConv"),
        vec![Argument::positional(expr)],
    )
}

fn maybe_rewrite_vb_loose_comparison(
    op: BinOp,
    left: &Expression,
    right: &Expression,
) -> Option<Expression> {
    let left_number = vb_loose_numeric_literal(left)?;
    let right_number = vb_loose_numeric_literal(right)?;
    Some(Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left_number),
        right: Box::new(right_number),
    }))
}

fn vb_loose_numeric_literal(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(_)) | ExprKind::Lit(Literal::Float(_)) => Some(expr.clone()),
        ExprKind::Lit(Literal::Bool(value)) => Some(Expression::int(if *value { -1 } else { 0 })),
        ExprKind::Lit(Literal::Str(text)) => {
            if let Ok(value) = text.parse::<i64>() {
                Some(Expression::int(value))
            } else if let Ok(value) = text.parse::<f64>() {
                Some(Expression::float(value))
            } else {
                None
            }
        }
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: inner,
        } => match &inner.kind {
            ExprKind::Lit(Literal::Int(value)) => Some(Expression::int(-*value)),
            ExprKind::Lit(Literal::Float(value)) => Some(Expression::float(-*value)),
            _ => None,
        },
        _ => None,
    }
}

fn vb_bool_expr(cond: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(vb_bool_condition_expr(cond)),
        then: Box::new(Expression::bool(true)),
        else_: Box::new(Expression::bool(false)),
    })
}

fn vb_bool_condition_expr(cond: Expression) -> Expression {
    call_expr(
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("System")),
                field: "Convert".into(),
                null_safe: false,
            })),
            field: "ToBoolean".into(),
            null_safe: false,
        }),
        vec![Argument::positional(cond)],
    )
}

fn literal_number(expr: &Expression) -> Option<f64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(n)) => Some(*n as f64),
        ExprKind::Lit(Literal::Float(n)) => Some(*n),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => literal_number(expr).map(|value| -value),
        ExprKind::Unary {
            op: UnaryOp::Pos,
            expr,
        } => literal_number(expr),
        ExprKind::Cast { expr, .. } => literal_number(expr),
        ExprKind::Binary { op, left, right } => {
            let left = literal_number(left)?;
            let right = literal_number(right)?;
            match op {
                BinOp::Add => Some(left + right),
                BinOp::Sub => Some(left - right),
                BinOp::Mul => Some(left * right),
                BinOp::Div => Some(left / right),
                BinOp::IDiv | BinOp::FloorDiv => Some((left / right).trunc()),
                BinOp::Mod => Some(left % right),
                _ => None,
            }
        }
        _ => None,
    }
}

fn literal_i64(expr: &Expression) -> Option<i64> {
    literal_number(expr).map(|n| n.trunc() as i64)
}

fn literal_bool(expr: &Expression) -> Option<bool> {
    match &expr.kind {
        ExprKind::Lit(Literal::Bool(value)) => Some(*value),
        _ => None,
    }
}

fn literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(value)) => Some(value.clone()),
        ExprKind::Cast { expr, type_name } if type_name.eq_ignore_ascii_case("Date") => {
            literal_string(expr)
        }
        _ => None,
    }
}

fn vb_known_string_value(expr: &Expression, locals: &HashMap<String, String>) -> Option<String> {
    literal_string(expr).or_else(|| match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(&format!("$string:{}", name.to_ascii_lowercase()))
            .cloned(),
        _ => None,
    })
}

fn vb_regex_literal_pattern(expr: &Expression, locals: &HashMap<String, String>) -> Option<String> {
    let ExprKind::New { class, args } = &expr.kind else {
        return None;
    };
    let class_name = dotted_expr_name(class)?;
    if !class_name.eq_ignore_ascii_case("Regex")
        && !class_name.eq_ignore_ascii_case("System.Text.RegularExpressions.Regex")
    {
        return None;
    }
    args.first()
        .and_then(|arg| vb_known_string_value(&arg.value, locals))
}

fn vb_regex_group_names(pattern: &str) -> Vec<String> {
    let mut names = vec!["0".to_string()];
    let chars: Vec<char> = pattern.chars().collect();
    let mut i = 0;
    let mut in_class = false;
    while i < chars.len() {
        match chars[i] {
            '\\' => i += 2,
            '[' => {
                in_class = true;
                i += 1;
            }
            ']' if in_class => {
                in_class = false;
                i += 1;
            }
            '(' if !in_class => {
                if chars.get(i + 1) == Some(&'?') {
                    if chars.get(i + 2) == Some(&'<')
                        && !matches!(chars.get(i + 3), Some('=') | Some('!'))
                    {
                        let start = i + 3;
                        if let Some(end_offset) = chars[start..].iter().position(|ch| *ch == '>') {
                            names.push(chars[start..start + end_offset].iter().collect());
                        }
                    } else if chars.get(i + 2) == Some(&'\'') {
                        let start = i + 3;
                        if let Some(end_offset) = chars[start..].iter().position(|ch| *ch == '\'') {
                            names.push(chars[start..start + end_offset].iter().collect());
                        }
                    }
                } else {
                    names.push(names.len().to_string());
                }
                i += 1;
            }
            _ => i += 1,
        }
    }
    names
}

fn vb_regex_group_names_array(pattern: &str) -> Expression {
    Expression::new(ExprKind::Array(
        vb_regex_group_names(pattern)
            .into_iter()
            .map(|name| ArrayElement {
                key: None,
                value: Expression::string(&name),
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn vb_regex_group_name_from_number(pattern: &str, number: i64) -> Expression {
    let names = vb_regex_group_names(pattern);
    names
        .get(usize::try_from(number).unwrap_or(usize::MAX))
        .map(|name| Expression::string(name))
        .unwrap_or_else(|| Expression::string(""))
}

fn vb_regex_group_number_from_name(pattern: &str, name: &str) -> Expression {
    let names = vb_regex_group_names(pattern);
    let number = names
        .iter()
        .position(|candidate| candidate == name)
        .map(|idx| idx as i64)
        .unwrap_or(-1);
    Expression::int(number)
}

fn vb_type_name_from_suffix(suffix: &str, has_decimal: bool) -> Option<&'static str> {
    match suffix.to_ascii_uppercase().as_str() {
        "S" => Some("Int16"),
        "I" | "%" => Some("Int32"),
        "L" | "&" => Some("Int64"),
        "US" => Some("UInt16"),
        "UI" => Some("UInt32"),
        "UL" => Some("UInt64"),
        "F" | "!" => Some("Single"),
        "R" | "#" => Some("Double"),
        "D" | "@" => Some("Decimal"),
        "" if has_decimal => Some("Double"),
        _ => None,
    }
}

fn split_vb_numeric_suffix(raw: &str) -> (&str, &str) {
    let upper = raw.to_ascii_uppercase();
    let suffixes: &[&str] =
        if upper.starts_with("&H") || upper.starts_with("&O") || upper.starts_with("&B") {
            &["US", "UI", "UL", "L", "S", "I", "%", "&"]
        } else {
            &[
                "US", "UI", "UL", "F", "D", "L", "R", "S", "I", "!", "#", "@", "%", "&",
            ]
        };
    for suffix in suffixes {
        if raw.len() > suffix.len() && raw.ends_with(suffix) {
            return (&raw[..raw.len() - suffix.len()], suffix);
        }
        if raw.len() > suffix.len() && raw.to_ascii_uppercase().ends_with(suffix) {
            return (
                &raw[..raw.len() - suffix.len()],
                &raw[raw.len() - suffix.len()..],
            );
        }
    }
    (raw, "")
}

fn parse_vb_numeric_literal(raw: &str) -> Expression {
    let raw = raw.trim();
    let (body, suffix) = split_vb_numeric_suffix(raw);
    let body_clean = body.replace('_', "");
    let upper = body_clean.to_ascii_uppercase();
    let has_decimal = body_clean.contains('.') || upper.contains('E');
    let kind = if let Some(digits) = upper.strip_prefix("&H") {
        let value = u64::from_str_radix(digits, 16).unwrap_or(0);
        ExprKind::Lit(Literal::Int(i64::try_from(value).unwrap_or(i64::MAX)))
    } else if let Some(digits) = upper.strip_prefix("&O") {
        let value = u64::from_str_radix(digits, 8).unwrap_or(0);
        ExprKind::Lit(Literal::Int(i64::try_from(value).unwrap_or(i64::MAX)))
    } else if let Some(digits) = upper.strip_prefix("&B") {
        let value = u64::from_str_radix(digits, 2).unwrap_or(0);
        ExprKind::Lit(Literal::Int(i64::try_from(value).unwrap_or(i64::MAX)))
    } else if has_decimal {
        let mut value: f64 = body_clean.parse().unwrap_or(0.0);
        if suffix.eq_ignore_ascii_case("D") || suffix == "@" {
            if value.fract() == 0.0 {
                if let Some((_, frac)) = body_clean.split_once('.') {
                    if frac.chars().any(|ch| ch != '0') {
                        value += 1e-12;
                    }
                }
            }
        }
        ExprKind::Lit(Literal::Float(value))
    } else {
        let value = body_clean.parse::<i128>().unwrap_or(0);
        ExprKind::Lit(Literal::Int(i64::try_from(value).unwrap_or(i64::MAX)))
    };
    let expr = Expression::new(kind);
    if let Some(type_name) = vb_type_name_from_suffix(suffix, has_decimal) {
        Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        })
    } else {
        expr
    }
}

fn last_day_of_month(year: i32, month: u32) -> Option<u32> {
    let (next_year, next_month) = if month == 12 {
        (year + 1, 1)
    } else {
        (year, month + 1)
    };
    let first_of_next = NaiveDate::from_ymd_opt(next_year, next_month, 1)?;
    Some((first_of_next - Duration::days(1)).day())
}

fn normalize_year_month(year: i64, month: i64) -> Option<(i32, u32)> {
    let total_months = year.checked_mul(12)?.checked_add(month - 1)?;
    let normalized_year = total_months.div_euclid(12);
    let normalized_month = total_months.rem_euclid(12) + 1;
    Some((
        i32::try_from(normalized_year).ok()?,
        u32::try_from(normalized_month).ok()?,
    ))
}

fn build_date_serial(year: i64, month: i64, day: i64) -> Option<NaiveDate> {
    let (normalized_year, normalized_month) = normalize_year_month(year, month)?;
    let first_of_month = NaiveDate::from_ymd_opt(normalized_year, normalized_month, 1)?;
    Some(first_of_month + Duration::days(day - 1))
}

fn build_time_serial(hour: i64, minute: i64, second: i64) -> Option<NaiveTime> {
    let total_seconds = hour
        .checked_mul(3600)?
        .checked_add(minute.checked_mul(60)?)?
        .checked_add(second)?;
    let seconds_in_day = 24 * 3600;
    let normalized = total_seconds.rem_euclid(seconds_in_day);
    let hours = (normalized / 3600) as u32;
    let minutes = ((normalized % 3600) / 60) as u32;
    let seconds = (normalized % 60) as u32;
    NaiveTime::from_hms_opt(hours, minutes, seconds)
}

fn parse_vb_date_text(text: &str) -> Option<VbDateValue> {
    let text = text.trim();

    for pattern in [
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d %b %y %I:%M:%S %p",
        "%d %b %y %I:%M %p",
        "%d %b %Y %I:%M:%S %p",
        "%d %b %Y %I:%M %p",
    ] {
        if let Ok(value) = NaiveDateTime::parse_from_str(text, pattern) {
            return Some(VbDateValue::DateTime(value));
        }
    }

    for pattern in ["%m/%d/%Y", "%Y-%m-%d", "%d %b %y", "%d %b %Y"] {
        if let Ok(value) = NaiveDate::parse_from_str(text, pattern) {
            return Some(VbDateValue::Date(value));
        }
    }

    for pattern in ["%I:%M:%S %p", "%I:%M %p", "%H:%M:%S", "%H:%M"] {
        if let Ok(value) = NaiveTime::parse_from_str(text, pattern) {
            return Some(VbDateValue::Time(value));
        }
    }

    None
}

fn parse_vb_date_expr(expr: &Expression) -> Option<VbDateValue> {
    if let Some(value) = literal_string(expr).and_then(|text| parse_vb_date_text(&text)) {
        return Some(value);
    }
    match &expr.kind {
        ExprKind::Call { callee, args, .. } if args.is_empty() => {
            let name = dotted_expr_name(callee)?;
            let today = Local::now().naive_local().date();
            match name.to_ascii_lowercase().as_str() {
                "now" | "date" | "time" | "datetime.now" | "system.datetime.now" => {
                    Some(VbDateValue::DateTime(Local::now().naive_local()))
                }
                "today" | "datetime.today" | "system.datetime.today" => {
                    Some(VbDateValue::Date(today))
                }
                "timeofday" => Some(VbDateValue::DateTime(
                    NaiveDate::from_ymd_opt(1, 1, 1)?.and_time(Local::now().naive_local().time()),
                )),
                _ => None,
            }
        }
        ExprKind::Member { .. } => {
            let name = dotted_expr_name(expr)?;
            let today = Local::now().naive_local().date();
            match name.to_ascii_lowercase().as_str() {
                "datetime.now" | "system.datetime.now" => {
                    Some(VbDateValue::DateTime(Local::now().naive_local()))
                }
                "datetime.today" | "system.datetime.today" => Some(VbDateValue::Date(today)),
                _ => None,
            }
        }
        ExprKind::New { class, args } => {
            let class_name = dotted_expr_name(class)?;
            if !matches!(
                class_name.to_ascii_lowercase().as_str(),
                "date" | "datetime" | "system.datetime"
            ) {
                return None;
            }
            let year = i32::try_from(literal_i64(&args.get(0)?.value)?).ok()?;
            let month = u32::try_from(literal_i64(&args.get(1)?.value)?).ok()?;
            let day = u32::try_from(literal_i64(&args.get(2)?.value)?).ok()?;
            let date = NaiveDate::from_ymd_opt(year, month, day)?;
            if args.len() >= 6 {
                let hour = u32::try_from(literal_i64(&args.get(3)?.value)?).ok()?;
                let minute = u32::try_from(literal_i64(&args.get(4)?.value)?).ok()?;
                let second = u32::try_from(literal_i64(&args.get(5)?.value)?).ok()?;
                Some(VbDateValue::DateTime(
                    date.and_hms_opt(hour, minute, second)?,
                ))
            } else {
                Some(VbDateValue::Date(date))
            }
        }
        _ => None,
    }
}

fn is_vb_date_literal_expr(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Cast { expr, type_name }
            if type_name.eq_ignore_ascii_case("Date")
                && matches!(expr.kind, ExprKind::Lit(Literal::Str(_)))
    )
}

fn partition_literal_args(arguments: &[Argument]) -> Option<(i64, i64, i64, i64)> {
    if arguments.len() != 4 {
        return None;
    }
    let mut values = [0i64; 4];
    for (index, arg) in arguments.iter().enumerate() {
        let ExprKind::Lit(Literal::Int(value)) = arg.value.kind else {
            return None;
        };
        values[index] = value;
    }
    Some((values[0], values[1], values[2], values[3]))
}

fn fold_partition(arguments: &[Argument]) -> Option<Expression> {
    let (number, start, stop, interval) = partition_literal_args(arguments)?;
    if interval <= 0 {
        return None;
    }

    let (low, high) = if number > stop {
        let overflow = stop.saturating_add(1);
        (overflow, overflow)
    } else {
        let bucket = ((number - start).max(0)) / interval;
        let low = start + (bucket * interval);
        let high = (low + interval - 1).min(stop);
        (low, high)
    };

    let width = (stop.saturating_add(1))
        .abs()
        .to_string()
        .len()
        .max(start.abs().to_string().len());
    Some(Expression::string(&format!("{low:>width$}:{high:>width$}")))
}

fn format_vb_time(time: NaiveTime) -> String {
    let mut hour = time.hour() % 12;
    if hour == 0 {
        hour = 12;
    }
    let suffix = if time.hour() < 12 { "AM" } else { "PM" };
    format!(
        "{}:{:02}:{:02} {}",
        hour,
        time.minute(),
        time.second(),
        suffix
    )
}

fn format_vb_date(value: NaiveDate) -> String {
    format!("{}/{}/{}", value.month(), value.day(), value.year())
}

fn format_vb_date_value(value: &VbDateValue) -> String {
    match value {
        VbDateValue::Date(date) => format_vb_date(*date),
        VbDateValue::Time(time) => format_vb_time(*time),
        VbDateValue::DateTime(datetime) => {
            format!(
                "{} {}",
                format_vb_date(datetime.date()),
                format_vb_time(datetime.time())
            )
        }
    }
}

fn format_vb_datetime_custom(value: &VbDateValue, format: &str) -> Option<String> {
    let datetime = date_value_as_datetime(value)?;
    match format {
        "yyyy-MM-dd" => Some(datetime.format("%Y-%m-%d").to_string()),
        "HH:mm:ss" => Some(datetime.format("%H:%M:%S").to_string()),
        "yyyy-MM-dd HH:mm:ss" => Some(datetime.format("%Y-%m-%d %H:%M:%S").to_string()),
        _ => None,
    }
}

fn date_value_as_datetime(value: &VbDateValue) -> Option<NaiveDateTime> {
    match value {
        VbDateValue::Date(date) => date.and_hms_opt(0, 0, 0),
        VbDateValue::Time(time) => NaiveDate::from_ymd_opt(1970, 1, 1)?.and_hms_opt(
            time.hour(),
            time.minute(),
            time.second(),
        ),
        VbDateValue::DateTime(datetime) => Some(*datetime),
    }
}

fn add_months_to_date(date: NaiveDate, months: i64) -> Option<NaiveDate> {
    let total_months = i64::from(date.year())
        .checked_mul(12)?
        .checked_add(i64::from(date.month()) - 1)?
        .checked_add(months)?;
    let year = total_months.div_euclid(12);
    let month = total_months.rem_euclid(12) + 1;
    let year = i32::try_from(year).ok()?;
    let month = u32::try_from(month).ok()?;
    let day = date.day().min(last_day_of_month(year, month)?);
    NaiveDate::from_ymd_opt(year, month, day)
}

fn add_interval(value: &VbDateValue, interval: &str, amount: i64) -> Option<VbDateValue> {
    match interval {
        "day" | "dayofyear" | "weekday" => match value {
            VbDateValue::Date(date) => Some(VbDateValue::Date(*date + Duration::days(amount))),
            VbDateValue::DateTime(datetime) => {
                Some(VbDateValue::DateTime(*datetime + Duration::days(amount)))
            }
            VbDateValue::Time(time) => Some(VbDateValue::Time(*time)),
        },
        "week" => add_interval(value, "day", amount.checked_mul(7)?),
        "hour" => match value {
            VbDateValue::Date(date) => Some(VbDateValue::DateTime(
                date.and_hms_opt(0, 0, 0)? + Duration::hours(amount),
            )),
            VbDateValue::DateTime(datetime) => {
                Some(VbDateValue::DateTime(*datetime + Duration::hours(amount)))
            }
            VbDateValue::Time(time) => Some(VbDateValue::Time(build_time_serial(
                i64::from(time.hour()) + amount,
                i64::from(time.minute()),
                i64::from(time.second()),
            )?)),
        },
        "month" => match value {
            VbDateValue::Date(date) => Some(VbDateValue::Date(add_months_to_date(*date, amount)?)),
            VbDateValue::DateTime(datetime) => {
                let next_date = add_months_to_date(datetime.date(), amount)?;
                Some(VbDateValue::DateTime(next_date.and_hms_opt(
                    datetime.hour(),
                    datetime.minute(),
                    datetime.second(),
                )?))
            }
            VbDateValue::Time(time) => Some(VbDateValue::Time(*time)),
        },
        "quarter" => add_interval(value, "month", amount.checked_mul(3)?),
        "year" => add_interval(value, "month", amount.checked_mul(12)?),
        "minute" => match value {
            VbDateValue::Date(date) => Some(VbDateValue::DateTime(
                date.and_hms_opt(0, 0, 0)? + Duration::minutes(amount),
            )),
            VbDateValue::DateTime(datetime) => {
                Some(VbDateValue::DateTime(*datetime + Duration::minutes(amount)))
            }
            VbDateValue::Time(time) => Some(VbDateValue::Time(build_time_serial(
                i64::from(time.hour()),
                i64::from(time.minute()) + amount,
                i64::from(time.second()),
            )?)),
        },
        "second" => match value {
            VbDateValue::Date(date) => Some(VbDateValue::DateTime(
                date.and_hms_opt(0, 0, 0)? + Duration::seconds(amount),
            )),
            VbDateValue::DateTime(datetime) => {
                Some(VbDateValue::DateTime(*datetime + Duration::seconds(amount)))
            }
            VbDateValue::Time(time) => Some(VbDateValue::Time(build_time_serial(
                i64::from(time.hour()),
                i64::from(time.minute()),
                i64::from(time.second()) + amount,
            )?)),
        },
        _ => None,
    }
}

fn normalize_vb_date_interval(raw: &str) -> Option<&'static str> {
    match raw.trim().to_ascii_lowercase().as_str() {
        "yyyy" | "year" => Some("year"),
        "q" | "quarter" => Some("quarter"),
        "m" | "month" => Some("month"),
        "y" | "dayofyear" => Some("dayofyear"),
        "d" | "day" => Some("day"),
        "w" | "weekday" => Some("weekday"),
        "ww" | "week" | "weekofyear" => Some("week"),
        "h" | "hour" => Some("hour"),
        "n" | "minute" => Some("minute"),
        "s" | "second" => Some("second"),
        _ => None,
    }
}

fn parse_interval_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Member { object, field, .. } => match &object.kind {
            ExprKind::Ident(name) if name.eq_ignore_ascii_case("DateInterval") => {
                normalize_vb_date_interval(field).map(str::to_string)
            }
            _ => None,
        },
        ExprKind::Lit(Literal::Str(value)) => normalize_vb_date_interval(value).map(str::to_string),
        _ => None,
    }
}

fn date_diff(interval: &str, start: &VbDateValue, end: &VbDateValue) -> Option<i64> {
    let start_dt = date_value_as_datetime(start)?;
    let end_dt = date_value_as_datetime(end)?;
    match interval {
        "day" | "dayofyear" | "weekday" => Some((end_dt - start_dt).num_days()),
        "week" => Some((end_dt - start_dt).num_days() / 7),
        "hour" => Some((end_dt - start_dt).num_hours()),
        "minute" => Some((end_dt - start_dt).num_minutes()),
        "second" => Some((end_dt - start_dt).num_seconds()),
        "month" => Some(
            i64::from(end_dt.year() - start_dt.year()) * 12 + i64::from(end_dt.month())
                - i64::from(start_dt.month()),
        ),
        "quarter" => Some(
            (i64::from(end_dt.year() - start_dt.year()) * 12 + i64::from(end_dt.month())
                - i64::from(start_dt.month()))
                / 3,
        ),
        "year" => Some(i64::from(end_dt.year() - start_dt.year())),
        _ => None,
    }
}

fn date_part(interval: &str, value: &VbDateValue) -> Option<i64> {
    let datetime = date_value_as_datetime(value)?;
    match interval {
        "year" => Some(i64::from(datetime.year())),
        "month" => Some(i64::from(datetime.month())),
        "quarter" => Some(i64::from(((datetime.month() - 1) / 3) + 1)),
        "dayofyear" => Some(i64::from(datetime.ordinal())),
        "day" => Some(i64::from(datetime.day())),
        "weekday" => weekday_index(value),
        "week" => Some(i64::from(((datetime.ordinal() - 1) / 7) + 1)),
        "hour" => Some(i64::from(datetime.hour())),
        "minute" => Some(i64::from(datetime.minute())),
        "second" => Some(i64::from(datetime.second())),
        _ => None,
    }
}

fn weekday_index(value: &VbDateValue) -> Option<i64> {
    Some(i64::from(
        date_value_as_datetime(value)?
            .weekday()
            .number_from_sunday(),
    ))
}

fn weekday_name(value: &VbDateValue) -> Option<&'static str> {
    let names = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let index = usize::try_from(weekday_index(value)?.checked_sub(1)?).ok()?;
    names.get(index).copied()
}

fn fold_date_member_field(value: &VbDateValue, field: &str) -> Option<Expression> {
    let datetime = date_value_as_datetime(value)?;
    match field.to_ascii_lowercase().as_str() {
        "year" => Some(Expression::int(i64::from(datetime.year()))),
        "month" => Some(Expression::int(i64::from(datetime.month()))),
        "day" => Some(Expression::int(i64::from(datetime.day()))),
        "hour" => Some(Expression::int(i64::from(datetime.hour()))),
        "minute" => Some(Expression::int(i64::from(datetime.minute()))),
        "second" => Some(Expression::int(i64::from(datetime.second()))),
        "dayofweek" => Some(Expression::string(weekday_name(value)?)),
        "timeofday" => Some(build_vb_timespan_expr(VbTimeSpanValue {
            days: 0,
            total_seconds: i64::from(datetime.time().num_seconds_from_midnight()),
        })),
        _ => None,
    }
}

fn build_vb_timespan_expr(value: VbTimeSpanValue) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("__type"),
            value: Expression::string("TimeSpan"),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Days"),
            value: Expression::int(value.days),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("TotalSeconds"),
            value: Expression::int(value.total_seconds),
        },
    ]))
}

fn parse_vb_timespan_expr(expr: &Expression) -> Option<VbTimeSpanValue> {
    match &expr.kind {
        ExprKind::Object(props) => {
            let mut days = None;
            let mut total_seconds = None;
            for prop in props {
                if let ObjectProperty::KeyValue { key, value } = prop {
                    let Some(key) = literal_string(key) else {
                        continue;
                    };
                    match key.to_ascii_lowercase().as_str() {
                        "days" => days = literal_i64(value),
                        "totalseconds" => total_seconds = literal_i64(value),
                        _ => {}
                    }
                }
            }
            Some(VbTimeSpanValue {
                days: days?,
                total_seconds: total_seconds?,
            })
        }
        ExprKind::New { class, args } => {
            let class_name = dotted_expr_name(class)?;
            if !matches!(
                class_name.to_ascii_lowercase().as_str(),
                "timespan" | "system.timespan"
            ) {
                return None;
            }
            let (days, hours, minutes, seconds) = match args.len() {
                4 => (
                    literal_i64(&args[0].value)?,
                    literal_i64(&args[1].value)?,
                    literal_i64(&args[2].value)?,
                    literal_i64(&args[3].value)?,
                ),
                3 => (
                    0,
                    literal_i64(&args[0].value)?,
                    literal_i64(&args[1].value)?,
                    literal_i64(&args[2].value)?,
                ),
                _ => return None,
            };
            let total_seconds = days
                .checked_mul(86_400)?
                .checked_add(hours.checked_mul(3_600)?)?
                .checked_add(minutes.checked_mul(60)?)?
                .checked_add(seconds)?;
            Some(VbTimeSpanValue {
                days: total_seconds.div_euclid(86_400),
                total_seconds,
            })
        }
        _ => None,
    }
}

fn fold_timespan_member_field(value: &VbTimeSpanValue, field: &str) -> Option<Expression> {
    match field.to_ascii_lowercase().as_str() {
        "days" => Some(Expression::int(value.days)),
        "totalseconds" => Some(Expression::int(value.total_seconds)),
        _ => None,
    }
}

fn compare_vb_dates(op: BinOp, left: &Expression, right: &Expression) -> Option<Expression> {
    let left = date_value_as_datetime(&parse_vb_date_expr(left)?)?;
    let right = date_value_as_datetime(&parse_vb_date_expr(right)?)?;
    let value = match op {
        BinOp::Eq => left == right,
        BinOp::NotEq => left != right,
        BinOp::Lt => left < right,
        BinOp::Gt => left > right,
        BinOp::LtEq => left <= right,
        BinOp::GtEq => left >= right,
        _ => return None,
    };
    Some(Expression::bool(value))
}

fn compare_vb_runtime_dates(
    op: BinOp,
    left: &Expression,
    right: &Expression,
) -> Option<Expression> {
    if !matches!(
        op,
        BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq
    ) {
        return None;
    }
    let empty = HashMap::new();
    if vb_infer_expr_type(left, &empty).as_deref() != Some("DateTime")
        || vb_infer_expr_type(right, &empty).as_deref() != Some("DateTime")
    {
        return None;
    }
    let cmp = call_expr(
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("DateTime")),
            field: "Compare".into(),
            null_safe: false,
        }),
        vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
        ],
    );
    Some(Expression::new(ExprKind::Binary {
        op,
        left: Box::new(cmp),
        right: Box::new(Expression::int(0)),
    }))
}

fn rewrite_vb_runtime_timespan_binary(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    let left_type = vb_infer_expr_type(left, locals)?;
    let right_type = vb_infer_expr_type(right, locals)?;
    if left_type.as_str() != "TimeSpan" || right_type.as_str() != "TimeSpan" {
        return None;
    }

    match op {
        BinOp::Add | BinOp::Sub => {
            let field = if op == BinOp::Add { "Add" } else { "Subtract" };
            Some(call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("TimeSpan")),
                    field: field.into(),
                    null_safe: false,
                }),
                vec![
                    Argument::positional(left.clone()),
                    Argument::positional(right.clone()),
                ],
            ))
        }
        BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq => {
            let cmp = call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("TimeSpan")),
                    field: "Compare".into(),
                    null_safe: false,
                }),
                vec![
                    Argument::positional(left.clone()),
                    Argument::positional(right.clone()),
                ],
            );
            Some(Expression::new(ExprKind::Binary {
                op,
                left: Box::new(cmp),
                right: Box::new(Expression::int(0)),
            }))
        }
        _ => None,
    }
}

fn fold_vb_date_arithmetic(op: BinOp, left: &Expression, right: &Expression) -> Option<Expression> {
    match op {
        BinOp::Sub => {
            let left_dt = date_value_as_datetime(&parse_vb_date_expr(left)?)?;
            let right_dt = date_value_as_datetime(&parse_vb_date_expr(right)?)?;
            let seconds = left_dt.signed_duration_since(right_dt).num_seconds();
            Some(build_vb_timespan_expr(VbTimeSpanValue {
                days: seconds.div_euclid(86_400),
                total_seconds: seconds,
            }))
        }
        BinOp::Add => {
            let (date, span) = parse_vb_date_expr(left)
                .zip(parse_vb_timespan_expr(right))
                .or_else(|| parse_vb_date_expr(right).zip(parse_vb_timespan_expr(left)))?;
            let datetime = date_value_as_datetime(&date)? + Duration::seconds(span.total_seconds);
            Some(Expression::string(&format_vb_date_value(
                &VbDateValue::DateTime(datetime),
            )))
        }
        _ => None,
    }
}

fn vb_date_method_receiver_name(expr: &Expression) -> Option<String> {
    if let Some(path) = dotted_expr_name(expr) {
        return Some(path);
    }
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if args.is_empty() {
            return dotted_expr_name(callee);
        }
    }
    None
}

fn parse_vb_date_with_format(text: &str, format: &str) -> Option<VbDateValue> {
    match format {
        "yyyyMMdd" => {
            let date = NaiveDate::parse_from_str(text, "%Y%m%d").ok()?;
            Some(VbDateValue::Date(date))
        }
        _ => parse_vb_date_text(text),
    }
}

fn fold_vb_date_try_parse_assignment(expr: &mut Expression) -> Option<(String, Expression)> {
    let ExprKind::Call { callee, args, .. } = &mut expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !field.eq_ignore_ascii_case("TryParseExact") || args.len() < 5 {
        return None;
    }
    let receiver = vb_date_method_receiver_name(object)?;
    if !matches!(
        receiver.to_ascii_lowercase().as_str(),
        "date" | "datetime" | "system.datetime"
    ) {
        return None;
    }
    let text = literal_string(&args[0].value)?;
    let format = literal_string(&args[1].value)?;
    let parsed = parse_vb_date_with_format(&text, &format)?;
    let ExprKind::Ident(target) = &args[4].value.kind else {
        return None;
    };
    let target = target.to_ascii_lowercase();
    let value = Expression::string(&format_vb_date_value(&parsed));
    *expr = Expression::bool(true);
    Some((target, value))
}

fn fold_date_constructor(args: &[Argument], is_time: bool) -> Option<Expression> {
    if is_time {
        let hour = literal_i64(&args.get(0)?.value)?;
        let minute = literal_i64(&args.get(1)?.value)?;
        let second = literal_i64(&args.get(2)?.value)?;
        let time = build_time_serial(hour, minute, second)?;
        return Some(Expression::string(&format_vb_date_value(
            &VbDateValue::Time(time),
        )));
    }

    let year = literal_i64(&args.get(0)?.value)?;
    let month = literal_i64(&args.get(1)?.value)?;
    let day = literal_i64(&args.get(2)?.value)?;
    let date = build_date_serial(year, month, day)?;
    Some(Expression::string(&format_vb_date_value(
        &VbDateValue::Date(date),
    )))
}

fn fold_month_name(args: &[Argument]) -> Option<Expression> {
    let month = literal_i64(&args.get(0)?.value)?;
    let abbreviate = args
        .get(1)
        .and_then(|arg| literal_bool(&arg.value))
        .unwrap_or(false);
    let names = if abbreviate {
        [
            "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
        ]
    } else {
        [
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December",
        ]
    };
    let index = usize::try_from(month.checked_sub(1)?).ok()?;
    Some(Expression::string(names.get(index)?))
}

fn fold_weekday_name(args: &[Argument]) -> Option<Expression> {
    let weekday = literal_i64(&args.get(0)?.value)?;
    let abbreviate = args
        .get(1)
        .and_then(|arg| literal_bool(&arg.value))
        .unwrap_or(false);
    let names = if abbreviate {
        ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    } else {
        [
            "Sunday",
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
        ]
    };
    let index = usize::try_from(weekday.checked_sub(1)?).ok()?;
    Some(Expression::string(names.get(index)?))
}

fn fold_date_value(name: &str, args: &[Argument]) -> Option<Expression> {
    let interval = args.get(0).and_then(|arg| parse_interval_expr(&arg.value));
    match name.to_ascii_lowercase().as_str() {
        "dateserial" => fold_date_constructor(args, false),
        "timeserial" => fold_date_constructor(args, true),
        "datevalue" | "timevalue" | "cdate" => {
            let parsed = parse_vb_date_text(&literal_string(&args.get(0)?.value)?)?;
            Some(Expression::string(&format_vb_date_value(&parsed)))
        }
        "dateadd" => {
            let amount = literal_i64(&args.get(1)?.value)?;
            let input = parse_vb_date_expr(&args.get(2)?.value)?;
            let value = add_interval(&input, interval.as_deref()?, amount)?;
            Some(Expression::string(&format_vb_date_value(&value)))
        }
        "datediff" => {
            let start = parse_vb_date_expr(&args.get(1)?.value)?;
            let end = parse_vb_date_expr(&args.get(2)?.value)?;
            Some(Expression::int(date_diff(
                interval.as_deref()?,
                &start,
                &end,
            )?))
        }
        "datepart" => {
            let value = parse_vb_date_expr(&args.get(1)?.value)?;
            Some(Expression::int(date_part(interval.as_deref()?, &value)?))
        }
        "year" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.year(),
        ))),
        "month" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.month(),
        ))),
        "day" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.day(),
        ))),
        "hour" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.hour(),
        ))),
        "minute" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.minute(),
        ))),
        "second" => Some(Expression::int(i64::from(
            date_value_as_datetime(&parse_vb_date_expr(&args.get(0)?.value)?)?.second(),
        ))),
        "weekday" => Some(Expression::int(weekday_index(&parse_vb_date_expr(
            &args.get(0)?.value,
        )?)?)),
        "monthname" => fold_month_name(args),
        "weekdayname" => fold_weekday_name(args),
        "isdate" => {
            let is_valid = parse_vb_date_text(&literal_string(&args.get(0)?.value)?).is_some();
            Some(Expression::bool(is_valid))
        }
        _ => None,
    }
}

fn fold_vb_isnumeric(args: &[Argument]) -> Option<Expression> {
    let value = args.get(0)?;
    let is_numeric = match &value.value.kind {
        ExprKind::Lit(Literal::Int(_)) | ExprKind::Lit(Literal::Float(_)) => true,
        ExprKind::Lit(Literal::Bool(_)) => false,
        ExprKind::Lit(Literal::Str(text)) => text.trim().parse::<f64>().is_ok(),
        ExprKind::Cast { expr, .. } => literal_number(expr).is_some(),
        _ => return None,
    };
    Some(Expression::bool(is_numeric))
}

fn canonicalize_special_identifier(name: &str) -> Option<Expression> {
    match name.to_ascii_lowercase().as_str() {
        "vbcr" => Some(Expression::string("\r")),
        "vblf" => Some(Expression::string("\n")),
        "vbcrlf" | "vbnewline" => Some(Expression::string("\r\n")),
        "vbtab" => Some(Expression::string("\t")),
        "vbback" => Some(Expression::string("\u{0008}")),
        "vbformfeed" => Some(Expression::string("\u{000c}")),
        "vbverticaltab" => Some(Expression::string("\u{000b}")),
        "vbnullchar" => Some(Expression::string("\0")),
        "vbnullstring" => Some(Expression::string("")),
        "now" => Some(zero_arg_call("now")),
        "today" => Some(zero_arg_call("today")),
        "timeofday" => Some(zero_arg_call("timeofday")),
        "timer" => Some(zero_arg_call("timer")),
        _ => None,
    }
}

fn canonicalize_call(name: &str, arguments: &[Argument]) -> Option<Expression> {
    match name.to_ascii_lowercase().as_str() {
        "await" if arguments.len() == 1 => Some(Expression::new(ExprKind::Await(Box::new(
            arguments[0].value.clone(),
        )))),
        "fix" if arguments.len() == 1 => Some(call_expr(
            build_dotted_expr("System.Math.Truncate"),
            arguments.to_vec(),
        )),
        "int" if arguments.len() == 1 => Some(call_expr(
            build_dotted_expr("System.Math.Floor"),
            arguments.to_vec(),
        )),
        // VB `Rnd()` → Math.random() (ecma:math.random via System.Math routing)
        "rnd" if arguments.is_empty() => {
            Some(call_expr(build_dotted_expr("System.Math.Random"), vec![]))
        }
        // VB `Rnd(seed)` — seed arg is ignored per VB spec (Rnd always 0..1)
        "rnd" => Some(call_expr(build_dotted_expr("System.Math.Random"), vec![])),
        "isnothing" if arguments.len() == 1 => {
            Some(vb_bool_expr(Expression::new(ExprKind::Binary {
                op: BinOp::Is,
                left: Box::new(arguments[0].value.clone()),
                right: Box::new(Expression::null()),
            })))
        }
        "isnumeric" if arguments.len() == 1 => fold_vb_isnumeric(arguments),
        // VB `Sgn(x)` → Math.sign(x)
        "sgn" if arguments.len() == 1 => Some(call_expr(
            build_dotted_expr("System.Math.Sign"),
            arguments.to_vec(),
        )),
        // VB `Array(a, b, c)` is an array literal of its arguments — normalize
        // to a common `ExprKind::Array` here so the shared compiler needs no
        // VB-specific `Array` intercept. `Array.Empty()` / `Array.IndexOf(...)`
        // are MEMBER calls and never reach this bare-identifier path.
        "array" => Some(Expression::new(ExprKind::Array(
            arguments
                .iter()
                .map(|arg| ArrayElement {
                    key: None,
                    value: arg.value.clone(),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        ))),
        // VB `Randomize` / `Randomize(seed)` — no-op; VB `Rnd()` routes to
        // wasi:random/random.random so seeding is the engine's concern.
        "randomize" => Some(Expression::new(ExprKind::Lit(Literal::Null))),
        "round" if arguments.len() == 1 => {
            Some(build_vb_bankers_round_expr(arguments[0].value.clone()))
        }
        "round" if arguments.len() >= 2 => Some(build_vb_precision_round_expr(
            arguments[0].value.clone(),
            arguments[1].value.clone(),
        )),
        "vartype"
            if arguments.len() == 1
                && matches!(arguments[0].value.kind, ExprKind::Lit(Literal::Null)) =>
        {
            Some(Expression::int(0))
        }
        "vartype" if arguments.len() == 1 && is_vb_date_literal_expr(&arguments[0].value) => {
            Some(Expression::int(7))
        }
        "partition" if arguments.len() == 4 => fold_partition(arguments),
        "cbool" if arguments.len() == 1 => {
            if let Some(value) = literal_bool(&arguments[0].value) {
                Some(Expression::bool(value))
            } else if let Some(value) = literal_i64(&arguments[0].value) {
                Some(Expression::bool(value != 0))
            } else {
                None
            }
        }
        "clng" | "cuint" | "culng" | "cushort" | "cshort" if arguments.len() == 1 => {
            literal_i64(&arguments[0].value).map(Expression::int)
        }
        "cobj" if arguments.len() == 1 => Some(arguments[0].value.clone()),
        "dateserial" | "timeserial" | "dateadd" | "datediff" | "datepart" | "datevalue"
        | "timevalue" | "cdate" | "year" | "month" | "day" | "hour" | "minute" | "second"
        | "weekday" | "monthname" | "weekdayname" | "isdate" => fold_date_value(name, arguments),
        _ => None,
    }
}

fn canonicalize_member_access(object: Expression, name: &str) -> Expression {
    let object = canonicalize_date_receiver(object);
    let is_class_static = matches!(
        &object.kind,
        ExprKind::Ident(n) if n.chars().next().is_some_and(|c| c.is_ascii_uppercase())
    );

    if let Some(path) = dotted_expr_name(&object) {
        if matches!(
            path.to_ascii_lowercase().as_str(),
            "double" | "system.double" | "single" | "system.single"
        ) {
            match name.to_ascii_lowercase().as_str() {
                "nan" => return Expression::float(f64::NAN),
                "positiveinfinity" => return Expression::float(f64::INFINITY),
                "negativeinfinity" => return Expression::float(f64::NEG_INFINITY),
                "epsilon" => return Expression::float(f64::MIN_POSITIVE),
                _ => {}
            }
        }
        if (path.eq_ignore_ascii_case("Thread")
            || path.eq_ignore_ascii_case("System.Threading.Thread"))
            && name.eq_ignore_ascii_case("CurrentThread")
        {
            return Expression::new(ExprKind::Object(vec![]));
        }
        if let Some(value) = vybe_platform_dotnet::emitter::static_member_constant(&path, name) {
            return Expression::string(value);
        }
        if vybe_platform_dotnet::emitter::static_member_parameterless_call(&path, name) {
            return call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(object),
                    field: name.to_string(),
                    null_safe: false,
                }),
                vec![],
            );
        }
    }

    let date_field_fn = match name.to_ascii_lowercase().as_str() {
        "year" => Some("year"),
        "month" => Some("month"),
        "day" => Some("day"),
        "date" => Some("date"),
        "hour" => Some("hour"),
        "minute" => Some("minute"),
        "second" => Some("second"),
        _ => None,
    };
    if let Some(function) = date_field_fn {
        if let Some(value) = parse_vb_date_expr(&object) {
            if let Some(folded) = fold_date_value(
                function,
                &[Argument::positional(Expression::string(
                    &format_vb_date_value(&value),
                ))],
            ) {
                return folded;
            }
        }
        if let Some(folded) = fold_special_date_field(&object, function) {
            return folded;
        }
        if is_dotnet_datetime_expr(&object) {
            return vb_object_get_expr(object, dotnet_vb::datetime_field_name(function));
        }
        if is_vb_date_producing_expr(&object) {
            let object = canonicalize_special_date_identifier(object);
            return vb_object_get_expr(object, dotnet_vb::datetime_field_name(function));
        }
    }

    if matches!(name, "Keys" | "Values") && !is_class_static {
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

fn canonicalize_date_receiver(object: Expression) -> Expression {
    if let ExprKind::Ident(name) = &object.kind {
        if matches!(
            name.to_ascii_lowercase().as_str(),
            "now" | "date" | "today" | "time" | "timeofday"
        ) {
            return zero_arg_call(&name.to_ascii_lowercase());
        }
    }

    let ExprKind::Call {
        callee,
        args,
        optional,
    } = &object.kind
    else {
        return object;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return object;
    };
    canonicalize_call(name, args).unwrap_or_else(|| {
        Expression::new(ExprKind::Call {
            callee: callee.clone(),
            args: args.clone(),
            optional: *optional,
        })
    })
}

fn fold_special_date_field(object: &Expression, field: &str) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if !args.is_empty() {
        return None;
    }
    let name = dotted_expr_name(callee)?.to_ascii_lowercase();
    match name.as_str() {
        "now" => {
            let now = Local::now();
            match field {
                "year" => Some(Expression::int(i64::from(now.year()))),
                "month" => Some(Expression::int(i64::from(now.month()))),
                "day" => Some(Expression::int(i64::from(now.day()))),
                "hour" => Some(Expression::int(i64::from(now.hour()))),
                "minute" => Some(Expression::int(i64::from(now.minute()))),
                "second" => Some(Expression::int(i64::from(now.second()))),
                _ => None,
            }
        }
        "today" | "date" => {
            let now = Local::now();
            match field {
                "year" => Some(Expression::int(i64::from(now.year()))),
                "month" => Some(Expression::int(i64::from(now.month()))),
                "day" => Some(Expression::int(i64::from(now.day()))),
                "hour" | "minute" | "second" => Some(Expression::int(0)),
                _ => None,
            }
        }
        "timeofday" | "time" => {
            let now = Local::now();
            match field {
                "year" => Some(Expression::int(1)),
                "month" | "day" => Some(Expression::int(1)),
                "hour" => Some(Expression::int(i64::from(now.hour()))),
                "minute" => Some(Expression::int(i64::from(now.minute()))),
                "second" => Some(Expression::int(i64::from(now.second()))),
                _ => None,
            }
        }
        _ => None,
    }
}

fn vb_object_get_expr(object: Expression, field: &str) -> Expression {
    call_expr(
        Expression::ident("__vb_object_get"),
        vec![
            Argument::positional(object),
            Argument::positional(Expression::string(field)),
        ],
    )
}

fn is_dotnet_datetime_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call {
            callee, args: _, ..
        } => {
            let Some(name) = dotted_expr_name(callee) else {
                return false;
            };
            dotnet_vb::is_datetime_static_producer(&name)
        }
        _ => false,
    }
}

fn canonicalize_special_date_identifier(expr: Expression) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name)
            if matches!(
                name.to_ascii_lowercase().as_str(),
                "now" | "date" | "today" | "time" | "timeofday"
            ) =>
        {
            zero_arg_call(&name.to_ascii_lowercase())
        }
        _ => expr,
    }
}

fn is_vb_date_producing_expr(expr: &Expression) -> bool {
    let (name, args_len) = match &expr.kind {
        ExprKind::Ident(name) => (name.clone(), Some(0)),
        ExprKind::Call { callee, args, .. } => {
            let Some(name) = dotted_expr_name(callee) else {
                return false;
            };
            (name, Some(args.len()))
        }
        _ => return false,
    };
    if matches!(
        name.to_ascii_lowercase().as_str(),
        "now" | "date" | "today" | "time" | "timeofday"
    ) {
        return args_len == Some(0);
    }
    matches!(
        name.to_ascii_lowercase().as_str(),
        "dateserial" | "timeserial" | "dateadd" | "datevalue" | "timevalue" | "cdate"
    )
}

fn emit_vb_object_init_iife(new_call: Expression, props: Vec<(String, Expression)>) -> Expression {
    let type_hint = match &new_call.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        },
        _ => None,
    };
    let mut body: Vec<Statement> = Vec::new();
    body.push(Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint,
                init: Some(new_call),
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

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![],
        optional: false,
    })
}

fn emit_vb_collection_init_iife(new_call: Expression, elements: Vec<Expression>) -> Expression {
    let type_hint = match &new_call.kind {
        ExprKind::New { class, .. } => match &class.kind {
            ExprKind::Ident(name) => Some(name.clone()),
            _ => None,
        },
        _ => None,
    };
    let mut body: Vec<Statement> = vec![Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__obj".into()),
                type_hint,
                init: Some(new_call),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    )];

    for element in elements {
        let args = match element.kind {
            ExprKind::Array(items) if items.len() >= 2 => items
                .into_iter()
                .map(|item| Argument::positional(item.value))
                .collect(),
            _ => vec![Argument::positional(element)],
        };
        body.push(Statement::with_span(
            StmtKind::Expr(call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("__obj")),
                    field: "Add".to_string(),
                    null_safe: false,
                }),
                args,
            )),
            Span::default(),
        ));
    }

    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__obj"))),
        Span::default(),
    ));

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![],
        optional: false,
    })
}

fn parse_vb_member_initializer(pair: Pair<Rule>) -> Result<(String, Expression, bool), String> {
    let mut name = None;
    let mut value = None;
    let mut is_key = false;
    for item in pair.into_inner() {
        match item.as_rule() {
            Rule::key_initializer_modifier => is_key = true,
            Rule::member_identifier => name = Some(item.as_str().to_ascii_lowercase()),
            Rule::expression => value = Some(parse_expression(item)?),
            _ => {}
        }
    }
    Ok((
        name.ok_or_else(|| "Missing object initializer member name".to_string())?,
        value.ok_or_else(|| "Missing object initializer member value".to_string())?,
        is_key,
    ))
}

fn emit_vb_anonymous_object_expr(props: Vec<(String, Expression, bool)>) -> Expression {
    let mut object_props = Vec::new();
    let mut key_names = Vec::new();
    for (name, value, is_key) in props {
        if is_key {
            key_names.push(name.clone());
        }
        object_props.push(ObjectProperty::KeyValue {
            key: Expression::string(&name),
            value,
        });
    }
    object_props.push(ObjectProperty::KeyValue {
        key: Expression::string("__vb_anonymous_keys"),
        value: Expression::new(ExprKind::Array(
            key_names
                .into_iter()
                .map(|name| ArrayElement {
                    key: None,
                    value: Expression::string(&name),
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )),
    });
    Expression::new(ExprKind::Object(object_props))
}

fn parse_anonymous_new_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut props = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() != Rule::with_initializer {
            continue;
        }
        for mi in p.into_inner() {
            if mi.as_rule() != Rule::member_initializer {
                continue;
            }
            props.push(parse_vb_member_initializer(mi)?);
        }
    }
    Ok(Expression::with_span(
        emit_vb_anonymous_object_expr(props).kind,
        span,
    ))
}

fn normalize_vb_anonymous_equals(module: &mut Module) {
    normalize_vb_anonymous_equals_statements(&mut module.body, &mut HashMap::new());
}

fn normalize_vb_anonymous_equals_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, Vec<String>>,
) {
    for stmt in body {
        normalize_vb_anonymous_equals_statement(stmt, locals);
    }
}

fn normalize_vb_anonymous_equals_statement(
    stmt: &mut Statement,
    locals: &mut HashMap<String, Vec<String>>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_anonymous_equals_expr(expr, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_anonymous_equals_expr(init, locals);
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(keys) = vb_anonymous_key_names(init) {
                            locals.insert(name.to_ascii_lowercase(), keys);
                        }
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in &mut *targets {
                normalize_vb_anonymous_equals_expr(target, locals);
            }
            normalize_vb_anonymous_equals_expr(value, locals);
            if let Some(Expression {
                kind: ExprKind::Ident(name),
                ..
            }) = targets.first()
            {
                if let Some(keys) = vb_anonymous_key_names(value) {
                    locals.insert(name.to_ascii_lowercase(), keys);
                } else {
                    locals.remove(&name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_anonymous_equals_expr(target, locals);
            normalize_vb_anonymous_equals_expr(value, locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_anonymous_equals_expr(cond, locals);
            normalize_vb_anonymous_equals_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_anonymous_equals_expr(elif_cond, locals);
                normalize_vb_anonymous_equals_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_anonymous_equals_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_anonymous_equals_statement(init, &mut loop_locals);
            }
            if let Some(cond) = cond {
                normalize_vb_anonymous_equals_expr(cond, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_anonymous_equals_expr(update, &loop_locals);
            }
            normalize_vb_anonymous_equals_statements(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_anonymous_equals_expr(iter, locals);
            normalize_vb_anonymous_equals_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_anonymous_equals_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_anonymous_equals_expr(cond, locals);
            normalize_vb_anonymous_equals_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_anonymous_equals_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_anonymous_equals_statements(body, &mut locals.clone());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_anonymous_equals_member(member);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_anonymous_equals_statements(body, &mut locals.clone());
        }
        _ => {}
    }
}

fn normalize_vb_anonymous_equals_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_anonymous_equals_statement(stmt, &mut HashMap::new());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_anonymous_equals_statements(body, &mut HashMap::new());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_anonymous_equals_statements(getter, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                normalize_vb_anonymous_equals_statements(&mut setter.body, &mut HashMap::new());
            }
        }
        _ => {}
    }
}

fn normalize_vb_anonymous_equals_expr(
    expr: &mut Expression,
    locals: &HashMap<String, Vec<String>>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_anonymous_equals_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_anonymous_equals_expr(&mut arg.value, locals);
            }
            if args.len() == 2
                && dotted_expr_name(callee)
                    .is_some_and(|name| name.eq_ignore_ascii_case("__vb_object_equals"))
            {
                if let ExprKind::Ident(name) = &args[0].value.kind {
                    if let Some(keys) = locals.get(&name.to_ascii_lowercase()) {
                        *expr = build_vb_anonymous_key_equals(&args[0].value, &args[1].value, keys);
                        return;
                    }
                }
            }
            if args.len() == 1 {
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if field.eq_ignore_ascii_case("Equals") {
                        if let ExprKind::Ident(name) = &object.kind {
                            if let Some(keys) = locals.get(&name.to_ascii_lowercase()) {
                                *expr = build_vb_anonymous_key_equals(
                                    &Expression::ident(name),
                                    &args[0].value,
                                    keys,
                                );
                            }
                        }
                    }
                }
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_anonymous_equals_expr(left, locals);
            normalize_vb_anonymous_equals_expr(right, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_anonymous_equals_expr(expr, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_anonymous_equals_expr(cond, locals);
            normalize_vb_anonymous_equals_expr(then, locals);
            normalize_vb_anonymous_equals_expr(else_, locals);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_anonymous_equals_expr(object, locals);
            normalize_vb_anonymous_equals_expr(index, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_anonymous_equals_expr(&mut item.value, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_anonymous_equals_expr(key, locals);
                        normalize_vb_anonymous_equals_expr(value, locals);
                    }
                    ObjectProperty::Spread(value) => {
                        normalize_vb_anonymous_equals_expr(value, locals)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_anonymous_equals_statement(value, &mut locals.clone());
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                normalize_vb_anonymous_equals_expr(&mut arg.value, locals);
            }
        }
        _ => {}
    }
}

fn vb_anonymous_key_names(expr: &Expression) -> Option<Vec<String>> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        if literal_string(key).is_some_and(|name| name == "__vb_anonymous_keys") {
            let ExprKind::Array(items) = &value.kind else {
                return Some(Vec::new());
            };
            return Some(
                items
                    .iter()
                    .filter_map(|item| literal_string(&item.value))
                    .collect(),
            );
        }
    }
    None
}

fn build_vb_anonymous_key_equals(
    left: &Expression,
    right: &Expression,
    keys: &[String],
) -> Expression {
    keys.iter()
        .map(|key| {
            Expression::new(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(left.clone()),
                    field: key.clone(),
                    null_safe: false,
                })),
                right: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(right.clone()),
                    field: key.clone(),
                    null_safe: false,
                })),
            })
        })
        .reduce(|left, right| {
            Expression::new(ExprKind::Binary {
                op: BinOp::And,
                left: Box::new(left),
                right: Box::new(right),
            })
        })
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true))))
}

fn normalize_vb_custom_collection_for_each(module: &mut Module) {
    let mut classes = HashMap::new();
    collect_vb_custom_collection_classes(&module.body, &mut classes);
    if classes.is_empty() {
        return;
    }
    normalize_vb_custom_collection_for_each_statements(
        &mut module.body,
        &classes,
        &mut HashMap::new(),
    );
}

fn collect_vb_custom_collection_classes(body: &[Statement], classes: &mut HashMap<String, String>) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                let has_get_enumerator = members.iter().any(|member| {
                    matches!(
                        member,
                        ClassMember::Method(method)
                            if matches!(
                                &method.kind,
                                StmtKind::FunctionDecl { name, .. }
                                    if name.eq_ignore_ascii_case("GetEnumerator")
                                        || name.eq_ignore_ascii_case("iterator")
                            )
                    )
                });
                if has_get_enumerator {
                    if let Some(field) = members.iter().find_map(vb_collection_backing_field_name) {
                        classes.insert(name.to_ascii_lowercase(), field);
                    }
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_custom_collection_classes(std::slice::from_ref(nested), classes);
                    }
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_custom_collection_classes(std::slice::from_ref(nested), classes);
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_vb_custom_collection_classes(body, classes)
            }
            _ => {}
        }
    }
}

fn vb_collection_backing_field_name(member: &ClassMember) -> Option<String> {
    let ClassMember::Field {
        name,
        type_hint,
        init,
        ..
    } = member
    else {
        return None;
    };
    let type_name = type_hint
        .clone()
        .or_else(|| init.as_ref().and_then(vb_new_expr_type_name))?;
    let base = dotnet_vb::collection_base_type_name(&type_name).to_ascii_lowercase();
    matches!(
        base.as_str(),
        "list" | "collection" | "linkedlist" | "ienumerable" | "arraylist"
    )
    .then(|| name.clone())
}

fn normalize_vb_custom_collection_for_each_statements(
    body: &mut [Statement],
    classes: &HashMap<String, String>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_custom_collection_for_each_statement(stmt, classes, locals);
    }
}

fn normalize_vb_custom_collection_for_each_statement(
    stmt: &mut Statement,
    classes: &HashMap<String, String>,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let class_name = decl
                        .type_hint
                        .as_deref()
                        .map(dotnet_vb::collection_base_type_name)
                        .or_else(|| decl.init.as_ref().and_then(vb_new_expr_type_name));
                    if let Some(class_name) = class_name {
                        if let Some(field) = classes.get(&class_name.to_ascii_lowercase()) {
                            locals.insert(name.to_ascii_lowercase(), field.clone());
                        }
                    }
                }
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            if let ExprKind::Ident(name) = &iter.kind {
                if let Some(field) = locals.get(&name.to_ascii_lowercase()) {
                    *iter = Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(name)),
                        field: field.clone(),
                        null_safe: false,
                    });
                }
            }
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_custom_collection_for_each_statements(
                    else_body,
                    classes,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_custom_collection_for_each_statements(
                then_body,
                classes,
                &mut locals.clone(),
            );
            for (_, elif_body) in elifs {
                normalize_vb_custom_collection_for_each_statements(
                    elif_body,
                    classes,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_custom_collection_for_each_statements(
                    else_body,
                    classes,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::For { init, body, .. } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_custom_collection_for_each_statement(init, classes, &mut loop_locals);
            }
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut loop_locals);
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_custom_collection_for_each_statements(
                    else_body,
                    classes,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::DoWhile { body, .. } => {
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut locals.clone());
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut locals.clone());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                        normalize_vb_custom_collection_for_each_statement(
                            stmt,
                            classes,
                            &mut HashMap::new(),
                        );
                    }
                    ClassMember::Constructor { body, .. } => {
                        normalize_vb_custom_collection_for_each_statements(
                            body,
                            classes,
                            &mut HashMap::new(),
                        );
                    }
                    ClassMember::Property { getter, setter, .. } => {
                        if let Some(getter) = getter {
                            normalize_vb_custom_collection_for_each_statements(
                                getter,
                                classes,
                                &mut HashMap::new(),
                            );
                        }
                        if let Some(setter) = setter {
                            normalize_vb_custom_collection_for_each_statements(
                                &mut setter.body,
                                classes,
                                &mut HashMap::new(),
                            );
                        }
                    }
                    _ => {}
                }
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_custom_collection_for_each_statements(body, classes, &mut locals.clone());
        }
        _ => {}
    }
}

fn try_parse_declaration(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    let stmt = match pair.as_rule() {
        Rule::sub_decl => Some(parse_sub_decl(pair)?),
        Rule::function_decl => Some(parse_function_decl(pair)?),
        Rule::module_decl => Some(parse_module_decl(pair)?),
        Rule::namespace_decl => Some(parse_namespace_decl(pair)?),
        Rule::class_decl => Some(parse_class_decl(pair)?),
        Rule::interface_decl => Some(parse_interface_decl(pair)?),
        Rule::structure_decl => Some(parse_structure_decl(pair)?),
        Rule::enum_decl => Some(parse_enum_decl(pair)?),
        _ => None,
    };
    Ok(stmt)
}

fn rewrite_vb_import_aliases(module: &mut Module) {
    let mut aliases: HashMap<String, String> = HashMap::new();
    // VB has `System` imported by default; keep unqualified BCL roots on the
    // shared Dotnet component path instead of falling through to host mapping.
    aliases.insert("Environment".into(), "System.Environment".into());
    for import in &module.imports {
        if let ImportKind::Simple {
            path,
            alias: Some(alias),
        } = &import.kind
        {
            aliases.insert(alias.clone(), path.clone());
        }
    }
    if aliases.is_empty() {
        return;
    }
    rewrite_vb_aliases_in_statements(&mut module.body, &aliases);
}

fn rewrite_vb_aliases_in_statements(body: &mut [Statement], aliases: &HashMap<String, String>) {
    for stmt in body {
        rewrite_vb_aliases_in_statement(stmt, aliases);
    }
}

fn rewrite_vb_aliases_in_statement(stmt: &mut Statement, aliases: &HashMap<String, String>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::CompoundAssign { value: expr, .. } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
        }
        StmtKind::Using { resource, body, .. } => {
            rewrite_vb_aliases_in_expr(resource, aliases);
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        StmtKind::Lock { expr, body } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
            rewrite_vb_aliases_in_expr(cause, aliases);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                rewrite_vb_alias_in_type_hint(&mut decl.type_hint, aliases);
                if let Some(init) = &mut decl.init {
                    rewrite_vb_aliases_in_expr(init, aliases);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        rewrite_vb_aliases_in_expr(bound, aliases);
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_aliases_in_expr(target, aliases);
            }
            rewrite_vb_aliases_in_expr(value, aliases);
        }
        StmtKind::FunctionDecl {
            params,
            return_type,
            body,
            ..
        } => {
            for param in params {
                rewrite_vb_alias_in_param(param, aliases);
            }
            rewrite_vb_alias_in_type_hint(return_type, aliases);
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_aliases_in_member(member, aliases);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_aliases_in_expr(cond, aliases);
            rewrite_vb_aliases_in_statements(then_body, aliases);
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_aliases_in_expr(elif_cond, aliases);
                rewrite_vb_aliases_in_statements(elif_body, aliases);
            }
            if let Some(else_body) = else_body {
                rewrite_vb_aliases_in_statements(else_body, aliases);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_vb_aliases_in_statement(init, aliases);
            }
            if let Some(cond) = cond {
                rewrite_vb_aliases_in_expr(cond, aliases);
            }
            if let Some(update) = update {
                rewrite_vb_aliases_in_expr(update, aliases);
            }
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_aliases_in_expr(iter, aliases);
            rewrite_vb_aliases_in_statements(body, aliases);
            if let Some(else_body) = else_body {
                rewrite_vb_aliases_in_statements(else_body, aliases);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_aliases_in_expr(cond, aliases);
            rewrite_vb_aliases_in_statements(body, aliases);
            if let Some(else_body) = else_body {
                rewrite_vb_aliases_in_statements(else_body, aliases);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            rewrite_vb_aliases_in_statements(body, aliases);
            rewrite_vb_aliases_in_expr(cond, aliases);
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            rewrite_vb_aliases_in_statements(body, aliases);
            for catch in catches {
                for ty in &mut catch.types {
                    if let Some(rewritten) = rewrite_vb_alias_name(ty, aliases) {
                        *ty = rewritten;
                    }
                }
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_vb_aliases_in_expr(when_clause, aliases);
                }
                rewrite_vb_aliases_in_statements(&mut catch.body, aliases);
            }
            if let Some(finally) = finally {
                rewrite_vb_aliases_in_statements(finally, aliases);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            rewrite_vb_aliases_in_expr(expr, aliases)
                        }
                        CaseCondition::Range { from, to } => {
                            rewrite_vb_aliases_in_expr(from, aliases);
                            rewrite_vb_aliases_in_expr(to, aliases);
                        }
                    }
                }
                rewrite_vb_aliases_in_statements(&mut case.body, aliases);
            }
            if let Some(default) = default {
                rewrite_vb_aliases_in_statements(default, aliases);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_aliases_in_member(member: &mut ClassMember, aliases: &HashMap<String, String>) {
    match member {
        ClassMember::Field {
            type_hint,
            init,
            array_bounds,
            ..
        } => {
            rewrite_vb_alias_in_type_hint(type_hint, aliases);
            if let Some(init) = init {
                rewrite_vb_aliases_in_expr(init, aliases);
            }
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    rewrite_vb_aliases_in_expr(bound, aliases);
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_aliases_in_statement(stmt, aliases);
        }
        ClassMember::Constructor {
            params,
            body,
            base_args,
            ..
        } => {
            for param in params {
                rewrite_vb_alias_in_param(param, aliases);
            }
            if let Some(base_args) = base_args {
                for arg in base_args {
                    rewrite_vb_aliases_in_expr(arg, aliases);
                }
            }
            rewrite_vb_aliases_in_statements(body, aliases);
        }
        ClassMember::Property {
            type_hint,
            getter,
            setter,
            ..
        } => {
            rewrite_vb_alias_in_type_hint(type_hint, aliases);
            if let Some(getter) = getter {
                rewrite_vb_aliases_in_statements(getter, aliases);
            }
            if let Some(setter) = setter {
                rewrite_vb_alias_in_param(&mut setter.param, aliases);
                rewrite_vb_aliases_in_statements(&mut setter.body, aliases);
            }
        }
        ClassMember::Event {
            type_hint, params, ..
        } => {
            rewrite_vb_alias_in_type_hint(type_hint, aliases);
            for param in params {
                rewrite_vb_alias_in_param(param, aliases);
            }
        }
        ClassMember::Const {
            type_hint, value, ..
        } => {
            rewrite_vb_alias_in_type_hint(type_hint, aliases);
            rewrite_vb_aliases_in_expr(value, aliases);
        }
        // VB declares no augmentations; nothing to rewrite.
        ClassMember::Augment(_) => {}
    }
}

fn rewrite_vb_alias_in_param(param: &mut Param, aliases: &HashMap<String, String>) {
    rewrite_vb_alias_in_type_hint(&mut param.type_hint, aliases);
    if let Some(default) = &mut param.default {
        rewrite_vb_aliases_in_expr(default, aliases);
    }
}

fn rewrite_vb_alias_in_type_hint(
    type_hint: &mut Option<String>,
    aliases: &HashMap<String, String>,
) {
    if let Some(current) = type_hint.as_ref() {
        if let Some(rewritten) = rewrite_vb_alias_type_name(current, aliases) {
            *type_hint = Some(rewritten);
        }
    }
}

fn rewrite_vb_alias_type_name(name: &str, aliases: &HashMap<String, String>) -> Option<String> {
    if let Some(path) = aliases.get(name) {
        return Some(vb_alias_target_type_name(path));
    }
    let (head, tail) = name.split_once('.')?;
    aliases
        .get(head)
        .map(|path| format!("{}.{}", vb_alias_target_type_name(path), tail))
}

fn vb_alias_target_type_name(path: &str) -> String {
    path.rsplit('.').next().unwrap_or(path).to_string()
}

fn rewrite_vb_alias_name(name: &str, aliases: &HashMap<String, String>) -> Option<String> {
    if let Some(path) = aliases.get(name) {
        return Some(path.clone());
    }
    let (head, tail) = name.split_once('.')?;
    aliases.get(head).map(|path| format!("{}.{}", path, tail))
}

fn rewrite_vb_aliases_in_expr(expr: &mut Expression, aliases: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(path) = aliases.get(name) {
                *expr = build_dotted_expr(path);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_aliases_in_expr(left, aliases);
            rewrite_vb_aliases_in_expr(right, aliases);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_vb_aliases_in_expr(expr, aliases);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } => rewrite_vb_aliases_in_expr(object, aliases),
            PlaceExpr::Index { object, index, .. } => {
                rewrite_vb_aliases_in_expr(object, aliases);
                rewrite_vb_aliases_in_expr(index, aliases);
            }
            PlaceExpr::Deref(expr) => rewrite_vb_aliases_in_expr(expr, aliases),
        },
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_aliases_in_expr(cond, aliases);
            rewrite_vb_aliases_in_expr(then, aliases);
            rewrite_vb_aliases_in_expr(else_, aliases);
        }
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } => {
            rewrite_vb_aliases_in_expr(object, aliases);
            let rewritten = canonicalize_member_access((**object).clone(), field);
            *expr = rewritten;
        }
        ExprKind::Member { object, .. } => rewrite_vb_aliases_in_expr(object, aliases),
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_aliases_in_expr(object, aliases);
            rewrite_vb_aliases_in_expr(index, aliases);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_aliases_in_expr(callee, aliases);
            for arg in args {
                rewrite_vb_aliases_in_expr(&mut arg.value, aliases);
            }
        }
        ExprKind::New { class, args } => {
            if let ExprKind::Ident(name) = &class.kind {
                if let Some(path) = aliases.get(name) {
                    **class = Expression::ident(&vb_alias_target_type_name(path));
                } else {
                    rewrite_vb_aliases_in_expr(class, aliases);
                }
            } else {
                rewrite_vb_aliases_in_expr(class, aliases);
            }
            for arg in args {
                rewrite_vb_aliases_in_expr(&mut arg.value, aliases);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_vb_aliases_in_expr(target, aliases);
            rewrite_vb_aliases_in_expr(value, aliases);
        }
        ExprKind::Lambda { params, body, .. } => {
            for param in params {
                rewrite_vb_alias_in_param(param, aliases);
            }
            match body {
                LambdaBody::Expr(expr) => rewrite_vb_aliases_in_expr(expr, aliases),
                LambdaBody::Block(body) => rewrite_vb_aliases_in_statements(body, aliases),
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    rewrite_vb_aliases_in_expr(key, aliases);
                }
                rewrite_vb_aliases_in_expr(&mut item.value, aliases);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_vb_aliases_in_expr(item, aliases);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                rewrite_vb_aliases_in_expr(value, aliases);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_vb_aliases_in_expr(key, aliases);
                        rewrite_vb_aliases_in_expr(value, aliases);
                    }
                    ObjectProperty::Spread(expr) => rewrite_vb_aliases_in_expr(expr, aliases),
                    ObjectProperty::Computed { key, value } => {
                        rewrite_vb_aliases_in_expr(key, aliases);
                        rewrite_vb_aliases_in_expr(value, aliases);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_aliases_in_statement(value, aliases);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                if let InterpolPart::Expr(expr) = part {
                    rewrite_vb_aliases_in_expr(expr, aliases);
                }
            }
        }
        ExprKind::IsType { expr, type_name } | ExprKind::Cast { expr, type_name } => {
            rewrite_vb_aliases_in_expr(expr, aliases);
            if let Some(rewritten) = rewrite_vb_alias_name(type_name, aliases) {
                *type_name = rewritten;
            }
        }
        ExprKind::DefaultOf(type_name) => {
            if let Some(rewritten) = rewrite_vb_alias_name(type_name, aliases) {
                *type_name = rewritten;
            }
        }
        ExprKind::Yield(Some(expr)) => rewrite_vb_aliases_in_expr(expr, aliases),
        ExprKind::Yield(None)
        | ExprKind::AddressOf(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::Lit(_) => {}
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_vb_aliases_in_expr(&mut arg.value, aliases);
            }
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            rewrite_vb_aliases_in_expr(element, aliases);
            for generator in generators {
                rewrite_vb_aliases_in_expr(&mut generator.iter, aliases);
                for condition in &mut generator.conditions {
                    rewrite_vb_aliases_in_expr(condition, aliases);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            if let Some(lower) = lower {
                rewrite_vb_aliases_in_expr(lower, aliases);
            }
            if let Some(upper) = upper {
                rewrite_vb_aliases_in_expr(upper, aliases);
            }
            if let Some(step) = step {
                rewrite_vb_aliases_in_expr(step, aliases);
            }
        }
        ExprKind::Destructure(_) => {}
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                rewrite_vb_aliases_in_expr(parent, aliases);
            }
            for member in members {
                rewrite_vb_aliases_in_member(member, aliases);
            }
        }
        ExprKind::FunctionExpr(stmt) => rewrite_vb_aliases_in_statement(stmt, aliases),
        ExprKind::Range { start, end, .. } => {
            rewrite_vb_aliases_in_expr(start, aliases);
            rewrite_vb_aliases_in_expr(end, aliases);
        }
        ExprKind::StaticAccess { class, member } => {
            rewrite_vb_aliases_in_expr(class, aliases);
            rewrite_vb_aliases_in_expr(member, aliases);
        }
        ExprKind::Match { subject, arms } => {
            rewrite_vb_aliases_in_expr(subject, aliases);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        rewrite_vb_aliases_in_expr(condition, aliases);
                    }
                }
                rewrite_vb_aliases_in_expr(&mut arm.body, aliases);
            }
        }
    }
}

fn normalize_vb_date_literal_body(body: &mut Vec<Statement>) {
    let mut dates = HashMap::new();
    normalize_vb_date_literal_statements(body, &mut dates);
}

fn normalize_vb_date_literal_statements(
    body: &mut [Statement],
    dates: &mut HashMap<String, Expression>,
) {
    for stmt in body {
        normalize_vb_date_literal_statement(stmt, dates);
    }
}

fn normalize_vb_date_literal_statement(
    stmt: &mut Statement,
    dates: &mut HashMap<String, Expression>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_date_literal_expr(expr, dates);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        normalize_vb_date_literal_expr(bound, dates);
                    }
                }
                if let Some(init) = &mut decl.init {
                    normalize_vb_date_literal_expr(init, dates);
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if decl
                            .type_hint
                            .as_deref()
                            .is_some_and(|hint| vb_canonical_type_name(hint) == "DateTime")
                        {
                            if parse_vb_date_expr(init).is_some() {
                                dates.insert(
                                    name.to_ascii_lowercase(),
                                    Expression::new(ExprKind::Cast {
                                        expr: Box::new(init.clone()),
                                        type_name: "DateTime".into(),
                                    }),
                                );
                            } else {
                                dates.remove(&name.to_ascii_lowercase());
                            }
                        } else if parse_vb_date_expr(init).is_some()
                            || parse_vb_timespan_expr(init).is_some()
                        {
                            dates.insert(name.to_ascii_lowercase(), init.clone());
                        } else {
                            dates.remove(&name.to_ascii_lowercase());
                        }
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            normalize_vb_date_literal_expr(value, dates);
            let mut assigned_date = parse_vb_date_expr(value).map(|_| value.clone());
            for target in targets {
                normalize_vb_date_literal_expr(target, dates);
                if let ExprKind::Ident(name) = &target.kind {
                    if let Some(value) = assigned_date.take() {
                        dates.insert(name.to_ascii_lowercase(), value);
                    } else {
                        dates.remove(&name.to_ascii_lowercase());
                    }
                }
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_date_literal_expr(target, dates);
            normalize_vb_date_literal_expr(value, dates);
            if let ExprKind::Ident(name) = &target.kind {
                dates.remove(&name.to_ascii_lowercase());
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_date_literal_expr(cond, dates);
            let mut then_dates = dates.clone();
            if let Some((target, value)) = fold_vb_date_try_parse_assignment(cond) {
                then_dates.insert(target, value);
            }
            normalize_vb_date_literal_statements(then_body, &mut then_dates);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_date_literal_expr(elif_cond, dates);
                let mut elif_dates = dates.clone();
                if let Some((target, value)) = fold_vb_date_try_parse_assignment(elif_cond) {
                    elif_dates.insert(target, value);
                }
                normalize_vb_date_literal_statements(elif_body, &mut elif_dates);
            }
            if let Some(else_body) = else_body {
                let mut else_dates = dates.clone();
                normalize_vb_date_literal_statements(else_body, &mut else_dates);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_vb_date_literal_statement(init, dates);
            }
            if let Some(cond) = cond {
                normalize_vb_date_literal_expr(cond, dates);
            }
            if let Some(update) = update {
                normalize_vb_date_literal_expr(update, dates);
            }
            let mut loop_dates = dates.clone();
            normalize_vb_date_literal_statements(body, &mut loop_dates);
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_date_literal_expr(iter, dates);
            let mut loop_dates = dates.clone();
            loop_dates.remove(&var.to_ascii_lowercase());
            if let Some(key) = key {
                loop_dates.remove(&key.to_ascii_lowercase());
            }
            normalize_vb_date_literal_statements(body, &mut loop_dates);
            if let Some(else_body) = else_body {
                let mut else_dates = dates.clone();
                normalize_vb_date_literal_statements(else_body, &mut else_dates);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_date_literal_expr(cond, dates);
            let mut loop_dates = dates.clone();
            normalize_vb_date_literal_statements(body, &mut loop_dates);
            if let Some(else_body) = else_body {
                let mut else_dates = dates.clone();
                normalize_vb_date_literal_statements(else_body, &mut else_dates);
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            let mut loop_dates = dates.clone();
            normalize_vb_date_literal_statements(body, &mut loop_dates);
            normalize_vb_date_literal_expr(cond, dates);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            normalize_vb_date_literal_expr(expr, dates);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            normalize_vb_date_literal_expr(expr, dates)
                        }
                        CaseCondition::Range { from, to } => {
                            normalize_vb_date_literal_expr(from, dates);
                            normalize_vb_date_literal_expr(to, dates);
                        }
                    }
                }
                let mut case_dates = dates.clone();
                normalize_vb_date_literal_statements(&mut case.body, &mut case_dates);
            }
            if let Some(default) = default {
                let mut default_dates = dates.clone();
                normalize_vb_date_literal_statements(default, &mut default_dates);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let mut try_dates = dates.clone();
            normalize_vb_date_literal_statements(body, &mut try_dates);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_date_literal_expr(when_clause, dates);
                }
                let mut catch_dates = dates.clone();
                if let Some(var_name) = &catch.var_name {
                    catch_dates.remove(&var_name.to_ascii_lowercase());
                }
                normalize_vb_date_literal_statements(&mut catch.body, &mut catch_dates);
            }
            if let Some(else_body) = else_body {
                let mut else_dates = dates.clone();
                normalize_vb_date_literal_statements(else_body, &mut else_dates);
            }
            if let Some(finally) = finally {
                let mut finally_dates = dates.clone();
                normalize_vb_date_literal_statements(finally, &mut finally_dates);
            }
        }
        StmtKind::Block(body)
        | StmtKind::NamespaceDecl { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => {
            let mut block_dates = dates.clone();
            normalize_vb_date_literal_statements(body, &mut block_dates);
        }
        StmtKind::FunctionDecl { body, .. } => {
            let mut function_dates = HashMap::new();
            normalize_vb_date_literal_statements(body, &mut function_dates);
        }
        StmtKind::Throw {
            expr: Some(expr),
            cause,
        } => {
            normalize_vb_date_literal_expr(expr, dates);
            if let Some(cause) = cause {
                normalize_vb_date_literal_expr(cause, dates);
            }
        }
        _ => {}
    }
}

fn normalize_vb_date_literal_expr(expr: &mut Expression, dates: &HashMap<String, Expression>) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(value) = dates.get(&name.to_ascii_lowercase()) {
                *expr = value.clone();
            }
        }
        ExprKind::Binary { op, left, right } => {
            normalize_vb_date_literal_expr(left, dates);
            normalize_vb_date_literal_expr(right, dates);
            if let Some(folded) = compare_vb_dates(*op, left, right) {
                *expr = folded;
            } else if let Some(folded) = compare_vb_runtime_dates(*op, left, right) {
                *expr = folded;
            } else if let Some(folded) = fold_vb_date_arithmetic(*op, left, right) {
                *expr = folded;
            }
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::RefLoad(inner)
        | ExprKind::Await(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Spread(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner)
        | ExprKind::TypeOf(inner) => normalize_vb_date_literal_expr(inner, dates),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_date_literal_expr(cond, dates);
            normalize_vb_date_literal_expr(then, dates);
            normalize_vb_date_literal_expr(else_, dates);
        }
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } => {
            normalize_vb_date_literal_expr(object, dates);
            if let Some(value) = parse_vb_date_expr(object) {
                if let Some(folded) = fold_date_member_field(&value, field) {
                    *expr = folded;
                    return;
                }
            }
            if let Some(value) = parse_vb_timespan_expr(object) {
                if let Some(folded) = fold_timespan_member_field(&value, field) {
                    *expr = folded;
                    return;
                }
            }
            *expr = canonicalize_member_access((**object).clone(), field);
        }
        ExprKind::Member { object, .. } => normalize_vb_date_literal_expr(object, dates),
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_date_literal_expr(callee, dates);
            for arg in &mut *args {
                normalize_vb_date_literal_expr(&mut arg.value, dates);
            }
            if args.is_empty() {
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if field.eq_ignore_ascii_case("ToString") {
                        if matches!(object.kind, ExprKind::Lit(Literal::Str(_))) {
                            *expr = (**object).clone();
                            return;
                        }
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("ToString") && args.len() == 1 {
                    if let Some(value) = parse_vb_date_expr(object) {
                        if let Some(format) = literal_string(&args[0].value) {
                            if let Some(text) = format_vb_datetime_custom(&value, &format) {
                                *expr = Expression::string(&text);
                                return;
                            }
                        }
                    }
                }
            }
            if let ExprKind::Ident(name) = &callee.kind {
                if let Some(folded) = canonicalize_call(name, args) {
                    *expr = folded;
                    return;
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("Parse") && !args.is_empty() {
                    if let Some(path) = vb_date_method_receiver_name(object) {
                        if matches!(
                            path.to_ascii_lowercase().as_str(),
                            "date" | "datetime" | "system.datetime"
                        ) {
                            if let Some(text) = literal_string(&args[0].value) {
                                if let Some(parsed) = parse_vb_date_text(&text) {
                                    *expr = Expression::string(&format_vb_date_value(&parsed));
                                    return;
                                }
                            }
                        }
                    }
                }
                if field.eq_ignore_ascii_case("Compare") && args.len() == 2 {
                    if let Some(path) = vb_date_method_receiver_name(object) {
                        if matches!(
                            path.to_ascii_lowercase().as_str(),
                            "date" | "datetime" | "system.datetime"
                        ) {
                            let left = parse_vb_date_expr(&args[0].value)
                                .and_then(|value| date_value_as_datetime(&value));
                            let right = parse_vb_date_expr(&args[1].value)
                                .and_then(|value| date_value_as_datetime(&value));
                            if let (Some(left), Some(right)) = (left, right) {
                                let value = if left < right {
                                    -1
                                } else if left > right {
                                    1
                                } else {
                                    0
                                };
                                *expr = Expression::int(value);
                            }
                        }
                    }
                }
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_date_literal_expr(class, dates);
            for arg in args {
                normalize_vb_date_literal_expr(&mut arg.value, dates);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_date_literal_expr(target, dates);
            normalize_vb_date_literal_expr(value, dates);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_date_literal_expr(expr, dates),
            LambdaBody::Block(body) => {
                let mut lambda_dates = HashMap::new();
                normalize_vb_date_literal_statements(body, &mut lambda_dates);
            }
        },
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    normalize_vb_date_literal_expr(key, dates);
                }
                normalize_vb_date_literal_expr(&mut item.value, dates);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_vb_date_literal_expr(item, dates);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                normalize_vb_date_literal_expr(value, dates);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_date_literal_expr(key, dates);
                        normalize_vb_date_literal_expr(value, dates);
                    }
                    ObjectProperty::Spread(expr) => normalize_vb_date_literal_expr(expr, dates),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_date_literal_statement(value, &mut dates.clone());
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        normalize_vb_date_literal_expr(expr, dates)
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::IsType { expr: inner, .. } | ExprKind::Cast { expr: inner, .. } => {
            normalize_vb_date_literal_expr(inner, dates);
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_date_literal_expr(left, dates);
            normalize_vb_date_literal_expr(right, dates);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => {}
            PlaceExpr::Member { object, .. } => normalize_vb_date_literal_expr(object, dates),
            PlaceExpr::Index { object, index, .. } => {
                normalize_vb_date_literal_expr(object, dates);
                normalize_vb_date_literal_expr(index, dates);
            }
            PlaceExpr::Deref(expr) => normalize_vb_date_literal_expr(expr, dates),
        },
        ExprKind::Index { object, index, .. } => {
            normalize_vb_date_literal_expr(object, dates);
            normalize_vb_date_literal_expr(index, dates);
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            normalize_vb_date_literal_expr(element, dates);
            for generator in generators {
                normalize_vb_date_literal_expr(&mut generator.iter, dates);
                for filter in &mut generator.conditions {
                    normalize_vb_date_literal_expr(filter, dates);
                }
            }
        }
        ExprKind::Slice { lower, upper, step } => {
            for part in [lower, upper, step] {
                if let Some(expr) = part {
                    normalize_vb_date_literal_expr(expr, dates);
                }
            }
        }
        ExprKind::Range { start, end, .. } => {
            normalize_vb_date_literal_expr(start, dates);
            normalize_vb_date_literal_expr(end, dates);
        }
        ExprKind::StaticAccess { class, member } => {
            normalize_vb_date_literal_expr(class, dates);
            normalize_vb_date_literal_expr(member, dates);
        }
        ExprKind::Match { subject, arms } => {
            normalize_vb_date_literal_expr(subject, dates);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        normalize_vb_date_literal_expr(condition, dates);
                    }
                }
                normalize_vb_date_literal_expr(&mut arm.body, dates);
            }
        }
        ExprKind::ClassExpr {
            parent, members, ..
        } => {
            if let Some(parent) = parent {
                normalize_vb_date_literal_expr(parent, dates);
            }
            for member in members {
                normalize_vb_date_literal_member(member);
            }
        }
        ExprKind::FunctionExpr(stmt) => {
            normalize_vb_date_literal_statement(stmt, &mut HashMap::new())
        }
        ExprKind::Yield(Some(expr)) => normalize_vb_date_literal_expr(expr, dates),
        ExprKind::Lit(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::Yield(None)
        | ExprKind::AddressOf(_)
        | ExprKind::SuperCall { .. }
        | ExprKind::DefaultOf(_)
        | ExprKind::Destructure(_) => {}
    }
}

fn normalize_vb_date_literal_member(member: &mut ClassMember) {
    match member {
        ClassMember::Field {
            init, array_bounds, ..
        } => {
            if let Some(init) = init {
                normalize_vb_date_literal_expr(init, &HashMap::new());
            }
            if let Some(bounds) = array_bounds {
                for bound in bounds {
                    normalize_vb_date_literal_expr(bound, &HashMap::new());
                }
            }
        }
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_date_literal_statement(stmt, &mut HashMap::new());
        }
        ClassMember::Constructor {
            body, base_args, ..
        } => {
            if let Some(base_args) = base_args {
                for arg in base_args {
                    normalize_vb_date_literal_expr(arg, &HashMap::new());
                }
            }
            normalize_vb_date_literal_body(body);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_date_literal_body(getter);
            }
            if let Some(setter) = setter {
                normalize_vb_date_literal_body(&mut setter.body);
            }
        }
        ClassMember::Const { value, .. } => {
            normalize_vb_date_literal_expr(value, &HashMap::new());
        }
        ClassMember::Event { .. } | ClassMember::Augment(_) => {}
    }
}

fn vb_canonical_type_name(raw: &str) -> String {
    vybe_platform_dotnet::emitter::canonical_type_name(raw)
}

fn vb_gettype_type_name(raw: &str) -> String {
    let trimmed = raw.trim();
    if trimmed.contains('.') && !trimmed.to_ascii_lowercase().starts_with("system.") {
        strip_vb_generic_suffix(trimmed)
    } else {
        vb_canonical_type_name(trimmed)
    }
}

fn vb_infer_expr_type(expr: &Expression, locals: &HashMap<String, String>) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => locals.get(&name.to_ascii_lowercase()).cloned(),
        ExprKind::Lit(Literal::Int(value)) => Some(
            if *value >= i64::from(i32::MIN) && *value <= i64::from(i32::MAX) {
                "Int32".into()
            } else {
                "Int64".into()
            },
        ),
        ExprKind::Lit(Literal::Float(_)) => Some("Double".into()),
        ExprKind::Lit(Literal::Str(_)) => Some("String".into()),
        ExprKind::Lit(Literal::Bool(_)) => Some("Boolean".into()),
        ExprKind::Lit(Literal::Char(_)) => Some("Char".into()),
        ExprKind::Lit(Literal::Null) => Some("Object".into()),
        ExprKind::Call { callee, args, .. } => {
            if vb_call_returns_bool(callee) {
                return Some("Boolean".into());
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if let Some(receiver_type) = vb_infer_expr_type(object, locals) {
                    if let Some(return_type) = vybe_platform_dotnet::emitter::surface()
                        .lookup_instance_method_return_type(&receiver_type, field, args.len() as u8)
                    {
                        return Some(return_type);
                    }
                }
                if let Some(class_name) = dotted_expr_name(object) {
                    if let Some(return_type) =
                        vybe_platform_dotnet::emitter::static_method_return_type(&class_name, field)
                    {
                        return Some(return_type.to_string());
                    }
                }
            }
            parse_vb_date_expr(expr).map(|_| "DateTime".into())
        }
        ExprKind::Cast { type_name, .. } => {
            let cast_type = type_name.split(':').next_back().unwrap_or(type_name);
            Some(vb_canonical_type_name(cast_type))
        }
        ExprKind::IsType { .. } => Some("Boolean".into()),
        ExprKind::Binary { op, left, right } => match op {
            BinOp::Eq
            | BinOp::NotEq
            | BinOp::StrictEq
            | BinOp::StrictNotEq
            | BinOp::Lt
            | BinOp::LtEq
            | BinOp::Gt
            | BinOp::GtEq
            | BinOp::Is
            | BinOp::IsNot
            | BinOp::Like
            | BinOp::In
            | BinOp::NotIn
            | BinOp::InstanceOf => Some("Boolean".into()),
            BinOp::And | BinOp::Or | BinOp::Xor | BinOp::Eqv | BinOp::Imp
                if vb_expr_is_boolish(left, locals) && vb_expr_is_boolish(right, locals) =>
            {
                Some("Boolean".into())
            }
            BinOp::BitAnd | BinOp::BitOr | BinOp::BitXor
                if vb_expr_is_boolish(left, locals) && vb_expr_is_boolish(right, locals) =>
            {
                Some("Boolean".into())
            }
            _ => None,
        },
        ExprKind::Unary {
            op: UnaryOp::Not,
            expr,
        } if vb_expr_is_boolish(expr, locals) => Some("Boolean".into()),
        ExprKind::Ternary { then, else_, .. }
            if vb_expr_is_boolish(then, locals) && vb_expr_is_boolish(else_, locals) =>
        {
            Some("Boolean".into())
        }
        ExprKind::New { class, .. } => {
            dotted_expr_name(class).map(|name| vb_canonical_type_name(&name))
        }
        ExprKind::Member { object, field, .. } => {
            if let Some(receiver_type) = vb_infer_expr_type(object, locals) {
                if let Some(return_type) =
                    vybe_platform_dotnet::emitter::instance_property_type(&receiver_type, field)
                {
                    return Some(return_type.to_string());
                }
                if let Some(return_type) = vybe_platform_dotnet::emitter::surface()
                    .lookup_instance_method_return_type(&receiver_type, field, 0)
                {
                    return Some(return_type);
                }
            }
            if let Some(class_name) = dotted_expr_name(object) {
                if let Some(return_type) =
                    vybe_platform_dotnet::emitter::static_property_type(&class_name, field)
                {
                    return Some(return_type.to_string());
                }
                if let Some(return_type) =
                    vybe_platform_dotnet::emitter::static_method_return_type(&class_name, field)
                {
                    return Some(return_type.to_string());
                }
            }
            None
        }
        ExprKind::Array(_) => Some("Array".into()),
        _ => parse_vb_date_expr(expr).map(|_| "DateTime".into()),
    }
}

fn vb_apply_known_local_value(expr: &mut Expression, locals: &HashMap<String, String>) {
    if let ExprKind::Ident(name) = &expr.kind {
        let key = name.to_ascii_lowercase();
        if let Some(value) = locals.get(&format!("$bool:{key}")) {
            *expr = Expression::bool(value == "true");
        } else if let Some(value) = locals.get(&format!("$value:{key}")) {
            if let Ok(number) = value.parse::<f64>() {
                if number.fract() == 0.0 {
                    *expr = Expression::int(number as i64);
                } else {
                    *expr = Expression::float(number);
                }
            }
        } else if let Some(value) = locals.get(&format!("$string:{key}")) {
            *expr = Expression::string(value);
        }
    }
}

fn vb_integral_type_name(name: &str) -> bool {
    matches!(
        vb_canonical_type_name(name).as_str(),
        "Int16" | "Int32" | "Int64" | "Byte" | "SByte" | "UInt16" | "UInt32" | "UInt64"
    )
}

fn vb_coerce_literal_to_type(expr: &mut Expression, target_type: &str) {
    let target = vb_canonical_type_name(target_type);
    match target.as_str() {
        "String" => {
            if let ExprKind::Lit(Literal::Int(value)) = expr.kind {
                *expr = Expression::string(&value.to_string());
            } else if let ExprKind::Lit(Literal::Float(value)) = expr.kind {
                *expr = Expression::string(&format_vb_number(value));
            } else if let ExprKind::Lit(Literal::Bool(value)) = expr.kind {
                *expr = Expression::string(if value { "True" } else { "False" });
            } else if parse_vb_date_expr(expr).is_some() {
                if let Some(value) = parse_vb_date_expr(expr) {
                    *expr = Expression::string(&format_vb_date_value(&value));
                }
            }
        }
        "Boolean" => match &expr.kind {
            ExprKind::Lit(Literal::Int(value)) => *expr = Expression::bool(*value != 0),
            ExprKind::Lit(Literal::Null) => *expr = Expression::bool(false),
            _ => {}
        },
        _ if vb_integral_type_name(&target) => match &expr.kind {
            ExprKind::Lit(Literal::Bool(value)) => {
                *expr = Expression::int(if *value { -1 } else { 0 })
            }
            ExprKind::Lit(Literal::Str(value)) => {
                if let Ok(parsed) = value.trim().parse::<i64>() {
                    *expr = Expression::int(parsed);
                }
            }
            ExprKind::Lit(Literal::Float(value)) => {
                *expr = build_vb_bankers_round_expr(Expression::float(*value));
            }
            ExprKind::Lit(Literal::Null) => *expr = Expression::int(0),
            ExprKind::Cast { expr: inner, .. } => {
                if let Some(value) = literal_number(inner) {
                    *expr = build_vb_bankers_round_expr(Expression::float(value));
                }
            }
            _ => {}
        },
        _ => {}
    }
}

fn vb_default_value_for_type(target_type: &str) -> Expression {
    match vb_array_element_type_name(target_type).as_str() {
        "Boolean" => Expression::bool(false),
        "Int16" | "Int32" | "Int64" | "Byte" | "SByte" | "UInt16" | "UInt32" | "UInt64" => {
            Expression::int(0)
        }
        "Single" | "Double" | "Decimal" => Expression::float(0.0),
        "Char" | "String" | "Object" | "DateTime" => Expression::null(),
        _ => Expression::null(),
    }
}

fn vb_array_element_type_name(target_type: &str) -> String {
    let mut canonical = vb_canonical_type_name(target_type);
    while canonical.trim_end().ends_with("()") {
        canonical = canonical
            .trim_end()
            .strip_suffix("()")
            .unwrap_or(canonical.trim_end())
            .trim()
            .to_string();
    }
    canonical
}

fn vb_array_length_from_upper_bound(upper_bound: Expression) -> Expression {
    if let Some((lower, upper)) = vb_array_bound_parts(&upper_bound) {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(upper),
                right: Box::new(lower),
            })),
            right: Box::new(Expression::int(1)),
        });
    }
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(upper_bound),
        right: Box::new(Expression::int(1)),
    })
}

fn vb_array_bound_parts(expr: &Expression) -> Option<(Expression, Expression)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(&callee.kind, ExprKind::Ident(name) if name == "__vb_array_bound")
        || args.len() != 2
    {
        return None;
    }
    Some((args[0].value.clone(), args[1].value.clone()))
}

fn vb_array_bound_lower(expr: &Expression) -> Expression {
    vb_array_bound_parts(expr)
        .map(|(lower, _)| lower)
        .unwrap_or_else(|| Expression::int(0))
}

fn vb_array_bound_upper(expr: &Expression) -> Expression {
    vb_array_bound_parts(expr)
        .map(|(_, upper)| upper)
        .unwrap_or_else(|| expr.clone())
}

fn vb_literal_i64(expr: &Expression) -> Option<i64> {
    literal_number(expr).and_then(|value| {
        if value.fract() == 0.0 {
            Some(value as i64)
        } else {
            None
        }
    })
}

fn vb_array_bound_lower_i64(expr: &Expression) -> Option<i64> {
    vb_literal_i64(&vb_array_bound_lower(expr))
}

fn vb_array_bound_upper_i64(expr: &Expression) -> Option<i64> {
    vb_literal_i64(&vb_array_bound_upper(expr))
}

fn vb_array_bound_length_i64(expr: &Expression) -> Option<i64> {
    Some(vb_array_bound_upper_i64(expr)? - vb_array_bound_lower_i64(expr)? + 1)
}

fn record_vb_array_bounds_metadata(
    locals: &mut HashMap<String, String>,
    name: &str,
    bounds: &[Expression],
) {
    if bounds.is_empty() {
        return;
    }
    let key = name.to_ascii_lowercase();
    locals.insert(format!("$array_rank:{key}"), bounds.len().to_string());
    let mut total: Option<i64> = Some(1);
    for (idx, bound) in bounds.iter().enumerate() {
        if let Some(lower) = vb_array_bound_lower_i64(bound) {
            locals.insert(format!("$array_lower:{key}:{idx}"), lower.to_string());
        }
        if let Some(upper) = vb_array_bound_upper_i64(bound) {
            locals.insert(format!("$array_upper:{key}:{idx}"), upper.to_string());
        }
        if let Some(length) = vb_array_bound_length_i64(bound) {
            locals.insert(format!("$array_length:{key}:{idx}"), length.to_string());
            total = total.map(|value| value.saturating_mul(length));
        } else {
            total = None;
        }
    }
    if let Some(total) = total {
        locals.insert(format!("$array_total_length:{key}"), total.to_string());
    }
}

fn vb_filled_array_expr(length: Expression, default_value: Expression) -> Expression {
    let array_expr = call_expr(
        Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Array")),
            field: "CreateInstance".to_string(),
            null_safe: false,
        }),
        vec![
            Argument::positional(Expression::null()),
            Argument::positional(length),
        ],
    );
    let mut body = vec![Statement::with_span(
        StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__arr".into()),
                type_hint: None,
                init: Some(array_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        },
        Span::default(),
    )];
    if !matches!(default_value.kind, ExprKind::Lit(Literal::Null)) {
        body.push(Statement::with_span(
            StmtKind::Expr(call_expr(
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident("Array")),
                    field: "Fill".to_string(),
                    null_safe: false,
                }),
                vec![
                    Argument::positional(Expression::ident("__arr")),
                    Argument::positional(default_value),
                ],
            )),
            Span::default(),
        ));
    }
    body.push(Statement::with_span(
        StmtKind::Return(Some(Expression::ident("__arr"))),
        Span::default(),
    ));
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: vec![],
        optional: false,
    })
}

fn vb_multidim_array_expr(bounds: &[Expression], default_value: Expression) -> Expression {
    let mut value = default_value;
    for bound in bounds.iter().rev() {
        value = vb_filled_array_expr(vb_array_length_from_upper_bound(bound.clone()), value);
    }
    value
}

fn vb_expr_has_decimal(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Cast { expr, type_name } => {
            vb_canonical_type_name(type_name) == "Decimal" || vb_expr_has_decimal(expr)
        }
        ExprKind::Binary { left, right, .. } => {
            vb_expr_has_decimal(left) || vb_expr_has_decimal(right)
        }
        ExprKind::Unary { expr, .. } => vb_expr_has_decimal(expr),
        _ => false,
    }
}

fn vb_expr_has_decimal_with_locals(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    match &expr.kind {
        ExprKind::Ident(name) => {
            locals.contains_key(&format!("$decimal:{}", name.to_ascii_lowercase()))
        }
        _ => vb_expr_has_decimal(expr),
    }
}

fn literal_number_with_locals(expr: &Expression, locals: &HashMap<String, String>) -> Option<f64> {
    match &expr.kind {
        ExprKind::Ident(name) => locals
            .get(&format!("$value:{}", name.to_ascii_lowercase()))
            .and_then(|value| value.parse().ok()),
        _ => literal_number(expr),
    }
}

fn vb_fold_decimal_comparison(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    if !vb_expr_has_decimal_with_locals(left, locals)
        && !vb_expr_has_decimal_with_locals(right, locals)
    {
        return None;
    }
    let left = literal_number_with_locals(left, locals)?;
    let right = literal_number_with_locals(right, locals)?;
    let epsilon = 1e-13;
    let value = match op {
        BinOp::Eq => (left - right).abs() <= epsilon,
        BinOp::NotEq => (left - right).abs() > epsilon,
        BinOp::Lt => left < right && (left - right).abs() > epsilon,
        BinOp::Gt => left > right && (left - right).abs() > epsilon,
        BinOp::LtEq => left < right || (left - right).abs() <= epsilon,
        BinOp::GtEq => left > right || (left - right).abs() <= epsilon,
        _ => return None,
    };
    Some(Expression::bool(value))
}

fn format_vb_number(value: f64) -> String {
    if value.is_finite() && value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn normalize_vb_local_type_body(body: &mut Vec<Statement>) {
    let mut locals = HashMap::new();
    normalize_vb_local_type_statements(body, &mut locals);
}

fn normalize_vb_local_type_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_local_type_statement(stmt, locals);
    }
}

fn normalize_vb_generic_new_factory_calls(module: &mut Module) {
    let mut factories = HashMap::new();
    collect_vb_generic_new_factories(&module.body, &mut factories);
    if factories.is_empty() {
        return;
    }
    rewrite_vb_generic_new_factory_statements(&mut module.body, &factories, &mut HashMap::new());
}

fn collect_vb_generic_new_factories(
    body: &[Statement],
    factories: &mut HashMap<(String, String), String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                let owner = vb_canonical_type_name(name).to_ascii_lowercase();
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            let StmtKind::FunctionDecl {
                                name: method_name,
                                params,
                                body,
                                return_type,
                                ..
                            } = &method.kind
                            else {
                                continue;
                            };
                            if !params.is_empty() || body.len() != 1 {
                                continue;
                            }
                            let Some(return_type) = return_type.as_ref() else {
                                continue;
                            };
                            let StmtKind::Return(Some(expr)) = &body[0].kind else {
                                continue;
                            };
                            let ExprKind::New { class, args } = &expr.kind else {
                                continue;
                            };
                            if !args.is_empty() {
                                continue;
                            }
                            let Some(type_param) = dotted_expr_name(class) else {
                                continue;
                            };
                            if return_type.eq_ignore_ascii_case(&type_param) {
                                factories.insert(
                                    (owner.clone(), method_name.to_ascii_lowercase()),
                                    type_param,
                                );
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_vb_generic_new_factories(
                                std::slice::from_ref(nested),
                                factories,
                            );
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } => {
                collect_vb_generic_new_factories(body, factories)
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    match member {
                        ClassMember::Method(method) => {
                            let StmtKind::FunctionDecl {
                                name: method_name,
                                params,
                                body,
                                return_type,
                                ..
                            } = &method.kind
                            else {
                                continue;
                            };
                            if !params.is_empty() || body.len() != 1 {
                                continue;
                            }
                            let Some(return_type) = return_type.as_ref() else {
                                continue;
                            };
                            let StmtKind::Return(Some(expr)) = &body[0].kind else {
                                continue;
                            };
                            let ExprKind::New { class, args } = &expr.kind else {
                                continue;
                            };
                            if !args.is_empty() {
                                continue;
                            }
                            let Some(type_param) = dotted_expr_name(class) else {
                                continue;
                            };
                            if return_type.eq_ignore_ascii_case(&type_param) {
                                factories.insert(
                                    ("".to_string(), method_name.to_ascii_lowercase()),
                                    type_param,
                                );
                            }
                        }
                        ClassMember::NestedType(nested) => {
                            collect_vb_generic_new_factories(
                                std::slice::from_ref(nested),
                                factories,
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

fn rewrite_vb_generic_new_factory_statements(
    body: &mut [Statement],
    factories: &HashMap<(String, String), String>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_generic_new_factory_statement(stmt, factories, locals);
    }
}

fn rewrite_vb_generic_new_factory_statement(
    stmt: &mut Statement,
    factories: &HashMap<(String, String), String>,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_generic_new_factory_expr(init, factories, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let ty = decl
                        .type_hint
                        .as_ref()
                        .map(|hint| hint.trim().to_string())
                        .or_else(|| {
                            decl.init
                                .as_ref()
                                .and_then(|init| vb_infer_expr_type(init, locals))
                        });
                    if let Some(ty) = ty {
                        locals.insert(name.to_ascii_lowercase(), ty);
                    }
                }
            }
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_generic_new_factory_expr(expr, factories, locals);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_generic_new_factory_expr(target, factories, locals);
            }
            rewrite_vb_generic_new_factory_expr(value, factories, locals);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scoped = locals.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    scoped.insert(param.name.to_ascii_lowercase(), type_hint.clone());
                }
            }
            rewrite_vb_generic_new_factory_statements(body, factories, &mut scoped);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                rewrite_vb_generic_new_factory_member(member, factories);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_generic_new_factory_expr(cond, factories, locals);
            rewrite_vb_generic_new_factory_statements(then_body, factories, &mut locals.clone());
            for (cond, body) in elifs {
                rewrite_vb_generic_new_factory_expr(cond, factories, locals);
                rewrite_vb_generic_new_factory_statements(body, factories, &mut locals.clone());
            }
            if let Some(body) = else_body {
                rewrite_vb_generic_new_factory_statements(body, factories, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_generic_new_factory_expr(cond, factories, locals);
            rewrite_vb_generic_new_factory_statements(body, factories, &mut locals.clone());
            if let Some(body) = else_body {
                rewrite_vb_generic_new_factory_statements(body, factories, &mut locals.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_generic_new_factory_statements(body, factories, &mut locals.clone());
        }
        _ => {}
    }
}

fn rewrite_vb_generic_new_factory_member(
    member: &mut ClassMember,
    factories: &HashMap<(String, String), String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_generic_new_factory_statement(stmt, factories, &mut HashMap::new());
        }
        ClassMember::Constructor { params, body, .. } => {
            let mut locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    locals.insert(param.name.to_ascii_lowercase(), type_hint.clone());
                }
            }
            rewrite_vb_generic_new_factory_statements(body, factories, &mut locals);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_generic_new_factory_statements(getter, factories, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                rewrite_vb_generic_new_factory_statements(
                    &mut setter.body,
                    factories,
                    &mut HashMap::new(),
                );
            }
        }
        _ => {}
    }
}

fn rewrite_vb_generic_new_factory_expr(
    expr: &mut Expression,
    factories: &HashMap<(String, String), String>,
    locals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_generic_new_factory_expr(callee, factories, locals);
            for arg in &mut *args {
                rewrite_vb_generic_new_factory_expr(&mut arg.value, factories, locals);
            }
            if args.is_empty() {
                if let ExprKind::Ident(name) = &callee.kind {
                    if let Some((method_name, actual_type)) = vb_generic_call_marker_parts(name) {
                        if factories
                            .contains_key(&("".to_string(), method_name.to_ascii_lowercase()))
                        {
                            *expr = Expression::new(ExprKind::New {
                                class: Box::new(build_dotted_expr(&actual_type)),
                                args: Vec::new(),
                            });
                            return;
                        }
                    }
                }
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if let Some(actual_type) =
                        vb_generic_factory_actual_type(object, field, locals, factories)
                    {
                        *expr = Expression::new(ExprKind::New {
                            class: Box::new(build_dotted_expr(&actual_type)),
                            args: Vec::new(),
                        });
                    }
                }
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_vb_generic_new_factory_expr(object, factories, locals);
        }
        ExprKind::New { class, args } => {
            rewrite_vb_generic_new_factory_expr(class, factories, locals);
            for arg in args {
                rewrite_vb_generic_new_factory_expr(&mut arg.value, factories, locals);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_generic_new_factory_expr(left, factories, locals);
            rewrite_vb_generic_new_factory_expr(right, factories, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::RefLoad(expr) => {
            rewrite_vb_generic_new_factory_expr(expr, factories, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_generic_new_factory_expr(&mut item.value, factories, locals);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                rewrite_vb_generic_new_factory_expr(item, factories, locals);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                rewrite_vb_generic_new_factory_expr(value, factories, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_vb_generic_new_factory_expr(key, factories, locals);
                        rewrite_vb_generic_new_factory_expr(value, factories, locals);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_vb_generic_new_factory_expr(value, factories, locals);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_generic_new_factory_statement(
                            value,
                            factories,
                            &mut locals.clone(),
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn vb_generic_factory_actual_type(
    object: &Expression,
    method_name: &str,
    locals: &HashMap<String, String>,
    factories: &HashMap<(String, String), String>,
) -> Option<String> {
    let receiver_type = vb_infer_expr_type(object, locals)?;
    let base = strip_vb_generic_suffix(&receiver_type);
    let _type_param =
        factories.get(&(base.to_ascii_lowercase(), method_name.to_ascii_lowercase()))?;
    vb_generic_suffix_first_type(receiver_type.trim().strip_prefix(&base).unwrap_or_default())
        .or_else(|| vb_call_generic_first_type(&receiver_type))
}

fn normalize_vb_option_compare_text(module: &mut Module) {
    normalize_vb_option_compare_text_statements(&mut module.body, &mut HashMap::new());
}

fn normalize_vb_option_compare_text_class_members(
    members: &mut [ClassMember],
    locals: &HashMap<String, String>,
) {
    for member in members {
        match member {
            ClassMember::Field {
                name,
                type_hint,
                init,
                ..
            } => {
                if let Some(init) = init {
                    normalize_vb_option_compare_text_expr(init, locals);
                }
                let _ = (name, type_hint);
            }
            ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
                normalize_vb_option_compare_text_statement(stmt, &mut locals.clone());
            }
            ClassMember::Constructor { body, .. } => {
                normalize_vb_option_compare_text_statements(body, &mut locals.clone());
            }
            ClassMember::Property { getter, setter, .. } => {
                if let Some(getter) = getter {
                    normalize_vb_option_compare_text_statements(getter, &mut locals.clone());
                }
                if let Some(setter) = setter {
                    normalize_vb_option_compare_text_statements(
                        &mut setter.body,
                        &mut locals.clone(),
                    );
                }
            }
            ClassMember::Const { value, .. } => {
                normalize_vb_option_compare_text_expr(value, locals);
            }
            ClassMember::Event { .. } | ClassMember::Augment(_) => {}
        }
    }
}

fn normalize_vb_option_compare_text_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_option_compare_text_statement(stmt, locals);
    }
}

fn normalize_vb_option_compare_text_statement(
    stmt: &mut Statement,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_option_compare_text_expr(expr, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_option_compare_text_expr(init, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(type_hint) = &decl.type_hint {
                        locals.insert(name.to_ascii_lowercase(), vb_canonical_type_name(type_hint));
                    } else if let Some(init) = &decl.init {
                        if let Some(type_name) = vb_infer_expr_type(init, locals) {
                            locals.insert(name.to_ascii_lowercase(), type_name);
                        }
                    }
                }
            }
        }
        StmtKind::FunctionDecl { body, params, .. } => {
            let mut fn_locals = HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    fn_locals.insert(
                        param.name.to_ascii_lowercase(),
                        vb_canonical_type_name(type_hint),
                    );
                }
            }
            normalize_vb_option_compare_text_statements(body, &mut fn_locals);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            normalize_vb_option_compare_text_class_members(members, locals);
        }
        StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
        }
        StmtKind::Assign { targets, value } => {
            normalize_vb_option_compare_text_expr(value, locals);
            for target in targets {
                normalize_vb_option_compare_text_expr(target, locals);
            }
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_option_compare_text_expr(target, locals);
            normalize_vb_option_compare_text_expr(value, locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_option_compare_text_expr(cond, locals);
            normalize_vb_option_compare_text_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_option_compare_text_expr(elif_cond, locals);
                normalize_vb_option_compare_text_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_option_compare_text_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_option_compare_text_statement(init, &mut loop_locals);
            }
            if let Some(cond) = cond {
                normalize_vb_option_compare_text_expr(cond, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_option_compare_text_expr(update, &loop_locals);
            }
            normalize_vb_option_compare_text_statements(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_option_compare_text_expr(iter, locals);
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_option_compare_text_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_option_compare_text_expr(cond, locals);
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_option_compare_text_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
            normalize_vb_option_compare_text_expr(cond, locals);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            normalize_vb_option_compare_text_expr(expr, locals);
            for case in cases {
                for condition in &mut case.conditions {
                    match condition {
                        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
                            normalize_vb_option_compare_text_expr(expr, locals);
                        }
                        CaseCondition::Range { from, to } => {
                            normalize_vb_option_compare_text_expr(from, locals);
                            normalize_vb_option_compare_text_expr(to, locals);
                        }
                    }
                }
                normalize_vb_option_compare_text_statements(&mut case.body, &mut locals.clone());
            }
            if let Some(default) = default {
                normalize_vb_option_compare_text_statements(default, &mut locals.clone());
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_option_compare_text_expr(when_clause, locals);
                }
                normalize_vb_option_compare_text_statements(&mut catch.body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_option_compare_text_statements(else_body, &mut locals.clone());
            }
            if let Some(finally) = finally {
                normalize_vb_option_compare_text_statements(finally, &mut locals.clone());
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                normalize_vb_option_compare_text_expr(&mut item.expr, locals);
            }
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
        }
        StmtKind::Using { resource, body, .. } => {
            normalize_vb_option_compare_text_expr(resource, locals);
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
        }
        StmtKind::Lock { expr, body } => {
            normalize_vb_option_compare_text_expr(expr, locals);
            normalize_vb_option_compare_text_statements(body, &mut locals.clone());
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                normalize_vb_option_compare_text_expr(expr, locals);
            }
            if let Some(cause) = cause {
                normalize_vb_option_compare_text_expr(cause, locals);
            }
        }
        StmtKind::AddHandler {
            control, handler, ..
        }
        | StmtKind::RemoveHandler {
            control, handler, ..
        } => {
            normalize_vb_option_compare_text_expr(control, locals);
            normalize_vb_option_compare_text_expr(handler, locals);
        }
        StmtKind::RaiseEvent { args, .. } => {
            for arg in args {
                normalize_vb_option_compare_text_expr(arg, locals);
            }
        }
        StmtKind::ReDim { bounds, .. } => {
            for bound in bounds {
                normalize_vb_option_compare_text_expr(bound, locals);
            }
        }
        StmtKind::OpenFile {
            path, file_number, ..
        } => {
            normalize_vb_option_compare_text_expr(path, locals);
            normalize_vb_option_compare_text_expr(file_number, locals);
        }
        StmtKind::CloseFile(Some(expr))
        | StmtKind::LineInput {
            file_number: expr, ..
        } => {
            normalize_vb_option_compare_text_expr(expr, locals);
        }
        StmtKind::PrintFile { file_number, items } | StmtKind::WriteFile { file_number, items } => {
            normalize_vb_option_compare_text_expr(file_number, locals);
            for item in items {
                normalize_vb_option_compare_text_expr(item, locals);
            }
        }
        StmtKind::InputFile {
            file_number,
            variables,
        } => {
            normalize_vb_option_compare_text_expr(file_number, locals);
            for variable in variables {
                normalize_vb_option_compare_text_expr(variable, locals);
            }
        }
        _ => {}
    }
}

fn normalize_vb_option_compare_text_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Binary { op, left, right } => {
            normalize_vb_option_compare_text_expr(left, locals);
            normalize_vb_option_compare_text_expr(right, locals);
            if let Some(rewritten) = rewrite_vb_text_compare_binary(*op, left, right, locals) {
                *expr = rewritten;
            }
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_option_compare_text_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_option_compare_text_expr(&mut arg.value, locals);
            }
            if args.len() == 2
                && dotted_expr_name(callee).as_deref().is_some_and(|name| {
                    name.eq_ignore_ascii_case("__vb_like_ismatch")
                        || name.eq_ignore_ascii_case("Regex.IsMatch")
                        || name.eq_ignore_ascii_case("System.Text.RegularExpressions.Regex.IsMatch")
                })
            {
                args.push(Argument::positional(build_dotted_expr(
                    "RegexOptions.IgnoreCase",
                )));
            }
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_option_compare_text_expr(expr, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_option_compare_text_expr(cond, locals);
            normalize_vb_option_compare_text_expr(then, locals);
            normalize_vb_option_compare_text_expr(else_, locals);
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_option_compare_text_expr(left, locals);
            normalize_vb_option_compare_text_expr(right, locals);
        }
        ExprKind::Member { object, .. } => normalize_vb_option_compare_text_expr(object, locals),
        ExprKind::Index { object, index, .. } => {
            normalize_vb_option_compare_text_expr(object, locals);
            normalize_vb_option_compare_text_expr(index, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_option_compare_text_expr(&mut item.value, locals);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_option_compare_text_expr(class, locals);
            for arg in args {
                normalize_vb_option_compare_text_expr(&mut arg.value, locals);
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_option_compare_text_expr(expr, locals),
            LambdaBody::Block(body) => {
                normalize_vb_option_compare_text_statements(body, &mut locals.clone())
            }
        },
        ExprKind::ClassExpr { members, .. } => {
            normalize_vb_option_compare_text_class_members(members, locals);
        }
        ExprKind::FunctionExpr(stmt) => {
            normalize_vb_option_compare_text_statement(stmt, &mut locals.clone());
        }
        ExprKind::Yield(Some(expr)) => normalize_vb_option_compare_text_expr(expr, locals),
        ExprKind::Sequence(exprs) => {
            for expr in exprs {
                normalize_vb_option_compare_text_expr(expr, locals);
            }
        }
        _ => {}
    }
}

fn rewrite_vb_text_compare_binary(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    if !matches!(
        op,
        BinOp::Eq | BinOp::NotEq | BinOp::Lt | BinOp::LtEq | BinOp::Gt | BinOp::GtEq
    ) || !vb_option_compare_operand_is_string(left, locals)
        || !vb_option_compare_operand_is_string(right, locals)
    {
        return None;
    }

    let compare = call_expr(
        Expression::ident("__dotnet_string_compare"),
        vec![
            Argument::positional(left.clone()),
            Argument::positional(right.clone()),
            Argument::positional(Expression::string(
                "__dotnet_stringcomparison_ordinalignorecase",
            )),
        ],
    );
    Some(Expression::new(ExprKind::Binary {
        op,
        left: Box::new(compare),
        right: Box::new(Expression::int(0)),
    }))
}

fn vb_option_compare_operand_is_string(
    expr: &Expression,
    locals: &HashMap<String, String>,
) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(_)) => true,
        ExprKind::Ident(name) => locals
            .get(&name.to_ascii_lowercase())
            .is_some_and(|type_name| type_name == "String"),
        ExprKind::Cast { type_name, .. } => vb_canonical_type_name(type_name) == "String",
        _ => vb_infer_expr_type(expr, locals).is_some_and(|type_name| type_name == "String"),
    }
}

fn normalize_vb_dotnet_collection_calls(module: &mut Module) {
    normalize_vb_dotnet_collection_statements(&mut module.body, &mut HashMap::new());
}

fn normalize_vb_dotnet_collection_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_dotnet_collection_statement(stmt, locals);
    }
}

fn normalize_vb_dotnet_collection_statement(
    stmt: &mut Statement,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            normalize_vb_dotnet_collection_expr(expr, locals);
        }
        StmtKind::Return(Some(expr)) => {
            normalize_vb_dotnet_collection_expr(expr, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_dotnet_collection_expr(init, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(bounds) = decl.array_bounds.as_ref() {
                        record_vb_array_bounds_metadata(locals, name, bounds);
                    }
                    if let Some(type_name) = decl
                        .type_hint
                        .as_deref()
                        .and_then(dotnet_vb::collection_local_type)
                        .or_else(|| {
                            decl.init
                                .as_ref()
                                .and_then(vb_new_expr_type_name)
                                .as_deref()
                                .and_then(dotnet_vb::collection_local_type)
                        })
                    {
                        let mut local_type = if type_name == "Dictionary"
                            && decl
                                .init
                                .as_ref()
                                .is_some_and(vb_new_dictionary_uses_ignorecase)
                        {
                            "DictionaryIgnoreCase".to_string()
                        } else {
                            type_name
                        };
                        if decl.array_bounds.is_some() && !local_type.trim().ends_with("()") {
                            local_type.push_str("()");
                        }
                        let storage_type = dotnet_vb::collection_storage_type(&local_type);
                        if vybe_platform_dotnet::emitter::is_component_descriptor_class(
                            storage_type,
                        ) {
                            decl.type_hint = Some(storage_type.into());
                            if let Some(init) = &mut decl.init {
                                vb_normalize_new_class_name(init, storage_type);
                            }
                        }
                        locals.insert(name.to_ascii_lowercase(), local_type.into());
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_dotnet_collection_expr(target, locals);
            }
            normalize_vb_dotnet_collection_expr(value, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_dotnet_collection_expr(target, locals);
            normalize_vb_dotnet_collection_expr(value, locals);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut function_locals = locals.clone();
            for param in params {
                if let Some(type_name) = param
                    .type_hint
                    .as_deref()
                    .and_then(dotnet_vb::collection_local_type)
                {
                    function_locals.insert(param.name.to_ascii_lowercase(), type_name);
                }
            }
            normalize_vb_dotnet_collection_statements(body, &mut function_locals);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            let member_locals = dotnet_collection_member_locals(members);
            for member in members {
                normalize_vb_dotnet_collection_member(member, &member_locals);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_dotnet_collection_expr(cond, locals);
            normalize_vb_dotnet_collection_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_dotnet_collection_expr(elif_cond, locals);
                normalize_vb_dotnet_collection_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_dotnet_collection_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_dotnet_collection_statement(init, &mut loop_locals);
            }
            if let Some(cond) = cond {
                normalize_vb_dotnet_collection_expr(cond, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_dotnet_collection_expr(update, &loop_locals);
            }
            normalize_vb_dotnet_collection_statements(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            var,
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_dotnet_collection_expr(iter, locals);
            let mut loop_locals = locals.clone();
            if matches!(
                &iter.kind,
                ExprKind::Ident(name)
                    if locals
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|type_name| dotnet_vb::collection_type_is_dictionary(type_name))
            ) {
                *iter = call_expr(
                    Expression::new(ExprKind::Member {
                        object: Box::new(iter.clone()),
                        field: "Entries".into(),
                        null_safe: false,
                    }),
                    Vec::new(),
                );
                loop_locals.insert(var.to_ascii_lowercase(), "KeyValuePair".into());
            } else if matches!(
                &iter.kind,
                ExprKind::Ident(name)
                    if locals
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|type_name| dotnet_vb::collection_base_type_name(type_name).eq_ignore_ascii_case("Collection"))
            ) {
                *iter = call_expr(
                    Expression::new(ExprKind::Member {
                        object: Box::new(iter.clone()),
                        field: "ToArray".into(),
                        null_safe: false,
                    }),
                    Vec::new(),
                );
            }
            normalize_vb_dotnet_collection_statements(body, &mut loop_locals);
            if let Some(else_body) = else_body {
                normalize_vb_dotnet_collection_statements(else_body, &mut loop_locals);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_dotnet_collection_expr(cond, locals);
            normalize_vb_dotnet_collection_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_dotnet_collection_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_dotnet_collection_statements(body, &mut locals.clone());
        }
        _ => {}
    }
}

fn dotnet_collection_member_locals(members: &[ClassMember]) -> HashMap<String, String> {
    let mut locals = HashMap::new();
    for member in members {
        let ClassMember::Field {
            name,
            type_hint: Some(type_hint),
            ..
        } = member
        else {
            continue;
        };
        if let Some(type_name) = dotnet_vb::collection_local_type(type_hint) {
            locals.insert(name.to_ascii_lowercase(), type_name);
        }
    }
    locals
}

fn normalize_vb_dotnet_collection_member(
    member: &mut ClassMember,
    member_locals: &HashMap<String, String>,
) {
    match member {
        ClassMember::Method(stmt) => {
            normalize_vb_dotnet_collection_statement(stmt, &mut member_locals.clone());
        }
        ClassMember::NestedType(stmt) => {
            normalize_vb_dotnet_collection_statement(stmt, &mut HashMap::new());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_dotnet_collection_statements(body, &mut member_locals.clone());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_dotnet_collection_statements(getter, &mut member_locals.clone());
            }
            if let Some(setter) = setter {
                normalize_vb_dotnet_collection_statements(
                    &mut setter.body,
                    &mut member_locals.clone(),
                );
            }
        }
        _ => {}
    }
}

fn normalize_vb_dotnet_collection_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_dotnet_collection_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_dotnet_collection_expr(&mut arg.value, locals);
            }

            loop {
                let replacement = if args.is_empty() {
                    if let ExprKind::Call {
                        callee: inner_callee,
                        args: inner_args,
                        ..
                    } = &callee.kind
                    {
                        inner_args.is_empty().then(|| inner_callee.clone())
                    } else {
                        None
                    }
                } else {
                    None
                };
                let Some(new_callee) = replacement else {
                    break;
                };
                *callee = new_callee;
            }

            if !args.is_empty() {
                if let ExprKind::Ident(name) = &callee.kind {
                    if vb_local_is_array_like(name, args.len(), locals) {
                        *expr = vb_array_index_chain_for_local(name, args, locals);
                        return;
                    }
                    if locals
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|type_name| type_name.trim().ends_with("()"))
                        && args.len() == 1
                    {
                        *expr = Expression::new(ExprKind::Index {
                            object: Box::new(Expression::ident(name)),
                            index: Box::new(args[0].value.clone()),
                            null_safe: false,
                        });
                        return;
                    } else if locals
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|type_name| type_name == "Collection")
                    {
                        *callee = Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(Expression::ident(name)),
                            field: "Item".into(),
                            null_safe: false,
                        }));
                    } else if locals
                        .get(&name.to_ascii_lowercase())
                        .is_some_and(|type_name| {
                            dotnet_vb::collection_type_is_dictionary(type_name)
                        })
                        && args.len() == 1
                    {
                        *expr = Expression::new(ExprKind::Index {
                            object: Box::new(Expression::ident(name)),
                            index: Box::new(vb_normalize_dictionary_key(
                                &Expression::ident(name),
                                args[0].value.clone(),
                                locals,
                            )),
                            null_safe: false,
                        });
                        return;
                    }
                }
            }

            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field.eq_ignore_ascii_case("GetLength") && args.len() == 1 {
                    if let Some((name, dim)) = vb_array_local_dim(object, &args[0].value) {
                        if let Some(length) = locals.get(&format!("$array_length:{name}:{dim}")) {
                            if let Ok(value) = length.parse::<i64>() {
                                *expr = Expression::int(value);
                                return;
                            }
                        }
                    }
                    if vb_is_zero_literal(&args[0].value) {
                        *expr = Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: "Length".into(),
                            null_safe: false,
                        });
                        return;
                    }
                    if vb_is_one_literal(&args[0].value) {
                        *expr = Expression::new(ExprKind::Member {
                            object: Box::new(Expression::new(ExprKind::Index {
                                object: object.clone(),
                                index: Box::new(Expression::int(0)),
                                null_safe: false,
                            })),
                            field: "Length".into(),
                            null_safe: false,
                        });
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("GetLowerBound") && args.len() == 1 {
                    if let Some((name, dim)) = vb_array_local_dim(object, &args[0].value) {
                        if let Some(lower) = locals.get(&format!("$array_lower:{name}:{dim}")) {
                            if let Ok(value) = lower.parse::<i64>() {
                                *expr = Expression::int(value);
                                return;
                            }
                        }
                    }
                    *expr = Expression::int(0);
                    return;
                }
                if field.eq_ignore_ascii_case("GetUpperBound") && args.len() == 1 {
                    if let Some((name, dim)) = vb_array_local_dim(object, &args[0].value) {
                        if let Some(upper) = locals.get(&format!("$array_upper:{name}:{dim}")) {
                            if let Ok(value) = upper.parse::<i64>() {
                                *expr = Expression::int(value);
                                return;
                            }
                        }
                    }
                    if vb_is_zero_literal(&args[0].value) {
                        *expr = Expression::new(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(Expression::new(ExprKind::Member {
                                object: object.clone(),
                                field: "Length".into(),
                                null_safe: false,
                            })),
                            right: Box::new(Expression::int(1)),
                        });
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("GetValue") && !args.is_empty() {
                    *expr = vb_array_index_chain((**object).clone(), args);
                    return;
                }
                if field.eq_ignore_ascii_case("SetValue") && args.len() >= 2 {
                    let value = args[0].value.clone();
                    let target = vb_array_index_chain((**object).clone(), &args[1..]);
                    *expr = Expression::new(ExprKind::Assign {
                        target: Box::new(target),
                        value: Box::new(value),
                    });
                    return;
                }
                if field.eq_ignore_ascii_case("CompareTo") && args.len() == 1 {
                    *expr = call_expr(
                        Expression::ident("__dotnet_string_compare"),
                        vec![
                            Argument::positional((**object).clone()),
                            Argument::positional(args[0].value.clone()),
                        ],
                    );
                    return;
                }
                if field.eq_ignore_ascii_case("Item")
                    && args.len() == 1
                    && matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if locals
                                .get(&name.to_ascii_lowercase())
                                .is_some_and(|type_name| dotnet_vb::collection_type_is_dictionary(type_name))
                    )
                {
                    *expr = Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: Box::new(vb_normalize_dictionary_key(
                            object,
                            args[0].value.clone(),
                            locals,
                        )),
                        null_safe: false,
                    });
                    return;
                }
                if args.len() == 1
                    && (field.eq_ignore_ascii_case("OldItems")
                        || field.eq_ignore_ascii_case("NewItems"))
                {
                    *expr = Expression::new(ExprKind::Index {
                        object: Box::new(Expression::new(ExprKind::Member {
                            object: object.clone(),
                            field: field.clone(),
                            null_safe: false,
                        })),
                        index: Box::new(args[0].value.clone()),
                        null_safe: false,
                    });
                    return;
                }
                if !args.is_empty() {
                    if let ExprKind::Ident(name) = &object.kind {
                        if locals
                            .get(&name.to_ascii_lowercase())
                            .is_some_and(|type_name| {
                                dotnet_vb::collection_type_is_dictionary(type_name)
                                    && dotnet_vb::collection_method_takes_dictionary_key(
                                        type_name,
                                        field,
                                        args.len(),
                                    )
                            })
                        {
                            args[0].value =
                                vb_normalize_dictionary_key(object, args[0].value.clone(), locals);
                        }
                        if field.eq_ignore_ascii_case("Add")
                            && args.len() == 1
                            && locals
                                .get(&name.to_ascii_lowercase())
                                .is_some_and(|type_name| {
                                    dotnet_vb::collection_base_type_name(type_name)
                                        .eq_ignore_ascii_case("List")
                                })
                        {
                            *expr = call_expr(
                                Expression::ident("__dotnet_list_add"),
                                vec![Argument::positional((**object).clone()), args[0].clone()],
                            );
                            return;
                        }
                    }
                }
                if field.eq_ignore_ascii_case("CopyTo")
                    && args.len() == 2
                    && vb_is_zero_literal(&args[1].value)
                {
                    let source = (**object).clone();
                    let dest = args[0].clone();
                    *expr = call_expr(
                        vb_system_array_member("Copy"),
                        vec![
                            Argument::positional(source.clone()),
                            dest,
                            Argument::positional(Expression::new(ExprKind::Member {
                                object: Box::new(source),
                                field: "Length".into(),
                                null_safe: false,
                            })),
                        ],
                    );
                }
            }
            if let Some(flattened) = flatten_vb_append_chain(expr) {
                *expr = flattened;
                normalize_vb_dotnet_collection_expr(expr, locals);
                return;
            }
        }
        ExprKind::Member { object, field, .. } => {
            normalize_vb_dotnet_collection_expr(object, locals);
            if let ExprKind::Ident(name) = &object.kind {
                let key = name.to_ascii_lowercase();
                if field.eq_ignore_ascii_case("Rank") {
                    if let Some(rank) = locals.get(&format!("$array_rank:{key}")) {
                        if let Ok(value) = rank.parse::<i64>() {
                            *expr = Expression::int(value);
                            return;
                        }
                    }
                }
                if field.eq_ignore_ascii_case("LongLength") || field.eq_ignore_ascii_case("Length")
                {
                    if let Some(length) = locals.get(&format!("$array_total_length:{key}")) {
                        if let Ok(value) = length.parse::<i64>() {
                            *expr = Expression::int(value);
                            return;
                        }
                    }
                }
            }
            if field.eq_ignore_ascii_case("IsFixedSize") {
                *expr = Expression::bool(true);
                return;
            }
            if field.eq_ignore_ascii_case("IsReadOnly")
                || field.eq_ignore_ascii_case("IsSynchronized")
            {
                *expr = Expression::bool(false);
                return;
            }
            if field.eq_ignore_ascii_case("SyncRoot") {
                *expr = (**object).clone();
                return;
            }
            if field.eq_ignore_ascii_case("HasValue") {
                *expr = Expression::new(ExprKind::Binary {
                    op: BinOp::IsNot,
                    left: Box::new((**object).clone()),
                    right: Box::new(Expression::null()),
                });
                return;
            }
            if let ExprKind::Ident(name) = &object.kind {
                if let Some(type_name) = locals.get(&name.to_ascii_lowercase()) {
                    if type_name == "KeyValuePair" {
                        let index = if field.eq_ignore_ascii_case("Key") {
                            Some(0)
                        } else if field.eq_ignore_ascii_case("Value") {
                            Some(1)
                        } else {
                            None
                        };
                        if let Some(index) = index {
                            *expr = Expression::new(ExprKind::Index {
                                object: Box::new(Expression::ident(name)),
                                index: Box::new(Expression::int(index)),
                                null_safe: false,
                            });
                            return;
                        }
                    }
                    let storage_type = dotnet_vb::collection_storage_type(type_name);
                    let should_call = dotnet_vb::collection_property_method(storage_type, field);
                    if should_call {
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident(name)),
                                field: field.clone(),
                                null_safe: false,
                            }),
                            Vec::new(),
                        );
                    }
                }
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_dotnet_collection_expr(left, locals);
            normalize_vb_dotnet_collection_expr(right, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_dotnet_collection_expr(expr, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_dotnet_collection_expr(cond, locals);
            normalize_vb_dotnet_collection_expr(then, locals);
            normalize_vb_dotnet_collection_expr(else_, locals);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_dotnet_collection_expr(object, locals);
            normalize_vb_dotnet_collection_expr(index, locals);
            *index = Box::new(vb_normalize_dictionary_key(
                object,
                (**index).clone(),
                locals,
            ));
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_dotnet_collection_expr(&mut item.value, locals);
            }
        }
        ExprKind::New { class, args } => {
            for arg in &mut *args {
                normalize_vb_dotnet_collection_expr(&mut arg.value, locals);
            }
            if dotted_expr_name(class).as_deref().is_some_and(|name| {
                dotnet_vb::collection_base_type_name(name).eq_ignore_ascii_case("HashSet")
            }) && args.len() == 1
                && literal_string(&args[0].value)
                    .is_some_and(|text| text.starts_with("__dotnet_stringcomparer_"))
            {
                args.clear();
                return;
            }
            if dotted_expr_name(class).as_deref().is_some_and(|name| {
                dotnet_vb::collection_base_type_name(name).eq_ignore_ascii_case("KeyValuePair")
            }) && args.len() >= 2
            {
                *expr = Expression::new(ExprKind::Array(vec![
                    ArrayElement {
                        key: None,
                        value: args[0].value.clone(),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: args[1].value.clone(),
                        spread: false,
                        by_ref: false,
                    },
                ]));
            }
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(expr) => normalize_vb_dotnet_collection_expr(expr, locals),
            LambdaBody::Block(body) => {
                normalize_vb_dotnet_collection_statements(body, &mut locals.clone())
            }
        },
        _ => {}
    }
}

fn vb_new_dictionary_uses_ignorecase(expr: &Expression) -> bool {
    let ExprKind::New { class, args } = &expr.kind else {
        return false;
    };
    dotted_expr_name(class).as_deref().is_some_and(|name| {
        dotnet_vb::collection_base_type_name(name).eq_ignore_ascii_case("Dictionary")
    }) && args.iter().any(|arg| {
        literal_string(&arg.value).is_some_and(|text| {
            text.eq_ignore_ascii_case("__dotnet_stringcomparer_ordinalignorecase")
        })
    })
}

fn flatten_vb_append_chain(expr: &Expression) -> Option<Expression> {
    let (root, parts) = collect_vb_append_chain(expr)?;
    if parts.len() <= 1 {
        return None;
    }
    let mut iter = parts.into_iter();
    let mut combined = iter.next()?;
    for part in iter {
        combined = Expression::new(ExprKind::Binary {
            op: BinOp::Concat,
            left: Box::new(combined),
            right: Box::new(part),
        });
    }
    Some(call_expr(
        Expression::new(ExprKind::Member {
            object: Box::new(root),
            field: "Append".into(),
            null_safe: false,
        }),
        vec![Argument::positional(combined)],
    ))
}

fn collect_vb_append_chain(expr: &Expression) -> Option<(Expression, Vec<Expression>)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if args.len() != 1 {
        return None;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !field.eq_ignore_ascii_case("Append") {
        return None;
    }
    if let Some((root, mut parts)) = collect_vb_append_chain(object) {
        parts.push(args[0].value.clone());
        Some((root, parts))
    } else {
        Some(((**object).clone(), vec![args[0].value.clone()]))
    }
}

fn vb_normalize_dictionary_key(
    object: &Expression,
    key: Expression,
    locals: &HashMap<String, String>,
) -> Expression {
    let ExprKind::Ident(name) = &object.kind else {
        return key;
    };
    if !locals
        .get(&name.to_ascii_lowercase())
        .is_some_and(|type_name| type_name == "DictionaryIgnoreCase")
    {
        return key;
    }
    match literal_string(&key) {
        Some(text) => Expression::string(&text.to_ascii_lowercase()),
        None => call_expr(
            Expression::new(ExprKind::Member {
                object: Box::new(key),
                field: "ToLower".into(),
                null_safe: false,
            }),
            Vec::new(),
        ),
    }
}

fn vb_normalize_new_class_name(expr: &mut Expression, class_name: &str) {
    if let ExprKind::New { class, .. } = &mut expr.kind {
        *class = Box::new(Expression::ident(class_name));
    }
}

fn vb_is_zero_literal(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Int(0)))
}

fn vb_is_one_literal(expr: &Expression) -> bool {
    matches!(expr.kind, ExprKind::Lit(Literal::Int(1)))
}

fn vb_array_local_dim(object: &Expression, dim: &Expression) -> Option<(String, usize)> {
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    Some((name.to_ascii_lowercase(), usize::try_from(vb_literal_i64(dim)?).ok()?))
}

fn vb_array_index_chain(mut object: Expression, args: &[Argument]) -> Expression {
    for arg in args {
        object = Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(arg.value.clone()),
            null_safe: false,
        });
    }
    object
}

fn vb_local_is_array_like(name: &str, _argc: usize, locals: &HashMap<String, String>) -> bool {
    let key = name.to_ascii_lowercase();
    locals.contains_key(&format!("$array_rank:{key}"))
        || locals
            .get(&key)
            .is_some_and(|type_name| type_name.trim().ends_with("()") || type_name.contains("(,)"))
}

fn vb_array_index_chain_for_local(
    name: &str,
    args: &[Argument],
    locals: &HashMap<String, String>,
) -> Expression {
    let key = name.to_ascii_lowercase();
    let mut object = Expression::ident(name);
    for (dim, arg) in args.iter().enumerate() {
        let mut index = arg.value.clone();
        if let Some(lower) = locals
            .get(&format!("$array_lower:{key}:{dim}"))
            .and_then(|value| value.parse::<i64>().ok())
            .filter(|value| *value != 0)
        {
            index = Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(index),
                right: Box::new(Expression::int(lower)),
            });
        }
        object = Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        });
    }
    object
}

fn vb_system_array_member(member: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("System")),
            field: "Array".into(),
            null_safe: false,
        })),
        field: member.into(),
        null_safe: false,
    })
}

fn normalize_vb_nested_member_arg_calls(module: &mut Module) {
    normalize_vb_nested_member_arg_call_statements(&mut module.body);
}

fn normalize_vb_nested_member_arg_call_statements(body: &mut [Statement]) {
    for stmt in body {
        normalize_vb_nested_member_arg_call_statement(stmt);
    }
}

fn normalize_vb_nested_member_arg_call_statement(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_nested_member_arg_call_expr(expr);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_nested_member_arg_call_expr(init);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_nested_member_arg_call_expr(target);
            }
            normalize_vb_nested_member_arg_call_expr(value);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_nested_member_arg_call_expr(target);
            normalize_vb_nested_member_arg_call_expr(value);
        }
        StmtKind::FunctionDecl { body, .. } => {
            normalize_vb_nested_member_arg_call_statements(body);
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_nested_member_arg_call_member(member);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_nested_member_arg_call_expr(cond);
            normalize_vb_nested_member_arg_call_statements(then_body);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_nested_member_arg_call_expr(elif_cond);
                normalize_vb_nested_member_arg_call_statements(elif_body);
            }
            if let Some(else_body) = else_body {
                normalize_vb_nested_member_arg_call_statements(else_body);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_vb_nested_member_arg_call_statement(init);
            }
            if let Some(cond) = cond {
                normalize_vb_nested_member_arg_call_expr(cond);
            }
            if let Some(update) = update {
                normalize_vb_nested_member_arg_call_expr(update);
            }
            normalize_vb_nested_member_arg_call_statements(body);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_nested_member_arg_call_expr(iter);
            normalize_vb_nested_member_arg_call_statements(body);
            if let Some(else_body) = else_body {
                normalize_vb_nested_member_arg_call_statements(else_body);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_nested_member_arg_call_expr(cond);
            normalize_vb_nested_member_arg_call_statements(body);
            if let Some(else_body) = else_body {
                normalize_vb_nested_member_arg_call_statements(else_body);
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_nested_member_arg_call_statements(body);
        }
        _ => {}
    }
}

fn normalize_vb_nested_member_arg_call_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_nested_member_arg_call_statement(stmt);
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_nested_member_arg_call_statements(body);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_nested_member_arg_call_statements(getter);
            }
            if let Some(setter) = setter {
                normalize_vb_nested_member_arg_call_statements(&mut setter.body);
            }
        }
        _ => {}
    }
}

fn normalize_vb_nested_member_arg_call_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_nested_member_arg_call_expr(callee);
            for arg in &mut *args {
                normalize_vb_nested_member_arg_call_expr(&mut arg.value);
            }
            if args.is_empty() && matches!(callee.kind, ExprKind::New { .. }) {
                *expr = (**callee).clone();
                return;
            }
            if !args.is_empty() {
                if let ExprKind::Call {
                    callee: inner_callee,
                    args: inner_args,
                    ..
                } = &callee.kind
                {
                    if inner_args.is_empty() && matches!(inner_callee.kind, ExprKind::Member { .. })
                    {
                        *callee = inner_callee.clone();
                    }
                }
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_nested_member_arg_call_expr(object),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_nested_member_arg_call_expr(left);
            normalize_vb_nested_member_arg_call_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_nested_member_arg_call_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_nested_member_arg_call_expr(cond);
            normalize_vb_nested_member_arg_call_expr(then);
            normalize_vb_nested_member_arg_call_expr(else_);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_nested_member_arg_call_expr(object);
            normalize_vb_nested_member_arg_call_expr(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_nested_member_arg_call_expr(&mut item.value);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_nested_member_arg_call_expr(class);
            for arg in args {
                normalize_vb_nested_member_arg_call_expr(&mut arg.value);
            }
        }
        _ => {}
    }
}

fn normalize_vb_array_paren_indexes(module: &mut Module) {
    normalize_vb_array_paren_index_statements(&mut module.body, &mut HashSet::new());
}

fn normalize_vb_array_paren_index_statements(body: &mut [Statement], arrays: &mut HashSet<String>) {
    for stmt in body {
        normalize_vb_array_paren_index_statement(stmt, arrays);
    }
}

fn normalize_vb_array_paren_index_statement(stmt: &mut Statement, arrays: &mut HashSet<String>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_array_paren_index_expr(expr, arrays);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_array_paren_index_expr(init, arrays);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    let type_is_array = decl
                        .type_hint
                        .as_deref()
                        .is_some_and(|hint| hint.trim().ends_with("()"));
                    let type_is_string = decl
                        .type_hint
                        .as_deref()
                        .is_some_and(|hint| vb_canonical_type_name(hint) == "String");
                    if decl.array_bounds.is_some() || type_is_array || type_is_string {
                        arrays.insert(name.to_ascii_lowercase());
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_array_paren_index_expr(target, arrays);
            }
            normalize_vb_array_paren_index_expr(value, arrays);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_array_paren_index_expr(target, arrays);
            normalize_vb_array_paren_index_expr(value, arrays);
        }
        StmtKind::FunctionDecl { body, .. } => {
            normalize_vb_array_paren_index_statements(body, &mut HashSet::new());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_array_paren_index_member(member);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_array_paren_index_expr(cond, arrays);
            normalize_vb_array_paren_index_statements(then_body, &mut arrays.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_array_paren_index_expr(elif_cond, arrays);
                normalize_vb_array_paren_index_statements(elif_body, &mut arrays.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_array_paren_index_statements(else_body, &mut arrays.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_arrays = arrays.clone();
            if let Some(init) = init {
                normalize_vb_array_paren_index_statement(init, &mut loop_arrays);
            }
            if let Some(cond) = cond {
                normalize_vb_array_paren_index_expr(cond, &loop_arrays);
            }
            if let Some(update) = update {
                normalize_vb_array_paren_index_expr(update, &loop_arrays);
            }
            normalize_vb_array_paren_index_statements(body, &mut loop_arrays);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_array_paren_index_expr(iter, arrays);
            normalize_vb_array_paren_index_statements(body, &mut arrays.clone());
            if let Some(else_body) = else_body {
                normalize_vb_array_paren_index_statements(else_body, &mut arrays.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_array_paren_index_expr(cond, arrays);
            normalize_vb_array_paren_index_statements(body, &mut arrays.clone());
            if let Some(else_body) = else_body {
                normalize_vb_array_paren_index_statements(else_body, &mut arrays.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_array_paren_index_statements(body, &mut arrays.clone());
        }
        _ => {}
    }
}

fn normalize_vb_array_paren_index_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_array_paren_index_statement(stmt, &mut HashSet::new());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_array_paren_index_statements(body, &mut HashSet::new());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_array_paren_index_statements(getter, &mut HashSet::new());
            }
            if let Some(setter) = setter {
                normalize_vb_array_paren_index_statements(&mut setter.body, &mut HashSet::new());
            }
        }
        _ => {}
    }
}

fn normalize_vb_array_paren_index_expr(expr: &mut Expression, arrays: &HashSet<String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_array_paren_index_expr(callee, arrays);
            for arg in &mut *args {
                normalize_vb_array_paren_index_expr(&mut arg.value, arrays);
            }
            if args.len() == 1 {
                if let ExprKind::Ident(name) = &callee.kind {
                    if arrays.contains(&name.to_ascii_lowercase()) {
                        *expr = Expression::new(ExprKind::Index {
                            object: Box::new(Expression::ident(name)),
                            index: Box::new(args[0].value.clone()),
                            null_safe: false,
                        });
                    }
                } else if matches!(callee.kind, ExprKind::Lit(Literal::Str(_))) {
                    *expr = Expression::new(ExprKind::Index {
                        object: Box::new((**callee).clone()),
                        index: Box::new(args[0].value.clone()),
                        null_safe: false,
                    });
                }
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_array_paren_index_expr(object, arrays),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_array_paren_index_expr(left, arrays);
            normalize_vb_array_paren_index_expr(right, arrays);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_array_paren_index_expr(expr, arrays),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_array_paren_index_expr(cond, arrays);
            normalize_vb_array_paren_index_expr(then, arrays);
            normalize_vb_array_paren_index_expr(else_, arrays);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_array_paren_index_expr(object, arrays);
            normalize_vb_array_paren_index_expr(index, arrays);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_array_paren_index_expr(&mut item.value, arrays);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_array_paren_index_expr(class, arrays);
            for arg in args {
                normalize_vb_array_paren_index_expr(&mut arg.value, arrays);
            }
        }
        _ => {}
    }
}

fn normalize_vb_default_indexer_calls(module: &mut Module) {
    let mut default_indexer_types = HashMap::new();
    for name in [
        "array",
        "arraylist",
        "collection",
        "dictionary",
        "idictionary",
        "ilist",
        "list",
        "observablecollection",
        "readonlycollection",
        "readonlylist",
        "sorteddictionary",
    ] {
        default_indexer_types.insert(name.to_string(), true);
    }
    collect_vb_default_indexer_types(&module.body, &mut default_indexer_types);
    let mut globals = HashMap::new();
    collect_vb_default_indexer_static_field_types(&module.body, &mut globals);
    rewrite_vb_default_indexer_statements(&mut module.body, &default_indexer_types, &mut globals);
}

fn collect_vb_default_indexer_types(body: &[Statement], types: &mut HashMap<String, bool>) {
    let mut inherited = Vec::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl {
                name,
                parents,
                members,
                ..
            } => {
                if members.iter().any(|member| {
                    matches!(
                        member,
                        ClassMember::Method(method)
                            if matches!(
                                &method.kind,
                                StmtKind::FunctionDecl { name, .. }
                                    if name.eq_ignore_ascii_case("__getitem__")
                            )
                    )
                }) {
                    types.insert(vb_canonical_type_name(name).to_ascii_lowercase(), true);
                } else if parents.iter().any(|parent| {
                    types.contains_key(&vb_default_indexer_type_key(parent))
                        || inherited
                            .iter()
                            .any(|(child, ancestor): &(String, String)| {
                                child.eq_ignore_ascii_case(parent)
                                    && types.contains_key(&vb_default_indexer_type_key(ancestor))
                            })
                }) {
                    types.insert(vb_canonical_type_name(name).to_ascii_lowercase(), true);
                } else if let Some(parent) = parents.first() {
                    inherited.push((name.clone(), parent.clone()));
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_default_indexer_types(std::slice::from_ref(nested), types);
                    }
                }
            }
            StmtKind::StructDecl { name, members, .. } => {
                if members.iter().any(|member| {
                    matches!(
                        member,
                        ClassMember::Method(method)
                            if matches!(
                                &method.kind,
                                StmtKind::FunctionDecl { name, .. }
                                    if name.eq_ignore_ascii_case("__getitem__")
                            )
                    )
                }) {
                    types.insert(vb_canonical_type_name(name).to_ascii_lowercase(), true);
                }
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_default_indexer_types(std::slice::from_ref(nested), types);
                    }
                }
            }
            StmtKind::InterfaceDecl { name, members, .. } => {
                if members.iter().any(|member| {
                    matches!(
                        member,
                        InterfaceMember::Property { name, .. }
                            if name.eq_ignore_ascii_case("Item")
                    )
                }) {
                    types.insert(vb_canonical_type_name(name).to_ascii_lowercase(), true);
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_default_indexer_types(std::slice::from_ref(nested), types);
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                collect_vb_default_indexer_types(body, types);
            }
            _ => {}
        }
    }
    let mut changed = true;
    while changed {
        changed = false;
        for (child, parent) in &inherited {
            if !types.contains_key(&vb_default_indexer_type_key(child))
                && types.contains_key(&vb_default_indexer_type_key(parent))
            {
                types.insert(vb_canonical_type_name(child).to_ascii_lowercase(), true);
                changed = true;
            }
        }
    }
}

fn collect_vb_default_indexer_static_field_types(
    body: &[Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        match &stmt.kind {
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. } => {
                collect_vb_default_indexer_static_fields(name, members, locals);
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_default_indexer_static_field_types(
                            std::slice::from_ref(nested),
                            locals,
                        );
                    }
                }
            }
            StmtKind::ModuleDecl { members, .. } => {
                for member in members {
                    if let ClassMember::NestedType(nested) = member {
                        collect_vb_default_indexer_static_field_types(
                            std::slice::from_ref(nested),
                            locals,
                        );
                    }
                }
            }
            StmtKind::NamespaceDecl { body, .. } | StmtKind::Block(body) => {
                collect_vb_default_indexer_static_field_types(body, locals);
            }
            _ => {}
        }
    }
}

fn vb_default_indexer_type_key(type_name: &str) -> String {
    strip_vb_generic_suffix(type_name)
        .rsplit('.')
        .next()
        .unwrap_or(type_name)
        .to_ascii_lowercase()
}

fn vb_default_indexer_local_type(type_hint: &str) -> String {
    let canonical = vb_canonical_type_name(type_hint);
    if type_hint.trim().ends_with("()") && !canonical.trim().ends_with("()") {
        format!("{canonical}()")
    } else {
        canonical
    }
}

fn rewrite_vb_default_indexer_statements(
    body: &mut [Statement],
    default_indexer_types: &HashMap<String, bool>,
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        rewrite_vb_default_indexer_statement(stmt, default_indexer_types, locals);
    }
}

fn rewrite_vb_default_indexer_statement(
    stmt: &mut Statement,
    default_indexer_types: &HashMap<String, bool>,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_default_indexer_expr(expr, default_indexer_types, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_default_indexer_expr(init, default_indexer_types, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(type_hint) = &decl.type_hint {
                        locals.insert(
                            name.to_ascii_lowercase(),
                            vb_default_indexer_local_type(type_hint),
                        );
                    } else if let Some(init) = &decl.init {
                        if let Some(type_name) = vb_infer_expr_type(init, locals) {
                            locals.insert(name.to_ascii_lowercase(), type_name);
                        }
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_default_indexer_target_expr(target, default_indexer_types, locals);
            }
            rewrite_vb_default_indexer_expr(value, default_indexer_types, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_vb_default_indexer_target_expr(target, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(value, default_indexer_types, locals);
        }
        StmtKind::FunctionDecl { body, params, .. } => {
            let mut fn_locals = locals.clone();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    fn_locals.insert(
                        param.name.to_ascii_lowercase(),
                        vb_default_indexer_local_type(type_hint),
                    );
                }
            }
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut fn_locals);
        }
        StmtKind::ClassDecl { name, members, .. } | StmtKind::StructDecl { name, members, .. } => {
            let mut member_locals = locals.clone();
            collect_vb_default_indexer_static_fields(name, members, &mut member_locals);
            for member in members {
                rewrite_vb_default_indexer_member(member, default_indexer_types, &member_locals);
            }
        }
        StmtKind::ModuleDecl { members, .. } => {
            let member_locals = locals.clone();
            for member in members {
                rewrite_vb_default_indexer_member(member, default_indexer_types, &member_locals);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_default_indexer_expr(cond, default_indexer_types, locals);
            rewrite_vb_default_indexer_statements(
                then_body,
                default_indexer_types,
                &mut locals.clone(),
            );
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_default_indexer_expr(elif_cond, default_indexer_types, locals);
                rewrite_vb_default_indexer_statements(
                    elif_body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                rewrite_vb_default_indexer_statements(
                    else_body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                rewrite_vb_default_indexer_statement(init, default_indexer_types, &mut loop_locals);
            }
            if let Some(cond) = cond {
                rewrite_vb_default_indexer_expr(cond, default_indexer_types, &loop_locals);
            }
            if let Some(update) = update {
                rewrite_vb_default_indexer_expr(update, default_indexer_types, &loop_locals);
            }
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_default_indexer_expr(iter, default_indexer_types, locals);
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_default_indexer_statements(
                    else_body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_default_indexer_expr(cond, default_indexer_types, locals);
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut locals.clone());
            if let Some(else_body) = else_body {
                rewrite_vb_default_indexer_statements(
                    else_body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut locals.clone());
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            else_body,
        } => {
            rewrite_vb_default_indexer_statements(body, default_indexer_types, &mut locals.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    rewrite_vb_default_indexer_expr(when_clause, default_indexer_types, locals);
                }
                rewrite_vb_default_indexer_statements(
                    &mut catch.body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                rewrite_vb_default_indexer_statements(
                    else_body,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
            if let Some(finally) = finally {
                rewrite_vb_default_indexer_statements(
                    finally,
                    default_indexer_types,
                    &mut locals.clone(),
                );
            }
        }
        _ => {}
    }
}

fn rewrite_vb_default_indexer_target_expr(
    expr: &mut Expression,
    default_indexer_types: &HashMap<String, bool>,
    locals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_default_indexer_target_expr(callee, default_indexer_types, locals);
            for arg in args {
                rewrite_vb_default_indexer_expr(&mut arg.value, default_indexer_types, locals);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_vb_default_indexer_target_expr(object, default_indexer_types, locals);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_default_indexer_target_expr(object, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(index, default_indexer_types, locals);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_default_indexer_target_expr(left, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(right, default_indexer_types, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_vb_default_indexer_target_expr(expr, default_indexer_types, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_default_indexer_expr(cond, default_indexer_types, locals);
            rewrite_vb_default_indexer_target_expr(then, default_indexer_types, locals);
            rewrite_vb_default_indexer_target_expr(else_, default_indexer_types, locals);
        }
        _ => rewrite_vb_default_indexer_expr(expr, default_indexer_types, locals),
    }
}

fn rewrite_vb_default_indexer_member(
    member: &mut ClassMember,
    default_indexer_types: &HashMap<String, bool>,
    outer_locals: &HashMap<String, String>,
) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            rewrite_vb_default_indexer_statement(
                stmt,
                default_indexer_types,
                &mut outer_locals.clone(),
            );
        }
        ClassMember::Constructor { body, .. } => {
            rewrite_vb_default_indexer_statements(
                body,
                default_indexer_types,
                &mut outer_locals.clone(),
            );
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                rewrite_vb_default_indexer_statements(
                    getter,
                    default_indexer_types,
                    &mut outer_locals.clone(),
                );
            }
            if let Some(setter) = setter {
                rewrite_vb_default_indexer_statements(
                    &mut setter.body,
                    default_indexer_types,
                    &mut outer_locals.clone(),
                );
            }
        }
        _ => {}
    }
}

fn collect_vb_default_indexer_static_fields(
    owner_name: &str,
    members: &[ClassMember],
    locals: &mut HashMap<String, String>,
) {
    for member in members {
        if let ClassMember::Field {
            name,
            type_hint: Some(type_hint),
            modifiers,
            ..
        } = member
        {
            if modifiers.is_static || modifiers.is_shared {
                locals.insert(
                    format!("{}.{}", owner_name, name).to_ascii_lowercase(),
                    vb_canonical_type_name(type_hint),
                );
            }
        }
    }
}

fn normalize_vb_stringbuilder_member_access(module: &mut Module) {
    normalize_vb_stringbuilder_member_statements(&mut module.body, &mut HashMap::new());
}

fn normalize_vb_stringbuilder_member_statements(
    body: &mut [Statement],
    locals: &mut HashMap<String, String>,
) {
    for stmt in body {
        normalize_vb_stringbuilder_member_statement(stmt, locals);
    }
}

fn normalize_vb_stringbuilder_member_statement(
    stmt: &mut Statement,
    locals: &mut HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_stringbuilder_member_expr(expr, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_stringbuilder_member_expr(init, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(type_hint) = &decl.type_hint {
                        locals.insert(name.to_ascii_lowercase(), vb_canonical_type_name(type_hint));
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_stringbuilder_member_expr(target, locals);
            }
            normalize_vb_stringbuilder_member_expr(value, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_stringbuilder_member_expr(target, locals);
            normalize_vb_stringbuilder_member_expr(value, locals);
        }
        StmtKind::FunctionDecl { body, .. } => {
            normalize_vb_stringbuilder_member_statements(body, &mut HashMap::new());
        }
        StmtKind::ClassDecl { members, .. }
        | StmtKind::StructDecl { members, .. }
        | StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_stringbuilder_member_class_member(member);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_stringbuilder_member_expr(cond, locals);
            normalize_vb_stringbuilder_member_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_stringbuilder_member_expr(elif_cond, locals);
                normalize_vb_stringbuilder_member_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_stringbuilder_member_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_stringbuilder_member_statement(init, &mut loop_locals);
            }
            if let Some(cond) = cond {
                normalize_vb_stringbuilder_member_expr(cond, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_stringbuilder_member_expr(update, &loop_locals);
            }
            normalize_vb_stringbuilder_member_statements(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_stringbuilder_member_expr(iter, locals);
            normalize_vb_stringbuilder_member_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_stringbuilder_member_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_stringbuilder_member_expr(cond, locals);
            normalize_vb_stringbuilder_member_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_stringbuilder_member_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::Block(body) | StmtKind::NamespaceDecl { body, .. } => {
            normalize_vb_stringbuilder_member_statements(body, &mut locals.clone());
        }
        _ => {}
    }
}

fn normalize_vb_stringbuilder_member_class_member(member: &mut ClassMember) {
    match member {
        ClassMember::Method(stmt) | ClassMember::NestedType(stmt) => {
            normalize_vb_stringbuilder_member_statement(stmt, &mut HashMap::new());
        }
        ClassMember::Constructor { body, .. } => {
            normalize_vb_stringbuilder_member_statements(body, &mut HashMap::new());
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_stringbuilder_member_statements(getter, &mut HashMap::new());
            }
            if let Some(setter) = setter {
                normalize_vb_stringbuilder_member_statements(&mut setter.body, &mut HashMap::new());
            }
        }
        _ => {}
    }
}

fn normalize_vb_stringbuilder_member_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_stringbuilder_member_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_stringbuilder_member_expr(&mut arg.value, locals);
            }
            if args.len() == 1 {
                if let ExprKind::Member { object, field, .. } = &callee.kind {
                    if field.eq_ignore_ascii_case("Chars")
                        && vb_expr_is_stringbuilder_local(object, locals)
                    {
                        *expr = Expression::new(ExprKind::Index {
                            object: object.clone(),
                            index: Box::new(args[0].value.clone()),
                            null_safe: false,
                        });
                    }
                }
            }
        }
        ExprKind::Member { object, field, .. } => {
            normalize_vb_stringbuilder_member_expr(object, locals);
            if matches!(
                field.to_ascii_lowercase().as_str(),
                "length" | "capacity" | "maxcapacity"
            ) && vb_expr_is_stringbuilder_local(object, locals)
            {
                let original = std::mem::replace(object, Box::new(Expression::null()));
                *object = Box::new(Expression::new(ExprKind::Cast {
                    expr: original,
                    type_name: "StringBuilder".to_string(),
                }));
            }
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_stringbuilder_member_expr(object, locals);
            normalize_vb_stringbuilder_member_expr(index, locals);
            if let ExprKind::Member {
                object: receiver,
                field,
                ..
            } = &object.kind
            {
                if field.eq_ignore_ascii_case("Chars")
                    && vb_expr_is_stringbuilder_local(receiver, locals)
                {
                    *object = receiver.clone();
                }
            }
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_vb_stringbuilder_member_expr(left, locals);
            normalize_vb_stringbuilder_member_expr(right, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_stringbuilder_member_expr(expr, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_stringbuilder_member_expr(cond, locals);
            normalize_vb_stringbuilder_member_expr(then, locals);
            normalize_vb_stringbuilder_member_expr(else_, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_stringbuilder_member_expr(&mut item.value, locals);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                normalize_vb_stringbuilder_member_expr(item, locals);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_stringbuilder_member_expr(class, locals);
            for arg in args {
                normalize_vb_stringbuilder_member_expr(&mut arg.value, locals);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_stringbuilder_member_expr(target, locals);
            normalize_vb_stringbuilder_member_expr(value, locals);
        }
        ExprKind::Lambda { body, .. } => {
            if let LambdaBody::Expr(expr) = body {
                normalize_vb_stringbuilder_member_expr(expr, locals);
            }
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        normalize_vb_stringbuilder_member_expr(expr, locals);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            normalize_vb_stringbuilder_member_expr(left, locals);
            normalize_vb_stringbuilder_member_expr(right, locals);
        }
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Member { object, .. } => {
                normalize_vb_stringbuilder_member_expr(object, locals);
            }
            PlaceExpr::Index { object, index, .. } => {
                normalize_vb_stringbuilder_member_expr(object, locals);
                normalize_vb_stringbuilder_member_expr(index, locals);
            }
            PlaceExpr::Deref(expr) => normalize_vb_stringbuilder_member_expr(expr, locals),
            PlaceExpr::Ident(_) => {}
        },
        _ => {}
    }
}

fn vb_expr_is_stringbuilder_local(expr: &Expression, locals: &HashMap<String, String>) -> bool {
    let ExprKind::Ident(name) = &expr.kind else {
        return false;
    };
    locals
        .get(&name.to_ascii_lowercase())
        .map(|ty| {
            ty.rsplit('.')
                .next()
                .unwrap_or(ty)
                .eq_ignore_ascii_case("StringBuilder")
        })
        .unwrap_or(false)
}

fn rewrite_vb_default_indexer_expr(
    expr: &mut Expression,
    default_indexer_types: &HashMap<String, bool>,
    locals: &HashMap<String, String>,
) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_default_indexer_expr(callee, default_indexer_types, locals);
            for arg in &mut *args {
                rewrite_vb_default_indexer_expr(&mut arg.value, default_indexer_types, locals);
            }
            if !args.is_empty() {
                if args.len() == 1 {
                    if let ExprKind::Member { object, field, .. } = &callee.kind {
                        if field.eq_ignore_ascii_case("Groups") {
                            let groups = Expression::new(ExprKind::Member {
                                object: object.clone(),
                                field: "__groups".to_string(),
                                null_safe: false,
                            });
                            if let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind {
                                *expr = Expression::new(ExprKind::Member {
                                    object: Box::new(groups),
                                    field: name.to_string(),
                                    null_safe: false,
                                });
                                return;
                            }
                            *expr = Expression::new(ExprKind::Index {
                                object: Box::new(groups),
                                index: Box::new(args[0].value.clone()),
                                null_safe: false,
                            });
                            return;
                        }
                    }
                }
                if let ExprKind::Ident(name) = &callee.kind {
                    if let Some(type_name) = locals.get(&name.to_ascii_lowercase()) {
                        let base_type = dotnet_vb::collection_base_type_name(type_name);
                        if type_name.trim().ends_with("()")
                            || base_type.eq_ignore_ascii_case("Array")
                        {
                            *expr = call_expr(
                                Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident(name)),
                                    field: "GetValue".to_string(),
                                    null_safe: false,
                                }),
                                vec![Argument::positional(args[0].value.clone())],
                            );
                            return;
                        }
                        if vb_canonical_type_name(type_name).eq_ignore_ascii_case("String") {
                            *expr = call_expr(
                                Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::ident(name)),
                                    field: "Chars".to_string(),
                                    null_safe: false,
                                }),
                                vec![Argument::positional(args[0].value.clone())],
                            );
                            return;
                        }
                        if default_indexer_types
                            .contains_key(&vb_default_indexer_type_key(type_name))
                        {
                            let field = if dotnet_vb::collection_local_type(type_name).is_some()
                                || base_type.eq_ignore_ascii_case("Dictionary")
                            {
                                "Item"
                            } else {
                                "__getitem__"
                            };
                            *callee = Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(Expression::ident(name)),
                                field: field.to_string(),
                                null_safe: false,
                            }));
                        }
                    }
                } else if let Some(name) = dotted_expr_name(callee) {
                    if let Some(type_name) = locals.get(&name.to_ascii_lowercase()) {
                        if default_indexer_types
                            .contains_key(&vb_default_indexer_type_key(type_name))
                        {
                            *expr = Expression::new(ExprKind::Index {
                                object: Box::new((**callee).clone()),
                                index: Box::new(args[0].value.clone()),
                                null_safe: false,
                            });
                        }
                    }
                }
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_vb_default_indexer_expr(object, default_indexer_types, locals);
            let mut replacement = None;
            if let ExprKind::Member { object, field, .. } = &expr.kind {
                if field.eq_ignore_ascii_case("Value") {
                    if let ExprKind::Index {
                        object: groups_object,
                        index,
                        null_safe,
                    } = &object.kind
                    {
                        if let ExprKind::Member {
                            object: match_object,
                            field: groups_field,
                            ..
                        } = &groups_object.kind
                        {
                            if groups_field == "__groups" {
                                replacement = Some(Expression::new(ExprKind::Index {
                                    object: Box::new(Expression::new(ExprKind::Member {
                                        object: match_object.clone(),
                                        field: "__group_values".to_string(),
                                        null_safe: false,
                                    })),
                                    index: index.clone(),
                                    null_safe: *null_safe,
                                }));
                            }
                        }
                    }
                    if let ExprKind::Member {
                        object: groups_object,
                        field: group_name,
                        null_safe,
                    } = &object.kind
                    {
                        if let ExprKind::Member {
                            object: match_object,
                            field: groups_field,
                            ..
                        } = &groups_object.kind
                        {
                            if groups_field == "__groups" {
                                replacement = Some(Expression::new(ExprKind::Member {
                                    object: Box::new(Expression::new(ExprKind::Member {
                                        object: match_object.clone(),
                                        field: "__group_values".to_string(),
                                        null_safe: false,
                                    })),
                                    field: group_name.clone(),
                                    null_safe: *null_safe,
                                }));
                            }
                        }
                    }
                }
            }
            if let Some(replacement) = replacement {
                *expr = replacement;
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_default_indexer_expr(left, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(right, default_indexer_types, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => {
            rewrite_vb_default_indexer_expr(expr, default_indexer_types, locals)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_default_indexer_expr(cond, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(then, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(else_, default_indexer_types, locals);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_default_indexer_expr(object, default_indexer_types, locals);
            rewrite_vb_default_indexer_expr(index, default_indexer_types, locals);
            if let Some(name) = dotted_expr_name(object) {
                if let Some(type_name) = locals.get(&name.to_ascii_lowercase()) {
                    let base_type = dotnet_vb::collection_base_type_name(type_name);
                    if type_name.trim().ends_with("()") || base_type.eq_ignore_ascii_case("Array") {
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "GetValue".to_string(),
                                null_safe: false,
                            }),
                            vec![Argument::positional((**index).clone())],
                        );
                        return;
                    }
                    if vb_canonical_type_name(type_name).eq_ignore_ascii_case("String") {
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "Chars".to_string(),
                                null_safe: false,
                            }),
                            vec![Argument::positional((**index).clone())],
                        );
                        return;
                    }
                    if default_indexer_types.contains_key(&vb_default_indexer_type_key(type_name))
                        && base_type.eq_ignore_ascii_case("Dictionary")
                    {
                        *expr = call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new((**object).clone()),
                                field: "Item".to_string(),
                                null_safe: false,
                            }),
                            vec![Argument::positional((**index).clone())],
                        );
                        return;
                    }
                }
            }
            if let ExprKind::Member {
                object: match_object,
                field,
                ..
            } = &object.kind
            {
                if field.eq_ignore_ascii_case("Groups") {
                    let groups = Expression::new(ExprKind::Member {
                        object: match_object.clone(),
                        field: "__groups".to_string(),
                        null_safe: false,
                    });
                    if let ExprKind::Lit(Literal::Str(name)) = &index.kind {
                        *expr = Expression::new(ExprKind::Member {
                            object: Box::new(groups),
                            field: name.to_string(),
                            null_safe: false,
                        });
                    } else {
                        *expr = Expression::new(ExprKind::Index {
                            object: Box::new(groups),
                            index: index.clone(),
                            null_safe: false,
                        });
                    }
                }
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_default_indexer_expr(&mut item.value, default_indexer_types, locals);
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                rewrite_vb_default_indexer_expr(&mut arg.value, default_indexer_types, locals);
            }
        }
        _ => {}
    }
}

fn clear_vb_known_local_value(locals: &mut HashMap<String, String>, name: &str) {
    let key = name.to_ascii_lowercase();
    for prefix in [
        "$value:",
        "$bool:",
        "$string:",
        "$decimal:",
        "$nullstring:",
        "$regex_pattern:",
    ] {
        locals.remove(&format!("{prefix}{key}"));
    }
}

fn vb_for_init_decl_names(init: &Statement) -> Vec<String> {
    match &init.kind {
        StmtKind::VarDecl { declarations, .. } => declarations
            .iter()
            .filter_map(|decl| match &decl.pattern {
                BindingPattern::Ident(name) => Some(name.clone()),
                _ => None,
            })
            .collect(),
        _ => Vec::new(),
    }
}

fn normalize_vb_local_type_statement(stmt: &mut Statement, locals: &mut HashMap<String, String>) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => {
            if let Some((array, new_bound)) = lower_vb_array_resize_expr_stmt(expr) {
                stmt.kind = StmtKind::ReDim {
                    preserve: true,
                    array,
                    bounds: vec![new_bound],
                };
                normalize_vb_local_type_statement(stmt, locals);
                return;
            }
            normalize_vb_local_type_expr(expr, locals);
            record_vb_local_assignment_expr(expr, locals);
        }
        StmtKind::Return(Some(expr)) => {
            normalize_vb_local_type_expr(expr, locals);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                let init_type_before_normalize =
                    decl.init.as_ref().and_then(|init| match &init.kind {
                        ExprKind::Cast { type_name, .. } => Some(vb_canonical_type_name(
                            type_name.split(':').next_back().unwrap_or(type_name),
                        )),
                        _ => None,
                    });
                if let Some(init) = &mut decl.init {
                    normalize_vb_local_type_expr(init, locals);
                    vb_apply_known_local_value(init, locals);
                    if let Some(type_hint) = &decl.type_hint {
                        vb_coerce_literal_to_type(init, type_hint);
                    }
                } else if let (Some(type_hint), Some(bounds)) =
                    (decl.type_hint.as_ref(), decl.array_bounds.as_ref())
                {
                    let default_value = vb_default_value_for_type(type_hint);
                    decl.init = Some(if bounds.len() == 1 {
                        vb_filled_array_expr(
                            vb_array_length_from_upper_bound(bounds[0].clone()),
                            default_value,
                        )
                    } else {
                        vb_multidim_array_expr(bounds, default_value)
                    });
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    if let Some(bounds) = decl.array_bounds.as_ref() {
                        record_vb_array_bounds_metadata(locals, name, bounds);
                    }
                    let key = name.to_ascii_lowercase();
                    if let Some(init) = &decl.init {
                        if let Some(value) = literal_number(init) {
                            locals.insert(format!("$value:{key}"), value.to_string());
                        }
                        if let ExprKind::Lit(Literal::Bool(value)) = &init.kind {
                            locals.insert(
                                format!("$bool:{key}"),
                                if *value { "true" } else { "false" }.into(),
                            );
                        }
                        if let ExprKind::Lit(Literal::Str(value)) = &init.kind {
                            locals.insert(format!("$string:{key}"), value.clone());
                        }
                        if let Some(pattern) = vb_regex_literal_pattern(init, locals) {
                            locals.insert(format!("$regex_pattern:{key}"), pattern);
                        }
                        if matches!(init.kind, ExprKind::Lit(Literal::Null))
                            && decl
                                .type_hint
                                .as_ref()
                                .is_some_and(|ty| vb_canonical_type_name(ty) == "String")
                        {
                            locals.insert(format!("$nullstring:{key}"), "true".into());
                        }
                    }
                    if let Some(type_hint) = &decl.type_hint {
                        let canonical = vb_canonical_type_name(type_hint);
                        if canonical == "Decimal" {
                            if let Some(init) = &decl.init {
                                if let Some(value) = literal_number(init) {
                                    locals.insert(format!("$value:{key}"), value.to_string());
                                }
                                locals.insert(format!("$decimal:{key}"), "true".into());
                            }
                        }
                        locals.insert(key, canonical);
                    } else if let Some(init) = &decl.init {
                        if let Some(type_name) = init_type_before_normalize
                            .clone()
                            .or_else(|| vb_infer_expr_type(init, locals))
                        {
                            locals.insert(name.to_ascii_lowercase(), type_name);
                        }
                        if let Some(element_type) = vb_infer_array_element_type(init, locals) {
                            locals.insert(format!("$element:{key}"), element_type);
                        }
                    }
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            normalize_vb_local_type_expr(value, locals);
            for target in targets {
                normalize_vb_local_type_expr(target, locals);
                if let ExprKind::Ident(name) = &target.kind {
                    if let Some(type_name) = locals.get(&name.to_ascii_lowercase()).cloned() {
                        vb_coerce_literal_to_type(value, &type_name);
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
            normalize_vb_local_type_expr(cond, locals);
            normalize_vb_local_type_statements(then_body, &mut locals.clone());
            for (elif_cond, elif_body) in elifs {
                normalize_vb_local_type_expr(elif_cond, locals);
                normalize_vb_local_type_statements(elif_body, &mut locals.clone());
            }
            if let Some(else_body) = else_body {
                normalize_vb_local_type_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_local_type_statement(init, &mut loop_locals);
                for name in vb_for_init_decl_names(init) {
                    clear_vb_known_local_value(&mut loop_locals, &name);
                }
            }
            if let Some(cond) = cond {
                normalize_vb_local_type_expr(cond, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_local_type_expr(update, &loop_locals);
            }
            normalize_vb_local_type_statements(body, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_local_type_expr(iter, locals);
            normalize_vb_local_type_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_local_type_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_local_type_expr(cond, locals);
            normalize_vb_local_type_statements(body, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_local_type_statements(else_body, &mut locals.clone());
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            normalize_vb_local_type_statements(body, &mut locals.clone());
            normalize_vb_local_type_expr(cond, locals);
        }
        StmtKind::ReDim {
            array,
            bounds,
            preserve,
        } => {
            for bound in &mut *bounds {
                normalize_vb_local_type_expr(bound, locals);
            }
            if !*preserve && !bounds.is_empty() {
                if let Some(type_name) = locals.get(&array.to_ascii_lowercase()).cloned() {
                    let default_value = vb_default_value_for_type(&type_name);
                    let value = if bounds.len() == 1 {
                        let length = vb_array_length_from_upper_bound(bounds[0].clone());
                        vb_filled_array_expr(length, default_value)
                    } else {
                        vb_multidim_array_expr(bounds, default_value)
                    };
                    stmt.kind = StmtKind::Assign {
                        targets: vec![Expression::ident(array)],
                        value,
                    };
                }
            }
        }
        StmtKind::Block(body) | StmtKind::Lock { body, .. } => {
            normalize_vb_local_type_statements(body, &mut locals.clone());
        }
        StmtKind::Using {
            var,
            resource,
            body,
        } => {
            normalize_vb_local_type_expr(resource, locals);
            let mut using_locals = locals.clone();
            if let Some(type_name) = vb_infer_expr_type(resource, locals) {
                using_locals.insert(var.to_ascii_lowercase(), type_name);
            }
            normalize_vb_local_type_statements(body, &mut using_locals);
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            normalize_vb_local_type_statements(body, &mut locals.clone());
            for catch in catches {
                normalize_vb_local_type_statements(&mut catch.body, &mut locals.clone());
            }
            if let Some(finally) = finally {
                normalize_vb_local_type_statements(finally, &mut locals.clone());
            }
        }
        _ => {}
    }
}

fn lower_vb_array_resize_expr_stmt(expr: &Expression) -> Option<(String, Expression)> {
    let ExprKind::Call {
        callee,
        args,
        optional: false,
    } = &expr.kind
    else {
        return None;
    };
    if args.len() != 2 {
        return None;
    }
    let array = match &args[0].value.kind {
        ExprKind::Ident(name) => name.clone(),
        _ => return None,
    };
    let ExprKind::Member {
        object,
        field,
        null_safe: false,
    } = &callee.kind
    else {
        return None;
    };
    if !field.eq_ignore_ascii_case("Resize") {
        return None;
    }
    let is_array_resize = dotted_expr_name(object).is_some_and(|path| {
        matches!(path.to_ascii_lowercase().as_str(), "array" | "system.array")
    });
    if !is_array_resize {
        return None;
    }
    Some((
        array,
        Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(args[1].value.clone()),
            right: Box::new(Expression::int(1)),
        }),
    ))
}

fn record_vb_local_assignment_expr(expr: &Expression, locals: &mut HashMap<String, String>) {
    match &expr.kind {
        ExprKind::Assign { target, value } => {
            if let ExprKind::Ident(name) = &target.kind {
                let key = name.to_ascii_lowercase();
                if let Some(number) = literal_number(value) {
                    locals.insert(format!("$value:{key}"), number.to_string());
                }
                if let ExprKind::Lit(Literal::Bool(value)) = &value.kind {
                    locals.insert(
                        format!("$bool:{key}"),
                        if *value { "true" } else { "false" }.into(),
                    );
                }
                if let ExprKind::Lit(Literal::Str(value)) = &value.kind {
                    locals.insert(format!("$string:{key}"), value.clone());
                }
            }
            record_vb_local_assignment_expr(value, locals);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            record_vb_local_assignment_expr(left, locals);
            record_vb_local_assignment_expr(right, locals);
        }
        ExprKind::Call { callee, args, .. } => {
            record_vb_local_assignment_expr(callee, locals);
            for arg in args {
                record_vb_local_assignment_expr(&arg.value, locals);
            }
        }
        ExprKind::Unary { expr, .. } | ExprKind::Cast { expr, .. } => {
            record_vb_local_assignment_expr(expr, locals);
        }
        _ => {}
    }
}

fn vb_infer_array_element_type(
    expr: &Expression,
    locals: &HashMap<String, String>,
) -> Option<String> {
    let ExprKind::Array(items) = &expr.kind else {
        return None;
    };
    let mut inferred: Option<String> = None;
    for item in items {
        let item_type = vb_infer_expr_type(&item.value, locals)?;
        if let Some(existing) = &inferred {
            if !existing.eq_ignore_ascii_case(&item_type) {
                return None;
            }
        } else {
            inferred = Some(item_type);
        }
    }
    inferred
}

fn fold_vb_null_string_eq(
    op: BinOp,
    left: &Expression,
    right: &Expression,
    locals: &HashMap<String, String>,
) -> Option<Expression> {
    fn is_null_string(expr: &Expression, locals: &HashMap<String, String>) -> bool {
        matches!(
            expr.kind,
            ExprKind::Ident(ref name)
                if locals.contains_key(&format!("$nullstring:{}", name.to_ascii_lowercase()))
        )
    }
    fn is_empty_string(expr: &Expression) -> bool {
        matches!(expr.kind, ExprKind::Lit(Literal::Str(ref value)) if value.is_empty())
    }
    let equal = (is_null_string(left, locals) && is_empty_string(right))
        || (is_null_string(right, locals) && is_empty_string(left));
    if equal {
        Some(Expression::bool(matches!(op, BinOp::Eq)))
    } else {
        None
    }
}

fn normalize_vb_local_type_expr(expr: &mut Expression, locals: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Ident(_) => {}
        ExprKind::Member {
            object,
            field,
            null_safe: false,
        } => {
            if field.eq_ignore_ascii_case("GetType") {
                return;
            }
            normalize_vb_local_type_expr(object, locals);
            if field.eq_ignore_ascii_case("HasValue") {
                *expr = Expression::new(ExprKind::Binary {
                    op: BinOp::IsNot,
                    left: Box::new((**object).clone()),
                    right: Box::new(Expression::null()),
                });
                return;
            }
            let mut replacement = None;
            if field.eq_ignore_ascii_case("Name") {
                if let ExprKind::Call { callee, args, .. } = &object.kind {
                    if args.is_empty() {
                        if let ExprKind::Member {
                            object: gettype_object,
                            field: gettype_field,
                            ..
                        } = &callee.kind
                        {
                            if gettype_field.eq_ignore_ascii_case("GetType") {
                                if let Some(type_name) = vb_infer_expr_type(gettype_object, locals)
                                {
                                    replacement = Some(Expression::string(&type_name));
                                }
                            }
                        }
                    }
                }
            }
            if replacement.is_none() {
                if let Some(path) = dotted_expr_name(object) {
                    if matches!(
                        path.to_ascii_lowercase().as_str(),
                        "double" | "system.double" | "single" | "system.single"
                    ) {
                        replacement = Some(canonicalize_member_access((**object).clone(), field));
                    }
                }
            }
            if let Some(replacement) = replacement {
                *expr = replacement;
            }
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_local_type_expr(callee, locals);
            for arg in &mut *args {
                normalize_vb_local_type_expr(&mut arg.value, locals);
            }
            if args.len() >= 2
                && dotted_expr_name(callee).as_deref().is_some_and(|name| {
                    name.eq_ignore_ascii_case("System.Text.RegularExpressions.Regex.IsMatch")
                })
            {
                if let ExprKind::Cast { expr, type_name } = &args[1].value.kind {
                    if type_name == "__vb_like_pattern" {
                        args[1].value = if let Some(pattern) = vb_known_string_value(expr, locals) {
                            Expression::string(&vb_like_pattern_to_regex(&pattern))
                        } else {
                            (**expr).clone()
                        };
                    }
                }
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if matches!(
                    field.to_ascii_lowercase().as_str(),
                    "all" | "any" | "where" | "select"
                ) && !args.is_empty()
                    && matches!(
                        &object.kind,
                        ExprKind::Ident(name)
                            if locals
                                .get(&name.to_ascii_lowercase())
                                .is_some_and(|type_name| dotnet_vb::collection_type_is_dictionary(type_name))
                    )
                {
                    let dict_object = (**object).clone();
                    *callee = Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(call_expr(
                            Expression::new(ExprKind::Member {
                                object: Box::new(dict_object),
                                field: "Entries".into(),
                                null_safe: false,
                            }),
                            Vec::new(),
                        )),
                        field: field.clone(),
                        null_safe: false,
                    }));
                    if let ExprKind::Lambda { params, body, .. } = &mut args[0].value.kind {
                        if let Some(param) = params.first() {
                            let mut lambda_locals = locals.clone();
                            lambda_locals
                                .insert(param.name.to_ascii_lowercase(), "KeyValuePair".into());
                            match body {
                                LambdaBody::Expr(expr) => {
                                    normalize_vb_dotnet_collection_expr(expr, &lambda_locals)
                                }
                                LambdaBody::Block(body) => {
                                    normalize_vb_dotnet_collection_statements(
                                        body,
                                        &mut lambda_locals,
                                    )
                                }
                            }
                        }
                    }
                    return;
                }
                if field.eq_ignore_ascii_case("GetType") && args.is_empty() {
                    if let Some(type_name) = vb_infer_expr_type(object, locals) {
                        *expr = Expression::string(&format!("System.{type_name}"));
                        return;
                    }
                }
                if matches!(
                    field.to_ascii_lowercase().as_str(),
                    "all"
                        | "any"
                        | "where"
                        | "select"
                        | "selectmany"
                        | "distinctby"
                        | "orderby"
                        | "orderbydescending"
                        | "groupby"
                        | "sum"
                        | "average"
                        | "minby"
                        | "maxby"
                ) && !args.is_empty()
                {
                    if let ExprKind::Ident(receiver_name) = &object.kind {
                        if let Some(element_type) =
                            locals.get(&format!("$element:{}", receiver_name.to_ascii_lowercase()))
                        {
                            if let ExprKind::Lambda { params, body, .. } = &mut args[0].value.kind {
                                if let Some(param) = params.first() {
                                    let mut lambda_locals = locals.clone();
                                    lambda_locals.insert(
                                        param.name.to_ascii_lowercase(),
                                        element_type.clone(),
                                    );
                                    match body {
                                        LambdaBody::Expr(expr) => {
                                            normalize_vb_date_literal_expr(expr, &HashMap::new());
                                            normalize_vb_local_type_expr(expr, &lambda_locals);
                                        }
                                        LambdaBody::Block(body) => {
                                            normalize_vb_date_literal_statements(
                                                body,
                                                &mut HashMap::new(),
                                            );
                                            normalize_vb_local_type_statements(
                                                body,
                                                &mut lambda_locals,
                                            );
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if let ExprKind::Ident(name) = &object.kind {
                    if let Some(pattern) =
                        locals.get(&format!("$regex_pattern:{}", name.to_ascii_lowercase()))
                    {
                        if field.eq_ignore_ascii_case("GetGroupNames") && args.is_empty() {
                            *expr = vb_regex_group_names_array(pattern);
                            return;
                        }
                        if field.eq_ignore_ascii_case("GroupNameFromNumber") && args.len() == 1 {
                            if let Some(number) = literal_i64(&args[0].value) {
                                *expr = vb_regex_group_name_from_number(pattern, number);
                                return;
                            }
                        }
                        if field.eq_ignore_ascii_case("GroupNumberFromName") && args.len() == 1 {
                            if let Some(name) = vb_known_string_value(&args[0].value, locals) {
                                *expr = vb_regex_group_number_from_name(pattern, &name);
                                return;
                            }
                        }
                    }
                }
                if field.eq_ignore_ascii_case("Contains") && args.len() == 1 {
                    if let (Some(haystack), Some(needle)) =
                        (literal_string(object), literal_string(&args[0].value))
                    {
                        *expr = Expression::bool(haystack.contains(&needle));
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("ToCharArray") && args.is_empty() {
                    if let Some(text) = literal_string(object) {
                        *expr = Expression::new(ExprKind::Array(
                            text.chars()
                                .map(|ch| ArrayElement {
                                    key: None,
                                    value: Expression::string(&ch.to_string()),
                                    spread: false,
                                    by_ref: false,
                                })
                                .collect(),
                        ));
                        return;
                    }
                }
                if field.eq_ignore_ascii_case("Equals")
                    && args.len() == 1
                    && vb_infer_expr_type(object, locals).as_deref() != Some("Version")
                {
                    *expr = Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident("__vb_object_equals")),
                        args: vec![
                            Argument::positional((**object).clone()),
                            Argument::positional(args[0].value.clone()),
                        ],
                        optional: false,
                    });
                    return;
                }
                if args.len() == 1 {
                    if let Some(path) = dotted_expr_name(object) {
                        if matches!(
                            path.to_ascii_lowercase().as_str(),
                            "double" | "system.double" | "single" | "system.single"
                        ) {
                            if field.eq_ignore_ascii_case("IsNaN") {
                                let value = args[0].value.clone();
                                *expr = Expression::new(ExprKind::Binary {
                                    op: BinOp::NotEq,
                                    left: Box::new(value.clone()),
                                    right: Box::new(value),
                                });
                                return;
                            }
                            if field.eq_ignore_ascii_case("IsInfinity") {
                                let value = args[0].value.clone();
                                let pos = Expression::new(ExprKind::Binary {
                                    op: BinOp::Eq,
                                    left: Box::new(value.clone()),
                                    right: Box::new(Expression::float(f64::INFINITY)),
                                });
                                let neg = Expression::new(ExprKind::Binary {
                                    op: BinOp::Eq,
                                    left: Box::new(value),
                                    right: Box::new(Expression::float(f64::NEG_INFINITY)),
                                });
                                *expr = Expression::new(ExprKind::Binary {
                                    op: BinOp::Or,
                                    left: Box::new(pos),
                                    right: Box::new(neg),
                                });
                                return;
                            }
                        }
                    }
                }
                if field.eq_ignore_ascii_case("Parse") && args.len() == 1 {
                    if let Some(path) = dotted_expr_name(object) {
                        let target = vb_canonical_type_name(&path);
                        if let Some(text) = literal_string(&args[0].value) {
                            match target.as_str() {
                                "Int32" | "Int64" | "Int16" | "Byte" | "SByte" | "UInt16"
                                | "UInt32" | "UInt64" => {
                                    if let Ok(value) = text.trim().parse::<i64>() {
                                        *expr = Expression::int(value);
                                        return;
                                    }
                                }
                                "Double" | "Single" | "Decimal" => {
                                    if let Ok(value) = text.trim().parse::<f64>() {
                                        *expr = Expression::float(value);
                                        return;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
                if field.eq_ignore_ascii_case("TryParse") && args.len() == 2 {
                    if let Some(path) = dotted_expr_name(object) {
                        let target_type = vb_canonical_type_name(&path);
                        if let Some(text) = literal_string(&args[0].value) {
                            if let ExprKind::Ident(target) = &args[1].value.kind {
                                let parsed = match target_type.as_str() {
                                    "Int32" | "Int64" | "Int16" | "Byte" | "SByte" | "UInt16"
                                    | "UInt32" | "UInt64" => {
                                        text.trim().parse::<i64>().ok().map(Expression::int)
                                    }
                                    "Double" | "Single" | "Decimal" => {
                                        text.trim().parse::<f64>().ok().map(Expression::float)
                                    }
                                    _ => None,
                                };
                                if let Some(value) = parsed {
                                    *expr = Expression::new(ExprKind::Assign {
                                        target: Box::new(Expression::ident(target)),
                                        value: Box::new(value),
                                    });
                                    return;
                                }
                            }
                        }
                    }
                }
                if field.eq_ignore_ascii_case("TryParse") && args.len() == 1 {
                    if let Some(path) = dotted_expr_name(object) {
                        let target = vb_canonical_type_name(&path);
                        if let Some(text) = literal_string(&args[0].value) {
                            match target.as_str() {
                                "Int32" | "Int64" | "Int16" | "Byte" | "SByte" | "UInt16"
                                | "UInt32" | "UInt64" => {
                                    if let Ok(value) = text.trim().parse::<i64>() {
                                        *expr = Expression::int(value);
                                        return;
                                    }
                                }
                                "Double" | "Single" | "Decimal" => {
                                    if let Ok(value) = text.trim().parse::<f64>() {
                                        *expr = Expression::float(value);
                                        return;
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                }
            }
        }
        ExprKind::Binary { op, left, right } => {
            normalize_vb_local_type_expr(left, locals);
            normalize_vb_local_type_expr(right, locals);
            if *op == BinOp::Concat {
                let new_left = vb_stringify_bool_for_concat((**left).clone(), locals);
                let new_right = vb_stringify_bool_for_concat((**right).clone(), locals);
                *left = Box::new(new_left);
                *right = Box::new(new_right);
            }
            if matches!(op, BinOp::Eq | BinOp::NotEq) {
                if let Some(folded) = fold_vb_null_string_eq(*op, left, right, locals) {
                    *expr = folded;
                    return;
                }
            }
            if let Some(folded) = vb_fold_decimal_comparison(*op, left, right, locals) {
                *expr = folded;
            }
        }
        ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_local_type_expr(left, locals);
            normalize_vb_local_type_expr(right, locals);
        }
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => {
            normalize_vb_local_type_expr(inner, locals);
            if type_name == "__vb_like_pattern" {
                *expr = if let Some(pattern) = vb_known_string_value(inner, locals) {
                    Expression::string(&vb_like_pattern_to_regex(&pattern))
                } else {
                    (**inner).clone()
                };
                return;
            }
            let is_trycast = type_name.to_ascii_lowercase().starts_with("trycast:");
            let cast_type = type_name
                .split(':')
                .next_back()
                .unwrap_or(type_name)
                .to_string();
            let mut value = (**inner).clone();
            vb_apply_known_local_value(&mut value, locals);
            if is_trycast {
                let source_type = vb_infer_expr_type(&value, locals).unwrap_or_default();
                if source_type == vb_canonical_type_name(&cast_type) {
                    *expr = value;
                    return;
                }
                if !matches!(value.kind, ExprKind::Ident(_)) {
                    *expr = Expression::null();
                    return;
                }
                return;
            }
            vb_coerce_literal_to_type(&mut value, &cast_type);
            if !matches!(value.kind, ExprKind::Ident(_)) {
                *expr = value;
            }
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_local_type_expr(expr, locals),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_local_type_expr(cond, locals);
            normalize_vb_local_type_expr(then, locals);
            normalize_vb_local_type_expr(else_, locals);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_local_type_expr(object, locals);
            normalize_vb_local_type_expr(index, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_local_type_expr(&mut item.value, locals);
            }
        }
        ExprKind::New { args, .. } => {
            for arg in args {
                normalize_vb_local_type_expr(&mut arg.value, locals);
            }
        }
        _ => {}
    }
}

fn vb_err_state_expr() -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("message"),
            value: Expression::string(""),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("number"),
            value: Expression::int(0),
        },
    ]))
}

fn vb_err_state_ident() -> Expression {
    Expression::ident("__vb_err")
}

fn vb_err_member(field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(vb_err_state_ident()),
        field: field.to_string(),
        null_safe: false,
    })
}

fn vb_err_decl_stmt() -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident("__vb_err".into()),
            type_hint: Some("Object".into()),
            init: Some(vb_err_state_expr()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Dim,
    })
}

fn vb_err_clear_statements() -> Vec<Statement> {
    vec![
        Statement::new(StmtKind::Assign {
            targets: vec![vb_err_member("message")],
            value: Expression::string(""),
        }),
        Statement::new(StmtKind::Assign {
            targets: vec![vb_err_member("number")],
            value: Expression::int(0),
        }),
    ]
}

fn vb_err_capture_stmt(name: &str) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![vb_err_state_ident()],
        value: Expression::ident(name),
    })
}

fn is_vb_err_ident(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Err"))
}

fn is_vb_err_clear_call(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if matches!(&callee.kind,
                ExprKind::Member { object, field, .. }
                    if is_vb_err_ident(object) && field.eq_ignore_ascii_case("Clear"))
    )
}

fn find_vb_label(body: &[Statement], label: &str, start: usize) -> Option<usize> {
    body.iter()
        .enumerate()
        .skip(start)
        .find_map(|(idx, stmt)| match &stmt.kind {
            StmtKind::Label(name) if name.eq_ignore_ascii_case(label) => Some(idx),
            _ => None,
        })
}

fn wrap_vb_resume_next(stmt: Statement, span: Span) -> Statement {
    Statement::with_span(
        StmtKind::Try {
            body: vec![stmt],
            catches: vec![CatchClause {
                types: Vec::new(),
                var_name: Some("__vb_err_catch".into()),
                stack_var: None,
                body: vec![vb_err_capture_stmt("__vb_err_catch")],
                when_clause: None,
            }],
            else_body: None,
            finally: None,
        },
        span,
    )
}

fn rewrite_vb_err_expr(expr: &mut Expression) -> bool {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            let mut used = rewrite_vb_err_expr(object);
            if is_vb_err_ident(object) {
                if field.eq_ignore_ascii_case("Description") {
                    *expr = vb_err_member("message");
                    used = true;
                } else if field.eq_ignore_ascii_case("Number") {
                    *expr = vb_err_member("number");
                    used = true;
                }
            }
            used
        }
        ExprKind::Call { callee, args, .. } => {
            let mut used = false;
            used |= rewrite_vb_err_expr(callee);
            for arg in args {
                used |= rewrite_vb_err_expr(&mut arg.value);
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if is_vb_err_ident(object) && field.eq_ignore_ascii_case("Raise") {
                    *callee = Box::new(Expression::ident("__vb_err_raise"));
                    used = true;
                }
            }
            used
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => rewrite_vb_err_expr(left) | rewrite_vb_err_expr(right),
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => rewrite_vb_err_expr(expr),
        ExprKind::RefOf(place) => match place.as_mut() {
            PlaceExpr::Ident(_) => false,
            PlaceExpr::Member { object, .. } => rewrite_vb_err_expr(object),
            PlaceExpr::Index { object, index, .. } => {
                rewrite_vb_err_expr(object) | rewrite_vb_err_expr(index)
            }
            PlaceExpr::Deref(expr) => rewrite_vb_err_expr(expr),
        },
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_err_expr(cond) | rewrite_vb_err_expr(then) | rewrite_vb_err_expr(else_)
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_err_expr(object) | rewrite_vb_err_expr(index)
        }
        ExprKind::New { class, args } => {
            let mut used = rewrite_vb_err_expr(class);
            for arg in args {
                used |= rewrite_vb_err_expr(&mut arg.value);
            }
            used
        }
        ExprKind::Assign { target, value } => {
            rewrite_vb_err_expr(target) | rewrite_vb_err_expr(value)
        }
        ExprKind::Lambda { body, .. } => {
            if let LambdaBody::Expr(expr) = body {
                rewrite_vb_err_expr(expr)
            } else {
                false
            }
        }
        ExprKind::Array(items) => {
            let mut used = false;
            for item in items {
                if let Some(key) = &mut item.key {
                    used |= rewrite_vb_err_expr(key);
                }
                used |= rewrite_vb_err_expr(&mut item.value);
            }
            used
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            let mut used = false;
            for item in items {
                used |= rewrite_vb_err_expr(item);
            }
            used
        }
        ExprKind::NamedTuple { fields, .. } => {
            let mut used = false;
            for (_, value) in fields {
                used |= rewrite_vb_err_expr(value);
            }
            used
        }
        ExprKind::Object(props) => {
            let mut used = false;
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        used |= rewrite_vb_err_expr(key);
                        used |= rewrite_vb_err_expr(value);
                    }
                    ObjectProperty::Spread(expr) => used |= rewrite_vb_err_expr(expr),
                    ObjectProperty::Computed { key, value } => {
                        used |= rewrite_vb_err_expr(key);
                        used |= rewrite_vb_err_expr(value);
                    }
                    ObjectProperty::Method { .. }
                    | ObjectProperty::Accessor { .. }
                    | ObjectProperty::Shorthand(_) => {}
                }
            }
            used
        }
        ExprKind::Interpolation(parts) => {
            let mut used = false;
            for part in parts {
                if let InterpolPart::Expr(expr) = part {
                    used |= rewrite_vb_err_expr(expr);
                }
            }
            used
        }
        ExprKind::IsType { expr, .. } | ExprKind::Cast { expr, .. } => rewrite_vb_err_expr(expr),
        ExprKind::Yield(Some(expr)) => rewrite_vb_err_expr(expr),
        ExprKind::SuperCall { args, .. } => {
            let mut used = false;
            for arg in args {
                used |= rewrite_vb_err_expr(&mut arg.value);
            }
            used
        }
        ExprKind::Comprehension {
            element,
            generators,
            ..
        } => {
            let mut used = rewrite_vb_err_expr(element);
            for generator in generators {
                used |= rewrite_vb_err_expr(&mut generator.iter);
                for condition in &mut generator.conditions {
                    used |= rewrite_vb_err_expr(condition);
                }
            }
            used
        }
        ExprKind::Slice { lower, upper, step } => {
            let mut used = false;
            if let Some(lower) = lower {
                used |= rewrite_vb_err_expr(lower);
            }
            if let Some(upper) = upper {
                used |= rewrite_vb_err_expr(upper);
            }
            if let Some(step) = step {
                used |= rewrite_vb_err_expr(step);
            }
            used
        }
        ExprKind::Range { start, end, .. } => rewrite_vb_err_expr(start) | rewrite_vb_err_expr(end),
        ExprKind::StaticAccess { class, member } => {
            rewrite_vb_err_expr(class) | rewrite_vb_err_expr(member)
        }
        ExprKind::Match { subject, arms } => {
            let mut used = rewrite_vb_err_expr(subject);
            for arm in arms {
                if let Some(conditions) = &mut arm.conditions {
                    for condition in conditions {
                        used |= rewrite_vb_err_expr(condition);
                    }
                }
                used |= rewrite_vb_err_expr(&mut arm.body);
            }
            used
        }
        ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::DefaultOf(_)
        | ExprKind::Yield(None)
        | ExprKind::AddressOf(_)
        | ExprKind::Destructure(_)
        | ExprKind::ClassExpr { .. }
        | ExprKind::FunctionExpr(_) => false,
    }
}

fn normalize_vb_legacy_error_statement(stmt: &mut Statement) -> bool {
    if matches!(&stmt.kind, StmtKind::Expr(expr) if is_vb_err_clear_call(expr)) {
        stmt.kind = StmtKind::Block(vb_err_clear_statements());
        return true;
    }

    match &mut stmt.kind {
        StmtKind::Expr(expr)
        | StmtKind::Return(Some(expr))
        | StmtKind::CompoundAssign { value: expr, .. } => rewrite_vb_err_expr(expr),
        StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        } => rewrite_vb_err_expr(expr),
        StmtKind::Throw {
            expr: Some(expr),
            cause: Some(cause),
        } => rewrite_vb_err_expr(expr) | rewrite_vb_err_expr(cause),
        StmtKind::VarDecl { declarations, .. } => {
            let mut used = false;
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    used |= rewrite_vb_err_expr(init);
                }
                if let Some(bounds) = &mut decl.array_bounds {
                    for bound in bounds {
                        used |= rewrite_vb_err_expr(bound);
                    }
                }
            }
            used
        }
        StmtKind::Assign { targets, value } => {
            let mut used = rewrite_vb_err_expr(value);
            for target in targets {
                used |= rewrite_vb_err_expr(target);
            }
            used
        }
        StmtKind::Block(body) | StmtKind::Using { body, .. } => {
            normalize_vb_legacy_error_body(body)
        }
        StmtKind::Lock { expr, body } => {
            rewrite_vb_err_expr(expr) | normalize_vb_legacy_error_body(body)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let mut used = rewrite_vb_err_expr(cond) | normalize_vb_legacy_error_body(then_body);
            for (elif_cond, elif_body) in elifs {
                used |= rewrite_vb_err_expr(elif_cond);
                used |= normalize_vb_legacy_error_body(elif_body);
            }
            if let Some(else_body) = else_body {
                used |= normalize_vb_legacy_error_body(else_body);
            }
            used
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut used = false;
            if let Some(init) = init {
                used |= normalize_vb_legacy_error_statement(init);
            }
            if let Some(cond) = cond {
                used |= rewrite_vb_err_expr(cond);
            }
            if let Some(update) = update {
                used |= rewrite_vb_err_expr(update);
            }
            used | normalize_vb_legacy_error_body(body)
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            let mut used = rewrite_vb_err_expr(iter) | normalize_vb_legacy_error_body(body);
            if let Some(else_body) = else_body {
                used |= normalize_vb_legacy_error_body(else_body);
            }
            used
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            let mut used = rewrite_vb_err_expr(cond) | normalize_vb_legacy_error_body(body);
            if let Some(else_body) = else_body {
                used |= normalize_vb_legacy_error_body(else_body);
            }
            used
        }
        StmtKind::DoWhile { body, cond, .. } => {
            normalize_vb_legacy_error_body(body) | rewrite_vb_err_expr(cond)
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let mut used = normalize_vb_legacy_error_body(body);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    used |= rewrite_vb_err_expr(when_clause);
                }
                used |= normalize_vb_legacy_error_body(&mut catch.body);
            }
            if let Some(else_body) = else_body {
                used |= normalize_vb_legacy_error_body(else_body);
            }
            if let Some(finally) = finally {
                used |= normalize_vb_legacy_error_body(finally);
            }
            used
        }
        StmtKind::FunctionDecl { body, .. } => normalize_vb_legacy_error_body(body),
        _ => false,
    }
}

fn normalize_vb_legacy_error_body(body: &mut Vec<Statement>) -> bool {
    let original = std::mem::take(body);
    let mut rewritten = Vec::new();
    let mut uses_err_state = false;
    let mut resume_next = false;
    let mut index = 0usize;

    while index < original.len() {
        match &original[index].kind {
            StmtKind::OnErrorResumeNext => {
                uses_err_state = true;
                resume_next = true;
                index += 1;
            }
            StmtKind::OnErrorGoTo(target) if target == "0" => {
                uses_err_state = true;
                resume_next = false;
                index += 1;
            }
            StmtKind::OnErrorGoTo(target) if target == "-1" => {
                uses_err_state = true;
                rewritten.extend(vb_err_clear_statements());
                index += 1;
            }
            StmtKind::OnErrorGoTo(target) => {
                uses_err_state = true;
                if let Some(label_index) = find_vb_label(&original, target, index + 1) {
                    let mut try_body = original[index + 1..label_index].to_vec();
                    let mut handler_body = original[label_index + 1..].to_vec();
                    normalize_vb_legacy_error_body(&mut try_body);
                    normalize_vb_legacy_error_body(&mut handler_body);
                    let mut catch_body = vec![vb_err_capture_stmt("__vb_err_catch")];
                    catch_body.append(&mut handler_body);
                    rewritten.push(Statement::with_span(
                        StmtKind::Try {
                            body: try_body,
                            catches: vec![CatchClause {
                                types: Vec::new(),
                                var_name: Some("__vb_err_catch".into()),
                                stack_var: None,
                                body: catch_body,
                                when_clause: None,
                            }],
                            else_body: None,
                            finally: None,
                        },
                        original[index].span,
                    ));
                    index = original.len();
                } else {
                    index += 1;
                }
            }
            _ => {
                let mut stmt = original[index].clone();
                uses_err_state |= normalize_vb_legacy_error_statement(&mut stmt);
                if resume_next {
                    uses_err_state = true;
                    rewritten.push(wrap_vb_resume_next(stmt, original[index].span));
                } else {
                    rewritten.push(stmt);
                }
                index += 1;
            }
        }
    }

    if uses_err_state {
        rewritten.insert(0, vb_err_decl_stmt());
    }
    *body = rewritten;
    uses_err_state
}

fn rewrite_vb_bare_throws(stmts: &mut Vec<Statement>, var_name: &str) {
    for stmt in stmts.iter_mut() {
        rewrite_vb_bare_throws_in_stmt(stmt, var_name);
    }
}

fn rewrite_vb_bare_throws_in_stmt(stmt: &mut Statement, var_name: &str) {
    match &mut stmt.kind {
        StmtKind::Throw { expr, .. } if expr.is_none() => {
            *expr = Some(Expression::ident(var_name));
        }
        StmtKind::Block(inner) => rewrite_vb_bare_throws(inner, var_name),
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            rewrite_vb_bare_throws(then_body, var_name);
            for (_, body) in elifs {
                rewrite_vb_bare_throws(body, var_name);
            }
            if let Some(else_body) = else_body {
                rewrite_vb_bare_throws(else_body, var_name);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => {
            rewrite_vb_bare_throws(body, var_name);
        }
        StmtKind::Try { body, finally, .. } => {
            rewrite_vb_bare_throws(body, var_name);
            if let Some(finally) = finally {
                rewrite_vb_bare_throws(finally, var_name);
            }
        }
        _ => {}
    }
}

/*

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs = VbParser::parse(Rule::program, source)
        .map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::imports_statement => imports.push(parse_imports_statement(inner)?),
                Rule::statement_line => {
                    for stmt_pair in inner.into_inner() {
                        if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                            continue;
                            let elements: Vec<Expression> = p.into_inner()
                                .filter(|e| e.as_rule() == Rule::expression)
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            let mut all_args = args;
                            for elem in elements {
                                all_args.push(Argument::positional(elem));
                            }
                            return Ok(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args: all_args,
                            }, span));
                        }
                        Rule::with_initializer => {
                            let mut members = Vec::new();
                            for mi in p.into_inner() {
                                if mi.as_rule() != Rule::member_initializer {
                                    continue;
                                }
                                let mut mi_inner = mi.into_inner();
                                let prop_name = mi_inner.next().unwrap().as_str().to_ascii_lowercase();
                                let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                                members.push((prop_name, prop_expr));
                            }
                            return Ok(emit_vb_object_init_iife(Expression::with_span(ExprKind::New {
                                class: Box::new(Expression::ident(&class_name)),
                                args,
                            }, span), members));
                        }
                        _ => {}
                    }
                }
                if let Some(elements) = array_init {
                    ExprKind::Array(
                        elements.into_iter().map(|expr| ArrayElement {
                            key: None,
                            value: expr,
                            spread: false,
                            by_ref: false,
                        }).collect(),
                    )
                } else {
                    ExprKind::New {
                        class: Box::new(Expression::ident(&class_name)),
                        args,
                    }
                }
            }
            Rule::if_expression => {
                let mut inner = pair.into_inner();
                let first = parse_expression(inner.next().unwrap())?;
                let second = parse_expression(inner.next().unwrap())?;
                let third = inner.next().map(parse_expression).transpose()?;
                match third {
                    Some(else_expr) => ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    },
                    None => ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    },
                }
            }
            Rule::addressof_expr => {
                let inner = pair.into_inner();
                let mut name = String::new();
                for p in inner {
                    if p.as_rule() == Rule::dotted_identifier {
                        name = p.as_str().to_string();
                    }
                }
                ExprKind::AddressOf(name)
            }
            Rule::me_keyword => ExprKind::This,
            Rule::dot_call_statement => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments: Vec<Argument> = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("dot_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::dot_member_access => {
                let inner = pair.into_inner();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_access => {
                let mut inner = pair.into_inner();
                let _me = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::This);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::mybase_member_access => {
                let mut inner = pair.into_inner();
                let _mybase = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::Super);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = vec![];
                let mut arguments: Vec<Argument> = vec![];
                for p in inner {
                    match p.as_rule() {
                        Rule::me_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }

                if identifiers.is_empty() {
                    return Err("me_member_call needs at least one identifier".to_string());
                }

                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::This);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }

                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::mybase_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = vec![];
                let mut arguments: Vec<Argument> = vec![];
                for p in inner {
                    match p.as_rule() {
                        Rule::mybase_keyword => {}
                        Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }

                if identifiers.is_empty() {
                    return Err("mybase_member_call needs at least one identifier".to_string());
                }

                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Super);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }

                if identifiers.len() == 1 {
                    ExprKind::SuperCall {
                        method: Some(method_name),
                        args: arguments,
                    }
                } else {
                    let callee = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: method_name,
                        null_safe: false,
                    });
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args: arguments,
                        optional: false,
                    }
                }
            }
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };

        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "andalso" => BinOp::And,
                    "orelse" => BinOp::Or,
                    "and" => BinOp::BitAnd,
                    "or" => BinOp::BitOr,
                    "xor" => BinOp::BitXor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left),
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = maybe_rewrite_vb_binary(op, left, right);
    }

    Ok(left)
}

*/
fn parse_sub_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let raw = pair.as_str();
    let mut decorators = parse_vb_attribute_specs(raw);
    let mut is_partial_method = vb_decl_starts_with_partial(raw);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut body = Vec::new();
    let mut handles: Vec<String> = Vec::new();
    let mut is_async = false;
    let mut is_extension = false;
    let mut is_overridable = false;
    let mut is_overrides = false;
    let mut is_must_override = false;
    let mut is_shared = false;
    let mut is_not_overridable = false;
    let mut is_overloads = false;

    for p in inner {
        match p.as_rule() {
            Rule::extension_attribute => is_extension = true,
            Rule::attribute_line => {
                if vb_attribute_line_is_extension(p.as_str()) {
                    is_extension = true;
                }
            }
            Rule::visibility_modifier => visibility = parse_visibility(p.as_str()),
            Rule::partial_keyword => is_partial_method = true,
            Rule::async_kw => is_async = true,
            Rule::sub_modifier_keyword => {
                let kw = p.as_str().to_lowercase();
                match kw.as_str() {
                    "overrides" => is_overrides = true,
                    "overridable" | "virtual" => is_overridable = true,
                    "mustoverride" => is_must_override = true,
                    "shared" => is_shared = true,
                    "notoverridable" => is_not_overridable = true,
                    "shadows" => is_not_overridable = true,
                    "overloads" => is_overloads = true,
                    _ => {}
                }
            }
            Rule::identifier | Rule::member_identifier | Rule::sub_name => {
                name = p.as_str().to_string()
            }
            Rule::generic_suffix => consume_vb_generic_suffix(p.as_str()),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::sub_block_body => {
                body.extend(parse_block(p)?);
            }
            Rule::sub_inline_body => {
                for stmt_pair in p.into_inner() {
                    match stmt_pair.as_rule() {
                        Rule::sub_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => body.push(parse_statement(stmt_pair)?),
                    }
                }
            }
            Rule::handles_clause => {
                for hp in p.into_inner() {
                    if hp.as_rule() == Rule::dotted_identifier {
                        handles.push(hp.as_str().to_string());
                    }
                }
            }
            Rule::implements_member_clause => {}
            _ => {}
        }
    }

    normalize_vb_legacy_error_body(&mut body);
    normalize_vb_date_literal_body(&mut body);
    normalize_vb_local_type_body(&mut body);

    let is_generator = body_has_yield(&body);
    if is_partial_method {
        decorators.push(Expression::string(VB_PARTIAL_METHOD_MARKER));
    }

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: parameters,
            return_type: None,
            body,
            modifiers: Modifiers {
                visibility,
                is_static: is_shared,
                is_abstract: is_must_override,
                is_virtual: is_overridable,
                is_override: is_overrides,
                is_readonly: false,
                is_shared,
                is_extension,
                is_overloads,
                is_not_overridable,
                is_destructor: false,
                protocol_slot: None,
                decorators,
            },
            handles,
            is_async,
            is_generator,
            is_sub: true,
        },
        span,
    ))
}

fn parse_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let raw = pair.as_str();
    let mut decorators = parse_vb_attribute_specs(raw);
    let mut is_partial_method = vb_decl_starts_with_partial(raw);
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut handles: Vec<String> = Vec::new();
    let mut is_async = false;
    let mut is_extension = false;
    let mut is_overridable = false;
    let mut is_overrides = false;
    let mut is_must_override = false;
    let mut is_shared = false;
    let mut is_not_overridable = false;
    let mut is_overloads = false;

    for p in inner {
        match p.as_rule() {
            Rule::extension_attribute => is_extension = true,
            Rule::attribute_line => {
                if vb_attribute_line_is_extension(p.as_str()) {
                    is_extension = true;
                }
            }
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::partial_keyword => is_partial_method = true,
            Rule::async_kw => is_async = true,
            Rule::sub_modifier_keyword => {
                let kw = p.as_str().to_lowercase();
                match kw.as_str() {
                    "overrides" => is_overrides = true,
                    "overridable" | "virtual" => is_overridable = true,
                    "mustoverride" => is_must_override = true,
                    "shared" => is_shared = true,
                    "notoverridable" => is_not_overridable = true,
                    "shadows" => is_not_overridable = true,
                    "overloads" => is_overloads = true,
                    _ => {}
                }
            }
            Rule::identifier | Rule::member_identifier | Rule::function_name => {
                name = p.as_str().to_string()
            }
            Rule::generic_suffix => consume_vb_generic_suffix(p.as_str()),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::func_block_body => {
                body.extend(parse_block(p)?);
            }
            Rule::func_inline_body => {
                for stmt_pair in p.into_inner() {
                    match stmt_pair.as_rule() {
                        Rule::func_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => body.push(parse_statement(stmt_pair)?),
                    }
                }
            }
            Rule::handles_clause => {
                for hp in p.into_inner() {
                    if hp.as_rule() == Rule::dotted_identifier {
                        handles.push(hp.as_str().to_string());
                    }
                }
            }
            Rule::implements_member_clause => {}
            _ => {}
        }
    }

    normalize_vb_legacy_error_body(&mut body);
    normalize_vb_date_literal_body(&mut body);
    normalize_vb_local_type_body(&mut body);

    let is_generator = body_has_yield(&body);
    if is_partial_method {
        decorators.push(Expression::string(VB_PARTIAL_METHOD_MARKER));
    }

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: parameters,
            return_type,
            body,
            modifiers: Modifiers {
                visibility,
                is_static: is_shared,
                is_abstract: is_must_override,
                is_virtual: is_overridable,
                is_override: is_overrides,
                is_readonly: false,
                is_shared,
                is_extension,
                is_overloads,
                is_not_overridable,
                is_destructor: false,
                protocol_slot: None,
                decorators,
            },
            handles,
            is_async,
            is_generator,
            is_sub: false,
        },
        span,
    ))
}

fn parse_operator_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let decorators = parse_vb_attribute_specs(pair.as_str());
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut symbol = String::new();
    let mut parameters = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut is_shared = false;
    let mut is_overloads = false;

    for p in inner {
        match p.as_rule() {
            Rule::visibility_modifier => visibility = parse_visibility(p.as_str()),
            Rule::sub_modifier_keyword => match p.as_str().to_ascii_lowercase().as_str() {
                "shared" => is_shared = true,
                "overloads" => is_overloads = true,
                _ => {}
            },
            Rule::operator_symbol => symbol = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => return_type = Some(p.as_str().to_string()),
            Rule::operator_block_body => body.extend(parse_block(p)?),
            Rule::operator_inline_body => {
                for stmt_pair in p.into_inner() {
                    match stmt_pair.as_rule() {
                        Rule::operator_end | Rule::NEWLINE | Rule::EOI => {}
                        _ => body.push(parse_statement(stmt_pair)?),
                    }
                }
            }
            _ => {}
        }
    }

    normalize_vb_date_literal_body(&mut body);
    normalize_vb_local_type_body(&mut body);
    let source_arity = parameters.len();
    let is_conversion_operator = symbol.eq_ignore_ascii_case("CType");
    let instance_operator = if is_conversion_operator {
        false
    } else {
        normalize_vb_operator_as_instance(&mut parameters, &mut body)
    };

    let name = if is_conversion_operator {
        let from_type = parameters
            .first()
            .and_then(|param| param.type_hint.as_ref())
            .map(|ty| vb_canonical_type_name(ty))
            .unwrap_or_else(|| "Object".to_string());
        let to_type = return_type
            .as_ref()
            .map(|ty| vb_canonical_type_name(ty))
            .unwrap_or_else(|| "Object".to_string());
        format!(
            "__ctype_{}_to_{}",
            from_type.to_ascii_lowercase(),
            to_type.to_ascii_lowercase()
        )
    } else {
        vb_operator_method_name(&symbol, source_arity).to_string()
    };
    let protocol_slot = vb_operator_protocol_slot(&symbol, source_arity, is_conversion_operator);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params: parameters,
            return_type,
            body,
            modifiers: Modifiers {
                visibility,
                is_static: is_shared && !instance_operator,
                is_abstract: false,
                is_virtual: false,
                is_override: false,
                is_readonly: false,
                is_shared: is_shared && !instance_operator,
                is_extension: false,
                is_overloads,
                is_not_overridable: false,
                is_destructor: false,
                protocol_slot,
                decorators,
            },
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
        span,
    ))
}

fn vb_operator_protocol_slot(
    symbol: &str,
    arity: usize,
    is_conversion_operator: bool,
) -> Option<ProtocolSlot> {
    if is_conversion_operator {
        return None;
    }
    match (symbol, arity) {
        (s, 1) if s.eq_ignore_ascii_case("IsTrue") => Some(ProtocolSlot::Bool),
        (s, 1) if s.eq_ignore_ascii_case("IsFalse") => Some(ProtocolSlot::Bool),
        ("+", _) => Some(ProtocolSlot::Add),
        ("-", 1) => Some(ProtocolSlot::Neg),
        ("-", _) => Some(ProtocolSlot::Sub),
        ("*", _) => Some(ProtocolSlot::Mul),
        ("/", _) => Some(ProtocolSlot::Div),
        ("\\", _) => Some(ProtocolSlot::FloorDiv),
        ("Mod" | "mod", _) => Some(ProtocolSlot::Mod),
        ("=", _) => Some(ProtocolSlot::Eq),
        ("<>", _) => Some(ProtocolSlot::Ne),
        ("<", _) => Some(ProtocolSlot::Lt),
        ("<=", _) => Some(ProtocolSlot::Le),
        (">", _) => Some(ProtocolSlot::Gt),
        (">=", _) => Some(ProtocolSlot::Ge),
        (s, _) if s.eq_ignore_ascii_case("And") => Some(ProtocolSlot::And),
        (s, _) if s.eq_ignore_ascii_case("Or") => Some(ProtocolSlot::Or),
        (s, _) if s.eq_ignore_ascii_case("Xor") => Some(ProtocolSlot::Xor),
        (s, 1) if s.eq_ignore_ascii_case("Not") => Some(ProtocolSlot::Not),
        _ => None,
    }
}

fn vb_operator_method_name(symbol: &str, arity: usize) -> &'static str {
    match (symbol, arity) {
        (s, 1) if s.eq_ignore_ascii_case("IsTrue") => "__istrue__",
        (s, 1) if s.eq_ignore_ascii_case("IsFalse") => "__isfalse__",
        ("+", _) => "__add__",
        ("-", 1) => "__neg__",
        ("-", _) => "__sub__",
        ("*", _) => "__mul__",
        ("/", _) => "__truediv__",
        ("\\", _) => "__floordiv__",
        ("Mod" | "mod", _) => "__mod__",
        ("=", _) => "__eq__",
        ("<>", _) => "__ne__",
        ("<", _) => "__lt__",
        ("<=", _) => "__le__",
        (">", _) => "__gt__",
        (">=", _) => "__ge__",
        (s, _) if s.eq_ignore_ascii_case("Like") => "__like__",
        ("Not" | "not", 1) => "__bitnot__",
        _ => "operator",
    }
}

fn normalize_vb_operator_as_instance(params: &mut Vec<Param>, body: &mut Vec<Statement>) -> bool {
    let Some(left_param) = params.first().cloned() else {
        return false;
    };
    params.remove(0);
    let alias = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(left_param.name),
            type_hint: left_param.type_hint,
            init: Some(Expression::new(ExprKind::This)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Dim,
    });
    body.insert(0, alias);
    true
}

fn parse_module_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut members = Vec::new();
    let module_field_modifiers = || {
        let mut modifiers = Modifiers::default();
        modifiers.is_static = true;
        modifiers.is_shared = true;
        modifiers
    };

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::generic_suffix => consume_vb_generic_suffix(p.as_str()),
            Rule::property_decl => members.extend(parse_property_decl_to_members(p)?),
            Rule::auto_property_decl => {
                let d = parse_auto_property_as_field(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers: module_field_modifiers(),
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::sub_decl => {
                let sub_stmt = parse_sub_decl(p)?;
                let is_ctor = matches!(
                    &sub_stmt.kind,
                    StmtKind::FunctionDecl { name, .. } if name.eq_ignore_ascii_case("New")
                );
                if is_ctor {
                    if let StmtKind::FunctionDecl {
                        params,
                        body,
                        mut modifiers,
                        ..
                    } = sub_stmt.kind
                    {
                        modifiers.is_static = true;
                        modifiers.is_shared = true;
                        members.push(ClassMember::Method(Box::new(Statement::with_span(
                            StmtKind::FunctionDecl {
                                name: "__static_init__".to_string(),
                                params,
                                return_type: None,
                                body,
                                modifiers,
                                handles: Vec::new(),
                                is_async: false,
                                is_generator: false,
                                is_sub: true,
                            },
                            span.clone(),
                        ))));
                    }
                } else {
                    members.push(ClassMember::Method(Box::new(sub_stmt)));
                }
            }
            Rule::function_decl => {
                members.push(ClassMember::Method(Box::new(parse_function_decl(p)?)))
            }
            Rule::const_statement => {
                let (vis, decl) = parse_const_statement(p)?;
                let init = decl.init.unwrap_or_else(|| Expression::null());
                let name = match decl.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Const {
                    name,
                    type_hint: decl.type_hint,
                    value: init,
                    visibility: vis,
                });
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: module_field_modifiers(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let mut modifiers = parse_field_modifiers(&p);
                modifiers.is_static = true;
                modifiers.is_shared = true;
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::class_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_class_decl(p)?)));
            }
            Rule::interface_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_interface_decl(p)?)));
            }
            Rule::structure_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_structure_decl(p)?)));
            }
            Rule::enum_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_enum_decl(p)?)));
            }
            Rule::NEWLINE | Rule::module_end => {}
            _ => {}
        }
    }
    Ok(Statement::with_span(
        StmtKind::ModuleDecl {
            name,
            members,
            visibility: Visibility::Public,
        },
        span,
    ))
}

fn parse_field_modifiers(pair: &Pair<Rule>) -> Modifiers {
    let mut modifiers = Modifiers::default();

    for field_part in pair.clone().into_inner() {
        match field_part.as_rule() {
            Rule::visibility_modifier => {
                modifiers.visibility = parse_visibility(field_part.as_str());
            }
            Rule::sub_modifier_keyword if field_part.as_str().eq_ignore_ascii_case("shared") => {
                modifiers.is_static = true;
                modifiers.is_shared = true;
            }
            _ if field_part.as_str().eq_ignore_ascii_case("readonly") => {
                modifiers.is_readonly = true;
            }
            _ => {}
        }
    }

    modifiers
}

fn parse_property_modifiers(pair: &Pair<Rule>) -> Modifiers {
    let mut modifiers = Modifiers::default();
    modifiers.decorators = parse_vb_attribute_specs(pair.as_str());
    for part in pair.clone().into_inner() {
        match part.as_rule() {
            Rule::visibility_modifier => {
                modifiers.visibility = parse_visibility(part.as_str());
            }
            Rule::default_keyword => {}
            Rule::sub_modifier_keyword => match part.as_str().to_ascii_lowercase().as_str() {
                "shared" => {
                    modifiers.is_static = true;
                    modifiers.is_shared = true;
                }
                "overrides" => modifiers.is_override = true,
                "overridable" | "virtual" => modifiers.is_virtual = true,
                "mustoverride" => modifiers.is_abstract = true,
                "notoverridable" => modifiers.is_not_overridable = true,
                "shadows" => modifiers.is_not_overridable = true,
                "overloads" => modifiers.is_overloads = true,
                _ => {}
            },
            _ => match part.as_str().to_ascii_lowercase().as_str() {
                "readonly" => modifiers.is_readonly = true,
                "shared" => {
                    modifiers.is_static = true;
                    modifiers.is_shared = true;
                }
                _ => {}
            },
        }
    }
    modifiers
}

/// Parse `Imports [alias =] namespace.or.type`
fn parse_imports_statement(pair: Pair<Rule>) -> Result<Import, String> {
    let span = to_span(&pair);
    let mut alias: Option<String> = None;
    let mut path = String::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::imports_alias => {
                // imports_alias = { identifier ~ "=" }
                if let Some(id) = p.into_inner().next() {
                    alias = Some(id.as_str().to_string());
                }
            }
            Rule::dotted_identifier | Rule::type_name => {
                path = p.as_str().to_string();
            }
            Rule::NEWLINE => {}
            _ => {}
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span,
    })
}

/// Parse `Namespace dotted.name ... End Namespace`
fn parse_namespace_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::dotted_identifier => {
                name = p.as_str().to_string();
            }
            Rule::class_decl => {
                body.push(parse_class_decl(p)?);
            }
            Rule::module_decl => {
                body.push(parse_module_decl(p)?);
            }
            Rule::enum_decl => {
                body.push(parse_enum_decl(p)?);
            }
            Rule::namespace_decl => {
                // Nested namespace
                body.push(parse_namespace_decl(p)?);
            }
            Rule::interface_decl => {
                body.push(parse_interface_decl(p)?);
            }
            Rule::structure_decl => {
                body.push(parse_structure_decl(p)?);
            }
            Rule::NEWLINE | Rule::namespace_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::NamespaceDecl { name, body },
        span,
    ))
}

/// Parse an auto-implemented property into a VarDeclarator (field), since it's syntactic sugar.
fn parse_auto_property_as_field(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let pair_text = pair.as_str().to_ascii_lowercase();
    let mut name = String::new();
    let mut var_type: Option<String> = None;
    let mut initializer = None;
    let is_with_events = pair_text
        .split(|c: char| !c.is_ascii_alphanumeric())
        .any(|part| part == "withevents");

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_name => var_type = Some(p.as_str().to_string()),
            Rule::expression => initializer = Some(parse_expression(p)?),
            // Skip visibility, ReadOnly, WriteOnly keywords
            _ => {}
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint: var_type,
        init: initializer,
        array_bounds: None,
        with_events: is_with_events,
    })
}

fn split_vb_top_level_list(text: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut depth = 0i32;
    let mut start = 0usize;
    for (idx, ch) in text.char_indices() {
        match ch {
            '(' => depth += 1,
            ')' => depth -= 1,
            ',' if depth == 0 => {
                parts.push(text[start..idx].trim().to_string());
                start = idx + ch.len_utf8();
            }
            _ => {}
        }
    }
    parts.push(text[start..].trim().to_string());
    parts
}

fn last_vb_top_level_member(text: &str) -> Option<String> {
    let mut depth = 0i32;
    let mut last_dot = None;
    for (idx, ch) in text.char_indices() {
        match ch {
            '(' => depth += 1,
            ')' => depth -= 1,
            '.' if depth == 0 => last_dot = Some(idx),
            _ => {}
        }
    }
    let member = if let Some(idx) = last_dot {
        &text[idx + 1..]
    } else {
        text
    };
    let member = member.trim();
    if member.is_empty() {
        None
    } else {
        Some(member.to_string())
    }
}

#[derive(Clone)]
struct VbImplementsTarget {
    leaf: String,
    forwarder: String,
}

fn vb_interface_forwarder_name(interface: &str, member: &str) -> String {
    format!(
        "__vb_iface_{}_{}",
        sanitize_vb_static_key(interface),
        sanitize_vb_static_key(member)
    )
}

fn vb_implements_target_member_infos(pair: &Pair<Rule>) -> Vec<VbImplementsTarget> {
    let mut names = Vec::new();
    for p in pair.clone().into_inner() {
        if p.as_rule() != Rule::implements_member_clause {
            continue;
        }
        let text = p.as_str().trim();
        let Some(rest) = text.get("Implements".len()..) else {
            continue;
        };
        for target in split_vb_top_level_list(rest) {
            if let Some(member) = last_vb_top_level_member(&target) {
                let interface = target
                    .rsplit_once('.')
                    .map(|(prefix, _)| strip_vb_generic_suffix(prefix.trim()))
                    .unwrap_or_default();
                names.push(VbImplementsTarget {
                    forwarder: vb_interface_forwarder_name(&interface, &member),
                    leaf: member,
                });
            }
        }
    }
    names
}

fn vb_class_member_method_name(member: &ClassMember) -> Option<&str> {
    let ClassMember::Method(stmt) = member else {
        return None;
    };
    let StmtKind::FunctionDecl { name, .. } = &stmt.kind else {
        return None;
    };
    Some(name)
}

fn vb_class_has_method(members: &[ClassMember], name: &str) -> bool {
    members.iter().any(|member| {
        vb_class_member_method_name(member)
            .is_some_and(|member_name| member_name.eq_ignore_ascii_case(name))
    })
}

fn vb_class_member_is_forwarder_named(member: &ClassMember, name: &str) -> bool {
    let ClassMember::Method(stmt) = member else {
        return false;
    };
    let StmtKind::FunctionDecl {
        name: method_name,
        body,
        ..
    } = &stmt.kind
    else {
        return false;
    };
    if !method_name.eq_ignore_ascii_case(name) || body.len() != 1 {
        return false;
    }
    let call = match &body[0].kind {
        StmtKind::Expr(expr) => expr,
        StmtKind::Return(Some(expr)) => expr,
        _ => return false,
    };
    let ExprKind::Call { callee, .. } = &call.kind else {
        return false;
    };
    matches!(
        &callee.kind,
        ExprKind::Member { object, .. } if matches!(object.kind, ExprKind::This)
    )
}

fn vb_class_member_has_method_signature(member: &ClassMember, stmt: &Statement) -> bool {
    let ClassMember::Method(member_stmt) = member else {
        return false;
    };
    let StmtKind::FunctionDecl {
        params: member_params,
        ..
    } = &member_stmt.kind
    else {
        return false;
    };
    let StmtKind::FunctionDecl { params, .. } = &stmt.kind else {
        return false;
    };
    member_params.len() == params.len()
        && member_params
            .iter()
            .zip(params)
            .all(|(left_param, right_param)| {
                let same_type = match (&left_param.type_hint, &right_param.type_hint) {
                    (Some(left_type), Some(right_type)) => vb_canonical_type_name(left_type)
                        .eq_ignore_ascii_case(&vb_canonical_type_name(right_type)),
                    (None, None) => true,
                    _ => false,
                };
                same_type && left_param.pass_by == right_param.pass_by
            })
}

fn vb_interface_forwarder(stmt: &Statement, interface_name: &str) -> Option<ClassMember> {
    let StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        modifiers,
        body,
        is_async,
        is_generator,
        is_sub,
        ..
    } = &stmt.kind
    else {
        return None;
    };
    if name.eq_ignore_ascii_case(interface_name) {
        return None;
    }

    let mut forwarder_modifiers = modifiers.clone();
    forwarder_modifiers.visibility = Visibility::Public;

    Some(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: interface_name.to_string(),
            params: params.clone(),
            return_type: return_type.clone(),
            body: body.clone(),
            modifiers: forwarder_modifiers,
            handles: vec![],
            is_async: *is_async,
            is_generator: *is_generator,
            is_sub: *is_sub,
        },
    ))))
}

fn push_vb_interface_forwarders(
    members: &mut Vec<ClassMember>,
    stmt: &Statement,
    targets: &[String],
) {
    for target in targets {
        if vb_class_has_method(members, target) {
            let has_exact_forwarder = members.iter().any(|member| {
                vb_class_member_is_forwarder_named(member, target)
                    && vb_class_member_has_method_signature(member, stmt)
            });
            let has_any_forwarder = members
                .iter()
                .any(|member| vb_class_member_is_forwarder_named(member, target));
            if has_exact_forwarder {
                members.retain(|member| {
                    !(vb_class_member_is_forwarder_named(member, target)
                        && vb_class_member_has_method_signature(member, stmt))
                });
            } else if !has_any_forwarder {
                continue;
            } else {
                // Same interface slot, different signature: keep both and let
                // the shared class overload machinery publish signature slots.
            }
        }
        if let Some(forwarder) = vb_interface_forwarder(stmt, target) {
            members.push(forwarder);
        }
    }
}

fn extract_vb_this_constructor_call(
    body: &mut Vec<Statement>,
) -> (
    vybe_ast::ConstructorInitializerTarget,
    Option<Vec<Expression>>,
) {
    let Some(first) = body.first() else {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    };
    let StmtKind::Expr(expr) = &first.kind else {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    };
    if !field.eq_ignore_ascii_case("New") {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    }
    let is_this_ctor = matches!(object.kind, ExprKind::This)
        || matches!(object.kind, ExprKind::Ident(ref name) if name.eq_ignore_ascii_case("MyClass"));
    if !is_this_ctor {
        return (vybe_ast::ConstructorInitializerTarget::Base, None);
    }
    let args = args.iter().map(|arg| arg.value.clone()).collect();
    body.remove(0);
    (vybe_ast::ConstructorInitializerTarget::This, Some(args))
}

fn vb_myclass_alias_name(class_name: &str, method_name: &str) -> String {
    format!(
        "__vb_myclass_{}_{}",
        class_name.to_ascii_lowercase(),
        method_name.to_ascii_lowercase()
    )
}

fn normalize_vb_implicit_property_self(members: &mut Vec<ClassMember>) {
    let properties: HashSet<String> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Property { name, .. } => Some(name.to_ascii_lowercase()),
            _ => None,
        })
        .collect();
    if properties.is_empty() {
        for member in members {
            if let ClassMember::NestedType(stmt) = member {
                if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
                    normalize_vb_implicit_property_self(members);
                }
            }
        }
        return;
    }

    for member in members {
        match member {
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl { params, body, .. } = &mut stmt.kind {
                    let mut locals = params
                        .iter()
                        .map(|param| param.name.to_ascii_lowercase())
                        .collect();
                    normalize_vb_implicit_property_self_statements(body, &properties, &mut locals);
                }
            }
            ClassMember::Constructor { params, body, .. } => {
                let mut locals = params
                    .iter()
                    .map(|param| param.name.to_ascii_lowercase())
                    .collect();
                normalize_vb_implicit_property_self_statements(body, &properties, &mut locals);
            }
            ClassMember::Property {
                name,
                getter,
                setter,
                ..
            } => {
                let mut accessor_properties = properties.clone();
                accessor_properties.remove(&name.to_ascii_lowercase());
                if let Some(getter) = getter {
                    normalize_vb_mybase_property_access_statements(getter, name);
                    normalize_vb_implicit_property_self_statements(
                        getter,
                        &accessor_properties,
                        &mut HashSet::new(),
                    );
                }
                if let Some(setter) = setter {
                    normalize_vb_mybase_property_access_statements(&mut setter.body, name);
                    let mut locals = HashSet::from([setter.param.name.to_ascii_lowercase()]);
                    normalize_vb_implicit_property_self_statements(
                        &mut setter.body,
                        &accessor_properties,
                        &mut locals,
                    );
                }
            }
            ClassMember::NestedType(stmt) => {
                if let StmtKind::ClassDecl { members, .. } = &mut stmt.kind {
                    normalize_vb_implicit_property_self(members);
                }
            }
            _ => {}
        }
    }
}

fn normalize_vb_mybase_property_access_statements(body: &mut [Statement], property: &str) {
    for stmt in body {
        normalize_vb_mybase_property_access_statement(stmt, property);
    }
}

fn normalize_vb_mybase_property_access_statement(stmt: &mut Statement, property: &str) {
    match &mut stmt.kind {
        StmtKind::Return(Some(expr)) | StmtKind::Expr(expr) => {
            normalize_vb_mybase_property_access_expr(expr, property);
        }
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            normalize_vb_mybase_property_access_expr(value, property);
            if vb_expr_is_mybase_member(&targets[0], property) {
                stmt.kind = StmtKind::Assign {
                    targets: vec![vb_mybase_property_backing_expr(property)],
                    value: value.clone(),
                };
            } else {
                normalize_vb_mybase_property_access_expr(&mut targets[0], property);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_mybase_property_access_expr(target, property);
            }
            normalize_vb_mybase_property_access_expr(value, property);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_mybase_property_access_expr(target, property);
            normalize_vb_mybase_property_access_expr(value, property);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_mybase_property_access_expr(cond, property);
            normalize_vb_mybase_property_access_statements(then_body, property);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_mybase_property_access_expr(elif_cond, property);
                normalize_vb_mybase_property_access_statements(elif_body, property);
            }
            if let Some(else_body) = else_body {
                normalize_vb_mybase_property_access_statements(else_body, property);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_mybase_property_access_expr(cond, property);
            normalize_vb_mybase_property_access_statements(body, property);
            if let Some(else_body) = else_body {
                normalize_vb_mybase_property_access_statements(else_body, property);
            }
        }
        StmtKind::DoWhile { cond, body, .. } => {
            normalize_vb_mybase_property_access_expr(cond, property);
            normalize_vb_mybase_property_access_statements(body, property);
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_vb_mybase_property_access_statement(init, property);
            }
            if let Some(cond) = cond {
                normalize_vb_mybase_property_access_expr(cond, property);
            }
            if let Some(update) = update {
                normalize_vb_mybase_property_access_expr(update, property);
            }
            normalize_vb_mybase_property_access_statements(body, property);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_mybase_property_access_expr(iter, property);
            normalize_vb_mybase_property_access_statements(body, property);
            if let Some(else_body) = else_body {
                normalize_vb_mybase_property_access_statements(else_body, property);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_mybase_property_access_statements(body, property);
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_mybase_property_access_expr(when_clause, property);
                }
                normalize_vb_mybase_property_access_statements(&mut catch.body, property);
            }
            if let Some(else_body) = else_body {
                normalize_vb_mybase_property_access_statements(else_body, property);
            }
            if let Some(finally) = finally {
                normalize_vb_mybase_property_access_statements(finally, property);
            }
        }
        StmtKind::Block(body) => normalize_vb_mybase_property_access_statements(body, property),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_mybase_property_access_expr(init, property);
                }
            }
        }
        _ => {}
    }
}

fn normalize_vb_mybase_property_access_expr(expr: &mut Expression, property: &str) {
    if vb_expr_is_mybase_member(expr, property) {
        *expr = vb_mybase_property_backing_expr(property);
        return;
    }
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_mybase_property_access_expr(callee, property);
            for arg in args {
                normalize_vb_mybase_property_access_expr(&mut arg.value, property);
            }
        }
        ExprKind::Member { object, .. } => {
            normalize_vb_mybase_property_access_expr(object, property)
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_mybase_property_access_expr(left, property);
            normalize_vb_mybase_property_access_expr(right, property);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::RefLoad(expr) => normalize_vb_mybase_property_access_expr(expr, property),
        ExprKind::Assign { target, value } => {
            normalize_vb_mybase_property_access_expr(target, property);
            normalize_vb_mybase_property_access_expr(value, property);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_mybase_property_access_expr(cond, property);
            normalize_vb_mybase_property_access_expr(then, property);
            normalize_vb_mybase_property_access_expr(else_, property);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_mybase_property_access_expr(object, property);
            normalize_vb_mybase_property_access_expr(index, property);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_mybase_property_access_expr(&mut item.value, property);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_mybase_property_access_expr(item, property);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_mybase_property_access_expr(class, property);
            for arg in args {
                normalize_vb_mybase_property_access_expr(&mut arg.value, property);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_mybase_property_access_expr(key, property);
                        normalize_vb_mybase_property_access_expr(value, property);
                    }
                    ObjectProperty::Spread(value) => {
                        normalize_vb_mybase_property_access_expr(value, property);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_mybase_property_access_statement(value, property);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn vb_mybase_property_backing_expr(property: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: format!("__{}", property.to_ascii_lowercase()),
        null_safe: false,
    })
}

fn vb_expr_is_mybase_member(expr: &Expression, property: &str) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Member { object, field, .. }
            if field.eq_ignore_ascii_case(property) && matches!(object.kind, ExprKind::Super)
    )
}

fn vb_property_self_expr(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: name.to_string(),
        null_safe: false,
    })
}

fn normalize_vb_implicit_property_self_statements(
    body: &mut [Statement],
    properties: &HashSet<String>,
    locals: &mut HashSet<String>,
) {
    for stmt in body {
        normalize_vb_implicit_property_self_statement(stmt, properties, locals);
    }
}

fn normalize_vb_implicit_property_self_statement(
    stmt: &mut Statement,
    properties: &HashSet<String>,
    locals: &mut HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_implicit_property_self_expr(init, properties, locals);
                }
                if let BindingPattern::Ident(name) = &decl.pattern {
                    locals.insert(name.to_ascii_lowercase());
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_implicit_property_self_target(target, properties, locals);
            }
            normalize_vb_implicit_property_self_expr(value, properties, locals);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_implicit_property_self_target(target, properties, locals);
            normalize_vb_implicit_property_self_expr(value, properties, locals);
        }
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_implicit_property_self_expr(expr, properties, locals);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_implicit_property_self_expr(cond, properties, locals);
            normalize_vb_implicit_property_self_statements(
                then_body,
                properties,
                &mut locals.clone(),
            );
            for (elif_cond, elif_body) in elifs {
                normalize_vb_implicit_property_self_expr(elif_cond, properties, locals);
                normalize_vb_implicit_property_self_statements(
                    elif_body,
                    properties,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_implicit_property_self_statements(
                    else_body,
                    properties,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            normalize_vb_implicit_property_self_expr(cond, properties, locals);
            normalize_vb_implicit_property_self_statements(body, properties, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_implicit_property_self_statements(
                    else_body,
                    properties,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::DoWhile { cond, body, .. } => {
            normalize_vb_implicit_property_self_expr(cond, properties, locals);
            normalize_vb_implicit_property_self_statements(body, properties, &mut locals.clone());
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_locals = locals.clone();
            if let Some(init) = init {
                normalize_vb_implicit_property_self_statement(init, properties, &mut loop_locals);
            }
            if let Some(cond) = cond {
                normalize_vb_implicit_property_self_expr(cond, properties, &loop_locals);
            }
            if let Some(update) = update {
                normalize_vb_implicit_property_self_expr(update, properties, &loop_locals);
            }
            normalize_vb_implicit_property_self_statements(body, properties, &mut loop_locals);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_implicit_property_self_expr(iter, properties, locals);
            normalize_vb_implicit_property_self_statements(body, properties, &mut locals.clone());
            if let Some(else_body) = else_body {
                normalize_vb_implicit_property_self_statements(
                    else_body,
                    properties,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_implicit_property_self_statements(body, properties, &mut locals.clone());
            for catch in catches {
                if let Some(when_clause) = &mut catch.when_clause {
                    normalize_vb_implicit_property_self_expr(when_clause, properties, locals);
                }
                normalize_vb_implicit_property_self_statements(
                    &mut catch.body,
                    properties,
                    &mut locals.clone(),
                );
            }
            if let Some(else_body) = else_body {
                normalize_vb_implicit_property_self_statements(
                    else_body,
                    properties,
                    &mut locals.clone(),
                );
            }
            if let Some(finally) = finally {
                normalize_vb_implicit_property_self_statements(
                    finally,
                    properties,
                    &mut locals.clone(),
                );
            }
        }
        StmtKind::Block(body) => {
            normalize_vb_implicit_property_self_statements(body, properties, &mut locals.clone());
        }
        _ => {}
    }
}

fn normalize_vb_implicit_property_self_target(
    expr: &mut Expression,
    properties: &HashSet<String>,
    locals: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name)
            if properties.contains(&name.to_ascii_lowercase())
                && !locals.contains(&name.to_ascii_lowercase()) =>
        {
            *expr = vb_property_self_expr(name);
        }
        _ => normalize_vb_implicit_property_self_expr(expr, properties, locals),
    }
}

fn normalize_vb_implicit_property_self_expr(
    expr: &mut Expression,
    properties: &HashSet<String>,
    locals: &HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name)
            if properties.contains(&name.to_ascii_lowercase())
                && !locals.contains(&name.to_ascii_lowercase()) =>
        {
            *expr = vb_property_self_expr(name);
        }
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_implicit_property_self_expr(callee, properties, locals);
            for arg in args {
                normalize_vb_implicit_property_self_expr(&mut arg.value, properties, locals);
            }
        }
        ExprKind::Member { object, .. } => {
            normalize_vb_implicit_property_self_expr(object, properties, locals);
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_implicit_property_self_expr(left, properties, locals);
            normalize_vb_implicit_property_self_expr(right, properties, locals);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr))
        | ExprKind::RefLoad(expr) => {
            normalize_vb_implicit_property_self_expr(expr, properties, locals);
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_implicit_property_self_target(target, properties, locals);
            normalize_vb_implicit_property_self_expr(value, properties, locals);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_implicit_property_self_expr(cond, properties, locals);
            normalize_vb_implicit_property_self_expr(then, properties, locals);
            normalize_vb_implicit_property_self_expr(else_, properties, locals);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_implicit_property_self_expr(object, properties, locals);
            normalize_vb_implicit_property_self_expr(index, properties, locals);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_implicit_property_self_expr(&mut item.value, properties, locals);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_implicit_property_self_expr(item, properties, locals);
            }
        }
        ExprKind::New { class, args } => {
            normalize_vb_implicit_property_self_expr(class, properties, locals);
            for arg in args {
                normalize_vb_implicit_property_self_expr(&mut arg.value, properties, locals);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        normalize_vb_implicit_property_self_expr(key, properties, locals);
                        normalize_vb_implicit_property_self_expr(value, properties, locals);
                    }
                    ObjectProperty::Spread(value) => {
                        normalize_vb_implicit_property_self_expr(value, properties, locals);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        let mut nested_locals = locals.clone();
                        normalize_vb_implicit_property_self_statement(
                            value,
                            properties,
                            &mut nested_locals,
                        );
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn normalize_vb_myclass_calls(members: &mut Vec<ClassMember>, class_name: &str) {
    let mut aliases = Vec::new();
    for member in members.iter_mut() {
        match member {
            ClassMember::Method(stmt) => {
                let alias_source_name =
                    if let StmtKind::FunctionDecl { name, body, .. } = &mut stmt.kind {
                        rewrite_vb_myclass_statements(body, class_name);
                        if !name.starts_with("__vb_myclass_") {
                            Some(name.clone())
                        } else {
                            None
                        }
                    } else {
                        None
                    };
                if let Some(source_name) = alias_source_name {
                    let mut cloned = (**stmt).clone();
                    if let StmtKind::FunctionDecl {
                        name: cloned_name,
                        modifiers: cloned_modifiers,
                        ..
                    } = &mut cloned.kind
                    {
                        *cloned_name = vb_myclass_alias_name(class_name, &source_name);
                        cloned_modifiers.is_override = false;
                        cloned_modifiers.is_virtual = false;
                        cloned_modifiers.is_overloads = false;
                        cloned_modifiers.is_not_overridable = true;
                        cloned_modifiers.visibility = Visibility::Private;
                    }
                    aliases.push(ClassMember::Method(Box::new(cloned)));
                }
            }
            ClassMember::Constructor { body, .. } => {
                rewrite_vb_myclass_statements(body, class_name);
            }
            ClassMember::Property { getter, setter, .. } => {
                if let Some(getter) = getter {
                    rewrite_vb_myclass_statements(getter, class_name);
                }
                if let Some(setter) = setter {
                    rewrite_vb_myclass_statements(&mut setter.body, class_name);
                }
            }
            ClassMember::NestedType(stmt) => {
                if let StmtKind::ClassDecl { name, members, .. } = &mut stmt.kind {
                    normalize_vb_myclass_calls(members, name);
                }
            }
            _ => {}
        }
    }
    members.extend(aliases);
}

fn rewrite_vb_myclass_statements(body: &mut [Statement], class_name: &str) {
    for stmt in body {
        rewrite_vb_myclass_statement(stmt, class_name);
    }
}

fn rewrite_vb_myclass_statement(stmt: &mut Statement, class_name: &str) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_myclass_expr(expr, class_name);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_vb_myclass_expr(init, class_name);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_vb_myclass_expr(target, class_name);
            }
            rewrite_vb_myclass_expr(value, class_name);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_vb_myclass_expr(target, class_name);
            rewrite_vb_myclass_expr(value, class_name);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_vb_myclass_expr(cond, class_name);
            rewrite_vb_myclass_statements(then_body, class_name);
            for (elif_cond, elif_body) in elifs {
                rewrite_vb_myclass_expr(elif_cond, class_name);
                rewrite_vb_myclass_statements(elif_body, class_name);
            }
            if let Some(else_body) = else_body {
                rewrite_vb_myclass_statements(else_body, class_name);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                rewrite_vb_myclass_statement(init, class_name);
            }
            if let Some(cond) = cond {
                rewrite_vb_myclass_expr(cond, class_name);
            }
            if let Some(update) = update {
                rewrite_vb_myclass_expr(update, class_name);
            }
            rewrite_vb_myclass_statements(body, class_name);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_vb_myclass_expr(iter, class_name);
            rewrite_vb_myclass_statements(body, class_name);
            if let Some(else_body) = else_body {
                rewrite_vb_myclass_statements(else_body, class_name);
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_vb_myclass_expr(cond, class_name);
            rewrite_vb_myclass_statements(body, class_name);
            if let Some(else_body) = else_body {
                rewrite_vb_myclass_statements(else_body, class_name);
            }
        }
        StmtKind::Block(body) => rewrite_vb_myclass_statements(body, class_name),
        _ => {}
    }
}

fn rewrite_vb_myclass_expr(expr: &mut Expression, class_name: &str) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_myclass_expr(callee, class_name);
            for arg in args {
                rewrite_vb_myclass_expr(&mut arg.value, class_name);
            }
        }
        ExprKind::Member { object, field, .. } => {
            rewrite_vb_myclass_expr(object, class_name);
            if matches!(object.kind, ExprKind::Ident(ref name) if name.eq_ignore_ascii_case("MyClass"))
                && !field.eq_ignore_ascii_case("New")
            {
                *object = Box::new(Expression::ident("Me"));
                *field = vb_myclass_alias_name(class_name, field);
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_vb_myclass_expr(left, class_name);
            rewrite_vb_myclass_expr(right, class_name);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::TypeOf(expr)
        | ExprKind::Await(expr)
        | ExprKind::Yield(Some(expr)) => rewrite_vb_myclass_expr(expr, class_name),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_myclass_expr(cond, class_name);
            rewrite_vb_myclass_expr(then, class_name);
            rewrite_vb_myclass_expr(else_, class_name);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_myclass_expr(object, class_name);
            rewrite_vb_myclass_expr(index, class_name);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_myclass_expr(&mut item.value, class_name);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                rewrite_vb_myclass_expr(item, class_name);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_vb_myclass_expr(class, class_name);
            for arg in args {
                rewrite_vb_myclass_expr(&mut arg.value, class_name);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value } => {
                        rewrite_vb_myclass_expr(key, class_name);
                        rewrite_vb_myclass_expr(value, class_name);
                    }
                    ObjectProperty::Computed { key, value } => {
                        rewrite_vb_myclass_expr(key, class_name);
                        rewrite_vb_myclass_expr(value, class_name);
                    }
                    ObjectProperty::Spread(value) => rewrite_vb_myclass_expr(value, class_name),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_myclass_statement(value, class_name);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn substitute_vb_constructor_expr(
    expr: &mut Expression,
    replacements: &HashMap<String, Expression>,
) {
    match &mut expr.kind {
        ExprKind::Ident(name) => {
            if let Some(replacement) = replacements.get(&name.to_ascii_lowercase()) {
                *expr = replacement.clone();
            }
        }
        ExprKind::Member { object, .. } => substitute_vb_constructor_expr(object, replacements),
        ExprKind::Call { callee, args, .. } => {
            substitute_vb_constructor_expr(callee, replacements);
            for arg in args {
                substitute_vb_constructor_expr(&mut arg.value, replacements);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            substitute_vb_constructor_expr(left, replacements);
            substitute_vb_constructor_expr(right, replacements);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => substitute_vb_constructor_expr(expr, replacements),
        ExprKind::Ternary { cond, then, else_ } => {
            substitute_vb_constructor_expr(cond, replacements);
            substitute_vb_constructor_expr(then, replacements);
            substitute_vb_constructor_expr(else_, replacements);
        }
        ExprKind::Index { object, index, .. } => {
            substitute_vb_constructor_expr(object, replacements);
            substitute_vb_constructor_expr(index, replacements);
        }
        ExprKind::New { class, args } => {
            substitute_vb_constructor_expr(class, replacements);
            for arg in args {
                substitute_vb_constructor_expr(&mut arg.value, replacements);
            }
        }
        ExprKind::Array(items) => {
            for item in items {
                substitute_vb_constructor_expr(&mut item.value, replacements);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        substitute_vb_constructor_expr(key, replacements);
                        substitute_vb_constructor_expr(value, replacements);
                    }
                    ObjectProperty::Spread(expr) => {
                        substitute_vb_constructor_expr(expr, replacements);
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        substitute_vb_constructor_stmt(value, replacements);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn substitute_vb_constructor_stmt(
    stmt: &mut Statement,
    replacements: &HashMap<String, Expression>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            substitute_vb_constructor_expr(expr, replacements);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                substitute_vb_constructor_expr(target, replacements);
            }
            substitute_vb_constructor_expr(value, replacements);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    substitute_vb_constructor_expr(init, replacements);
                }
            }
        }
        _ => {}
    }
}

fn resolve_vb_this_constructor_chains(members: &mut [ClassMember]) {
    let ctors: Vec<(Vec<Param>, Vec<Statement>)> = members
        .iter()
        .filter_map(|member| {
            let ClassMember::Constructor { params, body, .. } = member else {
                return None;
            };
            Some((params.clone(), body.clone()))
        })
        .collect();

    for member in members {
        let ClassMember::Constructor {
            params,
            body,
            base_args,
            initializer_target,
            ..
        } = member
        else {
            continue;
        };
        if *initializer_target != vybe_ast::ConstructorInitializerTarget::This {
            continue;
        }
        let Some(args) = base_args.take() else {
            continue;
        };
        let Some((target_params, target_body)) = ctors
            .iter()
            .find(|(candidate_params, _)| {
                candidate_params.len() == args.len() && candidate_params.len() != params.len()
            })
            .cloned()
        else {
            continue;
        };
        let mut replacements = HashMap::new();
        for (param, arg) in target_params.iter().zip(args.into_iter()) {
            replacements.insert(param.name.to_ascii_lowercase(), arg);
        }
        let mut inlined = target_body;
        for stmt in &mut inlined {
            substitute_vb_constructor_stmt(stmt, &replacements);
        }
        inlined.extend(std::mem::take(body));
        *body = inlined;
        *initializer_target = vybe_ast::ConstructorInitializerTarget::Base;
    }
}

fn parse_class_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let decorators = parse_vb_attribute_specs(pair.as_str());
    let inner = pair.into_inner();
    let mut name = String::new();
    let mut is_partial = false;
    let mut visibility = Visibility::Public;
    let mut parents: Vec<String> = Vec::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();
    let mut is_must_inherit = false;
    let mut is_not_inheritable = false;

    for p in inner {
        match p.as_rule() {
            Rule::generic_suffix => consume_vb_generic_suffix(p.as_str()),
            Rule::partial_keyword => is_partial = true,
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::must_inherit_keyword => is_must_inherit = true,
            Rule::not_inheritable_keyword => is_not_inheritable = true,
            Rule::inherits_statement => {
                if let Some(type_pair) = p.into_inner().next() {
                    // The common class pipeline compares canonical base names.
                    // VB heritage clauses may be qualified and/or generic
                    // (`System.IFoo(Of T)`); normalize those at the walker
                    // boundary so classes.rs sees the same identity everywhere.
                    parents.push(vb_declared_base_type_name(type_pair.as_str()));
                }
            }
            Rule::implements_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        interfaces.push(vb_declared_base_type_name(tp.as_str()));
                    }
                }
            }
            Rule::identifier => name = p.as_str().to_string(),
            Rule::property_decl => {
                members.extend(parse_property_decl_to_members(p)?);
            }
            Rule::auto_property_decl => {
                // Auto-implemented property → treat as a field
                let d = parse_auto_property_as_field(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers: Modifiers::default(),
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::sub_decl => {
                let implemented_targets = vb_implements_target_member_infos(&p);
                let implemented_members: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.leaf.clone())
                    .collect();
                let sub_stmt = parse_sub_decl(p)?;
                // Check if this is a constructor (New)
                let is_ctor = match &sub_stmt.kind {
                    StmtKind::FunctionDecl { name, .. } => name == "New",
                    _ => false,
                };
                if is_ctor {
                    match sub_stmt.kind {
                        StmtKind::FunctionDecl {
                            params,
                            mut body,
                            modifiers,
                            ..
                        } => {
                            if modifiers.is_static {
                                members.push(ClassMember::Method(Box::new(Statement::with_span(
                                    StmtKind::FunctionDecl {
                                        name: "__static_init__".to_string(),
                                        params,
                                        return_type: None,
                                        body,
                                        modifiers,
                                        handles: Vec::new(),
                                        is_async: false,
                                        is_generator: false,
                                        is_sub: true,
                                    },
                                    span.clone(),
                                ))));
                                continue;
                            }
                            let (initializer_target, base_args) =
                                extract_vb_this_constructor_call(&mut body);
                            members.push(ClassMember::Constructor {
                                name: None,
                                params,
                                body,
                                base_args,
                                initializer_target,
                                visibility: modifiers.visibility,
                            });
                        }
                        _ => unreachable!(),
                    }
                } else {
                    push_vb_interface_forwarders(&mut members, &sub_stmt, &implemented_members);
                    let qualified_forwarders: Vec<String> = implemented_targets
                        .iter()
                        .map(|target| target.forwarder.clone())
                        .collect();
                    push_vb_interface_forwarders(&mut members, &sub_stmt, &qualified_forwarders);
                    members.push(ClassMember::Method(Box::new(sub_stmt)));
                }
            }
            Rule::function_decl => {
                let implemented_targets = vb_implements_target_member_infos(&p);
                let implemented_members: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.leaf.clone())
                    .collect();
                let fn_stmt = parse_function_decl(p)?;
                push_vb_interface_forwarders(&mut members, &fn_stmt, &implemented_members);
                let qualified_forwarders: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.forwarder.clone())
                    .collect();
                push_vb_interface_forwarders(&mut members, &fn_stmt, &qualified_forwarders);
                members.push(ClassMember::Method(Box::new(fn_stmt)));
            }
            Rule::operator_decl => {
                members.push(ClassMember::Method(Box::new(parse_operator_decl(p)?)));
            }
            Rule::const_statement => {
                let (vis, decl) = parse_const_statement(p)?;
                let init = decl.init.unwrap_or_else(|| Expression::null());
                let name = match decl.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Const {
                    name,
                    type_hint: decl.type_hint,
                    value: init,
                    visibility: vis,
                });
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: Modifiers::default(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let modifiers = parse_field_modifiers(&p);
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::event_decl | Rule::custom_event_decl => {
                members.extend(parse_event_decl_to_members(p)?);
            }
            Rule::class_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_class_decl(p)?)));
            }
            Rule::interface_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_interface_decl(p)?)));
            }
            Rule::structure_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_structure_decl(p)?)));
            }
            Rule::enum_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_enum_decl(p)?)));
            }
            Rule::NEWLINE | Rule::class_end => {}
            _ => {}
        }
    }

    resolve_vb_this_constructor_chains(&mut members);
    normalize_vb_implicit_property_self(&mut members);
    normalize_vb_myclass_calls(&mut members, &name);
    normalize_vb_withevents_assignments(&mut members);

    // Inject canonical AddHandler statements at the END of the constructor
    // body for every class method that has a `Handles` clause. This is the
    // walker normalization that turns VB-specific `Handles ctrl.Event` into
    // the same canonical `StmtKind::AddHandler` that C# `+=` (and JS / Dart /
    // Python frontends) will produce. The compiler then has a single emit
    // path for events regardless of source language.
    inject_handles_into_constructor(&mut members);

    // Inject implicit `MyBase.New()` at the START of every constructor body
    // when this class has an `Inherits` clause and the body doesn't already
    // start with an explicit `MyBase.New(...)`. This matches real VB.NET
    // semantics: the runtime implicitly calls the parameterless parent ctor
    // before the body runs, and the VB compiler errors if no parameterless
    // parent ctor is accessible. By doing this here we keep the canonical
    // AST uniform — compile_class sees a body that always starts with the
    // base call, the same as Pascal `inherited Create(...)` and JS
    // `super(...)`. The compiler-side logic doesn't need a VB-specific
    // case.
    //
    // Also stamp `Me.__control_name = "<lowercased class name>"` immediately
    // after the base call, so any subsequent property writes (e.g.
    // `Me.Text = "X"`) mirror to the gui state under the user-meaningful
    // key. The base ctor (e.g. `Form()`) wires the underlying widget which
    // has its own auto-generated `__control_name`; this re-stamp overrides
    // it with the canonical "lowercased subclass name" form that real
    // WinForms users (and the existing Vybe form runner) expect.
    if let Some(parent_name) = parents.first() {
        inject_implicit_mybase_new(&mut members, &name, parent_name);
    }

    Ok(Statement::with_span(
        StmtKind::ClassDecl {
            name,
            parents,
            interfaces,
            members,
            modifiers: ClassModifiers {
                visibility,
                is_partial,
                is_abstract: is_must_inherit,
                is_sealed: is_not_inheritable,
                is_static: false,
                kind: vybe_ast::ClassKind::Class,
            },
            decorators,
        },
        span,
    ))
}

/// Inject `MyBase.New()` at the start of every constructor body that doesn't
/// already start with an explicit base call. Matches real VB.NET semantics
/// for `Inherits` classes.
///
/// "Already starts with one" is checked structurally — only the FIRST
/// statement is examined, because that's the only legal position for an
/// explicit `MyBase.New(...)` in VB. (VB.NET errors if `MyBase.New` appears
/// anywhere else in the body.)
///
/// If the class has no constructor at all, a default one is synthesized
/// containing just the `MyBase.New()` call.
fn inject_implicit_mybase_new(members: &mut Vec<ClassMember>, class_name: &str, parent_name: &str) {
    if vb_dotnet_descriptor_parent_skips_mybase_new(parent_name) {
        return;
    }
    let lowered = class_name.to_lowercase();

    let mybase_new = || -> Statement {
        // SuperCall { method: Some("New"), args: [] } — same shape the
        // mybase_member_call walker arm produces for explicit `MyBase.New()`.
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::SuperCall {
            method: Some("New".to_string()),
            args: Vec::new(),
        })))
    };

    // Me.__control_name = "<lowercased class name>"
    // This is the .NET-canonical "self identity" stamp. The base ctor (e.g.
    // `Form()`) wired up the underlying widget with its own auto-generated
    // name; we override it here so user property writes mirror to gui state
    // under the subclass name the rest of the system uses.
    let stamp_control_name = || -> Statement {
        let target = Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident("Me")),
            field: "__control_name".to_string(),
            null_safe: false,
        });
        let value = Expression::string(&lowered);
        Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
        })
    };

    let starts_with_mybase_new = |body: &[Statement]| -> bool {
        match body.first().map(|s| &s.kind) {
            Some(StmtKind::Expr(e)) => matches!(&e.kind, ExprKind::SuperCall { .. }),
            _ => false,
        }
    };

    let has_ctor = members
        .iter()
        .any(|m| matches!(m, ClassMember::Constructor { .. }));
    if !has_ctor {
        // Synthesize a default ctor that just calls MyBase.New() and stamps
        // the canonical control name.
        members.push(ClassMember::Constructor {
            name: None,
            params: Vec::new(),
            body: vec![mybase_new(), stamp_control_name()],
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
        return;
    }
    for m in members.iter_mut() {
        if let ClassMember::Constructor { body, .. } = m {
            if !starts_with_mybase_new(body) {
                body.insert(0, mybase_new());
                body.insert(1, stamp_control_name());
            } else {
                // Body already starts with MyBase.New() — insert the stamp
                // immediately after it.
                body.insert(1, stamp_control_name());
            }
        }
    }
}

fn vb_dotnet_descriptor_parent_skips_mybase_new(parent_name: &str) -> bool {
    vybe_platform_dotnet::emitter::is_component_descriptor_class(parent_name)
        && vybe_compiler::compiler::gui::canonical_control_name(parent_name).is_empty()
}

fn normalize_vb_withevents_assignments(members: &mut Vec<ClassMember>) {
    let with_events: HashSet<String> = members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Field {
                name, with_events, ..
            } if *with_events => Some(name.to_ascii_lowercase()),
            _ => None,
        })
        .collect();
    if with_events.is_empty() {
        return;
    }

    let mut bindings: HashMap<String, Vec<(String, String)>> = HashMap::new();
    let mut seen_bindings: HashSet<(String, String, String)> = HashSet::new();
    for member in members.iter() {
        let ClassMember::Method(stmt) = member else {
            continue;
        };
        let StmtKind::FunctionDecl { name, handles, .. } = &stmt.kind else {
            continue;
        };
        if name.starts_with("__vb_myclass_") {
            continue;
        }
        for handle in handles {
            let (control, event) = split_event_target(handle);
            let Some(field_name) = withevents_field_name(&control) else {
                continue;
            };
            let field_key = field_name.to_ascii_lowercase();
            let event_key = event.to_ascii_lowercase();
            let method_key = name.to_ascii_lowercase();
            if with_events.contains(&field_key)
                && seen_bindings.insert((field_key.clone(), event_key, method_key))
            {
                bindings
                    .entry(field_key)
                    .or_default()
                    .push((event, name.clone()));
            }
        }
    }
    if bindings.is_empty() {
        return;
    }

    for member in members.iter_mut() {
        match member {
            ClassMember::Method(stmt) => normalize_vb_withevents_stmt(stmt, &bindings),
            ClassMember::Constructor { body, .. } => {
                normalize_vb_withevents_statements(body, &bindings);
            }
            ClassMember::Property { getter, setter, .. } => {
                if let Some(getter) = getter {
                    normalize_vb_withevents_statements(getter, &bindings);
                }
                if let Some(setter) = setter {
                    normalize_vb_withevents_statements(&mut setter.body, &bindings);
                }
            }
            ClassMember::NestedType(stmt) => normalize_vb_withevents_stmt(stmt, &bindings),
            _ => {}
        }
    }
}

fn withevents_field_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. }
            if matches!(&object.kind, ExprKind::Ident(name) if name.eq_ignore_ascii_case("Me"))
                || matches!(object.kind, ExprKind::This) =>
        {
            Some(field.clone())
        }
        _ => None,
    }
}

fn normalize_vb_withevents_statements(
    stmts: &mut Vec<Statement>,
    bindings: &HashMap<String, Vec<(String, String)>>,
) {
    for stmt in stmts {
        normalize_vb_withevents_stmt(stmt, bindings);
    }
}

fn normalize_vb_withevents_stmt(
    stmt: &mut Statement,
    bindings: &HashMap<String, Vec<(String, String)>>,
) {
    match &mut stmt.kind {
        StmtKind::Assign { targets, value } if targets.len() == 1 => {
            let target = targets[0].clone();
            if let Some(field_name) = withevents_field_name(&target) {
                if let Some(field_bindings) = bindings.get(&field_name.to_ascii_lowercase()) {
                    let original = StmtKind::Assign {
                        targets: targets.clone(),
                        value: value.clone(),
                    };
                    let mut body = Vec::new();
                    body.extend(vb_withevents_handler_block(&target, field_bindings, false));
                    body.push(Statement::new(original));
                    body.extend(vb_withevents_handler_block(&target, field_bindings, true));
                    stmt.kind = StmtKind::Block(body);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } | StmtKind::Block(body) => {
            normalize_vb_withevents_statements(body, bindings);
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            normalize_vb_withevents_statements(then_body, bindings);
            for (_, body) in elifs {
                normalize_vb_withevents_statements(body, bindings);
            }
            if let Some(body) = else_body {
                normalize_vb_withevents_statements(body, bindings);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => normalize_vb_withevents_statements(body, bindings),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_withevents_statements(body, bindings);
            for catch in catches {
                normalize_vb_withevents_statements(&mut catch.body, bindings);
            }
            if let Some(body) = else_body {
                normalize_vb_withevents_statements(body, bindings);
            }
            if let Some(body) = finally {
                normalize_vb_withevents_statements(body, bindings);
            }
        }
        _ => {}
    }
}

fn vb_withevents_handler_block(
    target: &Expression,
    bindings: &[(String, String)],
    add: bool,
) -> Vec<Statement> {
    let mut then_body = Vec::new();
    for (event, method) in bindings {
        let handler = Expression::new(ExprKind::AddressOf(method.clone()));
        let kind = if add {
            StmtKind::AddHandler {
                control: target.clone(),
                event: event.clone(),
                handler,
            }
        } else {
            StmtKind::RemoveHandler {
                control: target.clone(),
                event: event.clone(),
                handler,
            }
        };
        then_body.push(Statement::new(kind));
    }
    vec![Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::IsNot,
            left: Box::new(target.clone()),
            right: Box::new(Expression::null()),
        }),
        then_body,
        elifs: Vec::new(),
        else_body: None,
    })]
}

/// Walk the class members; for every Method with `handles: ["ctrl.Event", ...]`,
/// build a canonical `AddHandler { control, event, handler: Me.method_name }`
/// statement and append it to the constructor body. If no constructor exists,
/// inject an empty one. Strips the `handles` field from the method afterward
/// so the compiler doesn't double-process it.
fn inject_handles_into_constructor(members: &mut Vec<ClassMember>) {
    // First pass: collect (handler_method_name, handles_list) and clear them
    // off the methods so the compile_function_decl path doesn't re-emit.
    let mut to_inject: Vec<(String, Vec<String>)> = Vec::new();
    for m in members.iter_mut() {
        if let ClassMember::Method(stmt) = m {
            if let StmtKind::FunctionDecl {
                name: mname,
                handles,
                modifiers,
                ..
            } = &mut stmt.kind
            {
                if !handles.is_empty()
                    && !modifiers.is_static
                    && !mname.starts_with("__vb_myclass_")
                {
                    let mut constructor_handles = Vec::new();
                    for handle in std::mem::take(handles) {
                        constructor_handles.push(handle);
                    }
                    if !constructor_handles.is_empty() {
                        to_inject.push((mname.clone(), constructor_handles));
                    }
                }
            }
        }
    }
    if to_inject.is_empty() {
        return;
    }

    // Build the AddHandler statements.
    let mut new_stmts: Vec<Statement> = Vec::new();
    for (method_name, handles) in &to_inject {
        for h in handles {
            let (control, event) = split_event_target(h);
            let control = match control.kind {
                ExprKind::Ident(name)
                    if !name.eq_ignore_ascii_case("me")
                        && !name.eq_ignore_ascii_case("mybase")
                        && !name.eq_ignore_ascii_case("this") =>
                {
                    Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident("Me")),
                        field: name,
                        null_safe: false,
                    })
                }
                _ => control,
            };
            // The handler is `Me.<method>` — a Member access on the class self.
            let handler = Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("Me")),
                field: method_name.clone(),
                null_safe: false,
            });
            new_stmts.push(Statement::new(
                vybe_compiler::compiler::events::add_handler_stmt(control, event, handler),
            ));
        }
    }

    // Find the constructor (or create one) and append the AddHandler statements.
    // VB constructors can appear either as an explicit `Sub New()` method or as a
    // dedicated `ClassMember::Constructor` node, depending on which parser path
    // produced the member. Handles normalization must attach to whichever form the
    // class already uses; otherwise `compile_class` can end up ignoring the
    // injected AddHandler body by selecting the explicit `Sub New()` body first.
    let has_explicit_new = members.iter().any(|m| {
        matches!(m,
            ClassMember::Method(stmt)
                if matches!(&stmt.kind,
                    StmtKind::FunctionDecl { name, .. } if name.eq_ignore_ascii_case("new")
                )
        )
    });
    let has_ctor = members
        .iter()
        .any(|m| matches!(m, ClassMember::Constructor { .. }));
    if !has_ctor && !has_explicit_new {
        members.push(ClassMember::Constructor {
            name: None,
            params: Vec::new(),
            body: new_stmts,
            base_args: None, // VB walker injects MyBase.New() into the body; no base_args needed here
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        });
    } else {
        for m in members.iter_mut() {
            match m {
                ClassMember::Constructor { body, .. } => {
                    body.extend(new_stmts.drain(..));
                    break;
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl { name, body, .. } = &mut stmt.kind {
                        if name.eq_ignore_ascii_case("new") {
                            body.extend(new_stmts.drain(..));
                            break;
                        }
                    }
                }
                _ => {}
            }
        }
    }
}

fn parse_property_decl_to_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let modifiers = parse_property_modifiers(&pair);
    let inner = pair.into_inner();
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut return_type: Option<String> = None;
    let mut getter = None;
    let mut setter = None;

    for p in inner {
        match p.as_str().to_lowercase().as_str() {
            _ => match p.as_rule() {
                Rule::identifier => name = p.as_str().to_string(),
                Rule::param_list => parameters = parse_param_list(p)?,
                Rule::type_name => return_type = Some(p.as_str().to_string()),
                Rule::property_get => getter = Some(parse_property_get(p)?),
                Rule::property_set => setter = Some(parse_property_set(p)?),
                _ => {}
            },
        }
    }

    if !parameters.is_empty() {
        let mut members = Vec::new();
        if let Some(body) = getter {
            let getter_name = if name.eq_ignore_ascii_case("Item") {
                "__getitem__".to_string()
            } else {
                name.clone()
            };
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: getter_name,
                    params: parameters.clone(),
                    return_type: return_type.clone(),
                    body: body.clone(),
                    modifiers: modifiers.clone(),
                    handles: vec![],
                    is_async: false,
                    is_generator: false,
                    is_sub: false,
                },
            ))));
            if name.eq_ignore_ascii_case("Item") {
                members.push(ClassMember::Method(Box::new(Statement::new(
                    StmtKind::FunctionDecl {
                        name: "__call__".to_string(),
                        params: parameters.clone(),
                        return_type: return_type.clone(),
                        body,
                        modifiers: modifiers.clone(),
                        handles: vec![],
                        is_async: false,
                        is_generator: false,
                        is_sub: false,
                    },
                ))));
            }
        }
        if let Some(setter) = setter {
            let setter_name = if name.eq_ignore_ascii_case("Item") {
                "__setitem__".to_string()
            } else {
                format!("__set_{name}")
            };
            let mut setter_params = parameters;
            setter_params.push(setter.param);
            members.push(ClassMember::Method(Box::new(Statement::new(
                StmtKind::FunctionDecl {
                    name: setter_name,
                    params: setter_params,
                    return_type: None,
                    body: setter.body,
                    modifiers,
                    handles: vec![],
                    is_async: false,
                    is_generator: false,
                    is_sub: true,
                },
            ))));
        }
        return Ok(members);
    }

    let is_auto = getter.is_none() && setter.is_none();
    Ok(vec![ClassMember::Property {
        name,
        type_hint: return_type,
        getter,
        setter,
        is_auto,
        modifiers,
    }])
}

fn parse_property_get(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt_pair in pair.into_inner() {
        match stmt_pair.as_rule() {
            Rule::statement_line => {
                for s in stmt_pair.into_inner() {
                    if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                        body.push(parse_statement(s)?);
                    }
                }
            }
            Rule::statement => body.push(parse_statement(stmt_pair)?),
            Rule::get_end | Rule::NEWLINE | Rule::EOI => {}
            Rule::visibility_modifier => {}
            _ => body.push(parse_statement(stmt_pair)?),
        }
    }
    Ok(body)
}

fn parse_property_set(pair: Pair<Rule>) -> Result<PropertySetter, String> {
    let inner = pair.into_inner();
    let mut param = Param {
        name: "value".to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };

    let mut body = Vec::new();
    for stmt_pair in inner {
        match stmt_pair.as_rule() {
            Rule::param => {
                param = parse_parameter(stmt_pair)?;
            }
            Rule::statement => body.push(parse_statement(stmt_pair)?),
            Rule::statement_line => {
                for s in stmt_pair.into_inner() {
                    if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                        body.push(parse_statement(s)?);
                    }
                }
            }
            Rule::set_end | Rule::NEWLINE | Rule::EOI | Rule::visibility_modifier => {}
            _ => body.push(parse_statement(stmt_pair)?),
        }
    }

    Ok(PropertySetter { param, body })
}

fn parse_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    pair.into_inner().map(parse_parameter).collect()
}

fn parse_parameter(pair: Pair<Rule>) -> Result<Param, String> {
    let inner = pair.into_inner();
    let mut pass_by = PassBy::Value;
    let mut name = String::new();
    let mut param_type: Option<String> = None;
    let mut is_optional = false;
    let mut default_value = None;
    let mut is_nullable = false;
    let mut is_param_array = false;

    for p in inner {
        match p.as_rule() {
            Rule::pass_type_keyword => {
                let text = p.as_str().to_lowercase();
                if text == "byval" {
                    pass_by = PassBy::Value;
                } else {
                    pass_by = PassBy::Ref;
                }
            }
            Rule::optional_keyword => {
                is_optional = true;
            }
            Rule::paramarray_keyword => {
                is_param_array = true;
                pass_by = PassBy::Value; // ParamArray is always ByVal
            }
            Rule::identifier => {
                name = p.as_str().to_string();
            }
            Rule::type_name => param_type = Some(p.as_str().to_string()),
            Rule::nullable_marker => is_nullable = true,
            Rule::expression => default_value = Some(parse_expression(p)?),
            _ => {}
        }
    }

    Ok(Param {
        name,
        type_hint: param_type,
        default: default_value,
        pass_by,
        is_rest: is_param_array,
        is_kwargs: false,
        is_optional,
        is_nullable,
    })
}

fn parse_array_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let elements = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            parse_expression(p).map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(Expression::with_span(ExprKind::Array(elements), span))
}

fn parse_const_statement(pair: Pair<Rule>) -> Result<(Visibility, VarDeclarator), String> {
    let mut visibility = Visibility::Private;
    let mut name = String::new();
    let mut type_hint = None;
    let mut init = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::visibility_modifier => visibility = parse_visibility(p.as_str()),
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_name => type_hint = Some(p.as_str().to_string()),
            Rule::expression => init = Some(parse_expression(p)?),
            Rule::array_literal => init = Some(parse_array_literal(p)?),
            _ => {}
        }
    }

    Ok((
        visibility,
        VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds: None,
            with_events: false,
        },
    ))
}

fn parse_array_bounds_pair(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    match pair.as_rule() {
        Rule::array_rank_spec => pair
            .into_inner()
            .next()
            .map(parse_array_bounds_pair)
            .transpose()
            .map(|bounds| bounds.unwrap_or_default()),
        Rule::array_bounds => pair
            .into_inner()
            .map(parse_array_bound_pair)
            .collect::<Result<Vec<_>, _>>(),
        Rule::array_bound => parse_array_bound_pair(pair).map(|bound| vec![bound]),
        Rule::array_rank_commas => Ok(Vec::new()),
        _ => Err(format!(
            "Unexpected array bounds rule: {:?}",
            pair.as_rule()
        )),
    }
}

fn parse_array_bound_pair(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner
        .next()
        .ok_or_else(|| "array bound missing upper expression".to_string())?;
    let lower_or_upper = parse_expression(first)?;
    if let Some(upper_pair) = inner.next() {
        let upper = parse_expression(upper_pair)?;
        Ok(call_expr(
            Expression::ident("__vb_array_bound"),
            vec![
                Argument::positional(lower_or_upper),
                Argument::positional(upper),
            ],
        ))
    } else {
        Ok(lower_or_upper)
    }
}

fn parse_dim_statement(pair: Pair<Rule>) -> Result<Vec<VarDeclarator>, String> {
    let mut decls = Vec::new();
    for part in pair.into_inner() {
        if part.as_rule() != Rule::dim_declaration_part {
            continue;
        }

        let mut name = String::new();
        let mut type_hint = None;
        let mut init = None;
        let mut array_bounds = None;
        let mut array_rank_count = 0usize;
        let mut ctor_args = Vec::new();
        let mut is_new = false;

        for p in part.into_inner() {
            match p.as_rule() {
                Rule::identifier => name = p.as_str().to_string(),
                Rule::array_rank_spec => {
                    array_rank_count += 1;
                    if array_bounds.is_none() {
                        array_bounds = Some(parse_array_bounds_pair(p)?);
                    }
                }
                Rule::array_bounds => {
                    array_bounds = Some(parse_array_bounds_pair(p)?);
                }
                Rule::type_name => type_hint = Some(p.as_str().to_string()),
                Rule::dim_new_keyword => is_new = true,
                Rule::argument_list => ctor_args = parse_argument_list(p)?,
                Rule::expression => init = Some(parse_expression(p)?),
                Rule::array_literal => init = Some(parse_array_literal(p)?),
                Rule::from_initializer => {
                    if let Some(class_name) = &type_hint {
                        let elements = p
                            .into_inner()
                            .filter(|e| e.as_rule() == Rule::expression)
                            .map(parse_expression)
                            .collect::<Result<Vec<_>, _>>()?;
                        init = Some(emit_vb_collection_init_iife(
                            Expression::new(ExprKind::New {
                                class: Box::new(build_dotted_expr(
                                    &strip_vb_generic_suffixes_preserve_path(class_name),
                                )),
                                args: ctor_args.clone(),
                            }),
                            elements,
                        ));
                    }
                }
                Rule::with_initializer => {
                    let mut members = Vec::new();
                    for mi in p.into_inner() {
                        if mi.as_rule() != Rule::member_initializer {
                            continue;
                        }
                        let mut mi_inner = mi.into_inner();
                        let prop_name = mi_inner.next().unwrap().as_str().to_ascii_lowercase();
                        let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                        members.push((prop_name, prop_expr));
                    }
                    if let Some(class_name) = &type_hint {
                        init = Some(emit_vb_object_init_iife(
                            Expression::new(ExprKind::New {
                                class: Box::new(build_dotted_expr(
                                    &strip_vb_generic_suffixes_preserve_path(class_name),
                                )),
                                args: ctor_args.clone(),
                            }),
                            members,
                        ));
                    }
                }
                _ => {}
            }
        }

        if array_rank_count > 0 {
            if let Some(type_hint_value) = type_hint.as_mut() {
                for _ in 0..array_rank_count {
                    type_hint_value.push_str("()");
                }
            }
        }

        if is_new {
            if let Some(type_hint_value) = type_hint.as_mut() {
                *type_hint_value = type_hint_value
                    .trim()
                    .strip_suffix("()")
                    .unwrap_or(type_hint_value.trim())
                    .trim()
                    .to_string();
            }
        }

        if is_new && init.is_none() {
            if let Some(class_name) = &type_hint {
                init = Some(Expression::new(ExprKind::New {
                    class: Box::new(build_dotted_expr(&strip_vb_generic_suffixes_preserve_path(
                        class_name,
                    ))),
                    args: ctor_args,
                }));
            }
        }

        if type_hint.is_none() {
            if let Some(inferred) = init.as_ref().and_then(vb_type_hint_from_cast_expr) {
                type_hint = Some(inferred);
            }
        }

        decls.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init,
            array_bounds,
            with_events: false,
        });
    }
    Ok(decls)
}

fn vb_type_hint_from_cast_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Cast { type_name, .. } => Some(vb_canonical_type_name(type_name)),
        _ => None,
    }
}

fn rewrite_vb_array_clone_for_cast(expr: Expression, type_name: &str) -> Expression {
    if !type_name.trim().ends_with("()") {
        return expr;
    }
    let ExprKind::Call {
        callee,
        args,
        optional: false,
    } = &expr.kind
    else {
        return expr;
    };
    if !args.is_empty() {
        return expr;
    }
    let ExprKind::Member {
        object,
        field,
        null_safe: false,
    } = &callee.kind
    else {
        return expr;
    };
    if !field.eq_ignore_ascii_case("Clone") {
        return expr;
    }
    call_expr(
        Expression::ident("__vb_array_clone"),
        vec![Argument::positional((**object).clone())],
    )
}

fn vb_new_array_bound_from_type_text(type_text: &str) -> Option<(String, Expression)> {
    let open = type_text.find('(')?;
    let close = type_text.rfind(')')?;
    if close <= open {
        return None;
    }
    let base = type_text[..open].trim().to_string();
    let bound_text = type_text[open + 1..close].trim();
    if base.is_empty() || bound_text.is_empty() || bound_text.contains(',') {
        return None;
    }
    parse_expression_str(bound_text)
        .ok()
        .map(|bound| (base, bound))
}

fn vb_new_array_bound_from_new_expr_text(text: &str) -> Option<(String, Expression)> {
    let trimmed = text.trim();
    let rest = trimmed
        .strip_prefix("New ")
        .or_else(|| trimmed.strip_prefix("new "))?
        .trim();
    let before_init = rest.split('{').next()?.trim();
    vb_new_array_bound_from_type_text(before_init)
}

fn parse_redim_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut preserve = false;
    let mut array = String::new();
    let mut bounds = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::preserve_keyword => preserve = true,
            Rule::identifier => array = p.as_str().to_string(),
            Rule::array_bounds => {
                bounds = parse_array_bounds_pair(p)?;
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::ReDim {
            preserve,
            array,
            bounds,
        },
        span,
    ))
}

fn parse_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let kind = match pair.as_rule() {
        Rule::dim_statement => {
            let decls = parse_dim_statement(pair)?;
            StmtKind::VarDecl {
                declarations: decls,
                kind: VarDeclKind::Dim,
            }
        }
        Rule::const_statement => {
            let (_vis, decl) = parse_const_statement(pair)?;
            StmtKind::VarDecl {
                declarations: vec![decl],
                kind: VarDeclKind::Const,
            }
        }
        Rule::redim_statement => {
            return parse_redim_statement(pair);
        }
        Rule::erase_statement => {
            let array = pair
                .into_inner()
                .find(|p| p.as_rule() == Rule::identifier)
                .map(|p| p.as_str().to_string())
                .ok_or_else(|| "Erase statement missing array name".to_string())?;
            StmtKind::Erase { array }
        }
        Rule::select_statement => {
            return parse_select_statement(pair);
        }
        Rule::dot_assign_statement => {
            // .prop1.prop2 = value (inside With block)
            let inner = pair.into_inner();
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => {
                        members.push(p.as_str().to_string())
                    }
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value = value_expr.ok_or_else(|| "dot_assign missing value".to_string())?;
            if members.is_empty() {
                return Err("dot_assign needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            // Build target: WithTarget.member1.member2...lastMember
            let mut obj = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::me_assign_statement => {
            // Me.prop1.prop2 = value
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => {
                        members.push(p.as_str().to_string())
                    }
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value =
                value_expr.ok_or_else(|| "me_assign_statement missing value".to_string())?;
            if members.is_empty() {
                return Err("me_assign_statement needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            let mut obj = Expression::new(ExprKind::This);
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::mybase_assign_statement => {
            // MyBase.prop = value
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut members: Vec<String> = Vec::new();
            let mut value_expr = None;
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => {
                        members.push(p.as_str().to_string())
                    }
                    Rule::expression => value_expr = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let value =
                value_expr.ok_or_else(|| "mybase_assign_statement missing value".to_string())?;
            if members.is_empty() {
                return Err("mybase_assign_statement needs at least one member".to_string());
            }
            let last = members.pop().unwrap();
            let mut obj = Expression::new(ExprKind::Super);
            for m in members {
                obj = Expression::new(ExprKind::Member {
                    object: Box::new(obj),
                    field: m,
                    null_safe: false,
                });
            }
            let target = Expression::new(ExprKind::Member {
                object: Box::new(obj),
                field: last,
                null_safe: false,
            });
            StmtKind::Assign {
                targets: vec![target],
                value,
            }
        }
        Rule::assign_statement => {
            let mut inner = pair.into_inner();
            // First child is l_value_expression
            let lhs_pair = inner.next().unwrap();
            let lhs_expr = parse_l_value_expression(lhs_pair)?;
            let value_expr = parse_expression(inner.next().unwrap())?;

            StmtKind::Assign {
                targets: vec![lhs_expr],
                value: value_expr,
            }
        }
        Rule::set_statement => {
            let mut inner = pair.into_inner();
            let target_name = inner.next().unwrap().as_str().to_string();
            let value = parse_expression(inner.next().unwrap())?;

            StmtKind::Assign {
                targets: vec![Expression::ident(&target_name)],
                value,
            }
        }
        Rule::lset_statement | Rule::rset_statement => {
            let is_right = pair.as_rule() == Rule::rset_statement;
            let mut inner = pair.into_inner();
            let target = parse_l_value_expression(inner.next().unwrap())?;
            let value = parse_expression(inner.next().unwrap())?;

            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(if is_right {
                    "__vb_rset_stmt"
                } else {
                    "__vb_lset_stmt"
                })),
                args: vec![Argument::positional(target), Argument::positional(value)],
                optional: false,
            }))
        }
        Rule::mid_assign_statement => {
            let mut inner = pair.into_inner();
            let target = parse_l_value_expression(inner.next().unwrap())?;
            let start = parse_expression(inner.next().unwrap())?;
            let mut trailing: Vec<_> = inner.collect();
            let value = parse_expression(
                trailing
                    .pop()
                    .ok_or_else(|| "Mid statement missing value".to_string())?,
            )?;
            let count = if let Some(count_pair) = trailing.pop() {
                parse_expression(count_pair)?
            } else {
                Expression::int(-1)
            };

            StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__vb_mid_stmt")),
                args: vec![
                    Argument::positional(target),
                    Argument::positional(start),
                    Argument::positional(count),
                    Argument::positional(value),
                ],
                optional: false,
            }))
        }
        Rule::dot_compound_assign_statement => {
            let mut identifiers = Vec::new();
            let mut op = None;
            let mut value = None;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => {
                        identifiers.push(p.as_str().to_string())
                    }
                    Rule::compound_assign_op => {
                        op = Some(match p.as_str() {
                            "+=" => CompoundOp::Add,
                            "-=" => CompoundOp::Sub,
                            "*=" => CompoundOp::Mul,
                            "/=" => CompoundOp::Div,
                            "\\=" => CompoundOp::IDiv,
                            "&=" => CompoundOp::Concat,
                            "^=" => CompoundOp::Pow,
                            "<<=" => CompoundOp::Shl,
                            ">>=" => CompoundOp::Shr,
                            _ => {
                                return Err(format!(
                                    "Unknown compound assignment operator: {}",
                                    p.as_str()
                                ));
                            }
                        });
                    }
                    Rule::expression => value = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            let mut target = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for part in identifiers {
                target = Expression::new(ExprKind::Member {
                    object: Box::new(target),
                    field: part,
                    null_safe: false,
                });
            }
            StmtKind::CompoundAssign {
                target,
                op: op.ok_or_else(|| "Missing compound assignment operator".to_string())?,
                value: value.ok_or_else(|| "Missing compound assignment value".to_string())?,
            }
        }
        Rule::compound_assign_statement => {
            let mut inner = pair.into_inner();
            let lhs_pair = inner.next().unwrap();
            let lhs_expr = parse_l_value_expression(lhs_pair)?;

            let op_pair = inner.next().unwrap();
            let op = match op_pair.as_str() {
                "+=" => CompoundOp::Add,
                "-=" => CompoundOp::Sub,
                "*=" => CompoundOp::Mul,
                "/=" => CompoundOp::Div,
                "\\=" => CompoundOp::IDiv,
                "&=" => CompoundOp::Concat,
                "^=" => CompoundOp::Pow,
                "<<=" => CompoundOp::Shl,
                ">>=" => CompoundOp::Shr,
                _ => {
                    return Err(format!(
                        "Unknown compound assignment operator: {}",
                        op_pair.as_str()
                    ));
                }
            };

            let value = parse_expression(inner.next().unwrap())?;

            StmtKind::CompoundAssign {
                target: lhs_expr,
                op,
                value,
            }
        }
        Rule::raiseevent_statement => {
            let mut inner = pair.into_inner();
            let event_name = inner.next().unwrap().as_str().to_string();
            let mut args = Vec::new();
            for p in inner {
                if p.as_rule() == Rule::argument_list {
                    for arg in p.into_inner() {
                        args.push(parse_expression(arg)?);
                    }
                }
            }
            StmtKind::RaiseEvent { event_name, args }
        }
        Rule::if_statement => return parse_if_statement(pair),
        Rule::single_line_if_statement => return parse_single_line_if(pair),
        Rule::for_each_statement => return parse_for_each_statement(pair),
        Rule::for_statement => return parse_for_statement(pair),
        Rule::while_statement => return parse_while_statement(pair),
        Rule::do_loop_statement => return parse_do_loop_statement(pair),
        Rule::with_statement => return parse_with_statement(pair),
        Rule::using_statement => return parse_using_statement(pair),
        Rule::exit_statement => {
            let mut inner = pair.into_inner();
            let exit_type = inner
                .next()
                .ok_or_else(|| "Exit statement missing type".to_string())?
                .as_str()
                .to_lowercase();

            match exit_type.as_str() {
                "sub" => StmtKind::Break(BreakTarget::Kind(ExitKind::Sub)),
                "function" => StmtKind::Break(BreakTarget::Kind(ExitKind::Function)),
                "for" => StmtKind::Break(BreakTarget::Kind(ExitKind::For)),
                "do" => StmtKind::Break(BreakTarget::Kind(ExitKind::Do)),
                "while" => StmtKind::Break(BreakTarget::Kind(ExitKind::While)),
                "select" => StmtKind::Break(BreakTarget::Kind(ExitKind::Select)),
                "try" => StmtKind::Break(BreakTarget::Kind(ExitKind::Try)),
                "property" => StmtKind::Break(BreakTarget::Kind(ExitKind::Property)),
                _ => return Err(format!("Unknown exit type: {}", exit_type)),
            }
        }
        Rule::try_statement => return parse_try_statement(pair),
        Rule::throw_statement => {
            let mut inner = pair.into_inner();
            let expr = inner.next().map(parse_expression).transpose()?;
            StmtKind::Throw { expr, cause: None }
        }
        Rule::yield_statement => {
            let mut inner = pair.into_inner();
            let value = inner
                .next()
                .map(parse_expression)
                .transpose()?
                .map(Box::new);
            StmtKind::Expr(Expression::new(ExprKind::Yield(value)))
        }
        Rule::continue_statement => return parse_continue_statement(pair),
        Rule::open_statement => return parse_open_statement(pair),
        Rule::close_statement => return parse_close_statement(pair),
        Rule::print_file_statement => return parse_print_file_statement(pair),
        Rule::write_file_statement => return parse_write_file_statement(pair),
        Rule::input_file_statement => return parse_input_file_statement(pair),
        Rule::line_input_statement => return parse_line_input_statement(pair),
        Rule::return_statement => {
            let mut inner = pair.into_inner();
            let value = inner.next().map(parse_expression).transpose()?;
            StmtKind::Return(value)
        }
        Rule::call_statement => {
            let mut inner = pair.into_inner();
            let mut first = inner.next().unwrap();

            // Skip optional Call keyword
            if first.as_rule() == Rule::call_keyword {
                first = inner.next().unwrap();
            }

            // Check if it's a member_call, member_access, call_expression, me_member_call, cast_member_call, or simple identifier
            match first.as_rule() {
                Rule::postfix
                | Rule::cast_member_call
                | Rule::me_member_call
                | Rule::mybase_member_call
                | Rule::member_call
                | Rule::member_access
                | Rule::call_expression => {
                    // Parse as expression and convert to statement
                    let expr = parse_expression(first)?;
                    StmtKind::Expr(expr)
                }
                Rule::identifier => {
                    // Could be: identifier, identifier(args), or identifier args
                    let name = first.as_str().to_string();
                    let arguments = inner
                        .next()
                        .map(|p| {
                            if p.as_rule() == Rule::argument_list {
                                parse_argument_list(p)
                            } else {
                                // Single expression argument
                                parse_expression(p).map(|e| vec![Argument::positional(e)])
                            }
                        })
                        .transpose()?
                        .unwrap_or_default();

                    StmtKind::Expr(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(&name)),
                        args: arguments,
                        optional: false,
                    }))
                }
                _ => {
                    let name = first.as_str().to_string();
                    StmtKind::Expr(Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::ident(&name)),
                        args: vec![],
                        optional: false,
                    }))
                }
            }
        }
        Rule::expression_statement => {
            let expr = parse_expression(pair.into_inner().next().unwrap())?;
            StmtKind::Expr(expr)
        }
        Rule::addhandler_statement => {
            let mut inner = pair.into_inner();
            let event_target = parse_event_target(inner.next().unwrap())?;
            let handler = parse_expression(unwrap_argument_expr_pair(inner.next().unwrap())?)?;
            let (control, event) = event_target;
            StmtKind::AddHandler {
                control,
                event,
                handler,
            }
        }
        Rule::removehandler_statement => {
            let mut inner = pair.into_inner();
            let event_target = parse_event_target(inner.next().unwrap())?;
            let handler = parse_expression(unwrap_argument_expr_pair(inner.next().unwrap())?)?;
            let (control, event) = event_target;
            StmtKind::RemoveHandler {
                control,
                event,
                handler,
            }
        }
        Rule::static_statement => {
            let mut name = String::new();
            let mut var_type: Option<String> = None;
            let mut initializer = None;
            for p in pair.into_inner() {
                match p.as_rule() {
                    Rule::identifier => name = p.as_str().to_string(),
                    Rule::type_name => var_type = Some(p.as_str().to_string()),
                    Rule::expression => initializer = Some(parse_expression(p)?),
                    _ => {}
                }
            }
            StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: var_type,
                    init: initializer,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Static,
            }
        }
        Rule::goto_statement => {
            let label = pair.into_inner().next().unwrap().as_str().to_string();
            StmtKind::GoTo(label)
        }
        Rule::label_statement => {
            let label = pair.into_inner().next().unwrap().as_str().to_string();
            StmtKind::Label(label)
        }
        Rule::on_error_statement => {
            let text = pair.as_str().to_lowercase();
            if text.contains("resume") && text.contains("next") {
                StmtKind::OnErrorResumeNext
            } else {
                // On Error GoTo <label> or On Error GoTo 0
                let inner = pair.into_inner();
                let target = inner
                    .last()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_else(|| "0".to_string());
                StmtKind::OnErrorGoTo(target)
            }
        }
        Rule::resume_statement => {
            // Resume → Empty (simplified in common AST)
            StmtKind::Empty
        }
        // New declarations — parse gracefully as no-op statements for now
        Rule::interface_decl
        | Rule::structure_decl
        | Rule::event_decl
        | Rule::delegate_sub_decl
        | Rule::delegate_function_decl => StmtKind::Empty,
        Rule::namespace_decl => StmtKind::Empty,
        Rule::synclock_statement => return parse_synclock_statement(pair),
        _ => return Err(format!("Unexpected rule: {:?}", pair.as_rule())),
    };
    Ok(Statement::with_span(kind, span))
}

fn parse_if_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;
    let mut then_body = Vec::new();
    let mut elifs: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut else_body = None;

    for p in inner {
        match p.as_rule() {
            Rule::if_body => {
                if then_body.is_empty() {
                    then_body = parse_block(p)?;
                }
            }
            Rule::elseif_block => {
                let mut elseif_condition = None;
                let mut elseif_body = Vec::new();
                for p_inner in p.into_inner() {
                    match p_inner.as_rule() {
                        Rule::expression => elseif_condition = Some(parse_expression(p_inner)?),
                        Rule::if_body => {
                            elseif_body = parse_block(p_inner)?;
                            break;
                        }
                        _ => {}
                    }
                }
                if let Some(cond) = elseif_condition {
                    elifs.push((cond, elseif_body));
                }
            }
            Rule::else_block => {
                let mut body = Vec::new();
                for p_inner in p.into_inner() {
                    if p_inner.as_rule() == Rule::if_body {
                        body = parse_block(p_inner)?;
                        break;
                    }
                }
                else_body = Some(body);
            }
            Rule::NEWLINE | Rule::if_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        },
        span,
    ))
}

fn parse_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut statements = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    statements.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::statement => {
                statements.push(parse_statement(p)?);
            }
            Rule::NEWLINE | Rule::EOI => {}
            _ => {}
        }
    }
    Ok(statements)
}

/// Monotonic sequence for VB `For` loop-local temp names (`__vb_for_limit_N`,
/// `__vb_for_step_N`), keeping nested/sequential loops' hoisted temps distinct.
static FOR_TEMP_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn parse_for_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let variable = inner.next().unwrap().as_str().to_string();

    // Skip optional 'As type_name'
    let mut variable_type = None;
    let mut next = inner.next().unwrap();
    if next.as_rule() == Rule::type_name {
        variable_type = Some(next.as_str().to_string());
        next = inner.next().unwrap();
    }
    let start = parse_expression(next)?;
    let end = parse_expression(inner.next().unwrap())?;

    let mut step = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::expression => step = Some(parse_expression(p)?),
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::NEWLINE | Rule::for_end => {}
            _ => body.push(parse_statement(p)?),
        }
    }

    let step_val = step.unwrap_or_else(|| Expression::int(1));
    // VB evaluates the `To` limit and `Step` EXACTLY ONCE at loop entry — unlike
    // a C-style `for`, whose condition re-runs every iteration. Emitting `end`
    // directly into the condition re-evaluates it each pass; when it has side
    // effects that raise the ceiling (`For i = 1 To GetLimit()` where GetLimit
    // increments), the loop never terminates. So hoist the limit — and a
    // non-literal step — into loop-local temps evaluated once in the init.
    // A literal step is inlined (no side effects, and it drives the direction
    // choice above).
    let seq = FOR_TEMP_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let limit_name = format!("__vb_for_limit_{seq}");
    let mut declarations = vec![
        VarDeclarator {
            pattern: BindingPattern::Ident(variable.clone()),
            type_hint: variable_type,
            init: Some(start),
            array_bounds: None,
            with_events: false,
        },
        VarDeclarator {
            pattern: BindingPattern::Ident(limit_name.clone()),
            type_hint: None,
            init: Some(end),
            array_bounds: None,
            with_events: false,
        },
    ];

    let step_operand = if matches!(
        &step_val.kind,
        ExprKind::Lit(Literal::Int(_))
            | ExprKind::Lit(Literal::Float(_))
            | ExprKind::Unary {
                op: UnaryOp::Neg,
                ..
            }
    ) {
        step_val
    } else {
        let step_name = format!("__vb_for_step_{seq}");
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(step_name.clone()),
            type_hint: None,
            init: Some(step_val),
            array_bounds: None,
            with_events: false,
        });
        Expression::ident(&step_name)
    };

    // Convert VB For to C-style For:
    // init: Dim variable = start, __vb_for_limit = end [, __vb_for_step = step]
    // cond: step >= 0 ? variable <= __vb_for_limit : variable >= __vb_for_limit
    // update: variable = variable + step
    let init = Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Dim,
    });

    let step_nonnegative = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(step_operand.clone()),
        right: Box::new(Expression::int(0)),
    });
    let positive_cond = Expression::new(ExprKind::Binary {
        op: BinOp::LtEq,
        left: Box::new(Expression::ident(&variable)),
        right: Box::new(Expression::ident(&limit_name)),
    });
    let negative_cond = Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(Expression::ident(&variable)),
        right: Box::new(Expression::ident(&limit_name)),
    });
    let cond = Expression::new(ExprKind::Ternary {
        cond: Box::new(step_nonnegative),
        then: Box::new(positive_cond),
        else_: Box::new(negative_cond),
    });
    let update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&variable)),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&variable)),
            right: Box::new(step_operand),
        })),
    });

    Ok(Statement::with_span(
        StmtKind::For {
            init: Some(Box::new(init)),
            cond: Some(cond),
            update: Some(update),
            body,
        },
        span,
    ))
}

fn parse_while_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::NEWLINE | Rule::while_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::While {
            cond,
            body,
            else_body: None,
        },
        span,
    ))
}

fn parse_do_loop_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();
    let mut pre_condition: Option<(bool, Expression)> = None; // (is_until, expr)
    let mut post_condition: Option<(bool, Expression)> = None;
    let mut body = Vec::new();
    let mut current_is_until = false;

    for p in inner {
        match p.as_rule() {
            Rule::do_while_kw => current_is_until = false,
            Rule::do_until_kw => current_is_until = true,
            Rule::expression => {
                // Determine if it's pre or post condition based on position
                if body.is_empty() {
                    pre_condition = Some((current_is_until, parse_expression(p)?));
                } else {
                    post_condition = Some((current_is_until, parse_expression(p)?));
                }
            }
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::do_end => {
                // Parse post-condition from do_end children (Loop While/Until)
                for dp in p.into_inner() {
                    match dp.as_rule() {
                        Rule::do_while_kw => current_is_until = false,
                        Rule::do_until_kw => current_is_until = true,
                        Rule::expression => {
                            post_condition = Some((current_is_until, parse_expression(dp)?));
                        }
                        _ => {}
                    }
                }
            }
            Rule::NEWLINE => {}
            _ => {}
        }
    }

    // Map to common AST:
    // If there's a pre_condition: it's a While loop (with condition potentially inverted for Until)
    // If there's a post_condition: it's a DoWhile loop
    // If neither: infinite loop (DoWhile with true condition)
    if let Some((is_until, cond)) = post_condition {
        // Mixed forms like `Do Until a ... Loop While b` use the trailing
        // Loop condition as the continuation condition after the first pass.
        Ok(Statement::with_span(
            StmtKind::DoWhile {
                body,
                cond,
                until: is_until,
            },
            span,
        ))
    } else if let Some((is_until, cond)) = pre_condition {
        // Do While/Until <cond> ... Loop → While loop
        let effective_cond = if is_until {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(cond),
            })
        } else {
            cond
        };
        Ok(Statement::with_span(
            StmtKind::While {
                cond: effective_cond,
                body,
                else_body: None,
            },
            span,
        ))
    } else {
        // Do ... Loop (infinite)
        Ok(Statement::with_span(
            StmtKind::DoWhile {
                body,
                cond: Expression::bool(true),
                until: false,
            },
            span,
        ))
    }
}

fn parse_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut pair = pair;
    loop {
        let span = to_span(&pair);
        let kind = match pair.as_rule() {
            Rule::argument | Rule::first_argument | Rule::trailing_argument => {
                let inner = pair.into_inner().next().unwrap();
                pair = inner;
                continue;
            }
            Rule::named_argument => {
                let mut inner = pair.into_inner();
                let _name = inner.next();
                pair = inner
                    .next()
                    .ok_or_else(|| "Named argument missing value".to_string())?;
                continue;
            }
            Rule::expression
            | Rule::logical_imp
            | Rule::logical_eqv
            | Rule::logical_xor
            | Rule::logical_or
            | Rule::logical_and
            | Rule::equality
            | Rule::comparison
            | Rule::bit_shift
            | Rule::additive
            | Rule::multiplicative
            | Rule::exponent => {
                let mut probe = pair.clone().into_inner();
                let first = probe.next().unwrap();
                if probe.next().is_none() {
                    pair = first;
                    continue;
                }
                return parse_binary_expression(pair);
            }
            Rule::not_condition => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                if first.as_rule() == Rule::not_op {
                    let operand = parse_expression(inner.next().unwrap())?;
                    ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(operand),
                    }
                } else {
                    pair = first;
                    continue;
                }
            }
            Rule::lambda_expression => return parse_lambda_expression(pair),
            Rule::nameof_expression => {
                let name = pair
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::dotted_identifier)
                    .map(|p| {
                        let text = p.as_str();
                        text.rsplit('.').next().unwrap_or(text).to_string()
                    })
                    .unwrap_or_default();
                ExprKind::Lit(Literal::Str(name))
            }
            Rule::gettype_expression => {
                let type_name = pair
                    .into_inner()
                    .find(|p| p.as_rule() == Rule::type_name)
                    .map(|p| p.as_str().trim().to_string())
                    .unwrap_or_default();
                ExprKind::TypeOf(Box::new(build_dotted_expr(&vb_gettype_type_name(
                    &type_name,
                ))))
            }
            Rule::typeof_expression => {
                let is_not = pair.as_str().to_ascii_lowercase().contains(" isnot ");
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().trim().to_string();
                let is_type = Expression::new(ExprKind::IsType {
                    expr: Box::new(expr),
                    type_name,
                });
                if is_not {
                    ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(is_type),
                    }
                } else {
                    is_type.kind
                }
            }
            Rule::unary => return parse_unary_expression(pair),
            Rule::postfix => return parse_postfix_expression(pair),
            Rule::call_expression => {
                let call_text = pair.as_str().to_string();
                let mut name = String::new();
                let mut generic_target_type = None;
                let mut arguments = Vec::new();
                for part in pair.into_inner() {
                    match part.as_rule() {
                        Rule::identifier => name = strip_vb_generic_suffix(part.as_str()),
                        Rule::generic_suffix => {
                            generic_target_type = vb_generic_suffix_first_type(part.as_str())
                        }
                        Rule::argument_list => arguments = parse_argument_list(part)?,
                        _ => {}
                    }
                }

                if name.eq_ignore_ascii_case("CTypeDynamic") {
                    let value = arguments
                        .into_iter()
                        .next()
                        .map(|arg| arg.value)
                        .unwrap_or_else(Expression::null);
                    let type_name = generic_target_type
                        .or_else(|| vb_call_generic_first_type(&call_text))
                        .unwrap_or_else(|| "Object".to_string());
                    return Ok(Expression::with_span(
                        ExprKind::Cast {
                            expr: Box::new(value),
                            type_name,
                        },
                        span,
                    ));
                }

                if arguments.is_empty() {
                    if let Some(type_name) = generic_target_type {
                        return Ok(Expression::with_span(
                            ExprKind::Call {
                                callee: Box::new(Expression::ident(&vb_generic_call_marker(
                                    &name, &type_name,
                                ))),
                                args: Vec::new(),
                                optional: false,
                            },
                            span,
                        ));
                    }
                }

                if let Some(rewritten) = canonicalize_call(&name, &arguments) {
                    return Ok(rewritten);
                }

                ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::generic_type_reference => {
                let raw = pair.as_str();
                let type_name = if raw
                    .trim_start()
                    .to_ascii_lowercase()
                    .starts_with("ctypedynamic(of")
                {
                    raw.trim().to_string()
                } else if let Some(marker) = vb_generic_type_marker(raw) {
                    return Ok(Expression::with_span(ExprKind::Ident(marker), span));
                } else {
                    strip_vb_generic_suffix(raw)
                };
                build_dotted_expr(&type_name).kind
            }
            Rule::member_call => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }

                return Ok(expr);
            }
            Rule::tuple_literal => {
                let mut fields = Vec::new();
                for element in pair.into_inner() {
                    if element.as_rule() != Rule::tuple_element {
                        continue;
                    }
                    let mut explicit_name = None;
                    let mut value = None;
                    for part in element.into_inner() {
                        match part.as_rule() {
                            Rule::identifier => {
                                explicit_name = Some(normalize_vb_identifier(part.as_str()));
                            }
                            Rule::expression => {
                                value = Some(parse_expression(part)?);
                            }
                            _ => {}
                        }
                    }
                    let value = value.ok_or_else(|| "Tuple element missing value".to_string())?;
                    let name = explicit_name.or_else(|| vb_infer_tuple_element_name(&value));
                    fields.push((name, value));
                }
                ExprKind::NamedTuple {
                    fields,
                    type_name: None,
                }
            }
            Rule::member_access => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for p in inner {
                    expr = canonicalize_member_access(expr, &normalize_vb_identifier(p.as_str()));
                }

                return Ok(expr);
            }
            Rule::aggregate_expression => return parse_aggregate_expression(pair),
            Rule::query_expression => return parse_query_expression(pair),
            Rule::xml_literal => return parse_xml_literal(pair),
            Rule::identifier | Rule::member_identifier => {
                let name = normalize_vb_identifier(pair.as_str());
                if let Some(rewritten) = canonicalize_special_identifier(&name) {
                    return Ok(rewritten);
                }
                ExprKind::Ident(name)
            }
            Rule::cast_expression => {
                let text = pair.as_str();
                let cast_kind = if text.len() >= 10 && text[..10].eq_ignore_ascii_case("DirectCast")
                {
                    "DirectCast"
                } else if text.len() >= 7 && text[..7].eq_ignore_ascii_case("TryCast") {
                    "TryCast"
                } else {
                    "CType"
                };
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().to_string();
                let expr = rewrite_vb_array_clone_for_cast(expr, &type_name);
                let full_type = if cast_kind != "CType" {
                    format!("{}:{}", cast_kind, type_name)
                } else {
                    type_name
                };
                ExprKind::Cast {
                    expr: Box::new(expr),
                    type_name: full_type,
                }
            }
            Rule::cast_member_call => {
                let mut inner = pair.into_inner();
                let cast_pair = inner.next().unwrap();
                let mut expr = parse_expression(cast_pair)?;
                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }
                return Ok(expr);
            }
            Rule::interpolated_string => {
                let s = pair.as_str();
                let inner_str = s[2..s.len() - 1].replace("\"\"", "\"");
                let mut parts = Vec::new();
                let mut current_text = String::new();
                let mut chars = inner_str.chars().peekable();

                while let Some(ch) = chars.next() {
                    if ch == '{' {
                        if chars.peek() == Some(&'{') {
                            chars.next();
                            current_text.push('{');
                            continue;
                        }
                        if !current_text.is_empty() {
                            parts.push(InterpolPart::Text(current_text.clone()));
                            current_text.clear();
                        }
                        let mut expr_text = String::new();
                        let mut depth = 1;
                        while let Some(c) = chars.next() {
                            if c == '{' {
                                depth += 1;
                            }
                            if c == '}' {
                                depth -= 1;
                                if depth == 0 {
                                    break;
                                }
                            }
                            expr_text.push(c);
                        }
                        match parse_expression_str(&expr_text) {
                            Ok(expr) => parts.push(InterpolPart::Expr(expr)),
                            Err(_) => {
                                parts.push(InterpolPart::Expr(Expression::ident(expr_text.trim())))
                            }
                        }
                    } else if ch == '}' {
                        if chars.peek() == Some(&'}') {
                            chars.next();
                            current_text.push('}');
                        }
                    } else {
                        current_text.push(ch);
                    }
                }

                if !current_text.is_empty() {
                    parts.push(InterpolPart::Text(current_text));
                }

                if parts.is_empty() {
                    ExprKind::Lit(Literal::Str(String::new()))
                } else if parts.len() == 1 {
                    match parts.into_iter().next().unwrap() {
                        InterpolPart::Text(s) => ExprKind::Lit(Literal::Str(s)),
                        InterpolPart::Expr(expr) => return Ok(expr),
                        InterpolPart::Formatted(expr, _) => return Ok(expr),
                    }
                } else {
                    ExprKind::Interpolation(parts)
                }
            }
            Rule::string_literal => {
                let s = pair
                    .as_str()
                    .trim_end_matches(|c: char| c == 'c' || c == 'C');
                ExprKind::Lit(Literal::Str(s[1..s.len() - 1].replace("\"\"", "\"")))
            }
            Rule::numeric_literal => {
                return Ok(parse_vb_numeric_literal(pair.as_str()));
            }
            Rule::boolean_literal => {
                ExprKind::Lit(Literal::Bool(pair.as_str().eq_ignore_ascii_case("true")))
            }
            Rule::array_literal => {
                let elements = pair
                    .into_inner()
                    .filter(|p| p.as_rule() == Rule::expression)
                    .map(|p| {
                        parse_expression(p).map(|value| ArrayElement {
                            key: None,
                            value,
                            spread: false,
                            by_ref: false,
                        })
                    })
                    .collect::<Result<Vec<_>, _>>()?;
                ExprKind::Array(elements)
            }
            Rule::date_literal => {
                let s = pair.as_str();
                ExprKind::Cast {
                    expr: Box::new(Expression::string(&s[1..s.len() - 1].trim())),
                    type_name: "Date".to_string(),
                }
            }
            Rule::nothing_literal => ExprKind::Lit(Literal::Null),
            Rule::anonymous_new_expression => return parse_anonymous_new_expression(pair),
            Rule::new_expression => {
                let raw_new_text = pair.as_str().to_string();
                let mut inner = pair.into_inner();
                let id_pair = inner.next().unwrap();
                let mut class_name = id_pair.as_str().to_string();
                let mut args = Vec::new();
                let mut array_init: Option<Vec<Expression>> = None;
                for p in inner {
                    match p.as_rule() {
                        Rule::generic_suffix => class_name.push_str(p.as_str()),
                        Rule::argument_list => args = parse_argument_list(p)?,
                        Rule::array_literal => {
                            let elements = p
                                .into_inner()
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            array_init = Some(elements);
                        }
                        Rule::from_initializer => {
                            let elements = p
                                .into_inner()
                                .filter(|e| e.as_rule() == Rule::expression)
                                .map(parse_expression)
                                .collect::<Result<Vec<_>, _>>()?;
                            return Ok(Expression::with_span(
                                emit_vb_collection_init_iife(
                                    Expression::new(ExprKind::New {
                                        class: Box::new(Expression::ident(&class_name)),
                                        args,
                                    }),
                                    elements,
                                )
                                .kind,
                                span,
                            ));
                        }
                        Rule::with_initializer => {
                            let mut members = Vec::new();
                            for mi in p.into_inner() {
                                if mi.as_rule() != Rule::member_initializer {
                                    continue;
                                }
                                let mut mi_inner = mi.into_inner();
                                let prop_name =
                                    mi_inner.next().unwrap().as_str().to_ascii_lowercase();
                                let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                                members.push((prop_name, prop_expr));
                            }
                            return Ok(emit_vb_object_init_iife(
                                Expression::with_span(
                                    ExprKind::New {
                                        class: Box::new(Expression::ident(&class_name)),
                                        args,
                                    },
                                    span,
                                ),
                                members,
                            ));
                        }
                        _ => {}
                    }
                }
                if let Some(elements) = array_init {
                    if elements.is_empty() && args.len() == 1 {
                        vb_filled_array_expr(
                            vb_array_length_from_upper_bound(args[0].value.clone()),
                            vb_default_value_for_type(&class_name),
                        )
                        .kind
                    } else if elements.is_empty() {
                        if let Some((element_type, upper_bound)) =
                            vb_new_array_bound_from_type_text(&class_name)
                                .or_else(|| vb_new_array_bound_from_new_expr_text(&raw_new_text))
                        {
                            vb_filled_array_expr(
                                vb_array_length_from_upper_bound(upper_bound),
                                vb_default_value_for_type(&element_type),
                            )
                            .kind
                        } else {
                            ExprKind::Array(vec![])
                        }
                    } else {
                        ExprKind::Array(
                            elements
                                .into_iter()
                                .map(|value| ArrayElement {
                                    key: None,
                                    value,
                                    spread: false,
                                    by_ref: false,
                                })
                                .collect(),
                        )
                    }
                } else {
                    ExprKind::New {
                        class: Box::new(Expression::ident(&class_name)),
                        args,
                    }
                }
            }
            Rule::if_expression => {
                let mut inner = pair.into_inner();
                let first = parse_expression(inner.next().unwrap())?;
                let second = parse_expression(inner.next().unwrap())?;
                let third = inner.next().map(parse_expression).transpose()?;
                match third {
                    Some(else_expr) => ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    },
                    None => ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    },
                }
            }
            Rule::addressof_expr => {
                let mut name = String::new();
                for p in pair.into_inner() {
                    if p.as_rule() == Rule::dotted_identifier {
                        name = p.as_str().to_string();
                    }
                }
                ExprKind::AddressOf(name)
            }
            Rule::me_keyword => ExprKind::This,
            Rule::dot_call_statement => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::identifier | Rule::member_identifier => {
                            identifiers.push(p.as_str().to_string())
                        }
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("dot_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::dot_member_access => {
                let inner = pair.into_inner();
                let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_access => {
                let mut inner = pair.into_inner();
                let _me = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::This);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::mybase_member_access => {
                let mut inner = pair.into_inner();
                let _mybase = inner.next().unwrap();
                let mut expr = Expression::new(ExprKind::Super);
                for p in inner {
                    if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                        expr = Expression::new(ExprKind::Member {
                            object: Box::new(expr),
                            field: p.as_str().to_string(),
                            null_safe: false,
                        });
                    }
                }
                return Ok(expr);
            }
            Rule::me_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::me_keyword => {}
                        Rule::identifier | Rule::member_identifier => {
                            identifiers.push(p.as_str().to_string())
                        }
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("me_member_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::This);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::mybase_member_call => {
                let inner = pair.into_inner();
                let mut identifiers = Vec::new();
                let mut arguments = Vec::new();
                for p in inner {
                    match p.as_rule() {
                        Rule::mybase_keyword => {}
                        Rule::identifier | Rule::member_identifier => {
                            identifiers.push(p.as_str().to_string())
                        }
                        Rule::argument_list => arguments = parse_argument_list(p)?,
                        _ => {}
                    }
                }
                if identifiers.is_empty() {
                    return Err("mybase_member_call needs at least one identifier".to_string());
                }
                let method_name = identifiers.last().unwrap().clone();
                let mut expr = Expression::new(ExprKind::Super);
                for part in identifiers.iter().take(identifiers.len() - 1) {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: part.clone(),
                        null_safe: false,
                    });
                }
                if identifiers.len() == 1 {
                    ExprKind::SuperCall {
                        method: Some(method_name),
                        args: arguments,
                    }
                } else {
                    let callee = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: method_name,
                        null_safe: false,
                    });
                    ExprKind::Call {
                        callee: Box::new(callee),
                        args: arguments,
                        optional: false,
                    }
                }
            }
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };
        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut operands = vec![parse_expression(first)?];
    let mut ops = Vec::new();

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op
            | Rule::mult_op
            | Rule::eq_op
            | Rule::comp_op
            | Rule::and_op
            | Rule::or_op
            | Rule::xor_op
            | Rule::eqv_op
            | Rule::imp_op
            | Rule::shift_op
            | Rule::like_op
            | Rule::exp_op => match op_pair.as_str().to_lowercase().as_str() {
                "+" => BinOp::Add,
                "-" => BinOp::Sub,
                "*" => BinOp::Mul,
                "/" => BinOp::Div,
                "\\" => BinOp::IDiv,
                "mod" => BinOp::Mod,
                "^" => BinOp::Pow,
                "&" => BinOp::Concat,
                "=" => BinOp::Eq,
                "<>" => BinOp::NotEq,
                "<" => BinOp::Lt,
                "<=" => BinOp::LtEq,
                ">" => BinOp::Gt,
                ">=" => BinOp::GtEq,
                "andalso" => BinOp::And,
                "orelse" => BinOp::Or,
                "and" => BinOp::BitAnd,
                "or" => BinOp::BitOr,
                "xor" => BinOp::BitXor,
                "eqv" => BinOp::Eqv,
                "imp" => BinOp::Imp,
                "<<" => BinOp::Shl,
                ">>" => BinOp::Shr,
                "is" => BinOp::Is,
                "isnot" => BinOp::IsNot,
                "like" => BinOp::Like,
                _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
            },
            _ => return Ok(operands.pop().unwrap()),
        };

        let right_pair = inner.next().unwrap();
        ops.push(op);
        operands.push(parse_expression(right_pair)?);
    }

    if ops.is_empty() {
        return Ok(operands.pop().unwrap());
    }

    if ops.iter().all(|op| *op == BinOp::Pow) {
        let mut expr = operands.pop().unwrap();
        while let Some(left) = operands.pop() {
            expr = Expression::new(ExprKind::Binary {
                op: BinOp::Pow,
                left: Box::new(left),
                right: Box::new(expr),
            });
        }
        return Ok(expr);
    }

    let mut operands = operands.into_iter();
    let mut left = operands.next().unwrap();
    for (op, right) in ops.into_iter().zip(operands) {
        left = maybe_rewrite_vb_binary(op, left, right);
    }

    Ok(left)
}

/*
fn parse_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut pair = pair;
    loop {
        let span = to_span(&pair);
        let kind = match pair.as_rule() {
            Rule::expression | Rule::logical_xor | Rule::logical_or | Rule::logical_and |
            Rule::equality | Rule::comparison | Rule::bit_shift | Rule::additive |
            Rule::multiplicative | Rule::exponent => {
                let mut probe = pair.clone().into_inner();
                let first = probe.next().unwrap();
                if probe.next().is_none() {
                    pair = first;
                    continue;
                }
                return parse_binary_expression(pair);
            }
            Rule::not_condition => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                if first.as_rule() == Rule::not_op {
                    let operand = parse_expression(inner.next().unwrap())?;
                    ExprKind::Unary { op: UnaryOp::Not, expr: Box::new(operand) }
                } else {
                    pair = first;
                    continue;
                }
            }
            Rule::lambda_expression => {
                return parse_lambda_expression(pair);
            }
            Rule::typeof_expression => {
                let mut inner = pair.into_inner();
                let expr = parse_expression(inner.next().unwrap())?;
                let type_name = inner.next().unwrap().as_str().trim().to_string();
                ExprKind::IsType {
                    expr: Box::new(expr),
                    type_name,
                }
            }
            Rule::unary => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                match first.as_rule() {
                    Rule::neg_op => {
                        let operand = parse_expression(inner.next().unwrap())?;
                        ExprKind::Unary {
                            op: UnaryOp::Neg,
                            expr: Box::new(operand),
                        }
                    }
                    Rule::await_op => {
                        let operand = parse_expression(inner.next().unwrap())?;
                        ExprKind::Await(Box::new(operand))
                    }
                    _ => {
                        pair = first;
                        continue;
                    }
                }
            }
            Rule::postfix => {
                let mut inner = pair.into_inner();
                let primary = inner.next().unwrap();
                let Some(first_chain) = inner.next() else {
                    pair = primary;
                    continue;
                };

                let mut expr = parse_expression(primary)?;
                expr = parse_member_chain_node(first_chain, expr)?;
                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }
                return Ok(expr);
            }
            Rule::call_expression => {
                let call_text = pair.as_str().to_string();
                let mut name = String::new();
                let mut generic_target_type = None;
                let mut arguments = Vec::new();
                for part in pair.into_inner() {
                    match part.as_rule() {
                        Rule::identifier => name = strip_vb_generic_suffix(part.as_str()),
                        Rule::generic_suffix => {
                            generic_target_type = vb_generic_suffix_first_type(part.as_str())
                        }
                        Rule::argument_list => arguments = parse_argument_list(part)?,
                        _ => {}
                    }
                }

                if name.eq_ignore_ascii_case("CTypeDynamic") {
                    let value = arguments
                        .into_iter()
                        .next()
                        .map(|arg| arg.value)
                        .unwrap_or_else(Expression::null);
                    let type_name = generic_target_type
                        .or_else(|| vb_call_generic_first_type(&call_text))
                        .unwrap_or_else(|| "Object".to_string());
                    return Ok(Expression::with_span(
                        ExprKind::Cast {
                            expr: Box::new(value),
                            type_name,
                        },
                        span,
                    ));
                }

                if arguments.is_empty() {
                    if let Some(type_name) = generic_target_type {
                        return Ok(Expression::with_span(
                            ExprKind::Call {
                                callee: Box::new(Expression::ident(&vb_generic_call_marker(
                                    &name,
                                    &type_name,
                                ))),
                                args: Vec::new(),
                                optional: false,
                            },
                            span,
                        ));
                    }
                }

                if let Some(rewritten) = canonicalize_call(&name, &arguments) {
                    return Ok(rewritten);
                }

                ExprKind::Call {
                    callee: Box::new(Expression::ident(&name)),
                    args: arguments,
                    optional: false,
                }
            }
            Rule::member_call => {
                let mut inner = pair.into_inner();
                let first = inner.next().unwrap();
                let mut expr = Expression::with_span(
                    ExprKind::Ident(normalize_vb_identifier(first.as_str())),
                    to_span(&first),
                );

                for chain in inner {
                    expr = parse_member_chain_node(chain, expr)?;
                }

                return Ok(expr);
            }
            Rule::member_access => {

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let _span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "andalso" => BinOp::And,
                    "orelse" => BinOp::Or,
                    "and" => BinOp::BitAnd,
                    "or" => BinOp::BitOr,
                    "xor" => BinOp::BitXor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left),
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}
                    }
                } else if ch == '}' {
                    // Check for }} escape (literal brace)
                    if chars.peek() == Some(&'}') {
                        chars.next();
                        current_text.push('}');
                    }
                } else {
                    current_text.push(ch);
                }
            }
            // Flush remaining text
            if !current_text.is_empty() {
                parts.push(InterpolPart::Text(current_text));
            }

            if parts.is_empty() {
                ExprKind::Lit(Literal::Str(String::new()))
            } else if parts.len() == 1 {
                match parts.into_iter().next().unwrap() {
                    InterpolPart::Text(s) => ExprKind::Lit(Literal::Str(s)),
                    InterpolPart::Expr(e) => return Ok(e),
                    InterpolPart::Formatted(e, _) => return Ok(e),
                }
            } else {
                ExprKind::Interpolation(parts)
            }
        }
        Rule::string_literal => {
            let s = pair.as_str().trim_end_matches(|c: char| c == 'c' || c == 'C');
            // Strip outer quotes, then unescape VB-style doubled quotes ("" -> ")
            let inner = s[1..s.len()-1].replace("\"\"", "\"");
            ExprKind::Lit(Literal::Str(inner))
        }
        Rule::numeric_literal => {
            return Ok(parse_vb_numeric_literal(pair.as_str()));
        }
        Rule::boolean_literal => {
            ExprKind::Lit(Literal::Bool(pair.as_str().to_lowercase() == "true"))
        }
        Rule::array_literal => {
            return parse_array_literal(pair);
        }
        Rule::date_literal => {
            let s = pair.as_str();
            // Strip the surrounding # delimiters
            let inner = s[1..s.len()-1].trim().to_string();
            ExprKind::Cast {
                expr: Box::new(Expression::string(&inner)),
                type_name: "Date".to_string(),
            }
        }
        Rule::nothing_literal => ExprKind::Lit(Literal::Null),
        Rule::anonymous_new_expression => return parse_anonymous_new_expression(pair),
        Rule::new_expression => {
            let raw_new_text = pair.as_str().to_string();
            let mut inner = pair.into_inner();
            let id_pair = inner.next().unwrap();
            let mut class_name = id_pair.as_str().to_string();
            let mut args: Vec<Argument> = Vec::new();
            let mut array_init: Option<Vec<Expression>> = None;
            for p in inner {
                match p.as_rule() {
                    Rule::generic_suffix => consume_vb_generic_suffix(p.as_str()),
                    Rule::argument_list => {
                        args = parse_argument_list(p)?;
                    }
                    Rule::array_literal => {
                        // New Type() {elem1, elem2, ...} → array initializer
                        let elements: Vec<Expression> = p.into_inner()
                            .filter(|e| e.as_rule() == Rule::expression)
                            .map(|e| parse_expression(e))
                            .collect::<Result<Vec<_>, _>>()?;
                        array_init = Some(elements);
                    }
                    Rule::from_initializer => {
                        // New List(Of T) From { expr, expr, ... }
                        let elements: Vec<Expression> = p.into_inner()
                            .filter(|e| e.as_rule() == Rule::expression)
                            .map(|e| parse_expression(e))
                            .collect::<Result<Vec<_>, _>>()?;
                        return Ok(Expression::with_span(
                            emit_vb_collection_init_iife(
                                Expression::new(ExprKind::New {
                                    class: Box::new(Expression::ident(&class_name)),
                                    args,
                                }),
                                elements,
                            )
                            .kind,
                            span,
                        ));
                    }
                    Rule::with_initializer => {
                        // New Type() With { .Prop = expr, ... }
                        let mut members = Vec::new();
                        for mi in p.into_inner() {
                            if mi.as_rule() != Rule::member_initializer { continue; }
                            let mut mi_inner = mi.into_inner();
                            let prop_name = mi_inner.next().unwrap().as_str().to_ascii_lowercase();
                            let prop_expr = parse_expression(mi_inner.next().unwrap())?;
                            members.push((prop_name, prop_expr));
                        }
                        return Ok(emit_vb_object_init_iife(Expression::with_span(ExprKind::New {
                            class: Box::new(Expression::ident(&class_name)),
                            args,
                        }, span), members));
                    }
                    _ => {}
                }
            }
            // If there's an array initializer, return an Array instead of New
            if let Some(elements) = array_init {
                if elements.is_empty() && args.len() == 1 {
                    vb_filled_array_expr(
                        vb_array_length_from_upper_bound(args[0].value.clone()),
                        vb_default_value_for_type(&class_name),
                    )
                    .kind
                } else if elements.is_empty() {
                    if let Some((element_type, upper_bound)) =
                        vb_new_array_bound_from_type_text(&class_name)
                            .or_else(|| vb_new_array_bound_from_new_expr_text(&raw_new_text))
                    {
                        vb_filled_array_expr(
                            vb_array_length_from_upper_bound(upper_bound),
                            vb_default_value_for_type(&element_type),
                        )
                        .kind
                    } else {
                        ExprKind::Array(vec![])
                    }
                } else {
                    ExprKind::Array(
                        elements
                            .into_iter()
                            .map(|e| ArrayElement {
                                key: None,
                                value: e,
                                spread: false,
                                by_ref: false,
                            })
                            .collect(),
                    )
                }
            } else {
                ExprKind::New {
                    class: Box::new(Expression::ident(&class_name)),
                    args,
                }
            }
        }
        Rule::if_expression => {
            let mut inner = pair.into_inner();
            let first = parse_expression(inner.next().unwrap())?;
            let second = parse_expression(inner.next().unwrap())?;
            let third = inner.next().map(|p| parse_expression(p)).transpose()?;
            match third {
                Some(else_expr) => {
                    ExprKind::Ternary {
                        cond: Box::new(first),
                        then: Box::new(second),
                        else_: Box::new(else_expr),
                    }
                }
                None => {
                    // If(a, b) with no else → null coalesce
                    ExprKind::NullCoalesce {
                        left: Box::new(first),
                        right: Box::new(second),
                    }
                }
            }
        }
        Rule::addressof_expr => {
            let inner = pair.into_inner();
            let mut name = String::new();
            for p in inner {
                if p.as_rule() == Rule::dotted_identifier {
                    name = p.as_str().to_string();
                }
            }
            ExprKind::AddressOf(name)
        }
        Rule::me_keyword => {
            ExprKind::This
        }
        Rule::dot_call_statement => {
            // .Method(args) or .obj.Method(args) inside With block
            let inner = pair.into_inner();
            let mut identifiers = Vec::new();
            let mut arguments: Vec<Argument> = Vec::new();
            for p in inner {
                match p.as_rule() {
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }
            if identifiers.is_empty() {
                return Err("dot_call needs at least one identifier".to_string());
            }
            let method_name = identifiers.last().unwrap().clone();
            let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }
            // Build callee as member access, then call it
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe: false,
            });
            ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }
        }
        Rule::dot_member_access => {
            // .prop or .obj.prop inside With block
            let inner = pair.into_inner();
            let mut expr = Expression::new(ExprKind::Ident("__with_target".to_string()));
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::me_member_access => {
            let mut inner = pair.into_inner();
            let _me = inner.next().unwrap(); // me_keyword
            let mut expr = Expression::new(ExprKind::This);
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::mybase_member_access => {
            // MyBase.Property
            let mut inner = pair.into_inner();
            let _mybase = inner.next().unwrap(); // mybase_keyword
            let mut expr = Expression::new(ExprKind::Super);
            for p in inner {
                if p.as_rule() == Rule::identifier || p.as_rule() == Rule::member_identifier {
                    expr = Expression::new(ExprKind::Member {
                        object: Box::new(expr),
                        field: p.as_str().to_string(),
                        null_safe: false,
                    });
                }
            }
            return Ok(expr);
        }
        Rule::me_member_call => {
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments: Vec<Argument> = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::me_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err("me_member_call needs at least one identifier".to_string());
            }

            // Last identifier is the method name
            let method_name = identifiers.last().unwrap().clone();

            // Build object expression: Me.a.b... (all except last)
            let mut expr = Expression::new(ExprKind::This);
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }

            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: method_name,
                null_safe: false,
            });
            ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }
        }
        Rule::mybase_member_call => {
            // MyBase.Method()
            let inner = pair.into_inner();
            let mut identifiers = vec![];
            let mut arguments: Vec<Argument> = vec![];
            for p in inner {
                match p.as_rule() {
                    Rule::mybase_keyword => {},
                    Rule::identifier | Rule::member_identifier => identifiers.push(p.as_str().to_string()),
                    Rule::argument_list => arguments = parse_argument_list(p)?,
                    _ => {}
                }
            }

            if identifiers.is_empty() {
                return Err("mybase_member_call needs at least one identifier".to_string());
            }

            let method_name = identifiers.last().unwrap().clone();
            let mut expr = Expression::new(ExprKind::Super);
            for i in 0..identifiers.len() - 1 {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: identifiers[i].clone(),
                    null_safe: false,
                });
            }

            // MyBase.Method(args) → SuperCall
            if identifiers.len() == 1 {
                ExprKind::SuperCall {
                    method: Some(method_name),
                    args: arguments,
                }
            } else {
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: method_name,
                    null_safe: false,
                });
                ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                }
            }
        }
        _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
    };
    Ok(Expression::with_span(kind, span))
                ExprKind::Lit(Literal::Str(format!("{}{}", fmt, arg_expr)))
            }
            Rule::date_literal => {
                let s = pair.as_str();
                let inner = s.trim_matches('#').to_string();
                ExprKind::Cast {
                    expr: Box::new(Expression::string(&inner)),
                    type_name: "Date".to_string(),
                }
            }
            Rule::string_literal => {
                let raw = pair.as_str().trim_end_matches(|c: char| c == 'c' || c == 'C');
                let inner = &raw[1..raw.len()-1];
                let s = inner.replace("\"\"", "\"");
                ExprKind::Lit(Literal::Str(s))
            }
            Rule::numeric_literal => {
                return Ok(parse_vb_numeric_literal(pair.as_str()));
            }
            Rule::boolean_literal => {
                ExprKind::Lit(Literal::Bool(pair.as_str().eq_ignore_ascii_case("true")))
            }
            Rule::nothing_literal => ExprKind::Lit(Literal::Null),
            Rule::array_literal => {
                let elements = pair.into_inner()
                    .filter(|p| p.as_rule() == Rule::expression)
                    .map(|p| parse_expression(p).map(ArrayElement::value))
                    .collect::<Result<Vec<_>, _>>()?;
                ExprKind::Array(elements)
            }
            Rule::anonymous_new_expression => return parse_anonymous_new_expression(pair),
            Rule::new_expression => return parse_new_expression(pair),
            Rule::if_expression => return parse_if_expression(pair),
            _ => return Err(format!("Unexpected expression rule: {:?}", pair.as_rule())),
        };
        return Ok(Expression::with_span(kind, span));
    }
}

fn parse_binary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let _span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();
    let mut left = parse_expression(first)?;

    while let Some(op_pair) = inner.next() {
        let op = match op_pair.as_rule() {
            Rule::add_op | Rule::mult_op | Rule::eq_op | Rule::comp_op | Rule::and_op | Rule::or_op | Rule::xor_op | Rule::shift_op | Rule::like_op | Rule::exp_op => {
                match op_pair.as_str().to_lowercase().as_str() {
                    "+" => BinOp::Add,
                    "-" => BinOp::Sub,
                    "*" => BinOp::Mul,
                    "/" => BinOp::Div,
                    "\\" => BinOp::IDiv,
                    "mod" => BinOp::Mod,
                    "^" => BinOp::Pow,
                    "&" => BinOp::Concat,
                    "=" => BinOp::Eq,
                    "<>" => BinOp::NotEq,
                    "<" => BinOp::Lt,
                    "<=" => BinOp::LtEq,
                    ">" => BinOp::Gt,
                    ">=" => BinOp::GtEq,
                    "andalso" => BinOp::And,
                    "orelse" => BinOp::Or,
                    "and" => BinOp::BitAnd,
                    "or" => BinOp::BitOr,
                    "xor" => BinOp::BitXor,
                    "<<" => BinOp::Shl,
                    ">>" => BinOp::Shr,
                    "is" => BinOp::Is,
                    "isnot" => BinOp::IsNot,
                    "like" => BinOp::Like,
                    _ => return Err(format!("Unknown operator: {}", op_pair.as_str())),
                }
            }
            _ => return Ok(left), // Should not happen with current grammar
        };

        let right_pair = inner.next().unwrap();
        let right = parse_expression(right_pair)?;
        left = Expression::new(ExprKind::Binary {
            op,
            left: Box::new(left),
            right: Box::new(right),
        });
    }

    Ok(left)
}

*/
fn parse_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let first = inner.next().unwrap();

    match first.as_rule() {
        Rule::not_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(operand),
                },
                span,
            ))
        }
        Rule::pos_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Pos,
                    expr: Box::new(operand),
                },
                span,
            ))
        }
        Rule::neg_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(
                ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(operand),
                },
                span,
            ))
        }
        Rule::await_op => {
            let operand = parse_expression(inner.next().unwrap())?;
            Ok(Expression::with_span(
                ExprKind::Await(Box::new(operand)),
                span,
            ))
        }
        Rule::postfix => parse_postfix_expression(first),
        _ => {
            // Fallback: treat as primary
            parse_expression(first)
        }
    }
}

fn parse_postfix_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut inner = pair.into_inner();
    let primary = inner.next().unwrap();
    let mut expr = parse_expression(primary)?;

    // Apply member_chain postfix operations
    for chain in inner {
        expr = parse_member_chain_node(chain, expr)?;
    }

    Ok(expr)
}

fn parse_member_chain_node(chain: Pair<Rule>, expr: Expression) -> Result<Expression, String> {
    match chain.as_rule() {
        Rule::member_chain_invoke => {
            let arguments = chain
                .into_inner()
                .next()
                .map(parse_argument_list)
                .transpose()?
                .unwrap_or_default();
            if arguments.is_empty() {
                if let ExprKind::Ident(marker) = &expr.kind {
                    if let Some((name, type_name)) = vb_generic_type_marker_parts(marker) {
                        return Ok(Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::ident(&vb_generic_call_marker(
                                &name, &type_name,
                            ))),
                            args: Vec::new(),
                            optional: false,
                        }));
                    }
                }
            }
            if arguments.len() == 1 {
                if let Some(name) = dotted_expr_name(&expr) {
                    if name
                        .trim_start()
                        .to_ascii_lowercase()
                        .starts_with("ctypedynamic(of")
                    {
                        let type_name = vb_call_generic_first_type(&name)
                            .unwrap_or_else(|| "Object".to_string());
                        return Ok(Expression::new(ExprKind::Cast {
                            expr: Box::new(arguments[0].value.clone()),
                            type_name,
                        }));
                    }
                }
            }
            if arguments.len() == 1 && vb_expr_is_xml_axis_result(&expr) {
                return Ok(Expression::new(ExprKind::Index {
                    object: Box::new(expr),
                    index: Box::new(arguments[0].value.clone()),
                    null_safe: false,
                }));
            }
            if let ExprKind::Ident(name) = &expr.kind {
                if let Some(rewritten) = canonicalize_call(name, &arguments) {
                    return Ok(rewritten);
                }
            }
            // `X.TryParse(s, r)` / `X.TryGetValue(k, v)` out-param normalization
            // delegates to the shared Dotnet lowering helpers.
            // `expr` is the `Member` callee.
            if let ExprKind::Member { object, field, .. } = &expr.kind {
                if field.eq_ignore_ascii_case("TryParse") && arguments.len() == 2 {
                    let recv = dotted_expr_name(object);
                    if let Some(rewritten) = dotnet_vb::try_parse_desugar(
                        recv.as_deref(),
                        &expr,
                        &arguments[0].value,
                        &arguments[1].value,
                    ) {
                        return Ok(rewritten);
                    }
                }
                if field.eq_ignore_ascii_case("TryCreate") && arguments.len() == 3 {
                    let recv = dotted_expr_name(object);
                    if let Some(rewritten) = dotnet_vb::try_create_desugar(
                        recv.as_deref(),
                        &expr,
                        &arguments[0].value,
                        &arguments[1].value,
                        &arguments[2].value,
                    ) {
                        return Ok(rewritten);
                    }
                }
                if field.eq_ignore_ascii_case("TryGetValue") && arguments.len() == 2 {
                    return Ok(dotnet_vb::try_get_value_desugar(
                        object,
                        &arguments[0].value,
                        &arguments[1].value,
                    ));
                }
                if field.eq_ignore_ascii_case("GetOrAdd") && arguments.len() == 2 {
                    return Ok(dotnet_vb::get_or_add_desugar(
                        object,
                        &arguments[0].value,
                        &arguments[1].value,
                    ));
                }
                if field.eq_ignore_ascii_case("AddOrUpdate") && arguments.len() == 3 {
                    return Ok(dotnet_vb::add_or_update_desugar(
                        object,
                        &arguments[0].value,
                        &arguments[1].value,
                        &arguments[2].value,
                    ));
                }
                if field.eq_ignore_ascii_case("TryUpdate") && arguments.len() == 3 {
                    return Ok(dotnet_vb::try_update_desugar(
                        object,
                        &arguments[0].value,
                        &arguments[1].value,
                        &arguments[2].value,
                    ));
                }
                if field.eq_ignore_ascii_case("TryRemove") && arguments.len() == 2 {
                    return Ok(dotnet_vb::try_remove_desugar(
                        object,
                        &arguments[0].value,
                        &arguments[1].value,
                    ));
                }
                if dotnet_vb::is_hashset_relation_method(field) && arguments.len() == 1 {
                    return Ok(vb_bool_expr(Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: arguments,
                        optional: false,
                    })));
                }
                if (field.eq_ignore_ascii_case("TryDequeue")
                    || field.eq_ignore_ascii_case("TryPop")
                    || field.eq_ignore_ascii_case("TryPeek")
                    || field.eq_ignore_ascii_case("TryTake"))
                    && arguments.len() == 1
                {
                    return Ok(dotnet_vb::try_take_desugar(
                        object,
                        field,
                        &arguments[0].value,
                    ));
                }
            }
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(expr),
                args: arguments,
                optional: false,
            }))
        }
        Rule::member_chain_call => {
            let chain_inner = chain.into_inner();
            let mut name = String::new();
            let mut generic_type = None;
            let mut arguments = Vec::new();
            for part in chain_inner {
                match part.as_rule() {
                    Rule::member_identifier => name = strip_vb_generic_suffix(part.as_str()),
                    Rule::generic_suffix => {
                        generic_type = vb_generic_suffix_first_type(part.as_str())
                    }
                    Rule::argument_list => arguments = parse_argument_list(part)?,
                    _ => {}
                }
            }
            if name.eq_ignore_ascii_case("OfType") {
                if let Some(type_name) = generic_type {
                    arguments.insert(0, Argument::positional(Expression::string(&type_name)));
                }
            }
            if name.eq_ignore_ascii_case("Item") && !arguments.is_empty() {
                let mut indexed = expr;
                for arg in arguments {
                    indexed = Expression::new(ExprKind::Index {
                        object: Box::new(indexed),
                        index: Box::new(arg.value),
                        null_safe: false,
                    });
                }
                return Ok(indexed);
            }
            if name.eq_ignore_ascii_case("Round") {
                if let Some(path) = dotted_expr_name(&expr) {
                    if path.eq_ignore_ascii_case("Math") || path.eq_ignore_ascii_case("System.Math")
                    {
                        if arguments.len() >= 2 {
                            if let Some(folded) =
                                try_fold_vb_double_round(&arguments[0].value, &arguments[1].value)
                            {
                                return Ok(folded);
                            }
                        }
                        return Ok(if arguments.len() >= 2 {
                            build_vb_precision_round_expr(
                                arguments[0].value.clone(),
                                arguments[1].value.clone(),
                            )
                        } else {
                            build_vb_bankers_round_expr(arguments[0].value.clone())
                        });
                    }
                }
            }
            // `X.TryParse(s, r)` out-param normalization for the combined
            // name+args grammar form; `expr` is the receiver, `name` the method.
            if name.eq_ignore_ascii_case("TryParse") && arguments.len() == 2 {
                let recv = dotted_expr_name(&expr);
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr.clone()),
                    field: name.clone(),
                    null_safe: false,
                });
                if let Some(rewritten) = dotnet_vb::try_parse_desugar(
                    recv.as_deref(),
                    &callee,
                    &arguments[0].value,
                    &arguments[1].value,
                ) {
                    return Ok(rewritten);
                }
            }
            if name.eq_ignore_ascii_case("TryCreate") && arguments.len() == 3 {
                let recv = dotted_expr_name(&expr);
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr.clone()),
                    field: name.clone(),
                    null_safe: false,
                });
                if let Some(rewritten) = dotnet_vb::try_create_desugar(
                    recv.as_deref(),
                    &callee,
                    &arguments[0].value,
                    &arguments[1].value,
                    &arguments[2].value,
                ) {
                    return Ok(rewritten);
                }
            }
            if name.eq_ignore_ascii_case("TryGetValue") && arguments.len() == 2 {
                return Ok(dotnet_vb::try_get_value_desugar(
                    &expr,
                    &arguments[0].value,
                    &arguments[1].value,
                ));
            }
            if name.eq_ignore_ascii_case("GetOrAdd") && arguments.len() == 2 {
                return Ok(dotnet_vb::get_or_add_desugar(
                    &expr,
                    &arguments[0].value,
                    &arguments[1].value,
                ));
            }
            if name.eq_ignore_ascii_case("AddOrUpdate") && arguments.len() == 3 {
                return Ok(dotnet_vb::add_or_update_desugar(
                    &expr,
                    &arguments[0].value,
                    &arguments[1].value,
                    &arguments[2].value,
                ));
            }
            if name.eq_ignore_ascii_case("TryUpdate") && arguments.len() == 3 {
                return Ok(dotnet_vb::try_update_desugar(
                    &expr,
                    &arguments[0].value,
                    &arguments[1].value,
                    &arguments[2].value,
                ));
            }
            if name.eq_ignore_ascii_case("TryRemove") && arguments.len() == 2 {
                return Ok(dotnet_vb::try_remove_desugar(
                    &expr,
                    &arguments[0].value,
                    &arguments[1].value,
                ));
            }
            if dotnet_vb::is_hashset_relation_method(&name) && arguments.len() == 1 {
                let callee = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: name,
                    null_safe: false,
                });
                return Ok(vb_bool_expr(Expression::new(ExprKind::Call {
                    callee: Box::new(callee),
                    args: arguments,
                    optional: false,
                })));
            }
            if (name.eq_ignore_ascii_case("TryDequeue")
                || name.eq_ignore_ascii_case("TryPop")
                || name.eq_ignore_ascii_case("TryPeek")
                || name.eq_ignore_ascii_case("TryTake"))
                && arguments.len() == 1
            {
                return Ok(dotnet_vb::try_take_desugar(
                    &expr,
                    &name,
                    &arguments[0].value,
                ));
            }
            let callee = Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: name,
                null_safe: false,
            });
            Ok(Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args: arguments,
                optional: false,
            }))
        }
        Rule::member_chain_access => {
            let name = normalize_vb_identifier(chain.into_inner().next().unwrap().as_str());
            if let ExprKind::Ident(marker) = &expr.kind {
                if let Some(static_name) = vb_generic_static_name(marker, &name) {
                    return Ok(Expression::ident(&static_name));
                }
            }
            Ok(canonicalize_member_access(expr, &name))
        }
        Rule::member_chain_xml_child_axis | Rule::member_chain_xml_descendant_axis => {
            let name = chain
                .into_inner()
                .next()
                .map(|p| p.as_str().to_string())
                .unwrap_or_default();
            Ok(build_vb_xml_axis_call(expr, &name))
        }
        Rule::member_chain => {
            let inner_chain = chain.into_inner().next().unwrap();
            parse_member_chain_node(inner_chain, expr)
        }
        _ => Ok(expr),
    }
}

fn build_vb_xml_axis_call(receiver: Expression, name: &str) -> Expression {
    call_expr(
        build_dotted_expr("xml.elements"),
        vec![
            Argument::positional(receiver),
            Argument::positional(Expression::string(name)),
        ],
    )
}

fn vb_expr_is_xml_axis_result(expr: &Expression) -> bool {
    matches!(
        &expr.kind,
        ExprKind::Call { callee, .. }
            if dotted_expr_name(callee).is_some_and(|name| name.eq_ignore_ascii_case("xml.elements"))
    )
}

fn vb_expr_is_xml_elements_sequence(expr: &Expression) -> bool {
    vb_expr_is_xml_axis_result(expr)
        || matches!(
            &expr.kind,
            ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Elements"))
        )
}

fn parse_argument_list(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .map(|p| match p.as_rule() {
            Rule::omitted_argument => Ok(Argument::positional(Expression::null())),
            Rule::first_argument | Rule::trailing_argument => {
                let Some(inner) = p.into_inner().next() else {
                    return Ok(Argument::positional(Expression::null()));
                };
                match inner.as_rule() {
                    Rule::omitted_argument => Ok(Argument::positional(Expression::null())),
                    Rule::named_argument => {
                        let mut named_inner = inner.into_inner();
                        let name = normalize_vb_identifier(named_inner.next().unwrap().as_str());
                        let value = parse_expression(named_inner.next().unwrap())?;
                        Ok(Argument {
                            value,
                            name: Some(name),
                            by_ref: false,
                            spread: false,
                        })
                    }
                    _ => parse_expression(inner).map(Argument::positional),
                }
            }
            Rule::named_argument => {
                let mut inner = p.into_inner();
                let name = normalize_vb_identifier(inner.next().unwrap().as_str());
                let value = parse_expression(inner.next().unwrap())?;
                Ok(Argument {
                    value,
                    name: Some(name),
                    by_ref: false,
                    spread: false,
                })
            }
            Rule::argument => {
                let Some(inner) = p.into_inner().next() else {
                    return Ok(Argument::positional(Expression::null()));
                };
                match inner.as_rule() {
                    Rule::omitted_argument => Ok(Argument::positional(Expression::null())),
                    Rule::named_argument => {
                        let mut named_inner = inner.into_inner();
                        let name = normalize_vb_identifier(named_inner.next().unwrap().as_str());
                        let value = parse_expression(named_inner.next().unwrap())?;
                        Ok(Argument {
                            value,
                            name: Some(name),
                            by_ref: false,
                            spread: false,
                        })
                    }
                    _ => parse_expression(inner).map(Argument::positional),
                }
            }
            _ => parse_expression(p).map(Argument::positional),
        })
        .collect()
}

fn unwrap_argument_expr_pair(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    match pair.as_rule() {
        Rule::first_argument | Rule::trailing_argument => pair
            .into_inner()
            .next()
            .ok_or_else(|| "Missing argument expression".to_string()),
        _ => Ok(pair),
    }
}

fn parse_try_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner();

    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally = None;

    for p in inner {
        match p.as_rule() {
            Rule::try_body => {
                body = parse_block_body(p)?;
            }
            Rule::catch_block => catches.push(parse_catch_block(p)?),
            Rule::finally_block => {
                let f_inner = p.into_inner();
                for fp in f_inner {
                    if fp.as_rule() == Rule::try_body {
                        finally = Some(parse_block_body(fp)?);
                    }
                }
            }
            Rule::try_end => {}
            _ => {}
        }
    }

    if let Some(exit_index) = body
        .iter()
        .position(|stmt| matches!(stmt.kind, StmtKind::Break(BreakTarget::Kind(ExitKind::Try))))
    {
        body.truncate(exit_index);
    }

    for catch in catches.iter_mut() {
        if let Some(var_name) = catch.var_name.clone() {
            rewrite_vb_bare_throws(&mut catch.body, &var_name);
        }
    }

    Ok(Statement::with_span(
        StmtKind::Try {
            body,
            catches,
            else_body: None,
            finally,
        },
        span,
    ))
}

fn parse_catch_block(pair: Pair<Rule>) -> Result<CatchClause, String> {
    let inner = pair.into_inner();
    let mut var_name: Option<String> = None;
    let mut catch_type: Option<String> = None;
    let mut when_clause = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                var_name = Some(p.as_str().to_string());
            }
            Rule::type_name => {
                catch_type = Some(p.as_str().to_string());
            }
            Rule::expression => {
                when_clause = Some(parse_expression(p)?);
            }
            Rule::try_body => {
                body = parse_block_body(p)?;
            }
            _ => {}
        }
    }

    if var_name.is_none() {
        var_name = Some("__vb_caught_exception".to_string());
    }

    Ok(CatchClause {
        types: catch_type.into_iter().collect(),
        var_name,
        stack_var: None,
        body,
        when_clause,
    })
}

fn parse_continue_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let text = pair.as_str().to_lowercase();
    let target = if text.contains("do") {
        ContinueTarget::Kind(ContinueKind::Do)
    } else if text.contains("for") {
        ContinueTarget::Kind(ContinueKind::For)
    } else {
        ContinueTarget::Kind(ContinueKind::While)
    };

    Ok(Statement::with_span(StmtKind::Continue(target), span))
}

fn parse_lambda_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let text = pair.as_str().trim_start();
    let _is_function = text.to_lowercase().starts_with("function");

    let mut inner = pair.into_inner();
    let mut params = Vec::new();

    let mut next_pair = inner
        .next()
        .ok_or_else(|| "Lambda missing body".to_string())?;

    if next_pair.as_rule() == Rule::param_list {
        params = parse_param_list(next_pair)?;
        next_pair = inner
            .next()
            .ok_or_else(|| "Lambda missing body".to_string())?;
    }

    let body = match next_pair.as_rule() {
        Rule::expression => LambdaBody::Expr(Box::new(parse_expression(next_pair)?)),
        Rule::sub_end => LambdaBody::Block(Vec::new()),
        Rule::NEWLINE => {
            // Multiline block
            let mut body_stmts = Vec::new();
            for item in inner {
                match item.as_rule() {
                    Rule::statement_line => {
                        for stmt_pair in item.into_inner() {
                            if stmt_pair.as_rule() != Rule::NEWLINE
                                && stmt_pair.as_rule() != Rule::EOI
                            {
                                body_stmts.push(parse_statement(stmt_pair)?);
                            }
                        }
                    }
                    _ => {
                        if let Some(decl_stmt) = try_parse_declaration(item.clone())? {
                            body_stmts.push(decl_stmt);
                        }
                    }
                }
            }
            LambdaBody::Block(body_stmts)
        }
        _ => {
            // Any statement variant rule (call_statement, assign_statement, etc.)
            // These appear directly because `statement` is a silent rule in the grammar.
            let stmt = parse_statement(next_pair)?;
            LambdaBody::Block(vec![stmt])
        }
    };

    Ok(Expression::with_span(
        ExprKind::Lambda {
            params,
            body,
            is_async: false,
            captures: vec![],
        },
        span,
    ))
}

fn parse_block_body(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut body = Vec::new();
    for stmt_pair in pair.into_inner() {
        if stmt_pair.as_rule() == Rule::statement_line {
            for s in stmt_pair.into_inner() {
                if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                    body.push(parse_statement(s)?);
                }
            }
        }
    }
    Ok(body)
}

fn parse_for_each_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let hidden_suffix = pair.as_span().start();
    let mut inner = pair.into_inner();
    let variable = inner.next().unwrap().as_str().to_string();
    let mut variable_type = None;
    let mut collection = None;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::type_name => variable_type = Some(p.as_str().to_string()),
            Rule::expression => {
                if collection.is_none() {
                    collection = Some(parse_expression(p)?);
                }
            }
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::for_end => {}
            _ => body.push(parse_statement(p)?),
        }
    }

    let mut loop_var = variable.clone();
    if let Some(type_hint) = variable_type {
        let source_var = format!("__vb_foreach_item_{}", hidden_suffix);
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(variable),
                    type_hint: Some(type_hint),
                    init: Some(Expression::ident(&source_var)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Dim,
            }),
        );
        loop_var = source_var;
    }

    Ok(Statement::with_span(
        StmtKind::ForIn {
            var: loop_var,
            key: None,
            iter: collection.ok_or_else(|| "For Each missing collection".to_string())?,
            body,
            of: true, // VB For Each iterates values, like JS for...of
            else_body: None,
            is_async: false,
        },
        span,
    ))
}

fn parse_with_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let object = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::EOI | Rule::with_end => {}
            _ => body.push(parse_statement(p)?),
        }
    }

    let simple_target = matches!(object.kind, ExprKind::Ident(_));
    let target_expr = if simple_target {
        object.clone()
    } else {
        let temp_name = format!(
            "__vb_with_target_{}_{}",
            span.start_line.max(1),
            span.start_col.max(1)
        );
        Expression::ident(&temp_name)
    };
    let null_check = Statement::new(StmtKind::If {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::Is,
            left: Box::new(target_expr.clone()),
            right: Box::new(Expression::null()),
        }),
        then_body: vec![Statement::new(StmtKind::Throw {
            expr: Some(Expression::new(ExprKind::New {
                class: Box::new(Expression::ident("NullReferenceException")),
                args: Vec::new(),
            })),
            cause: None,
        })],
        elifs: Vec::new(),
        else_body: None,
    });
    let with_stmt = Statement::with_span(
        StmtKind::With {
            items: vec![WithItem {
                expr: target_expr.clone(),
                var: None,
            }],
            body,
            is_async: false,
        },
        Span::default(),
    );

    let mut lowered = Vec::new();
    if !simple_target {
        let ExprKind::Ident(temp_name) = &target_expr.kind else {
            unreachable!();
        };
        lowered.push(Statement::with_span(
            StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(temp_name.clone()),
                    type_hint: None,
                    init: Some(object),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Dim,
            },
            span,
        ));
    }
    lowered.push(null_check);
    lowered.push(with_stmt);
    Ok(Statement::with_span(
        StmtKind::Block(lowered),
        Span::default(),
    ))
}

fn parse_using_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner: Vec<_> = pair.into_inner().collect();
    let mut resources: Vec<(String, Expression)> = Vec::new();
    let mut body = Vec::new();

    for p in &inner {
        match p.as_rule() {
            Rule::using_resource_decl => {
                let mut var_name = String::new();
                let mut resource_expr = None;
                for rp in p.clone().into_inner() {
                    match rp.as_rule() {
                        Rule::identifier => var_name = rp.as_str().to_string(),
                        Rule::type_name => {}
                        Rule::new_expression | Rule::expression => {
                            resource_expr = Some(parse_expression(rp)?);
                        }
                        _ => {}
                    }
                }
                let resource = resource_expr
                    .ok_or_else(|| "Using statement missing resource expression".to_string())?;
                resources.push((var_name, resource));
            }
            Rule::statement_line => {
                for stmt_pair in p.clone().into_inner() {
                    if stmt_pair.as_rule() != Rule::NEWLINE && stmt_pair.as_rule() != Rule::EOI {
                        body.push(parse_statement(stmt_pair)?);
                    }
                }
            }
            Rule::NEWLINE | Rule::using_end => {}
            _ => {}
        }
    }

    let mut nested_body = body;
    let mut nested_stmt = None;
    for (var, resource) in resources.into_iter().rev() {
        let using_stmt = Statement::with_span(
            StmtKind::Using {
                var,
                resource,
                body: nested_body,
            },
            span.clone(),
        );
        nested_body = vec![using_stmt.clone()];
        nested_stmt = Some(using_stmt);
    }

    nested_stmt.ok_or_else(|| "Using statement missing resource declaration".to_string())
}

fn parse_enum_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let decorators = parse_vb_attribute_specs(pair.as_str());
    let is_flags = decorators.iter().any(|attr| {
        vb_attribute_leaf_name(attr).is_some_and(|short| {
            short.eq_ignore_ascii_case("Flags") || short.eq_ignore_ascii_case("FlagsAttribute")
        })
    });
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut members = Vec::new();
    let mut backing_type = None;

    for p in inner {
        match p.as_rule() {
            Rule::identifier => {
                let text = p.as_str().to_lowercase();
                match text.as_str() {
                    "public" => visibility = Visibility::Public,
                    "private" => visibility = Visibility::Private,
                    "protected" => visibility = Visibility::Protected,
                    "friend" => visibility = Visibility::Internal,
                    _ => name = p.as_str().to_string(),
                }
            }
            Rule::enum_member | Rule::enum_member_inline => {
                let mut member_inner = p.into_inner();
                let member_name = member_inner.next().unwrap().as_str().to_string();
                let value = member_inner
                    .find(|e| e.as_rule() == Rule::expression)
                    .map(|e| parse_expression(e))
                    .transpose()?;
                members.push(EnumMember {
                    name: member_name,
                    value,
                    constructor_args: Vec::new(),
                });
            }
            Rule::type_name => backing_type = Some(vb_canonical_type_name(p.as_str())),
            Rule::enum_end | Rule::NEWLINE => {}
            _ => {}
        }
    }

    normalize_vb_enum_member_values(&mut members);
    if backing_type.as_deref() == Some("Int64") {
        for member in &mut members {
            if member.value.as_ref().and_then(literal_int) == Some(i64::MAX) {
                member.value = Some(Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::int(i64::MAX)),
                    type_name: "Int64".to_string(),
                }));
            }
        }
    }

    Ok(Statement::with_span(
        StmtKind::EnumDecl {
            name,
            members,
            visibility,
            is_flags,
            backing_type,
            interfaces: Vec::new(),
            body_members: Vec::new(),
            decorators,
        },
        span,
    ))
}

fn normalize_vb_enum_member_values(members: &mut [EnumMember]) {
    let raw_values: Vec<Option<Expression>> =
        members.iter().map(|member| member.value.clone()).collect();
    let mut resolved: HashMap<String, i64> = HashMap::new();

    for _ in 0..members.len().saturating_add(1) {
        let mut changed = false;
        let mut previous = -1_i64;
        for (index, member) in members.iter().enumerate() {
            if let Some(value) = resolved.get(&member.name.to_ascii_lowercase()).copied() {
                previous = value;
                continue;
            }
            let value = match &raw_values[index] {
                Some(expr) => eval_vb_enum_const_expr(expr, &resolved),
                None => Some(previous.saturating_add(1)),
            };
            if let Some(value) = value {
                resolved.insert(member.name.to_ascii_lowercase(), value);
                previous = value;
                changed = true;
            }
        }
        if !changed {
            break;
        }
    }

    let mut previous = -1_i64;
    for member in members {
        let value = resolved
            .get(&member.name.to_ascii_lowercase())
            .copied()
            .unwrap_or_else(|| previous.saturating_add(1));
        member.value = Some(Expression::int(value));
        previous = value;
    }
}

fn eval_vb_enum_const_expr(expr: &Expression, values: &HashMap<String, i64>) -> Option<i64> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Some(*value),
        ExprKind::Lit(Literal::Float(value)) => Some(*value as i64),
        ExprKind::Ident(name) => values.get(&name.to_ascii_lowercase()).copied(),
        ExprKind::Member { field, .. } => values.get(&field.to_ascii_lowercase()).copied(),
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => eval_vb_enum_const_expr(expr, values).map(|value| -value),
        ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr,
        } => eval_vb_enum_const_expr(expr, values).map(|value| !value),
        ExprKind::Binary { op, left, right } => {
            let left = eval_vb_enum_const_expr(left, values)?;
            let right = eval_vb_enum_const_expr(right, values)?;
            match op {
                BinOp::Add => Some(left.saturating_add(right)),
                BinOp::Sub => Some(left.saturating_sub(right)),
                BinOp::Mul => Some(left.saturating_mul(right)),
                BinOp::Div | BinOp::IDiv => (right != 0).then_some(left / right),
                BinOp::Mod => (right != 0).then_some(left % right),
                BinOp::BitAnd | BinOp::And => Some(left & right),
                BinOp::BitOr | BinOp::Or => Some(left | right),
                BinOp::BitXor | BinOp::Xor => Some(left ^ right),
                BinOp::Shl => Some(left << right.max(0).min(63)),
                BinOp::Shr => Some(left >> right.max(0).min(63)),
                _ => None,
            }
        }
        _ => None,
    }
}

fn parse_single_line_if(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let cond = parse_expression(inner.next().unwrap())?;

    let then_body = vec![parse_statement(inner.next().unwrap())?];

    let else_body = if let Some(else_body_pair) = inner.next() {
        Some(vec![parse_statement(else_body_pair)?])
    } else {
        None
    };

    Ok(Statement::with_span(
        StmtKind::If {
            cond,
            then_body,
            elifs: Vec::new(),
            else_body,
        },
        span,
    ))
}

fn parse_field_decl(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let pair_text = pair.as_str().to_ascii_lowercase();
    let mut field_name = String::new();
    let mut field_type: Option<String> = None;
    let mut field_init = None;
    let mut field_bounds = None;
    let mut is_new = false;
    let mut ctor_args: Vec<Argument> = Vec::new();
    let mut is_with_events = pair_text
        .split(|c: char| !c.is_ascii_alphanumeric())
        .any(|part| part == "withevents");

    for fp in pair.into_inner() {
        match fp.as_rule() {
            Rule::withevents_keyword => {
                is_with_events = true;
            }
            Rule::visibility_modifier | Rule::sub_modifier_keyword | Rule::partial_keyword => {} // modifiers handled by caller
            Rule::dim_new_keyword => {
                is_new = true;
            }
            Rule::identifier => field_name = fp.as_str().to_string(),
            Rule::type_name => field_type = Some(fp.as_str().to_string()),
            Rule::array_rank_spec => {
                field_bounds = Some(parse_array_bounds_pair(fp)?);
            }
            Rule::array_bounds => {
                field_bounds = Some(parse_array_bounds_pair(fp)?);
            }
            Rule::argument_list => {
                for arg_pair in fp.into_inner() {
                    if arg_pair.as_rule() == Rule::expression {
                        ctor_args.push(Argument::positional(parse_expression(arg_pair)?));
                    }
                }
            }
            Rule::expression => field_init = Some(parse_expression(fp)?),
            Rule::array_literal => field_init = Some(parse_array_literal(fp)?),
            _ => {}
        }
    }

    // Handle "As New Type" syntax
    if is_new {
        if let Some(field_type_value) = field_type.as_mut() {
            *field_type_value = field_type_value
                .trim()
                .strip_suffix("()")
                .unwrap_or(field_type_value.trim())
                .trim()
                .to_string();
        }
    }

    if is_new && field_init.is_none() {
        if let Some(t) = &field_type {
            field_init = Some(Expression::new(ExprKind::New {
                class: Box::new(build_dotted_expr(&strip_vb_generic_suffix(t))),
                args: ctor_args,
            }));
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(field_name),
        type_hint: field_type,
        init: field_init,
        array_bounds: field_bounds,
        with_events: is_with_events,
    })
}

fn parse_open_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_path = parse_expression(inner.next().unwrap())?;
    let mode_pair = inner.next().unwrap(); // file_mode
    let mode = match mode_pair.as_str().to_lowercase().as_str() {
        "input" => FileMode::Input,
        "output" => FileMode::Output,
        "append" => FileMode::Append,
        "binary" => FileMode::Binary,
        "random" => FileMode::Random,
        _ => return Err(format!("Unknown file mode: {}", mode_pair.as_str())),
    };
    let file_number = parse_expression(inner.next().unwrap())?;
    Ok(Statement::with_span(
        StmtKind::OpenFile {
            path: file_path,
            mode,
            file_number,
        },
        span,
    ))
}

fn parse_close_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    // Close with no arguments closes all files
    let file_number = inner.next().map(|p| parse_expression(p)).transpose()?;
    Ok(Statement::with_span(StmtKind::CloseFile(file_number), span))
}

fn parse_print_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner
        .next()
        .map(|p| parse_argument_list(p).map(|args| args.into_iter().map(|a| a.value).collect()))
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::with_span(
        StmtKind::PrintFile { file_number, items },
        span,
    ))
}

fn parse_write_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let items = inner
        .next()
        .map(|p| parse_argument_list(p).map(|args| args.into_iter().map(|a| a.value).collect()))
        .transpose()?
        .unwrap_or_default();
    Ok(Statement::with_span(
        StmtKind::WriteFile { file_number, items },
        span,
    ))
}

fn parse_input_file_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let mut variables = Vec::new();
    for p in inner {
        if p.as_rule() == Rule::identifier {
            variables.push(Expression::ident(p.as_str()));
        }
    }
    Ok(Statement::with_span(
        StmtKind::InputFile {
            file_number,
            variables,
        },
        span,
    ))
}

fn parse_line_input_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let file_number = parse_expression(inner.next().unwrap())?;
    let variable = inner.next().unwrap().as_str().to_string();
    Ok(Statement::with_span(
        StmtKind::LineInput {
            file_number,
            variable,
        },
        span,
    ))
}

fn parse_select_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let expr = parse_expression(inner.next().unwrap())?;

    let mut cases = Vec::new();
    let mut default = None;

    for p in inner {
        match p.as_rule() {
            Rule::case_block => {
                let mut case_inner = p.into_inner();
                let conditions_pair = case_inner.next().unwrap();
                let mut conditions = Vec::new();

                for cond_pair in conditions_pair.into_inner() {
                    let mut cond_inner = cond_pair.into_inner();
                    let first = cond_inner.next().unwrap();

                    let condition = match first.as_rule() {
                        Rule::expression => {
                            let expr1 = parse_expression(first)?;
                            if let Some(next) = cond_inner.next() {
                                let expr2 = parse_expression(next)?;
                                CaseCondition::Range {
                                    from: expr1,
                                    to: expr2,
                                }
                            } else {
                                CaseCondition::Value(expr1)
                            }
                        }
                        Rule::comp_op => {
                            let op = match first.as_str() {
                                "=" => ComparisonOp::Eq,
                                "<>" => ComparisonOp::NotEq,
                                "<" => ComparisonOp::Lt,
                                "<=" => ComparisonOp::LtEq,
                                ">" => ComparisonOp::Gt,
                                ">=" => ComparisonOp::GtEq,
                                _ => {
                                    return Err(format!(
                                        "Unknown comparison operator: {}",
                                        first.as_str()
                                    ));
                                }
                            };
                            let expr = parse_expression(cond_inner.next().unwrap())?;
                            CaseCondition::Comparison { op, expr }
                        }
                        _ => {
                            return Err(format!(
                                "Unexpected rule in case condition: {:?}",
                                first.as_rule()
                            ));
                        }
                    };
                    conditions.push(condition);
                }

                let mut body = Vec::new();
                for stmt_pair in case_inner {
                    if stmt_pair.as_rule() == Rule::statement_line {
                        for inner in stmt_pair.into_inner() {
                            if inner.as_rule() != Rule::NEWLINE && inner.as_rule() != Rule::EOI {
                                body.push(parse_statement(inner)?);
                            }
                        }
                    }
                }
                cases.push(SwitchCase { conditions, body });
            }
            Rule::case_else => {
                let mut body = Vec::new();
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::statement_line {
                        for inner in stmt_pair.into_inner() {
                            if inner.as_rule() != Rule::NEWLINE && inner.as_rule() != Rule::EOI {
                                body.push(parse_statement(inner)?);
                            }
                        }
                    }
                }
                default = Some(body);
            }
            Rule::select_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::Switch {
            expr,
            cases,
            default,
        },
        span,
    ))
}

// ---------------------------------------------------------------------------
// Interface / Structure / Delegate / Event parsers
// ---------------------------------------------------------------------------

fn parse_interface_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let decorators = parse_vb_attribute_specs(pair.as_str());
    let inner = pair.into_inner();
    let mut _visibility = Visibility::Public;
    let mut name = String::new();
    let mut parents: Vec<String> = Vec::new();
    let mut members: Vec<InterfaceMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::inherits_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        parents.push(vb_declared_base_type_name(tp.as_str()));
                    }
                }
            }
            Rule::interface_sub => {
                let mut sname = String::new();
                let mut params = Vec::new();
                let is_shadows = p
                    .as_str()
                    .trim_start()
                    .to_ascii_lowercase()
                    .starts_with("shadows");
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => sname = sp.as_str().to_string(),
                        Rule::param_list => params = parse_param_list(sp)?,
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Method {
                    name: sname,
                    params,
                    return_type: None,
                    is_sub: true,
                    signature_source: is_shadows.then(|| "shadows".to_string()),
                });
            }
            Rule::interface_function => {
                let mut fname = String::new();
                let mut params = Vec::new();
                let mut ret: Option<String> = None;
                let is_shadows = p
                    .as_str()
                    .trim_start()
                    .to_ascii_lowercase()
                    .starts_with("shadows");
                for fp in p.into_inner() {
                    match fp.as_rule() {
                        Rule::identifier => fname = fp.as_str().to_string(),
                        Rule::param_list => params = parse_param_list(fp)?,
                        Rule::type_name => ret = Some(fp.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Method {
                    name: fname,
                    params,
                    return_type: ret,
                    is_sub: false,
                    signature_source: is_shadows.then(|| "shadows".to_string()),
                });
            }
            Rule::interface_property => {
                let mut pname = String::new();
                let mut ptype: Option<String> = None;
                let mut is_readonly = false;
                let mut is_writeonly = false;
                let txt = p.as_str().to_lowercase();
                if txt.starts_with("readonly") {
                    is_readonly = true;
                }
                if txt.starts_with("writeonly") {
                    is_writeonly = true;
                }
                for pp in p.into_inner() {
                    match pp.as_rule() {
                        Rule::identifier => pname = pp.as_str().to_string(),
                        Rule::type_name => ptype = Some(pp.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Property {
                    name: pname,
                    type_hint: ptype,
                    is_readonly,
                    is_writeonly,
                });
            }
            Rule::interface_event => {
                let mut ename = String::new();
                let mut etype: Option<String> = None;
                for ep in p.into_inner() {
                    match ep.as_rule() {
                        Rule::identifier => ename = ep.as_str().to_string(),
                        Rule::type_name => etype = Some(ep.as_str().to_string()),
                        _ => {}
                    }
                }
                members.push(InterfaceMember::Event {
                    name: ename,
                    type_hint: etype,
                });
            }
            Rule::visibility_modifier => {
                _visibility = parse_visibility(p.as_str());
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::InterfaceDecl {
            name,
            parents,
            members,
            decorators,
        },
        span,
    ))
}

fn parse_structure_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let decorators = parse_vb_attribute_specs(pair.as_str());
    let inner = pair.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut interfaces: Vec<String> = Vec::new();
    let mut members: Vec<ClassMember> = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::implements_statement => {
                for tp in p.into_inner() {
                    if tp.as_rule() == Rule::type_name {
                        interfaces.push(vb_declared_base_type_name(tp.as_str()));
                    }
                }
            }
            Rule::property_decl => {
                members.extend(parse_property_decl_to_members(p)?);
            }
            Rule::auto_property_decl => {
                let d = parse_auto_property_as_field(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers: Modifiers::default(),
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::sub_decl => {
                let implemented_targets = vb_implements_target_member_infos(&p);
                let implemented_members: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.leaf.clone())
                    .collect();
                let sub_stmt = parse_sub_decl(p)?;
                let is_ctor = match &sub_stmt.kind {
                    StmtKind::FunctionDecl { name, .. } => name == "New",
                    _ => false,
                };
                if is_ctor {
                    match sub_stmt.kind {
                        StmtKind::FunctionDecl {
                            params,
                            mut body,
                            modifiers,
                            ..
                        } => {
                            if modifiers.is_static {
                                members.push(ClassMember::Method(Box::new(Statement::with_span(
                                    StmtKind::FunctionDecl {
                                        name: "__static_init__".to_string(),
                                        params,
                                        return_type: None,
                                        body,
                                        modifiers,
                                        handles: Vec::new(),
                                        is_async: false,
                                        is_generator: false,
                                        is_sub: true,
                                    },
                                    span.clone(),
                                ))));
                                continue;
                            }
                            let (initializer_target, base_args) =
                                extract_vb_this_constructor_call(&mut body);
                            members.push(ClassMember::Constructor {
                                name: None,
                                params,
                                body,
                                base_args,
                                initializer_target,
                                visibility: modifiers.visibility,
                            });
                        }
                        _ => unreachable!(),
                    }
                } else {
                    push_vb_interface_forwarders(&mut members, &sub_stmt, &implemented_members);
                    let qualified_forwarders: Vec<String> = implemented_targets
                        .iter()
                        .map(|target| target.forwarder.clone())
                        .collect();
                    push_vb_interface_forwarders(&mut members, &sub_stmt, &qualified_forwarders);
                    members.push(ClassMember::Method(Box::new(sub_stmt)));
                }
            }
            Rule::function_decl => {
                let implemented_targets = vb_implements_target_member_infos(&p);
                let implemented_members: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.leaf.clone())
                    .collect();
                let fn_stmt = parse_function_decl(p)?;
                push_vb_interface_forwarders(&mut members, &fn_stmt, &implemented_members);
                let qualified_forwarders: Vec<String> = implemented_targets
                    .iter()
                    .map(|target| target.forwarder.clone())
                    .collect();
                push_vb_interface_forwarders(&mut members, &fn_stmt, &qualified_forwarders);
                members.push(ClassMember::Method(Box::new(fn_stmt)));
            }
            Rule::operator_decl => {
                members.push(ClassMember::Method(Box::new(parse_operator_decl(p)?)));
            }
            Rule::class_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_class_decl(p)?)));
            }
            Rule::interface_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_interface_decl(p)?)));
            }
            Rule::structure_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_structure_decl(p)?)));
            }
            Rule::enum_decl => {
                members.push(ClassMember::NestedType(Box::new(parse_enum_decl(p)?)));
            }
            Rule::dim_statement => {
                let decls = parse_dim_statement(p)?;
                for d in decls {
                    let field_name = match d.pattern {
                        BindingPattern::Ident(n) => n,
                        _ => String::new(),
                    };
                    members.push(ClassMember::Field {
                        name: field_name,
                        type_hint: d.type_hint,
                        init: d.init,
                        modifiers: Modifiers::default(),
                        with_events: d.with_events,
                        array_bounds: d.array_bounds,
                    });
                }
            }
            Rule::field_decl => {
                let modifiers = parse_field_modifiers(&p);
                let d = parse_field_decl(p)?;
                let field_name = match d.pattern {
                    BindingPattern::Ident(n) => n,
                    _ => String::new(),
                };
                members.push(ClassMember::Field {
                    name: field_name,
                    type_hint: d.type_hint,
                    init: d.init,
                    modifiers,
                    with_events: d.with_events,
                    array_bounds: d.array_bounds,
                });
            }
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
            }
            Rule::NEWLINE | Rule::structure_end => {}
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::StructDecl {
            name,
            interfaces,
            members,
            visibility,
            decorators,
        },
        span,
    ))
}

fn parse_event_decl_to_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let is_custom = pair.as_rule() == Rule::custom_event_decl;
    let mut header = None;
    let mut accessors = Vec::new();
    if is_custom {
        for p in pair.into_inner() {
            match p.as_rule() {
                Rule::event_header => header = Some(p),
                Rule::custom_event_accessor => accessors.push(p),
                _ => {}
            }
        }
    } else if pair.as_rule() == Rule::event_decl {
        header = pair.into_inner().next();
    } else if pair.as_rule() == Rule::event_header {
        header = Some(pair);
    }

    let header = header.ok_or_else(|| "Event missing header".to_string())?;
    let inner = header.into_inner();
    let mut visibility = Visibility::Public;
    let mut name = String::new();
    let mut parameters = Vec::new();
    let mut event_type: Option<String> = None;
    let mut is_shared = false;
    let mut modifiers = Modifiers::default();

    for p in inner {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::param_list => parameters = parse_param_list(p)?,
            Rule::type_name => event_type = Some(p.as_str().to_string()),
            Rule::visibility_modifier => {
                visibility = parse_visibility(p.as_str());
                modifiers.visibility = visibility;
            }
            Rule::sub_modifier_keyword => match p.as_str().to_ascii_lowercase().as_str() {
                "shared" => {
                    is_shared = true;
                    modifiers.is_static = true;
                    modifiers.is_shared = true;
                }
                "overrides" => modifiers.is_override = true,
                "overridable" | "virtual" => modifiers.is_virtual = true,
                "mustoverride" => modifiers.is_abstract = true,
                "notoverridable" => modifiers.is_not_overridable = true,
                "overloads" => modifiers.is_overloads = true,
                "shadows" => {}
                _ => {}
            },
            _ => {}
        }
    }

    if is_shared {
        let mut modifiers = Modifiers::default();
        modifiers.visibility = visibility;
        modifiers.is_static = true;
        modifiers.is_shared = true;
        return Ok(vec![ClassMember::Field {
            name,
            type_hint: event_type,
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
        }]);
    }

    let mut members = vec![ClassMember::Event {
        name: name.clone(),
        type_hint: event_type,
        params: parameters,
        visibility,
    }];

    if !accessors.is_empty() {
        VB_CUSTOM_EVENTS.with(|events| {
            events
                .borrow_mut()
                .insert(name.to_ascii_lowercase(), name.clone());
        });
        for accessor in accessors {
            if let Some(member) =
                parse_custom_event_accessor_to_method(&name, accessor, modifiers.clone())?
            {
                members.push(member);
            }
        }
    }

    Ok(members)
}

fn parse_custom_event_accessor_to_method(
    event_name: &str,
    pair: Pair<Rule>,
    modifiers: Modifiers,
) -> Result<Option<ClassMember>, String> {
    let text = pair.as_str().trim_start();
    let method_prefix = if text
        .get(..10)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("AddHandler"))
    {
        "add"
    } else if text
        .get(..13)
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("RemoveHandler"))
    {
        "remove"
    } else {
        return Ok(None);
    };

    let mut params = Vec::new();
    let mut body = Vec::new();
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_list => params = parse_param_list(p)?,
            Rule::statement_line => {
                for s in p.into_inner() {
                    if s.as_rule() != Rule::NEWLINE && s.as_rule() != Rule::EOI {
                        body.push(parse_statement(s)?);
                    }
                }
            }
            Rule::statement => body.push(parse_statement(p)?),
            Rule::NEWLINE | Rule::EOI => {}
            _ => {}
        }
    }

    if params.is_empty() {
        params.push(Param {
            name: "value".to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
    }

    Ok(Some(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name: format!("{method_prefix}_{event_name}"),
            params,
            return_type: None,
            body,
            modifiers,
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: true,
        },
    )))))
}

fn normalize_vb_custom_event_calls(module: &mut Module) {
    let custom_events = VB_CUSTOM_EVENTS.with(|events| events.borrow().clone());
    if custom_events.is_empty() {
        return;
    }
    for stmt in &mut module.body {
        normalize_vb_custom_event_stmt(stmt, &custom_events);
    }
}

fn normalize_vb_custom_event_member(member: &mut ClassMember, events: &HashMap<String, String>) {
    match member {
        ClassMember::Method(stmt) => normalize_vb_custom_event_stmt(stmt, events),
        ClassMember::Constructor { body, .. } => {
            normalize_vb_custom_event_statements(body, events);
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(getter) = getter {
                normalize_vb_custom_event_statements(getter, events);
            }
            if let Some(setter) = setter {
                normalize_vb_custom_event_statements(&mut setter.body, events);
            }
        }
        ClassMember::NestedType(stmt) => normalize_vb_custom_event_stmt(stmt, events),
        ClassMember::Field {
            init: Some(expr), ..
        }
        | ClassMember::Const { value: expr, .. } => {
            normalize_vb_custom_event_expr(expr, events);
        }
        _ => {}
    }
}

fn normalize_vb_custom_event_statements(
    stmts: &mut Vec<Statement>,
    events: &HashMap<String, String>,
) {
    for stmt in stmts {
        normalize_vb_custom_event_stmt(stmt, events);
    }
}

fn normalize_vb_custom_event_stmt(stmt: &mut Statement, events: &HashMap<String, String>) {
    let replacement = match &stmt.kind {
        StmtKind::CompoundAssign { target, op, value }
            if matches!(op, CompoundOp::Add | CompoundOp::Sub) =>
        {
            custom_event_accessor_call(target, *op, value, events).map(StmtKind::Expr)
        }
        StmtKind::AddHandler {
            control,
            event,
            handler,
        } => custom_event_accessor_call_from_parts(control, event, true, handler, events)
            .map(StmtKind::Expr),
        StmtKind::RemoveHandler {
            control,
            event,
            handler,
        } => custom_event_accessor_call_from_parts(control, event, false, handler, events)
            .map(StmtKind::Expr),
        _ => None,
    };
    if let Some(kind) = replacement {
        stmt.kind = kind;
        return;
    }

    match &mut stmt.kind {
        StmtKind::Expr(expr) => normalize_vb_custom_event_expr(expr, events),
        StmtKind::Block(body)
        | StmtKind::FunctionDecl { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => normalize_vb_custom_event_statements(body, events),
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_custom_event_expr(init, events);
                }
            }
        }
        StmtKind::ClassDecl {
            members,
            decorators,
            ..
        } => {
            for decorator in decorators {
                normalize_vb_custom_event_expr(decorator, events);
            }
            for member in members {
                normalize_vb_custom_event_member(member, events);
            }
        }
        StmtKind::StructDecl {
            members,
            decorators,
            ..
        } => {
            for decorator in decorators {
                normalize_vb_custom_event_expr(decorator, events);
            }
            for member in members {
                normalize_vb_custom_event_member(member, events);
            }
        }
        StmtKind::ModuleDecl { members, .. } => {
            for member in members {
                normalize_vb_custom_event_member(member, events);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => normalize_vb_custom_event_statements(body, events),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            normalize_vb_custom_event_expr(cond, events);
            normalize_vb_custom_event_statements(then_body, events);
            for (elif_cond, elif_body) in elifs {
                normalize_vb_custom_event_expr(elif_cond, events);
                normalize_vb_custom_event_statements(elif_body, events);
            }
            if let Some(else_body) = else_body {
                normalize_vb_custom_event_statements(else_body, events);
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                normalize_vb_custom_event_stmt(init, events);
            }
            if let Some(cond) = cond {
                normalize_vb_custom_event_expr(cond, events);
            }
            if let Some(update) = update {
                normalize_vb_custom_event_expr(update, events);
            }
            normalize_vb_custom_event_statements(body, events);
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            normalize_vb_custom_event_expr(iter, events);
            normalize_vb_custom_event_statements(body, events);
            if let Some(else_body) = else_body {
                normalize_vb_custom_event_statements(else_body, events);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            normalize_vb_custom_event_expr(expr, events);
            for case in cases {
                for cond in &mut case.conditions {
                    normalize_vb_custom_event_case_condition(cond, events);
                }
                normalize_vb_custom_event_statements(&mut case.body, events);
            }
            if let Some(default) = default {
                normalize_vb_custom_event_statements(default, events);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            normalize_vb_custom_event_statements(body, events);
            for catch in catches {
                normalize_vb_custom_event_statements(&mut catch.body, events);
            }
            if let Some(else_body) = else_body {
                normalize_vb_custom_event_statements(else_body, events);
            }
            if let Some(finally) = finally {
                normalize_vb_custom_event_statements(finally, events);
            }
        }
        StmtKind::With { items, body, .. } => {
            for item in items {
                normalize_vb_custom_event_expr(&mut item.expr, events);
            }
            normalize_vb_custom_event_statements(body, events);
        }
        StmtKind::Using { resource, body, .. } => {
            normalize_vb_custom_event_expr(resource, events);
            normalize_vb_custom_event_statements(body, events);
        }
        StmtKind::Lock { expr, body } => {
            normalize_vb_custom_event_expr(expr, events);
            normalize_vb_custom_event_statements(body, events);
        }
        StmtKind::Return(Some(expr))
        | StmtKind::Throw {
            expr: Some(expr), ..
        } => normalize_vb_custom_event_expr(expr, events),
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_custom_event_expr(target, events);
            }
            normalize_vb_custom_event_expr(value, events);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            normalize_vb_custom_event_expr(target, events);
            normalize_vb_custom_event_expr(value, events);
        }
        StmtKind::AddHandler {
            control, handler, ..
        }
        | StmtKind::RemoveHandler {
            control, handler, ..
        } => {
            normalize_vb_custom_event_expr(control, events);
            normalize_vb_custom_event_expr(handler, events);
        }
        StmtKind::RaiseEvent { args, .. }
        | StmtKind::PrintFile { items: args, .. }
        | StmtKind::WriteFile { items: args, .. } => {
            for arg in args {
                normalize_vb_custom_event_expr(arg, events);
            }
        }
        StmtKind::ReDim { bounds, .. } => {
            for bound in bounds {
                normalize_vb_custom_event_expr(bound, events);
            }
        }
        _ => {}
    }
}

fn custom_event_accessor_call(
    target: &Expression,
    op: CompoundOp,
    value: &Expression,
    events: &HashMap<String, String>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &target.kind else {
        return None;
    };
    let event_name = events.get(&field.to_ascii_lowercase())?;
    let prefix = if matches!(op, CompoundOp::Add) {
        "add"
    } else {
        "remove"
    };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: object.clone(),
            field: format!("{prefix}_{event_name}"),
            null_safe: false,
        })),
        args: vec![Argument::positional(value.clone())],
        optional: false,
    }))
}

fn custom_event_accessor_call_from_parts(
    control: &Expression,
    event: &str,
    add: bool,
    handler: &Expression,
    events: &HashMap<String, String>,
) -> Option<Expression> {
    let event_name = events.get(&event.to_ascii_lowercase())?;
    let prefix = if add { "add" } else { "remove" };
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(control.clone()),
            field: format!("{prefix}_{event_name}"),
            null_safe: false,
        })),
        args: vec![Argument::positional(handler.clone())],
        optional: false,
    }))
}

fn normalize_vb_custom_event_case_condition(
    cond: &mut CaseCondition,
    events: &HashMap<String, String>,
) {
    match cond {
        CaseCondition::Value(expr) | CaseCondition::Comparison { expr, .. } => {
            normalize_vb_custom_event_expr(expr, events);
        }
        CaseCondition::Range { from, to } => {
            normalize_vb_custom_event_expr(from, events);
            normalize_vb_custom_event_expr(to, events);
        }
    }
}

fn normalize_vb_custom_event_expr(expr: &mut Expression, events: &HashMap<String, String>) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } => {
            normalize_vb_custom_event_expr(callee, events);
            for arg in args {
                normalize_vb_custom_event_expr(&mut arg.value, events);
            }
        }
        ExprKind::Member { object, .. } => normalize_vb_custom_event_expr(object, events),
        ExprKind::Index { object, index, .. } => {
            normalize_vb_custom_event_expr(object, events);
            normalize_vb_custom_event_expr(index, events);
        }
        ExprKind::Binary { left, right, .. } => {
            normalize_vb_custom_event_expr(left, events);
            normalize_vb_custom_event_expr(right, events);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Cast { expr, .. }
        | ExprKind::Await(expr) => normalize_vb_custom_event_expr(expr, events),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_custom_event_expr(cond, events);
            normalize_vb_custom_event_expr(then, events);
            normalize_vb_custom_event_expr(else_, events);
        }
        ExprKind::New { class, args } => {
            normalize_vb_custom_event_expr(class, events);
            for arg in args {
                normalize_vb_custom_event_expr(&mut arg.value, events);
            }
        }
        ExprKind::Assign { target, value } => {
            normalize_vb_custom_event_expr(target, events);
            normalize_vb_custom_event_expr(value, events);
        }
        ExprKind::Lambda { body, .. } => match body {
            LambdaBody::Expr(value) => normalize_vb_custom_event_expr(value, events),
            LambdaBody::Block(body) => normalize_vb_custom_event_statements(body, events),
        },
        ExprKind::Array(items) => {
            for item in items {
                if let Some(key) = &mut item.key {
                    normalize_vb_custom_event_expr(key, events);
                }
                normalize_vb_custom_event_expr(&mut item.value, events);
            }
        }
        ExprKind::Tuple(items) => {
            for item in items {
                normalize_vb_custom_event_expr(item, events);
            }
        }
        ExprKind::NamedTuple { fields, .. } => {
            for (_, value) in fields {
                normalize_vb_custom_event_expr(value, events);
            }
        }
        ExprKind::Set(items) => {
            for item in items {
                normalize_vb_custom_event_expr(item, events);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { value, .. }
                    | ObjectProperty::Computed { value, .. }
                    | ObjectProperty::Spread(value) => {
                        normalize_vb_custom_event_expr(value, events)
                    }
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_custom_event_stmt(value, events);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

// ── Syntax Extensions Implementation ──

fn parse_synclock_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut inner = pair.into_inner();
    let lock_expr = parse_expression(inner.next().unwrap())?;
    let mut body = Vec::new();

    for p in inner {
        match p.as_rule() {
            Rule::statement | Rule::statement_line => {
                for stmt_pair in p.into_inner() {
                    if stmt_pair.as_rule() == Rule::NEWLINE || stmt_pair.as_rule() == Rule::EOI {
                        continue;
                    }
                    body.push(parse_statement(stmt_pair)?);
                }
            }
            Rule::synclock_end | Rule::NEWLINE => {}
            _ => {}
        }
    }
    Ok(Statement::with_span(
        StmtKind::Lock {
            expr: lock_expr,
            body,
        },
        span,
    ))
}

fn parse_query_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    // LINQ query expressions are VB-specific. Map to a Call expression with
    // a chain of method calls that the compiler can recognize.
    // For now, produce a placeholder that preserves the structure.
    let _span = to_span(&pair);
    let query_text = pair.as_str().to_string();
    let normalized_query_text = normalize_vb_inline_query_clauses(&query_text);
    if normalized_query_text
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .count()
        > 1
    {
        if let Some(rewritten) = parse_multiline_vb_query_text(&normalized_query_text)? {
            return Ok(rewritten);
        }
    }
    let mut inner = pair.into_inner();
    let from_clause_pair = inner.next().unwrap();

    // Parse From clause
    let mut from_inner = from_clause_pair.into_inner();
    let mut range_var = String::new();
    let mut collection_expr = Expression::null();

    while let Some(id_pair) = from_inner.next() {
        if id_pair.as_rule() == Rule::identifier {
            range_var = id_pair.as_str().to_string();

            while let Some(p) = from_inner.next() {
                match p.as_rule() {
                    Rule::expression => {
                        collection_expr = parse_expression(p)?;
                        break;
                    }
                    Rule::type_name => {} // skip
                    _ => {}
                }
            }
            break; // Take first range variable for ForIn mapping
        }
    }

    // Build the query as a series of method calls on the collection
    let query_body_pair = inner.next().unwrap();
    let body_inner = query_body_pair.into_inner();
    let mut result_expr = collection_expr.clone();

    for p in body_inner {
        match p.as_rule() {
            Rule::query_operator => {
                let op = p.into_inner().next().unwrap();
                result_expr = apply_vb_query_operator(result_expr, op, &range_var)?;
            }
            Rule::where_clause
            | Rule::skip_clause
            | Rule::take_clause
            | Rule::skip_while_clause
            | Rule::take_while_clause
            | Rule::distinct_clause
            | Rule::order_by_clause
            | Rule::let_clause => {
                result_expr = apply_vb_query_operator(result_expr, p, &range_var)?;
            }
            Rule::select_or_group_clause => {
                let inner_sg = p.into_inner().next().unwrap();
                match inner_sg.as_rule() {
                    Rule::select_clause => {
                        let exprs: Vec<Expression> = inner_sg
                            .into_inner()
                            .map(|x| parse_expression(x))
                            .collect::<Result<Vec<_>, _>>()?;
                        if !exprs.is_empty() {
                            // .Select(Function(x) expr)
                            let select_body = if exprs.len() == 1 {
                                exprs.into_iter().next().unwrap()
                            } else {
                                // Multiple select expressions → tuple-like
                                Expression::new(ExprKind::Array(
                                    exprs
                                        .into_iter()
                                        .map(|e| ArrayElement {
                                            key: None,
                                            value: e,
                                            spread: false,
                                            by_ref: false,
                                        })
                                        .collect(),
                                ))
                            };
                            let identity_select = matches!(
                                &select_body.kind,
                                ExprKind::Ident(name) if name.eq_ignore_ascii_case(&range_var)
                            );
                            if !identity_select {
                                result_expr = build_linq_lambda_call(
                                    result_expr,
                                    "Select",
                                    &range_var,
                                    select_body,
                                );
                            }
                        }
                    }
                    Rule::group_clause => {
                        let mut exprs = Vec::new();
                        for x in inner_sg.into_inner() {
                            if x.as_rule() == Rule::expression {
                                exprs.push(parse_expression(x)?);
                            }
                        }
                        if exprs.len() >= 2 {
                            let key = exprs.pop().unwrap();
                            let _item = exprs.pop().unwrap();
                            // .GroupBy(Function(x) key)
                            result_expr =
                                build_linq_lambda_call(result_expr, "GroupBy", &range_var, key);
                        }
                    }
                    _ => {}
                }
            }
            Rule::select_clause | Rule::group_clause => {
                result_expr = apply_vb_query_projection(result_expr, p, &range_var)?;
            }
            _ => {}
        }
    }

    if query_text.contains('\n') || query_text.contains('\r') {
        if let Some(rewritten) = parse_multiline_vb_query_text(&query_text)? {
            return Ok(rewritten);
        }
    }

    Ok(result_expr)
}

fn vb_infer_tuple_element_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { field, .. } => Some(field.clone()),
        _ => None,
    }
}

fn parse_aggregate_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut range_var_seen = false;
    let mut collection = None;
    let mut aggregates = Vec::new();

    for part in pair.into_inner() {
        match part.as_rule() {
            Rule::identifier if !range_var_seen => {
                range_var_seen = true;
            }
            Rule::expression => collection = Some(parse_expression(part)?),
            Rule::aggregate_call => {
                let mut inner = part.into_inner();
                let name = inner
                    .next()
                    .map(|p| p.as_str().to_string())
                    .unwrap_or_else(|| "Count".to_string());
                let args = inner
                    .next()
                    .map(parse_argument_list)
                    .transpose()?
                    .unwrap_or_default();
                aggregates.push((name, args));
            }
            Rule::type_name => {}
            _ => {}
        }
    }

    let receiver = collection.unwrap_or_else(Expression::null);
    if aggregates.is_empty() {
        aggregates.push(("Count".to_string(), Vec::new()));
    }
    if aggregates.len() == 1 {
        let (method, args) = aggregates.remove(0);
        return Ok(build_linq_call_args(receiver, &method, args));
    }

    let mut props = Vec::new();
    for (method, args) in aggregates {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&method),
            value: build_linq_call_args(receiver.clone(), &method, args),
        });
    }
    Ok(Expression::new(ExprKind::Object(props)))
}

fn parse_multiline_vb_query_text(source: &str) -> Result<Option<Expression>, String> {
    let lines: Vec<String> = source
        .lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .map(str::to_string)
        .collect();
    let Some(first) = lines.first() else {
        return Ok(None);
    };
    let from_ranges = parse_vb_query_from_ranges(first)?;
    if from_ranges.is_empty() {
        return Ok(None);
    }
    let range_var = from_ranges[0].0.clone();
    let mut result = parse_expression_str(&from_ranges[0].1)?;
    let mut let_bindings: HashMap<String, Expression> = HashMap::new();
    let mut lambda_param = range_var.to_string();
    let mut nested_froms: Vec<(String, Expression)> = Vec::new();
    for (name, source) in from_ranges.iter().skip(1) {
        nested_froms.push((name.clone(), parse_expression_str(source)?));
    }
    let mut pending_join: Option<VbQueryJoin> = None;
    let mut pending_group_join: Option<VbQueryGroupJoin> = None;

    for line in lines.iter().skip(1) {
        let lower = line.to_ascii_lowercase();
        if lower.starts_with("where ") {
            let mut expr = parse_expression_str(line[6..].trim())?;
            substitute_vb_constructor_expr(&mut expr, &let_bindings);
            result = build_linq_lambda_call(result, "Where", &lambda_param, expr);
        } else if lower.starts_with("skip while ") {
            let mut expr = parse_expression_str(line[11..].trim())?;
            substitute_vb_constructor_expr(&mut expr, &let_bindings);
            result = build_linq_lambda_call(result, "SkipWhile", &lambda_param, expr);
        } else if lower.starts_with("take while ") {
            let mut expr = parse_expression_str(line[11..].trim())?;
            substitute_vb_constructor_expr(&mut expr, &let_bindings);
            result = build_linq_lambda_call(result, "TakeWhile", &lambda_param, expr);
        } else if lower.starts_with("skip ") {
            let expr = parse_expression_str(line[5..].trim())?;
            result = build_linq_value_call(result, "Skip", expr);
        } else if lower.starts_with("take ") {
            let expr = parse_expression_str(line[5..].trim())?;
            result = build_linq_value_call(result, "Take", expr);
        } else if lower == "distinct" {
            result = Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(result),
                    field: "Distinct".to_string(),
                    null_safe: false,
                })),
                args: vec![],
                optional: false,
            });
        } else if lower.starts_with("order by ") {
            for ordering in split_vb_top_level_commas(line[9..].trim()) {
                let mut key_src = ordering.trim();
                let lower_key = key_src.to_ascii_lowercase();
                let descending = lower_key.ends_with(" descending");
                let ascending = lower_key.ends_with(" ascending");
                if descending {
                    key_src = key_src[..key_src.len() - " descending".len()].trim();
                } else if ascending {
                    key_src = key_src[..key_src.len() - " ascending".len()].trim();
                }
                let mut expr = parse_expression_str(key_src)?;
                substitute_vb_constructor_expr(&mut expr, &let_bindings);
                result = build_linq_lambda_call(
                    result,
                    if descending {
                        "OrderByDescending"
                    } else {
                        "OrderBy"
                    },
                    &lambda_param,
                    expr,
                );
            }
        } else if lower.starts_with("let ") {
            for binding in split_vb_top_level_commas(line[4..].trim()) {
                let Some(eq_idx) = binding.find('=') else {
                    continue;
                };
                let name = binding[..eq_idx].trim();
                let expr_src = binding[eq_idx + 1..].trim();
                if name.is_empty() || expr_src.is_empty() {
                    continue;
                }
                let mut expr = parse_expression_str(expr_src)?;
                substitute_vb_constructor_expr(&mut expr, &let_bindings);
                let_bindings.insert(name.to_ascii_lowercase(), expr);
            }
        } else if lower.starts_with("join ") {
            if let Some(join) = parse_vb_query_join_line(line)? {
                pending_join = Some(join);
            }
        } else if lower.starts_with("group join ") {
            if let Some(join) = parse_vb_query_group_join_line(line)? {
                pending_group_join = Some(join);
            }
        } else if lower.starts_with("group by ") {
            let Some(into_idx) = lower.find(" into ") else {
                continue;
            };
            let keys_src = line[9..into_idx].trim();
            let into_src = line[into_idx + 6..].trim();
            let key_parts = split_vb_top_level_commas(keys_src);
            let mut key_bindings = Vec::new();
            for key_part in key_parts {
                let (key_name, key_expr_src) =
                    if let Some(eq_idx) = find_vb_top_level_equals(key_part) {
                        (
                            Some(key_part[..eq_idx].trim().to_string()),
                            key_part[eq_idx + 1..].trim(),
                        )
                    } else {
                        (vb_query_projection_name(key_part), key_part.trim())
                    };
                let Some(key_name) = key_name.filter(|name| !name.is_empty()) else {
                    continue;
                };
                let mut key_expr = parse_expression_str(key_expr_src)?;
                substitute_vb_constructor_expr(&mut key_expr, &let_bindings);
                normalize_vb_query_member_fields(&mut key_expr);
                key_bindings.push((key_name, key_expr));
            }
            let key_expr = if key_bindings.len() == 1 {
                key_bindings[0].1.clone()
            } else {
                Expression::new(ExprKind::Object(
                    key_bindings
                        .iter()
                        .map(|(name, expr)| ObjectProperty::KeyValue {
                            key: Expression::string(&name.to_ascii_lowercase()),
                            value: expr.clone(),
                        })
                        .collect(),
                ))
            };
            result = build_linq_lambda_call(result, "GroupBy", &range_var, key_expr);

            lambda_param = "__vb_group".to_string();
            let_bindings.clear();
            for (name, _) in &key_bindings {
                let replacement = if key_bindings.len() == 1 {
                    build_vb_index_expr(Expression::ident("__vb_group"), "Key")
                } else {
                    build_vb_index_expr(
                        build_vb_index_expr(Expression::ident("__vb_group"), "Key"),
                        &name.to_ascii_lowercase(),
                    )
                };
                let_bindings.insert(name.to_ascii_lowercase(), replacement);
            }
            let group_name = into_src
                .split('=')
                .next()
                .map(str::trim)
                .filter(|name| !name.is_empty())
                .unwrap_or("Group");
            let_bindings.insert(
                group_name.to_ascii_lowercase(),
                build_vb_cast_expr(
                    build_vb_index_expr(Expression::ident("__vb_group"), "Items"),
                    "IEnumerable",
                ),
            );
        } else if lower.starts_with("select ") {
            if let Some(group_join) = pending_group_join.take() {
                result = build_vb_group_join_select(
                    result,
                    &range_var,
                    group_join,
                    line[7..].trim(),
                    &let_bindings,
                )?;
                continue;
            }
            if let Some(join) = pending_join.take() {
                result = build_vb_join_select(
                    result,
                    &range_var,
                    join,
                    line[7..].trim(),
                    &let_bindings,
                )?;
                continue;
            }
            let mut expr = parse_vb_query_select_expression(line[7..].trim(), &let_bindings)?;
            normalize_vb_query_member_fields(&mut expr);
            if !nested_froms.is_empty() {
                for (name, collection) in nested_froms.iter().rev() {
                    expr = build_linq_lambda_call(collection.clone(), "Select", name, expr);
                }
                result = build_linq_lambda_call(result, "SelectMany", &lambda_param, expr);
                continue;
            }
            let identity_select = matches!(
                &expr.kind,
                ExprKind::Ident(name) if name.eq_ignore_ascii_case(&lambda_param)
            );
            if !identity_select {
                result = build_linq_lambda_call(result, "Select", &lambda_param, expr);
            }
        }
    }

    Ok(Some(result))
}

#[derive(Clone)]
struct VbQueryJoin {
    var: String,
    collection: Expression,
    conditions: Vec<(Expression, Expression)>,
}

#[derive(Clone)]
struct VbQueryGroupJoin {
    var: String,
    collection: Expression,
    conditions: Vec<(Expression, Expression)>,
    group_name: String,
}

fn parse_vb_query_from_ranges(line: &str) -> Result<Vec<(String, String)>, String> {
    let trimmed = line.trim();
    let lower = trimmed.to_ascii_lowercase();
    if !lower.starts_with("from ") {
        return Ok(Vec::new());
    }
    let mut ranges = Vec::new();
    for part in split_vb_top_level_commas(trimmed[5..].trim()) {
        let lower_part = part.to_ascii_lowercase();
        let Some(in_pos) = lower_part.find(" in ") else {
            continue;
        };
        let name_part = part[..in_pos].trim();
        let name = name_part.split_whitespace().next().unwrap_or("").trim();
        let collection = part[in_pos + 4..].trim();
        if !name.is_empty() && !collection.is_empty() {
            ranges.push((name.to_string(), collection.to_string()));
        }
    }
    Ok(ranges)
}

fn parse_vb_query_join_line(line: &str) -> Result<Option<VbQueryJoin>, String> {
    let trimmed = line.trim();
    let lower = trimmed.to_ascii_lowercase();
    if !lower.starts_with("join ") {
        return Ok(None);
    }
    let Some(in_pos) = lower.find(" in ") else {
        return Ok(None);
    };
    let var = trimmed[5..in_pos]
        .split_whitespace()
        .next()
        .unwrap_or("")
        .trim();
    let Some(on_pos_rel) = lower[in_pos + 4..].find(" on ") else {
        return Ok(None);
    };
    let on_pos = in_pos + 4 + on_pos_rel;
    let collection_src = trimmed[in_pos + 4..on_pos].trim();
    let conditions_src = trimmed[on_pos + 4..].trim();
    if var.is_empty() || collection_src.is_empty() || conditions_src.is_empty() {
        return Ok(None);
    }
    Ok(Some(VbQueryJoin {
        var: var.to_string(),
        collection: parse_expression_str(collection_src)?,
        conditions: parse_vb_join_conditions(conditions_src)?,
    }))
}

fn parse_vb_query_group_join_line(line: &str) -> Result<Option<VbQueryGroupJoin>, String> {
    let trimmed = line.trim();
    let lower = trimmed.to_ascii_lowercase();
    if !lower.starts_with("group join ") {
        return Ok(None);
    }
    let Some(in_pos) = lower.find(" in ") else {
        return Ok(None);
    };
    let var = trimmed["group join ".len()..in_pos]
        .split_whitespace()
        .next()
        .unwrap_or("")
        .trim();
    let Some(on_pos_rel) = lower[in_pos + 4..].find(" on ") else {
        return Ok(None);
    };
    let on_pos = in_pos + 4 + on_pos_rel;
    let Some(into_pos_rel) = lower[on_pos + 4..].find(" into ") else {
        return Ok(None);
    };
    let into_pos = on_pos + 4 + into_pos_rel;
    let collection_src = trimmed[in_pos + 4..on_pos].trim();
    let conditions_src = trimmed[on_pos + 4..into_pos].trim();
    let group_name = trimmed[into_pos + 6..]
        .split('=')
        .next()
        .map(str::trim)
        .filter(|name| !name.is_empty())
        .unwrap_or("Group");
    if var.is_empty() || collection_src.is_empty() || conditions_src.is_empty() {
        return Ok(None);
    }
    Ok(Some(VbQueryGroupJoin {
        var: var.to_string(),
        collection: parse_expression_str(collection_src)?,
        conditions: parse_vb_join_conditions(conditions_src)?,
        group_name: group_name.to_string(),
    }))
}

fn parse_vb_join_conditions(source: &str) -> Result<Vec<(Expression, Expression)>, String> {
    let mut conditions = Vec::new();
    for part in split_vb_top_level_keyword(source, "And") {
        let lower = part.to_ascii_lowercase();
        let Some(eq_pos) = lower.find(" equals ") else {
            continue;
        };
        let left = parse_expression_str(part[..eq_pos].trim())?;
        let right = parse_expression_str(part[eq_pos + " equals ".len()..].trim())?;
        conditions.push((left, right));
    }
    Ok(conditions)
}

fn build_vb_join_select(
    receiver: Expression,
    outer_var: &str,
    join: VbQueryJoin,
    select_src: &str,
    replacements: &HashMap<String, Expression>,
) -> Result<Expression, String> {
    let mut pred = build_vb_join_condition_expr(join.conditions);
    substitute_vb_constructor_expr(&mut pred, replacements);
    let filtered = build_linq_lambda_call(join.collection, "Where", &join.var, pred);
    let select_expr = parse_vb_query_select_expression(select_src, replacements)?;
    let selected = build_linq_lambda_call(filtered, "Select", &join.var, select_expr);
    Ok(build_linq_lambda_call(
        receiver,
        "SelectMany",
        outer_var,
        selected,
    ))
}

fn build_vb_group_join_select(
    receiver: Expression,
    outer_var: &str,
    join: VbQueryGroupJoin,
    select_src: &str,
    replacements: &HashMap<String, Expression>,
) -> Result<Expression, String> {
    let mut pred = build_vb_join_condition_expr(join.conditions);
    substitute_vb_constructor_expr(&mut pred, replacements);
    let group_expr = build_linq_lambda_call(join.collection, "Where", &join.var, pred);
    let mut scoped = replacements.clone();
    scoped.insert(join.group_name.to_ascii_lowercase(), group_expr);
    let select_expr = parse_vb_query_select_expression(select_src, &scoped)?;
    Ok(build_linq_lambda_call(
        receiver,
        "Select",
        outer_var,
        select_expr,
    ))
}

fn build_vb_join_condition_expr(conditions: Vec<(Expression, Expression)>) -> Expression {
    let mut iter = conditions.into_iter().map(|(left, right)| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(left),
            right: Box::new(right),
        })
    });
    let Some(mut expr) = iter.next() else {
        return Expression::bool(true);
    };
    for next in iter {
        expr = Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(expr),
            right: Box::new(next),
        });
    }
    expr
}

fn parse_vb_query_select_expression(
    source: &str,
    replacements: &HashMap<String, Expression>,
) -> Result<Expression, String> {
    let parts = split_vb_top_level_commas(source);
    if parts.len() == 1 {
        let mut expr = parse_expression_str(parts[0])?;
        substitute_vb_constructor_expr(&mut expr, replacements);
        return Ok(expr);
    }

    let mut props = Vec::new();
    for part in parts {
        let (name, expr_src) = if let Some(eq_idx) = find_vb_top_level_equals(part) {
            (part[..eq_idx].trim().to_string(), part[eq_idx + 1..].trim())
        } else {
            (
                vb_query_projection_name(part).unwrap_or_else(|| "Value".to_string()),
                part.trim(),
            )
        };
        let mut expr =
            if let Some(group_sum) = parse_vb_query_group_sum_expr(expr_src, replacements)? {
                group_sum
            } else {
                let mut expr = parse_expression_str(expr_src)?;
                substitute_vb_constructor_expr(&mut expr, replacements);
                expr
            };
        rewrite_vb_group_items_aggregate_expr(&mut expr);
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&name.to_ascii_lowercase()),
            value: expr,
        });
    }
    Ok(Expression::new(ExprKind::Object(props)))
}

fn parse_vb_query_group_sum_expr(
    source: &str,
    replacements: &HashMap<String, Expression>,
) -> Result<Option<Expression>, String> {
    let trimmed = source.trim();
    let lower = trimmed.to_ascii_lowercase();
    for (name, items_expr) in replacements {
        let prefix = format!("{name}.sum");
        if !lower.starts_with(&prefix) {
            continue;
        }
        let rest = trimmed[prefix.len()..].trim();
        if rest == "()" {
            return Ok(Some(build_vb_index_expr(
                Expression::ident("__vb_group"),
                "Sum",
            )));
        }
        if rest.starts_with('(') && rest.ends_with(')') {
            let inner = rest[1..rest.len() - 1].trim();
            let inner_lower = inner.to_ascii_lowercase();
            if inner_lower.starts_with("function(") {
                let Some(close_idx) = inner.find(')') else {
                    return Ok(None);
                };
                let param = inner["Function(".len()..close_idx].trim();
                let body_src = inner[close_idx + 1..].trim();
                if !param.is_empty() && !body_src.is_empty() {
                    let mut body = parse_expression_str(body_src)?;
                    normalize_vb_query_member_fields(&mut body);
                    let lambda = Expression::new(ExprKind::Lambda {
                        params: vec![Param {
                            name: param.to_string(),
                            type_hint: None,
                            default: None,
                            pass_by: PassBy::Value,
                            is_rest: false,
                            is_kwargs: false,
                            is_optional: false,
                            is_nullable: false,
                        }],
                        body: LambdaBody::Expr(Box::new(body)),
                        is_async: false,
                        captures: vec![],
                    });
                    return Ok(Some(Expression::new(ExprKind::Call {
                        callee: Box::new(build_dotted_expr("System.Linq.Enumerable.Sum")),
                        args: vec![
                            Argument::positional(items_expr.clone()),
                            Argument::positional(lambda),
                        ],
                        optional: false,
                    })));
                }
            }
        }
    }
    Ok(None)
}

fn normalize_vb_query_member_fields(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Member { object, field, .. } => {
            normalize_vb_query_member_fields(object);
            *field = field.to_ascii_lowercase();
        }
        ExprKind::Call { callee, args, .. } => {
            if let ExprKind::Member { object, .. } = &mut callee.kind {
                normalize_vb_query_member_fields(object);
            } else {
                normalize_vb_query_member_fields(callee);
            }
            for arg in args {
                normalize_vb_query_member_fields(&mut arg.value);
            }
        }
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            normalize_vb_query_member_fields(left);
            normalize_vb_query_member_fields(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => normalize_vb_query_member_fields(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            normalize_vb_query_member_fields(cond);
            normalize_vb_query_member_fields(then);
            normalize_vb_query_member_fields(else_);
        }
        ExprKind::Index { object, index, .. } => {
            normalize_vb_query_member_fields(object);
            normalize_vb_query_member_fields(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                normalize_vb_query_member_fields(&mut item.value);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { value, .. }
                    | ObjectProperty::Computed { value, .. } => {
                        normalize_vb_query_member_fields(value);
                    }
                    ObjectProperty::Spread(expr) => normalize_vb_query_member_fields(expr),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        normalize_vb_query_member_fields_in_stmt(value);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        ExprKind::Lambda { body, .. } => {
            if let LambdaBody::Expr(body) = body {
                normalize_vb_query_member_fields(body);
            }
        }
        _ => {}
    }
}

fn normalize_vb_query_member_fields_in_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            normalize_vb_query_member_fields(expr);
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                normalize_vb_query_member_fields(target);
            }
            normalize_vb_query_member_fields(value);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    normalize_vb_query_member_fields(init);
                }
            }
        }
        _ => {}
    }
}

fn rewrite_vb_group_items_aggregate_expr(expr: &mut Expression) {
    match &mut expr.kind {
        ExprKind::Call { callee, args, .. } if args.is_empty() => {
            rewrite_vb_group_items_aggregate_expr(callee);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_vb_group_items_aggregate_expr(callee);
            for arg in args {
                rewrite_vb_group_items_aggregate_expr(&mut arg.value);
            }
        }
        ExprKind::Member { object, .. } => rewrite_vb_group_items_aggregate_expr(object),
        ExprKind::Binary { left, right, .. }
        | ExprKind::NullCoalesce { left, right }
        | ExprKind::Walrus {
            target: left,
            value: right,
        } => {
            rewrite_vb_group_items_aggregate_expr(left);
            rewrite_vb_group_items_aggregate_expr(right);
        }
        ExprKind::Unary { expr, .. }
        | ExprKind::Cast { expr, .. }
        | ExprKind::RefLoad(expr)
        | ExprKind::Await(expr)
        | ExprKind::YieldFrom(expr)
        | ExprKind::Spread(expr)
        | ExprKind::Void(expr)
        | ExprKind::Delete(expr)
        | ExprKind::TypeOf(expr) => rewrite_vb_group_items_aggregate_expr(expr),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_vb_group_items_aggregate_expr(cond);
            rewrite_vb_group_items_aggregate_expr(then);
            rewrite_vb_group_items_aggregate_expr(else_);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_vb_group_items_aggregate_expr(object);
            rewrite_vb_group_items_aggregate_expr(index);
        }
        ExprKind::Array(items) => {
            for item in items {
                rewrite_vb_group_items_aggregate_expr(&mut item.value);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_vb_group_items_aggregate_expr(key);
                        rewrite_vb_group_items_aggregate_expr(value);
                    }
                    ObjectProperty::Spread(value) => rewrite_vb_group_items_aggregate_expr(value),
                    ObjectProperty::Method { value, .. }
                    | ObjectProperty::Accessor { value, .. } => {
                        rewrite_vb_group_items_aggregate_stmt(value);
                    }
                    ObjectProperty::Shorthand(_) => {}
                }
            }
        }
        _ => {}
    }
}

fn rewrite_vb_group_items_aggregate_stmt(stmt: &mut Statement) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) | StmtKind::Return(Some(expr)) => {
            rewrite_vb_group_items_aggregate_expr(expr)
        }
        _ => {}
    }
}

fn vb_query_projection_name(source: &str) -> Option<String> {
    let trimmed = source.trim();
    if trimmed.is_empty() {
        return None;
    }
    let tail = trimmed.rsplit('.').next().unwrap_or(trimmed).trim();
    let valid = tail
        .chars()
        .next()
        .is_some_and(|ch| ch.is_ascii_alphabetic() || ch == '_')
        && tail
            .chars()
            .all(|ch| ch.is_ascii_alphanumeric() || ch == '_');
    valid.then(|| tail.to_string())
}

fn build_vb_index_expr(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(Expression::string(field)),
        null_safe: false,
    })
}

fn build_vb_cast_expr(expr: Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn find_vb_top_level_equals(source: &str) -> Option<usize> {
    let mut paren = 0usize;
    let mut brace = 0usize;
    let mut in_string = false;
    let bytes = source.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        let ch = bytes[i] as char;
        if in_string {
            if ch == '"' {
                if i + 1 < bytes.len() && bytes[i + 1] as char == '"' {
                    i += 1;
                } else {
                    in_string = false;
                }
            }
        } else {
            match ch {
                '"' => in_string = true,
                '(' => paren += 1,
                ')' => paren = paren.saturating_sub(1),
                '{' => brace += 1,
                '}' => brace = brace.saturating_sub(1),
                '=' if paren == 0 && brace == 0 => return Some(i),
                _ => {}
            }
        }
        i += 1;
    }
    None
}

fn split_vb_top_level_commas(source: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut paren = 0usize;
    let mut brace = 0usize;
    let mut in_string = false;
    let bytes = source.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        let ch = bytes[i] as char;
        if in_string {
            if ch == '"' {
                if i + 1 < bytes.len() && bytes[i + 1] as char == '"' {
                    i += 1;
                } else {
                    in_string = false;
                }
            }
        } else {
            match ch {
                '"' => in_string = true,
                '(' => paren += 1,
                ')' => paren = paren.saturating_sub(1),
                '{' => brace += 1,
                '}' => brace = brace.saturating_sub(1),
                ',' if paren == 0 && brace == 0 => {
                    parts.push(source[start..i].trim());
                    start = i + 1;
                }
                _ => {}
            }
        }
        i += 1;
    }
    parts.push(source[start..].trim());
    parts.into_iter().filter(|part| !part.is_empty()).collect()
}

fn split_vb_top_level_keyword<'a>(source: &'a str, keyword: &str) -> Vec<&'a str> {
    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut paren = 0usize;
    let mut brace = 0usize;
    let mut in_string = false;
    let lower = source.to_ascii_lowercase();
    let keyword_lower = keyword.to_ascii_lowercase();
    let bytes = source.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        let ch = bytes[i] as char;
        if in_string {
            if ch == '"' {
                if i + 1 < bytes.len() && bytes[i + 1] as char == '"' {
                    i += 1;
                } else {
                    in_string = false;
                }
            }
        } else {
            match ch {
                '"' => in_string = true,
                '(' => paren += 1,
                ')' => paren = paren.saturating_sub(1),
                '{' => brace += 1,
                '}' => brace = brace.saturating_sub(1),
                _ if paren == 0 && brace == 0 => {
                    let end = i + keyword_lower.len();
                    if end <= lower.len()
                        && lower[i..end] == keyword_lower
                        && source[..i].ends_with(char::is_whitespace)
                        && source[end..].starts_with(char::is_whitespace)
                    {
                        parts.push(source[start..i].trim());
                        start = end;
                        i = end;
                        continue;
                    }
                }
                _ => {}
            }
        }
        i += 1;
    }
    parts.push(source[start..].trim());
    parts.into_iter().filter(|part| !part.is_empty()).collect()
}

fn apply_vb_query_operator(
    result_expr: Expression,
    op: Pair<Rule>,
    range_var: &str,
) -> Result<Expression, String> {
    Ok(match op.as_rule() {
        Rule::where_clause => {
            let filter_expr = parse_expression(op.into_inner().next().unwrap())?;
            build_linq_lambda_call(result_expr, "Where", range_var, filter_expr)
        }
        Rule::skip_clause | Rule::take_clause => {
            let op_rule = op.as_rule();
            let value = parse_expression(op.into_inner().next().unwrap())?;
            let method = if op_rule == Rule::skip_clause {
                "Skip"
            } else {
                "Take"
            };
            build_linq_value_call(result_expr, method, value)
        }
        Rule::skip_while_clause | Rule::take_while_clause => {
            let op_rule = op.as_rule();
            let pred = parse_expression(op.into_inner().next().unwrap())?;
            let method = if op_rule == Rule::skip_while_clause {
                "SkipWhile"
            } else {
                "TakeWhile"
            };
            build_linq_lambda_call(result_expr, method, range_var, pred)
        }
        Rule::distinct_clause => Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(result_expr),
                field: "Distinct".to_string(),
                null_safe: false,
            })),
            args: vec![],
            optional: false,
        }),
        Rule::order_by_clause => {
            let mut out = result_expr;
            for ord in op.into_inner() {
                let raw_ordering = ord.as_str().trim().to_ascii_lowercase();
                let mut ord_inner = ord.into_inner();
                let key_expr = parse_expression(ord_inner.next().unwrap())?;
                let descending = raw_ordering.ends_with(" descending");
                let method = if descending {
                    "OrderByDescending"
                } else {
                    "OrderBy"
                };
                out = build_linq_lambda_call(out, method, range_var, key_expr);
            }
            out
        }
        Rule::let_clause => result_expr,
        _ => result_expr,
    })
}

fn apply_vb_query_projection(
    result_expr: Expression,
    inner_sg: Pair<Rule>,
    range_var: &str,
) -> Result<Expression, String> {
    Ok(match inner_sg.as_rule() {
        Rule::select_clause => {
            let exprs: Vec<Expression> = inner_sg
                .into_inner()
                .map(|x| parse_expression(x))
                .collect::<Result<Vec<_>, _>>()?;
            if exprs.is_empty() {
                result_expr
            } else {
                let select_body = if exprs.len() == 1 {
                    exprs.into_iter().next().unwrap()
                } else {
                    Expression::new(ExprKind::Array(
                        exprs
                            .into_iter()
                            .map(|e| ArrayElement {
                                key: None,
                                value: e,
                                spread: false,
                                by_ref: false,
                            })
                            .collect(),
                    ))
                };
                let identity_select = matches!(
                    &select_body.kind,
                    ExprKind::Ident(name) if name.eq_ignore_ascii_case(range_var)
                );
                if identity_select {
                    result_expr
                } else {
                    build_linq_lambda_call(result_expr, "Select", range_var, select_body)
                }
            }
        }
        Rule::group_clause => {
            let mut exprs = Vec::new();
            for x in inner_sg.into_inner() {
                if x.as_rule() == Rule::expression {
                    exprs.push(parse_expression(x)?);
                }
            }
            if exprs.len() >= 2 {
                let key = exprs.pop().unwrap();
                build_linq_lambda_call(result_expr, "GroupBy", range_var, key)
            } else {
                result_expr
            }
        }
        _ => result_expr,
    })
}

fn build_linq_value_call(receiver: Expression, method: &str, arg: Expression) -> Expression {
    build_linq_call_args(receiver, method, vec![Argument::positional(arg)])
}

fn build_linq_call_args(receiver: Expression, method: &str, args: Vec<Argument>) -> Expression {
    let callee = Expression::new(ExprKind::Member {
        object: Box::new(receiver),
        field: method.to_string(),
        null_safe: false,
    });
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args,
        optional: false,
    })
}

fn build_linq_lambda_call(
    receiver: Expression,
    method: &str,
    param: &str,
    body: Expression,
) -> Expression {
    let lambda = Expression::new(ExprKind::Lambda {
        params: vec![Param {
            name: param.to_string(),
            type_hint: None,
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        }],
        body: LambdaBody::Expr(Box::new(body)),
        is_async: false,
        captures: vec![],
    });
    build_linq_value_call(receiver, method, lambda)
}

fn parse_xml_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let xml_text = pair.as_str().to_string();
    if xml_text.contains("<%=") && !xml_text.contains("<![CDATA[") && !xml_text.contains("<!--") {
        if let (Some(root_name), Some(content_src)) = (
            vb_xml_literal_root_name_from_source(&xml_text),
            vb_xml_literal_embedded_expr_from_source(&xml_text),
        ) {
            let content = parse_expression_str(&content_src)?;
            return Ok(Expression::with_span(
                ExprKind::New {
                    class: Box::new(Expression::ident("XElement")),
                    args: vec![
                        Argument::positional(Expression::string(&root_name)),
                        Argument::positional(content),
                    ],
                },
                span,
            ));
        }
    }
    if xml_text.trim_start().starts_with("<?xml") {
        return Ok(Expression::with_span(
            ExprKind::Call {
                callee: Box::new(build_dotted_expr("xml.parse")),
                args: vec![Argument::positional(Expression::string(&xml_text))],
                optional: false,
            },
            span,
        ));
    }
    Ok(Expression::with_span(
        ExprKind::Member {
            object: Box::new(call_expr(
                build_dotted_expr("xml.parse"),
                vec![Argument::positional(Expression::string(&xml_text))],
            )),
            field: "documentElement".to_string(),
            null_safe: false,
        },
        span,
    ))
}

fn vb_xml_literal_root_name_from_source(source: &str) -> Option<String> {
    let trimmed = source.trim_start();
    let rest = trimmed.strip_prefix('<')?;
    if rest.starts_with('?') || rest.starts_with('!') || rest.starts_with('/') {
        return None;
    }
    let name: String = rest
        .chars()
        .take_while(|ch| !ch.is_whitespace() && *ch != '>' && *ch != '/')
        .collect();
    (!name.is_empty()).then_some(name)
}

fn vb_xml_literal_embedded_expr_from_source(source: &str) -> Option<String> {
    let start = source.find("<%=")? + 3;
    let mut cursor = start;
    let mut depth = 1usize;
    while cursor < source.len() {
        let rest = &source[cursor..];
        if rest.starts_with("<%=") {
            depth += 1;
            cursor += 3;
            continue;
        }
        if rest.starts_with("%>") {
            depth -= 1;
            if depth == 0 {
                return Some(source[start..cursor].trim().to_string());
            }
            cursor += 2;
            continue;
        }
        cursor += rest.chars().next()?.len_utf8();
    }
    None
}

fn parse_l_value_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let source = pair.as_str().trim();
    if source.to_ascii_lowercase().contains("(of ") || source.to_ascii_lowercase().contains("(of") {
        if let Ok(expr) = parse_expression_str(source) {
            return Ok(expr);
        }
    }
    let bytes = source.as_bytes();
    let mut cursor = 0usize;

    while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
        cursor += 1;
    }
    let start = cursor;
    while cursor < bytes.len() && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'_') {
        cursor += 1;
    }

    if start == cursor {
        return Err("l_value_expression missing identifier".to_string());
    }

    let mut expr = Expression::ident(&source[start..cursor]);

    while cursor < bytes.len() {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            break;
        }

        match bytes[cursor] {
            b'.' => {
                cursor += 1;
                while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
                    cursor += 1;
                }
                let name_start = cursor;
                while cursor < bytes.len()
                    && (bytes[cursor].is_ascii_alphanumeric() || bytes[cursor] == b'_')
                {
                    cursor += 1;
                }
                if name_start == cursor {
                    return Err("l_value_expression missing member name".to_string());
                }
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field: source[name_start..cursor].to_string(),
                    null_safe: false,
                });
            }
            b'(' => {
                let args_start = cursor + 1;
                let mut depth = 1usize;
                cursor += 1;
                while cursor < bytes.len() && depth > 0 {
                    match bytes[cursor] {
                        b'(' => depth += 1,
                        b')' => depth -= 1,
                        _ => {}
                    }
                    cursor += 1;
                }
                if depth != 0 {
                    return Err("l_value_expression missing closing ')'".to_string());
                }

                let args_text = source[args_start..cursor - 1].trim();
                let mut args = Vec::new();
                if !args_text.is_empty() {
                    let mut parsed = VbParser::parse(Rule::argument_list, args_text)
                        .map_err(|err| err.to_string())?;
                    let arg_list_pair = parsed
                        .next()
                        .ok_or_else(|| "l_value_expression missing argument list".to_string())?;
                    args = parse_argument_list(arg_list_pair)?
                        .into_iter()
                        .map(|arg| arg.value)
                        .collect();
                }

                if let ExprKind::Member { object, field, .. } = &expr.kind {
                    if field.eq_ignore_ascii_case("Item") && !args.is_empty() {
                        let mut indexed = (**object).clone();
                        for idx_expr in args {
                            indexed = Expression::new(ExprKind::Index {
                                object: Box::new(indexed),
                                index: Box::new(idx_expr),
                                null_safe: false,
                            });
                        }
                        expr = indexed;
                        continue;
                    }
                }

                if args.len() == 1 {
                    expr = Expression::new(ExprKind::Index {
                        object: Box::new(expr),
                        index: Box::new(args.into_iter().next().unwrap()),
                        null_safe: false,
                    });
                } else if !args.is_empty() {
                    for idx_expr in args {
                        expr = Expression::new(ExprKind::Index {
                            object: Box::new(expr),
                            index: Box::new(idx_expr),
                            null_safe: false,
                        });
                    }
                } else {
                    expr = Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: vec![],
                        optional: false,
                    });
                }
            }
            _ => break,
        }
    }

    Ok(expr)
}

// ── Helper functions ──

/// Split a `ctrl.Event` (or `obj.Sub.Event`) string into a control expression
/// and a lowercase event name. The last segment is the event; everything
/// before becomes the control expression (member chain).
fn split_event_target(s: &str) -> (Expression, String) {
    let parts: Vec<&str> = s.split('.').collect();
    if parts.len() < 2 {
        return (Expression::ident(s), String::new());
    }
    let event = parts[parts.len() - 1].to_lowercase();
    let control = build_dotted_expr(&parts[..parts.len() - 1].join("."));
    (control, event)
}

fn parse_event_target(pair: Pair<Rule>) -> Result<(Expression, String), String> {
    let raw = pair.as_str().to_string();
    if let Some((receiver, event)) = raw.rsplit_once('.') {
        let receiver = receiver.trim();
        if receiver
            .get(..5)
            .is_some_and(|prefix| prefix.eq_ignore_ascii_case("CType"))
            || receiver
                .get(..8)
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("TryCast"))
            || receiver
                .get(..10)
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("DirectCast"))
        {
            let mut receiver_expr = parse_expression_str(receiver)?;
            if let ExprKind::Cast { expr, .. } = receiver_expr.kind {
                receiver_expr = *expr;
            }
            return Ok((receiver_expr, event.trim().to_ascii_lowercase()));
        }
    }
    let mut inner = pair.into_inner();
    let Some(first) = inner.next() else {
        return Ok(split_event_target(&raw));
    };

    let mut expr = match first.as_rule() {
        Rule::cast_expression => {
            let expr = parse_expression(first)?;
            if let ExprKind::Cast { expr, .. } = expr.kind {
                *expr
            } else {
                expr
            }
        }
        Rule::dotted_identifier => return Ok(split_event_target(first.as_str())),
        _ => return Ok(split_event_target(&raw)),
    };

    let mut last_field = None;
    for item in inner {
        if item.as_rule() == Rule::member_identifier {
            if let Some(field) = last_field.replace(item.as_str().to_string()) {
                expr = Expression::new(ExprKind::Member {
                    object: Box::new(expr),
                    field,
                    null_safe: false,
                });
            }
        }
    }

    Ok((expr, last_field.unwrap_or_default().to_ascii_lowercase()))
}

/// Build an Expression from a dotted name like `me.btn1` or `obj.field.method`.
/// The first segment becomes an Ident; subsequent segments become Member access.
fn build_dotted_expr(s: &str) -> Expression {
    let parts: Vec<&str> = s.split('.').collect();
    let mut iter = parts.into_iter();
    let first = iter.next().unwrap_or("");
    let mut expr = Expression::ident(first);
    for seg in iter {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: seg.to_string(),
            null_safe: false,
        });
    }
    expr
}

fn parse_visibility(s: &str) -> Visibility {
    match s.to_lowercase().as_str() {
        "public" => Visibility::Public,
        "private" => Visibility::Private,
        "protected" => Visibility::Protected,
        "friend" => Visibility::Internal,
        _ => Visibility::Public,
    }
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
