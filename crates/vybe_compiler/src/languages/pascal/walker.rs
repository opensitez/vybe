use super::{PascalParser, Rule};
use crate::ast::*;
use pest::Parser;
use pest::iterators::Pair;

const PASCAL_HELPER_TARGET_PREFIX: &str = "__pascal_helper_target__:";
const PASCAL_VARIANT_FIELD_MARKER: &str = "__pascal_variant_field__";

pub fn parse(source: &str) -> Result<Module, String> {
    let source = source.trim_start_matches('\u{feff}');
    let pairs =
        PascalParser::parse(Rule::program, source).map_err(|e| format!("Parse error: {}", e))?;
    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut name = "main".to_string();

    for pair in pairs {
        if pair.as_rule() != Rule::program {
            continue;
        }
        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::program_heading => {
                    // program Foo; or unit Foo;
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::identifier {
                            name = p.as_str().to_string();
                        }
                    }
                }
                Rule::uses_clause => {
                    let span = to_span(&inner);
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::identifier {
                            imports.push(Import {
                                kind: ImportKind::Simple {
                                    path: p.as_str().to_string(),
                                    alias: None,
                                },
                                span,
                            });
                        }
                    }
                }
                Rule::interface_section | Rule::implementation_section => {
                    // Markers only — no content to walk
                }
                Rule::decl_section => {
                    walk_decl_section(inner, &mut body)?;
                }
                Rule::program_body => {
                    // compound_statement wrapping main body
                    for p in inner.into_inner() {
                        if p.as_rule() == Rule::compound_statement {
                            let stmts = walk_compound_statement(p)?;
                            body.extend(stmts);
                        }
                    }
                }
                Rule::EOI => {}
                _ => {}
            }
        }
    }

    // Pascal allows method bodies to be implemented outside the class declaration
    // (e.g. `constructor TFoo.Create(...) begin ... end;`). Merge those standalone
    // FunctionDecls back into the matching ClassDecl so the compiler sees them as
    // ordinary class members.
    merge_separated_methods(&mut body);
    lower_pascal_helpers(&mut body);

    // Synthesize minimal RTL classes at the top of every Pascal program
    // before constructor-call rewriting so known runtime base types exist
    // during later passes.
    body.insert(0, synthesize_tinterfacedobject_class());
    body.insert(0, synthesize_exception_class());

    // Now that class declarations are stable, rewrite `TFoo.Create(args)` (Pascal's
    // constructor invocation syntax) into the canonical `New { class: TFoo, args }`
    // AST so every language ends up with the same instantiation node.
    let class_names: std::collections::HashSet<String> = body
        .iter()
        .filter_map(|s| match &s.kind {
            StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                Some(name.to_lowercase())
            }
            _ => None,
        })
        .collect();
    for stmt in body.iter_mut() {
        rewrite_constructor_calls_stmt(stmt, &class_names);
    }

    // Default-initialize record / struct variables. `var p: TPoint;`
    // declares an uninitialised value-type local — without an init,
    // the compiler emits `null` and the first `p.X := 10` writes to
    // an undefined receiver. Pascal records are value types: the
    // declaration allocates a fresh instance. Mirror by emitting
    // `new <TypeName>()` for any var/local whose type_hint matches a
    // walked StructDecl.
    let struct_names: std::collections::HashSet<String> = body
        .iter()
        .filter_map(|s| {
            if let StmtKind::StructDecl { name, .. } = &s.kind {
                Some(name.to_lowercase())
            } else {
                None
            }
        })
        .collect();
    let explicit_ctor_record_names = collect_records_without_default_constructor(&body);
    let variant_record_names = collect_variant_record_names_and_clear_markers(&mut body);
    for stmt in body.iter_mut() {
        default_init_struct_locals_stmt(stmt, &struct_names, &explicit_ctor_record_names);
    }
    for stmt in body.iter_mut() {
        default_init_struct_results_stmt(stmt, &struct_names);
    }
    for stmt in body.iter_mut() {
        erase_variant_record_param_type_hints_stmt(stmt, &variant_record_names);
    }
    let struct_fields = collect_struct_fields(&body);
    lower_struct_copy_assignments(&mut body, &struct_fields);
    // Pascal allows user functions to shadow builtin type names. When that
    // happens, `Double(x)` should stay a function call rather than getting
    // frozen into a builtin cast during expression walking.
    let free_function_names: std::collections::HashSet<String> = body
        .iter()
        .filter_map(|stmt| {
            if let StmtKind::FunctionDecl { name, .. } = &stmt.kind {
                Some(name.to_lowercase())
            } else {
                None
            }
        })
        .collect();
    for stmt in body.iter_mut() {
        rewrite_shadowed_builtin_casts_stmt(stmt, &free_function_names);
    }

    Ok(Module {
        name,
        language: Lang::Pascal,
        body,
        imports,
    })
}

/// Walk a single statement and stamp `init: Some(new <Type>())` on
/// any declarator whose `type_hint` names a walked record. Recurses
/// into block / control structures.
fn default_init_struct_locals_stmt(
    stmt: &mut Statement,
    struct_names: &std::collections::HashSet<String>,
    explicit_ctor_record_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations.iter_mut() {
                if decl.init.is_none() {
                    if let Some(ref type_hint) = decl.type_hint {
                        let bare = bare_type_name(type_hint);
                        if struct_names.contains(&bare.to_lowercase())
                            && !explicit_ctor_record_names.contains(&bare.to_lowercase())
                        {
                            decl.init = Some(Expression::new(ExprKind::New {
                                class: Box::new(Expression::ident(bare)),
                                args: Vec::new(),
                            }));
                        } else if explicit_ctor_record_names.contains(&bare.to_lowercase()) {
                            decl.type_hint = None;
                        } else if let Some((count, element_type)) =
                            fixed_record_array(type_hint, struct_names, explicit_ctor_record_names)
                        {
                            decl.init = Some(record_array_initializer(count, &element_type));
                        }
                    }
                }
            }
        }
        StmtKind::Block(inner) => {
            for s in inner {
                default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for s in then_body {
                default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
            }
            for (_, body) in elifs {
                for s in body {
                    default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
                }
            }
            if let Some(eb) = else_body {
                for s in eb {
                    default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
                }
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            for s in body {
                default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            for s in body {
                default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
            }
            for c in catches.iter_mut() {
                for s in c.body.iter_mut() {
                    default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
                }
            }
            if let Some(f) = finally {
                for s in f {
                    default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body {
                default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for m in members {
                if let ClassMember::Method(box_stmt) = m {
                    default_init_struct_locals_stmt(
                        box_stmt,
                        struct_names,
                        explicit_ctor_record_names,
                    );
                } else if let ClassMember::Constructor { body, .. } = m {
                    for s in body {
                        default_init_struct_locals_stmt(s, struct_names, explicit_ctor_record_names);
                    }
                }
            }
        }
        _ => {}
    }
}

fn default_init_struct_results_stmt(
    stmt: &mut Statement,
    struct_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::FunctionDecl {
            return_type, body, ..
        } => {
            if let Some(return_type) = return_type {
                let bare = bare_type_name(return_type);
                if struct_names.contains(&bare.to_lowercase()) {
                    body.insert(0, assign_result_new_record(&bare));
                }
            }
            for s in body {
                default_init_struct_results_stmt(s, struct_names);
            }
        }
        StmtKind::Block(inner) => {
            for s in inner {
                default_init_struct_results_stmt(s, struct_names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for s in then_body {
                default_init_struct_results_stmt(s, struct_names);
            }
            for (_, body) in elifs {
                for s in body {
                    default_init_struct_results_stmt(s, struct_names);
                }
            }
            if let Some(else_body) = else_body {
                for s in else_body {
                    default_init_struct_results_stmt(s, struct_names);
                }
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            for s in body {
                default_init_struct_results_stmt(s, struct_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            for s in body {
                default_init_struct_results_stmt(s, struct_names);
            }
            for c in catches.iter_mut() {
                for s in c.body.iter_mut() {
                    default_init_struct_results_stmt(s, struct_names);
                }
            }
            if let Some(finally) = finally {
                for s in finally {
                    default_init_struct_results_stmt(s, struct_names);
                }
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) => {
                        default_init_struct_results_stmt(method, struct_names);
                    }
                    ClassMember::Constructor { body, .. } => {
                        for s in body {
                            default_init_struct_results_stmt(s, struct_names);
                        }
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn bare_type_name(type_hint: &str) -> &str {
    type_hint.split('<').next().unwrap_or(type_hint).trim()
}

fn collect_records_without_default_constructor(
    body: &[Statement],
) -> std::collections::HashSet<String> {
    body.iter()
        .filter_map(|stmt| {
            let StmtKind::StructDecl { name, members, .. } = &stmt.kind else {
                return None;
            };
            let mut has_constructor = false;
            let mut has_default_constructor = false;
            for member in members {
                if let ClassMember::Constructor { params, .. } = member {
                    has_constructor = true;
                    has_default_constructor |= params.is_empty();
                }
            }
            (has_constructor && !has_default_constructor).then(|| name.to_lowercase())
        })
        .collect()
}

fn assign_result_new_record(type_name: &str) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident("Result")],
        value: Expression::new(ExprKind::New {
            class: Box::new(Expression::ident(type_name)),
            args: Vec::new(),
        }),
    })
}

fn fixed_record_array(
    type_hint: &str,
    struct_names: &std::collections::HashSet<String>,
    explicit_ctor_record_names: &std::collections::HashSet<String>,
) -> Option<(usize, String)> {
    let trimmed = type_hint.trim();
    let lower = trimmed.to_ascii_lowercase();
    let rest = lower.strip_prefix("array[")?;
    let close = rest.find(']')?;
    let bounds = &trimmed["array[".len().."array[".len() + close];
    let after = trimmed["array[".len() + close + 1..].trim_start();
    let element_type = after.strip_prefix("of ")?.trim();
    let bare_element = bare_type_name(element_type);
    if !struct_names.contains(&bare_element.to_lowercase())
        || explicit_ctor_record_names.contains(&bare_element.to_lowercase())
    {
        return None;
    }

    let (lo, hi) = bounds.split_once("..")?;
    let lo = lo.trim().parse::<isize>().ok()?;
    let hi = hi.trim().parse::<isize>().ok()?;
    if lo < 0 || hi < lo {
        return None;
    }

    Some(((hi - lo + 1) as usize, bare_element.to_string()))
}

fn record_array_initializer(count: usize, element_type: &str) -> Expression {
    let elements = (0..count)
        .map(|_| ArrayElement {
            key: None,
            value: Expression::new(ExprKind::New {
                class: Box::new(Expression::ident(element_type)),
                args: Vec::new(),
            }),
            spread: false,
            by_ref: false,
        })
        .collect();
    Expression::new(ExprKind::Array(elements))
}

fn collect_struct_fields(
    body: &[Statement],
) -> std::collections::HashMap<String, Vec<String>> {
    let mut fields = std::collections::HashMap::new();
    for stmt in body {
        if let StmtKind::StructDecl { name, members, .. } = &stmt.kind {
            let names = members
                .iter()
                .filter_map(|member| match member {
                    ClassMember::Field { name, .. } => Some(name.clone()),
                    _ => None,
                })
                .collect::<Vec<_>>();
            fields.insert(name.to_lowercase(), names);
        }
    }
    fields
}

fn lower_struct_copy_assignments(
    body: &mut Vec<Statement>,
    struct_fields: &std::collections::HashMap<String, Vec<String>>,
) {
    let mut env = std::collections::HashMap::new();
    lower_struct_copy_assignments_in_block(body, struct_fields, &mut env);
}

fn lower_struct_copy_assignments_in_block(
    body: &mut Vec<Statement>,
    struct_fields: &std::collections::HashMap<String, Vec<String>>,
    env: &mut std::collections::HashMap<String, String>,
) {
    let mut lowered = Vec::with_capacity(body.len());
    for mut stmt in std::mem::take(body) {
        match &mut stmt.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for decl in declarations.iter() {
                    if let BindingPattern::Ident(name) = &decl.pattern {
                        if let Some(type_hint) = &decl.type_hint {
                            let bare = bare_type_name(type_hint).to_string();
                            if struct_fields.contains_key(&bare.to_lowercase()) {
                                env.insert(name.to_lowercase(), bare);
                            }
                        }
                    }
                }
                lowered.push(stmt);
            }
            StmtKind::Assign { targets, value } if targets.len() == 1 => {
                let copy = match (&targets[0].kind, &value.kind) {
                    (ExprKind::Ident(target), ExprKind::Ident(source)) => {
                        let target_type = env.get(&target.to_lowercase());
                        let source_type = env.get(&source.to_lowercase());
                        target_type
                            .zip(source_type)
                            .filter(|(left, right)| left.eq_ignore_ascii_case(right))
                            .and_then(|(type_name, _)| {
                                struct_fields
                                    .get(&type_name.to_lowercase())
                                    .map(|fields| build_struct_copy_statements(target, source, type_name, fields))
                            })
                    }
                    _ => None,
                };
                if let Some(stmts) = copy {
                    lowered.extend(stmts);
                } else {
                    lowered.push(stmt);
                }
            }
            _ => {
                lower_struct_copy_assignments_stmt(&mut stmt, struct_fields, env);
                lowered.push(stmt);
            }
        }
    }
    *body = lowered;
}

fn lower_struct_copy_assignments_stmt(
    stmt: &mut Statement,
    struct_fields: &std::collections::HashMap<String, Vec<String>>,
    env: &mut std::collections::HashMap<String, String>,
) {
    match &mut stmt.kind {
        StmtKind::Block(inner) => {
            let mut scope = env.clone();
            lower_struct_copy_assignments_in_block(inner, struct_fields, &mut scope);
        }
        StmtKind::FunctionDecl { params, body, .. } => {
            let mut scope = std::collections::HashMap::new();
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    let bare = bare_type_name(type_hint).to_string();
                    if struct_fields.contains_key(&bare.to_lowercase()) {
                        scope.insert(param.name.to_lowercase(), bare);
                    }
                }
            }
            lower_struct_copy_assignments_in_block(body, struct_fields, &mut scope);
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            let mut then_scope = env.clone();
            lower_struct_copy_assignments_in_block(then_body, struct_fields, &mut then_scope);
            for (_, body) in elifs {
                let mut elif_scope = env.clone();
                lower_struct_copy_assignments_in_block(body, struct_fields, &mut elif_scope);
            }
            if let Some(else_body) = else_body {
                let mut else_scope = env.clone();
                lower_struct_copy_assignments_in_block(else_body, struct_fields, &mut else_scope);
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            let mut scope = env.clone();
            lower_struct_copy_assignments_in_block(body, struct_fields, &mut scope);
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            let mut try_scope = env.clone();
            lower_struct_copy_assignments_in_block(body, struct_fields, &mut try_scope);
            for catch in catches {
                let mut catch_scope = env.clone();
                lower_struct_copy_assignments_in_block(&mut catch.body, struct_fields, &mut catch_scope);
            }
            if let Some(finally) = finally {
                let mut finally_scope = env.clone();
                lower_struct_copy_assignments_in_block(finally, struct_fields, &mut finally_scope);
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) => {
                        lower_struct_copy_assignments_stmt(method, struct_fields, env);
                    }
                    ClassMember::Constructor { body, .. } => {
                        let mut scope = env.clone();
                        lower_struct_copy_assignments_in_block(body, struct_fields, &mut scope);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn build_struct_copy_statements(
    target: &str,
    source: &str,
    type_name: &str,
    fields: &[String],
) -> Vec<Statement> {
    let mut stmts = vec![Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(target)],
        value: Expression::new(ExprKind::New {
            class: Box::new(Expression::ident(type_name)),
            args: Vec::new(),
        }),
    })];

    for field in fields {
        stmts.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(target)),
                field: field.clone(),
                null_safe: false,
            })],
            value: Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident(source)),
                field: field.clone(),
                null_safe: false,
            }),
        }));
    }

    stmts
}

fn collect_variant_record_names_and_clear_markers(
    body: &mut [Statement],
) -> std::collections::HashSet<String> {
    let mut names = std::collections::HashSet::new();
    for stmt in body {
        collect_variant_record_names_stmt(stmt, &mut names);
    }
    names
}

fn collect_variant_record_names_stmt(
    stmt: &mut Statement,
    names: &mut std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::StructDecl {
            name, members, ..
        } => {
            let mut has_variant_field = false;
            for member in members.iter_mut() {
                match member {
                    ClassMember::Field { modifiers, .. } => {
                        let before = modifiers.decorators.len();
                        modifiers.decorators.retain(|decorator| {
                            !matches!(
                                decorator.kind,
                                ExprKind::Ident(ref ident) if ident == PASCAL_VARIANT_FIELD_MARKER
                            )
                        });
                        has_variant_field |= modifiers.decorators.len() != before;
                    }
                    ClassMember::Method(method) => collect_variant_record_names_stmt(method, names),
                    ClassMember::Constructor { body, .. } => {
                        for s in body {
                            collect_variant_record_names_stmt(s, names);
                        }
                    }
                    _ => {}
                }
            }
            if has_variant_field {
                names.insert(name.to_lowercase());
            }
        }
        StmtKind::ClassDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) => collect_variant_record_names_stmt(method, names),
                    ClassMember::Constructor { body, .. } => {
                        for s in body {
                            collect_variant_record_names_stmt(s, names);
                        }
                    }
                    ClassMember::NestedType(nested) => {
                        collect_variant_record_names_stmt(nested, names);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}

fn erase_variant_record_param_type_hints_stmt(
    stmt: &mut Statement,
    variant_record_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::FunctionDecl { params, body, .. } => {
            for param in params {
                if let Some(type_hint) = &param.type_hint {
                    if variant_record_names.contains(&bare_type_name(type_hint).to_lowercase()) {
                        param.pass_by = PassBy::Value;
                        param.type_hint = None;
                    }
                }
            }
            for s in body {
                erase_variant_record_param_type_hints_stmt(s, variant_record_names);
            }
        }
        StmtKind::Block(inner) => {
            for s in inner {
                erase_variant_record_param_type_hints_stmt(s, variant_record_names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for s in then_body {
                erase_variant_record_param_type_hints_stmt(s, variant_record_names);
            }
            for (_, body) in elifs {
                for s in body {
                    erase_variant_record_param_type_hints_stmt(s, variant_record_names);
                }
            }
            if let Some(else_body) = else_body {
                for s in else_body {
                    erase_variant_record_param_type_hints_stmt(s, variant_record_names);
                }
            }
        }
        StmtKind::For { body, .. }
        | StmtKind::ForIn { body, .. }
        | StmtKind::While { body, .. }
        | StmtKind::DoWhile { body, .. } => {
            for s in body {
                erase_variant_record_param_type_hints_stmt(s, variant_record_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            finally,
            ..
        } => {
            for s in body {
                erase_variant_record_param_type_hints_stmt(s, variant_record_names);
            }
            for c in catches.iter_mut() {
                for s in c.body.iter_mut() {
                    erase_variant_record_param_type_hints_stmt(s, variant_record_names);
                }
            }
            if let Some(finally) = finally {
                for s in finally {
                    erase_variant_record_param_type_hints_stmt(s, variant_record_names);
                }
            }
        }
        StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
            for member in members {
                match member {
                    ClassMember::Method(method) => {
                        erase_variant_record_param_type_hints_stmt(method, variant_record_names);
                    }
                    ClassMember::Constructor { body, .. } => {
                        for s in body {
                            erase_variant_record_param_type_hints_stmt(s, variant_record_names);
                        }
                    }
                    ClassMember::NestedType(nested) => {
                        erase_variant_record_param_type_hints_stmt(nested, variant_record_names);
                    }
                    _ => {}
                }
            }
        }
        _ => {}
    }
}


fn rewrite_shadowed_builtin_casts_stmt(
    stmt: &mut Statement,
    free_function_names: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(expr) => rewrite_shadowed_builtin_casts_expr(expr, free_function_names),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                rewrite_shadowed_builtin_casts_expr(expr, free_function_names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                rewrite_shadowed_builtin_casts_expr(expr, free_function_names);
            }
            if let Some(cause) = cause {
                rewrite_shadowed_builtin_casts_expr(cause, free_function_names);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &mut decl.init {
                    rewrite_shadowed_builtin_casts_expr(init, free_function_names);
                }
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                rewrite_shadowed_builtin_casts_expr(target, free_function_names);
            }
            rewrite_shadowed_builtin_casts_expr(value, free_function_names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_shadowed_builtin_casts_expr(target, free_function_names);
            rewrite_shadowed_builtin_casts_expr(value, free_function_names);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_shadowed_builtin_casts_expr(cond, free_function_names);
            for stmt in then_body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
            for (cond, body) in elifs {
                rewrite_shadowed_builtin_casts_expr(cond, free_function_names);
                for stmt in body {
                    rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
                }
            }
        }
        StmtKind::While { cond, body, .. } | StmtKind::DoWhile { cond, body, .. } => {
            rewrite_shadowed_builtin_casts_expr(cond, free_function_names);
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
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
                rewrite_shadowed_builtin_casts_stmt(init, free_function_names);
            }
            if let Some(cond) = cond {
                rewrite_shadowed_builtin_casts_expr(cond, free_function_names);
            }
            if let Some(update) = update {
                rewrite_shadowed_builtin_casts_expr(update, free_function_names);
            }
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
        }
        StmtKind::ForIn { iter, body, .. } => {
            rewrite_shadowed_builtin_casts_expr(iter, free_function_names);
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
        }
        StmtKind::FunctionDecl { body, .. }
        | StmtKind::With { body, .. }
        | StmtKind::Using { body, .. }
        | StmtKind::Lock { body, .. } => {
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
            ..
        } => {
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
            for catch in catches {
                for stmt in &mut catch.body {
                    rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
                }
            }
            if let Some(else_body) = else_body {
                for stmt in else_body {
                    rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
                }
            }
            if let Some(finally_body) = finally {
                for stmt in finally_body {
                    rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
                }
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                rewrite_shadowed_builtin_casts_stmt(stmt, free_function_names);
            }
        }
        _ => {}
    }
}

fn rewrite_shadowed_builtin_casts_expr(
    expr: &mut Expression,
    free_function_names: &std::collections::HashSet<String>,
) {
    match &mut expr.kind {
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => {
            rewrite_shadowed_builtin_casts_expr(inner, free_function_names);
            if free_function_names.contains(&type_name.to_lowercase()) {
                let arg = (**inner).clone();
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::ident(type_name)),
                    args: vec![Argument::positional(arg)],
                    optional: false,
                };
            } else if matches!(
                type_name.to_lowercase().as_str(),
                "integer" | "int" | "longint"
            ) {
                let arg = (**inner).clone();
                expr.kind = ExprKind::Call {
                    callee: Box::new(Expression::ident("Trunc")),
                    args: vec![Argument::positional(arg)],
                    optional: false,
                };
            }
        }
        ExprKind::Binary { left, right, .. } => {
            rewrite_shadowed_builtin_casts_expr(left, free_function_names);
            rewrite_shadowed_builtin_casts_expr(right, free_function_names);
        }
        ExprKind::Unary { expr: inner, .. }
        | ExprKind::Await(inner)
        | ExprKind::Yield(Some(inner))
        | ExprKind::Spread(inner) => {
            rewrite_shadowed_builtin_casts_expr(inner, free_function_names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_shadowed_builtin_casts_expr(cond, free_function_names);
            rewrite_shadowed_builtin_casts_expr(then, free_function_names);
            rewrite_shadowed_builtin_casts_expr(else_, free_function_names);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_shadowed_builtin_casts_expr(callee, free_function_names);
            for arg in args {
                rewrite_shadowed_builtin_casts_expr(&mut arg.value, free_function_names);
            }
        }
        ExprKind::Member { object, .. } => {
            rewrite_shadowed_builtin_casts_expr(object, free_function_names);
        }
        ExprKind::Index { object, index, .. } => {
            rewrite_shadowed_builtin_casts_expr(object, free_function_names);
            rewrite_shadowed_builtin_casts_expr(index, free_function_names);
        }
        ExprKind::New { class, args } => {
            rewrite_shadowed_builtin_casts_expr(class, free_function_names);
            for arg in args {
                rewrite_shadowed_builtin_casts_expr(&mut arg.value, free_function_names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &mut element.key {
                    rewrite_shadowed_builtin_casts_expr(key, free_function_names);
                }
                rewrite_shadowed_builtin_casts_expr(&mut element.value, free_function_names);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for item in items {
                rewrite_shadowed_builtin_casts_expr(item, free_function_names);
            }
        }
        ExprKind::Object(props) => {
            for prop in props {
                match prop {
                    ObjectProperty::KeyValue { key, value }
                    | ObjectProperty::Computed { key, value } => {
                        rewrite_shadowed_builtin_casts_expr(key, free_function_names);
                        rewrite_shadowed_builtin_casts_expr(value, free_function_names);
                    }
                    ObjectProperty::Spread(value) => {
                        rewrite_shadowed_builtin_casts_expr(value, free_function_names);
                    }
                    ObjectProperty::Shorthand(_)
                    | ObjectProperty::Method { .. }
                    | ObjectProperty::Accessor { .. } => {}
                }
            }
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_shadowed_builtin_casts_expr(start, free_function_names);
            rewrite_shadowed_builtin_casts_expr(end, free_function_names);
        }
        ExprKind::TypeOf(inner)
        | ExprKind::YieldFrom(inner)
        | ExprKind::Void(inner)
        | ExprKind::Delete(inner) => {
            rewrite_shadowed_builtin_casts_expr(inner, free_function_names);
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_shadowed_builtin_casts_expr(left, free_function_names);
            rewrite_shadowed_builtin_casts_expr(right, free_function_names);
        }
        ExprKind::Interpolation(parts) => {
            for part in parts {
                match part {
                    InterpolPart::Expr(expr) | InterpolPart::Formatted(expr, _) => {
                        rewrite_shadowed_builtin_casts_expr(expr, free_function_names);
                    }
                    InterpolPart::Text(_) => {}
                }
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_shadowed_builtin_casts_expr(target, free_function_names);
            rewrite_shadowed_builtin_casts_expr(value, free_function_names);
        }
        ExprKind::SuperCall { args, .. } => {
            for arg in args {
                rewrite_shadowed_builtin_casts_expr(&mut arg.value, free_function_names);
            }
        }
        ExprKind::Yield(None)
        | ExprKind::Lit(_)
        | ExprKind::Ident(_)
        | ExprKind::This
        | ExprKind::Super
        | ExprKind::Lambda { .. }
        | ExprKind::ClassExpr { .. }
        | ExprKind::FunctionExpr(_)
        | ExprKind::IsType { .. } => {}
        _ => {}
    }
}

fn synthesize_exception_class() -> Statement {
    // class Exception { Message: String; constructor Create(msg: String); }
    // The Create body assigns `Self.Message := msg` so `e.Message`
    // returns the constructor argument inside catch handlers.
    let span = Span::default();
    let msg_param = Param {
        name: "msg".into(),
        type_hint: Some("String".into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    };
    // Self.Message := msg
    let assign_msg = Statement::with_span(
        StmtKind::Assign {
            targets: vec![Expression::with_span(
                ExprKind::Member {
                    object: Box::new(Expression::with_span(ExprKind::This, span.clone())),
                    field: "Message".into(),
                    null_safe: false,
                },
                span.clone(),
            )],
            value: Expression::with_span(ExprKind::Ident("msg".into()), span.clone()),
        },
        span.clone(),
    );
    Statement::with_span(
        StmtKind::ClassDecl {
            name: "Exception".into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members: vec![
                ClassMember::Field {
                    name: "Message".into(),
                    type_hint: Some("String".into()),
                    init: None,
                    modifiers: Modifiers::default(),
                    with_events: false,
                    array_bounds: None,
                },
                ClassMember::Constructor {
                    params: vec![msg_param],
                    body: vec![assign_msg],
                    base_args: None,
                    initializer_target: crate::ast::ConstructorInitializerTarget::Base,
                    visibility: Visibility::Public,
                },
            ],
            modifiers: ClassModifiers::default(),
            decorators: vec![],
        },
        span,
    )
}

fn synthesize_tinterfacedobject_class() -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: "TInterfacedObject".into(),
            parents: Vec::new(),
            interfaces: Vec::new(),
            members: Vec::new(),
            modifiers: ClassModifiers::default(),
            decorators: vec![],
        },
        Span::default(),
    )
}

/// Walk a statement and rewrite `ClassName.Create(args)` into `New { class, args }`
/// when `ClassName` matches a class declared in the same module.
fn rewrite_constructor_calls_stmt(
    stmt: &mut Statement,
    classes: &std::collections::HashSet<String>,
) {
    match &mut stmt.kind {
        StmtKind::Expr(e) => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Block(stmts) => {
            for s in stmts {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for d in declarations {
                if let Some(e) = &mut d.init {
                    rewrite_constructor_calls_expr(e, classes);
                }
            }
        }
        StmtKind::FunctionDecl { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        StmtKind::ClassDecl { members, .. } => {
            for m in members {
                rewrite_constructor_calls_member(m, classes);
            }
        }
        StmtKind::StructDecl { members, .. } | StmtKind::ModuleDecl { members, .. } => {
            for m in members {
                rewrite_constructor_calls_member(m, classes);
            }
        }
        StmtKind::NamespaceDecl { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            rewrite_constructor_calls_expr(cond, classes);
            for s in then_body {
                rewrite_constructor_calls_stmt(s, classes);
            }
            for (c, b) in elifs {
                rewrite_constructor_calls_expr(c, classes);
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(i) = init {
                rewrite_constructor_calls_stmt(i, classes);
            }
            if let Some(c) = cond {
                rewrite_constructor_calls_expr(c, classes);
            }
            if let Some(u) = update {
                rewrite_constructor_calls_expr(u, classes);
            }
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            rewrite_constructor_calls_expr(iter, classes);
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            rewrite_constructor_calls_expr(cond, classes);
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
            rewrite_constructor_calls_expr(cond, classes);
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            rewrite_constructor_calls_expr(expr, classes);
            for c in cases {
                for s in &mut c.body {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
            if let Some(b) = default {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
            for c in catches {
                for s in &mut c.body {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
            if let Some(b) = else_body {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
            if let Some(b) = finally {
                for s in b {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        StmtKind::With { items, body, .. } => {
            for it in items {
                rewrite_constructor_calls_expr(&mut it.expr, classes);
            }
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        StmtKind::Return(Some(e)) => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Throw { expr: Some(e), .. } => rewrite_constructor_calls_expr(e, classes),
        StmtKind::Assign { targets, value } => {
            for t in targets {
                rewrite_constructor_calls_expr(t, classes);
            }
            rewrite_constructor_calls_expr(value, classes);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            rewrite_constructor_calls_expr(target, classes);
            rewrite_constructor_calls_expr(value, classes);
        }
        _ => {}
    }
}

fn rewrite_constructor_calls_member(
    m: &mut ClassMember,
    classes: &std::collections::HashSet<String>,
) {
    match m {
        ClassMember::Field { init: Some(e), .. } => rewrite_constructor_calls_expr(e, classes),
        ClassMember::Method(stmt) => rewrite_constructor_calls_stmt(stmt, classes),
        ClassMember::Constructor { body, .. } => {
            for s in body {
                rewrite_constructor_calls_stmt(s, classes);
            }
        }
        ClassMember::Property { getter, setter, .. } => {
            if let Some(g) = getter {
                for s in g {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
            if let Some(set) = setter {
                for s in &mut set.body {
                    rewrite_constructor_calls_stmt(s, classes);
                }
            }
        }
        ClassMember::Const { value, .. } => rewrite_constructor_calls_expr(value, classes),
        ClassMember::NestedType(stmt) => rewrite_constructor_calls_stmt(stmt, classes),
        _ => {}
    }
}

fn rewrite_constructor_calls_expr(
    expr: &mut Expression,
    classes: &std::collections::HashSet<String>,
) {
    // Check Call(Member(ClassName, "Create"), args) BEFORE descending so the
    // Member-only rewrite below doesn't fire on the callee position first and
    // turn `TFoo.Create(42)` into a call on a New expression.
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if let ExprKind::Member { object, field, .. } = &callee.kind {
            if let ExprKind::Ident(class_name) = &object.kind {
                if classes.contains(&class_name.to_lowercase())
                    && field.eq_ignore_ascii_case("Create")
                {
                    let new_class = Box::new(Expression::ident(class_name));
                    let mut new_args = args.clone();
                    for a in new_args.iter_mut() {
                        rewrite_constructor_calls_expr(&mut a.value, classes);
                    }
                    expr.kind = ExprKind::New {
                        class: new_class,
                        args: new_args,
                    };
                    return;
                }
            }
        }
    }

    // First descend into children, then check this node so deeply-nested
    // patterns are also normalized.
    match &mut expr.kind {
        ExprKind::Binary { left, right, .. } => {
            rewrite_constructor_calls_expr(left, classes);
            rewrite_constructor_calls_expr(right, classes);
        }
        ExprKind::Unary { expr: e, .. } => rewrite_constructor_calls_expr(e, classes),
        ExprKind::Ternary { cond, then, else_ } => {
            rewrite_constructor_calls_expr(cond, classes);
            rewrite_constructor_calls_expr(then, classes);
            rewrite_constructor_calls_expr(else_, classes);
        }
        ExprKind::Member { object, .. } => rewrite_constructor_calls_expr(object, classes),
        ExprKind::Index { object, index, .. } => {
            rewrite_constructor_calls_expr(object, classes);
            rewrite_constructor_calls_expr(index, classes);
        }
        ExprKind::Call { callee, args, .. } => {
            rewrite_constructor_calls_expr(callee, classes);
            for a in args.iter_mut() {
                rewrite_constructor_calls_expr(&mut a.value, classes);
            }
        }
        ExprKind::New { class, args } => {
            rewrite_constructor_calls_expr(class, classes);
            for a in args.iter_mut() {
                rewrite_constructor_calls_expr(&mut a.value, classes);
            }
        }
        ExprKind::Assign { target, value } => {
            rewrite_constructor_calls_expr(target, classes);
            rewrite_constructor_calls_expr(value, classes);
        }
        ExprKind::Array(elems) => {
            for el in elems {
                rewrite_constructor_calls_expr(&mut el.value, classes);
            }
        }
        ExprKind::Tuple(items) | ExprKind::Set(items) | ExprKind::Sequence(items) => {
            for e in items {
                rewrite_constructor_calls_expr(e, classes);
            }
        }
        ExprKind::Object(props) => {
            for p in props {
                if let ObjectProperty::KeyValue { value, .. } = p {
                    rewrite_constructor_calls_expr(value, classes);
                }
            }
        }
        ExprKind::Interpolation(parts) => {
            for p in parts {
                match p {
                    InterpolPart::Expr(e) | InterpolPart::Formatted(e, _) => {
                        rewrite_constructor_calls_expr(e, classes)
                    }
                    _ => {}
                }
            }
        }
        ExprKind::NullCoalesce { left, right } => {
            rewrite_constructor_calls_expr(left, classes);
            rewrite_constructor_calls_expr(right, classes);
        }
        ExprKind::Range { start, end, .. } => {
            rewrite_constructor_calls_expr(start, classes);
            rewrite_constructor_calls_expr(end, classes);
        }
        ExprKind::IsType { expr: e, .. } | ExprKind::Cast { expr: e, .. } => {
            rewrite_constructor_calls_expr(e, classes)
        }
        _ => {}
    }

    // Pascal allows zero-arg constructor calls without parens: `f := TFoo.Create;`
    // Detect bare `ClassName.Create` member access on a known class and rewrite
    // it to a zero-arg `New { class, [] }`.
    if let ExprKind::Member { object, field, .. } = &expr.kind {
        if let ExprKind::Ident(class_name) = &object.kind {
            if classes.contains(&class_name.to_lowercase()) && field.eq_ignore_ascii_case("Create")
            {
                expr.kind = ExprKind::New {
                    class: Box::new(Expression::ident(class_name)),
                    args: Vec::new(),
                };
            }
        }
    }
}

/// Pascal post-processing: attach `ClassName.Method` standalone FunctionDecls
/// to their matching ClassDecl members. The constructor (`ClassName.Create`)
/// fills in `ClassMember::Constructor`. Other methods fill the body of the
/// matching `ClassMember::Method`.
fn merge_separated_methods(body: &mut Vec<Statement>) {
    use std::collections::HashMap;

    // Collect class indices by canonicalized (lowercase) name
    let mut class_idx: HashMap<String, usize> = HashMap::new();
    for (i, s) in body.iter().enumerate() {
        match &s.kind {
            StmtKind::ClassDecl { name, .. } | StmtKind::StructDecl { name, .. } => {
                class_idx.insert(name.to_lowercase(), i);
            }
            _ => {}
        }
    }

    // Walk in reverse so removals don't shift earlier indices
    let mut to_remove: Vec<usize> = Vec::new();
    for i in 0..body.len() {
        let (class_name, method_name, params, ret, mods, body_stmts, is_sub) = {
            let stmt = &body[i];
            let StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body: b,
                modifiers,
                is_sub,
                ..
            } = &stmt.kind
            else {
                continue;
            };
            let Some((cls, mth)) = name.split_once('.') else {
                continue;
            };
            (
                cls.to_string(),
                mth.to_string(),
                params.clone(),
                return_type.clone(),
                modifiers.clone(),
                b.clone(),
                *is_sub,
            )
        };

        let Some(&ci) = class_idx.get(&class_name.to_lowercase()) else {
            continue;
        };
        let members = match &mut body[ci].kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => members,
            _ => continue,
        };

        // Try constructor first: any ClassMember::Constructor whose params arity matches,
        // when the method name is "Create" (Pascal convention) — fall back to first ctor.
        let is_create = method_name.eq_ignore_ascii_case("Create");
        let mut attached = false;
        if is_create {
            for m in members.iter_mut() {
                if let ClassMember::Constructor {
                    params: cp,
                    body: cb,
                    base_args: ba,
                    ..
                } = m
                {
                    if cb.is_empty() {
                        *cp = params.clone();
                        let mut new_body = body_stmts.clone();
                        // Pascal pattern: `inherited Create(args)` as the FIRST statement
                        // is the base-constructor invocation. Lift it into `base_args`
                        // so the compiler runs the canonical C#-style path
                        // (parent ctor → field inits → method bindings → body).
                        // This keeps the AST uniform across languages.
                        if let Some(first) = new_body.first() {
                            let extracted = match &first.kind {
                                StmtKind::Expr(e) => match &e.kind {
                                    ExprKind::SuperCall { method, args } => {
                                        let is_ctor = method.is_none()
                                            || method.as_ref().map_or(false, |m| {
                                                m.eq_ignore_ascii_case("Create")
                                            });
                                        if is_ctor {
                                            Some(
                                                args.iter()
                                                    .map(|a| a.value.clone())
                                                    .collect::<Vec<_>>(),
                                            )
                                        } else {
                                            None
                                        }
                                    }
                                    _ => None,
                                },
                                _ => None,
                            };
                            if let Some(extracted_args) = extracted {
                                *ba = Some(extracted_args);
                                new_body.remove(0);
                            }
                        }
                        *cb = new_body;
                        attached = true;
                        break;
                    }
                }
            }
        }

        if !attached {
            // Find a Method with matching name and empty body
            for m in members.iter_mut() {
                if let ClassMember::Method(stmt) = m {
                    if let StmtKind::FunctionDecl {
                        name: mn,
                        params: mp,
                        body: mb,
                        return_type: mr,
                        modifiers: mm,
                        is_sub: ms,
                        ..
                    } = &mut stmt.kind
                    {
                        if mn.eq_ignore_ascii_case(&method_name) && mb.is_empty() {
                            *mp = params.clone();
                            *mb = body_stmts.clone();
                            *mr = ret.clone();
                            *mm = mods.clone();
                            *ms = is_sub;
                            let is_generator = body_has_yield(mb);
                            if let StmtKind::FunctionDecl {
                                is_generator: flag, ..
                            } = &mut stmt.kind
                            {
                                *flag = is_generator;
                            }
                            attached = true;
                            break;
                        }
                    }
                }
            }
        }

        if attached {
            to_remove.push(i);
        }
    }

    for i in to_remove.into_iter().rev() {
        body.remove(i);
    }
}

fn lower_pascal_helpers(body: &mut Vec<Statement>) {
    use std::collections::HashMap;

    let mut target_idx: HashMap<String, usize> = HashMap::new();
    for (idx, stmt) in body.iter().enumerate() {
        match &stmt.kind {
            StmtKind::ClassDecl { name, parents, .. } => {
                if !parents
                    .iter()
                    .any(|p| p.starts_with(PASCAL_HELPER_TARGET_PREFIX))
                {
                    target_idx.insert(name.to_lowercase(), idx);
                }
            }
            StmtKind::StructDecl {
                name, interfaces, ..
            } => {
                if !interfaces
                    .iter()
                    .any(|i| i.starts_with(PASCAL_HELPER_TARGET_PREFIX))
                {
                    target_idx.insert(name.to_lowercase(), idx);
                }
            }
            _ => {}
        }
    }

    let mut merges: Vec<(usize, Vec<ClassMember>)> = Vec::new();
    let mut removals = Vec::new();
    let mut lifted_functions = Vec::new();

    for (idx, stmt) in body.iter().enumerate() {
        let helper_target = match &stmt.kind {
            StmtKind::ClassDecl { parents, .. } => parents
                .iter()
                .find_map(|p| p.strip_prefix(PASCAL_HELPER_TARGET_PREFIX))
                .map(str::to_string),
            StmtKind::StructDecl { interfaces, .. } => interfaces
                .iter()
                .find_map(|p| p.strip_prefix(PASCAL_HELPER_TARGET_PREFIX))
                .map(str::to_string),
            _ => None,
        };
        let Some(target_name) = helper_target else {
            continue;
        };
        removals.push(idx);

        let Some(&target_stmt_idx) = target_idx.get(&target_name.to_lowercase()) else {
            let helper_members = match &stmt.kind {
                StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                    members
                }
                _ => continue,
            };
            lifted_functions.extend(lower_pascal_builtin_helper_members(
                &target_name,
                helper_members,
            ));
            continue;
        };
        let helper_members = match &stmt.kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                members.clone()
            }
            _ => Vec::new(),
        };
        merges.push((target_stmt_idx, helper_members));
    }

    for (idx, mut helper_members) in merges {
        match &mut body[idx].kind {
            StmtKind::ClassDecl { members, .. } | StmtKind::StructDecl { members, .. } => {
                members.append(&mut helper_members);
            }
            _ => {}
        }
    }

    for idx in removals.into_iter().rev() {
        body.remove(idx);
    }

    let insert_at = body
        .iter()
        .position(|stmt| {
            !matches!(
                stmt.kind,
                StmtKind::VarDecl { .. }
                    | StmtKind::FunctionDecl { .. }
                    | StmtKind::ClassDecl { .. }
                    | StmtKind::StructDecl { .. }
                    | StmtKind::InterfaceDecl { .. }
                    | StmtKind::EnumDecl { .. }
                    | StmtKind::ModuleDecl { .. }
                    | StmtKind::NamespaceDecl { .. }
                    | StmtKind::DelegateDecl { .. }
            )
        })
        .unwrap_or(body.len());
    body.splice(insert_at..insert_at, lifted_functions);
}

fn lower_pascal_builtin_helper_members(
    target_name: &str,
    members: &[ClassMember],
) -> Vec<Statement> {
    let mut lifted = Vec::new();
    for member in members {
        let ClassMember::Method(stmt) = member else {
            continue;
        };
        let StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            is_sub,
            ..
        } = &stmt.kind
        else {
            continue;
        };
        let helper_method = name.rsplit('.').next().unwrap_or(name);
        let mut lifted_params = Vec::with_capacity(params.len() + 1);
        lifted_params.push(Param {
            name: "Self".to_string(),
            type_hint: Some(target_name.to_string()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        lifted_params.extend(params.clone());
        lifted.push(Statement::with_span(
            StmtKind::FunctionDecl {
                name: pascal_helper_function_name(target_name, helper_method),
                params: lifted_params,
                return_type: return_type.clone(),
                body: body.clone(),
                modifiers: modifiers.clone(),
                handles: Vec::new(),
                is_async: false,
                is_generator: body_has_yield(body),
                is_sub: *is_sub,
            },
            stmt.span.clone(),
        ));
    }
    lifted
}

fn pascal_helper_function_name(target_name: &str, method_name: &str) -> String {
    let sanitize = |text: &str| {
        text.chars()
            .map(|ch| {
                if ch.is_ascii_alphanumeric() {
                    ch.to_ascii_lowercase()
                } else {
                    '_'
                }
            })
            .collect::<String>()
    };
    format!(
        "__pascal_helper_{}_{}",
        sanitize(target_name),
        sanitize(method_name)
    )
}

// ════════════════════════════════════════════════════════════════════════════
// Declaration section
// ════════════════════════════════════════════════════════════════════════════

fn walk_decl_section(pair: Pair<Rule>, body: &mut Vec<Statement>) -> Result<(), String> {
    for decl in pair.into_inner() {
        match decl.as_rule() {
            Rule::var_section => {
                body.extend(walk_var_section(decl)?);
            }
            Rule::const_section => {
                body.extend(walk_const_section(decl)?);
            }
            Rule::type_section => {
                body.extend(walk_type_section(decl)?);
            }
            Rule::procedure_decl_or_method => {
                body.push(walk_procedure_decl_or_method(decl)?);
            }
            Rule::function_decl_or_method => {
                body.push(walk_function_decl_or_method(decl)?);
            }
            Rule::constructor_method_impl => {
                body.push(walk_constructor_method_impl(decl)?);
            }
            Rule::destructor_method_impl => {
                body.push(walk_destructor_method_impl(decl)?);
            }
            _ => {}
        }
    }
    Ok(())
}

// ── Var section ────────────────────────────────────────────────────────────

fn walk_var_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::var_decl {
            let decls = walk_var_decl(p)?;
            stmts.push(Statement::with_span(
                StmtKind::VarDecl {
                    declarations: decls,
                    kind: VarDeclKind::Dim,
                },
                span,
            ));
        }
    }
    Ok(stmts)
}

fn walk_var_decl(pair: Pair<Rule>) -> Result<Vec<VarDeclarator>, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut array_bounds: Option<Vec<Expression>> = None;
    let mut init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => {
                array_bounds = extract_array_bounds(&p)?;
                type_hint = Some(type_ref_to_string(&p));
            }
            Rule::inline_var_init => {
                for inner in p.into_inner() {
                    if inner.as_rule() == Rule::expression {
                        init = Some(walk_expression(inner)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(build_var_declarators(names, type_hint, init, array_bounds))
}

fn build_var_declarators(
    names: Vec<String>,
    type_hint: Option<String>,
    init: Option<Expression>,
    array_bounds: Option<Vec<Expression>>,
) -> Vec<VarDeclarator> {
    names
        .into_iter()
        .map(|n| VarDeclarator {
            pattern: BindingPattern::Ident(n),
            type_hint: type_hint.clone(),
            init: init.clone(),
            array_bounds: array_bounds.clone(),
            with_events: false,
        })
        .collect()
}

// ── Const section ──────────────────────────────────────────────────────────

fn walk_const_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let span = to_span(&pair);
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::const_decl {
            let decl = walk_const_decl(p)?;
            stmts.push(Statement::with_span(
                StmtKind::VarDecl {
                    declarations: vec![decl],
                    kind: VarDeclKind::Const,
                },
                span,
            ));
        }
    }
    Ok(stmts)
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<VarDeclarator, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = p.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::expression => init = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(VarDeclarator {
        pattern: BindingPattern::Ident(name),
        type_hint,
        init,
        array_bounds: None,
        with_events: false,
    })
}

// ── Type section ───────────────────────────────────────────────────────────

fn walk_type_section(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::type_decl {
            stmts.push(walk_type_decl(p)?);
        }
    }
    Ok(stmts)
}

fn walk_type_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let mut name = String::new();
    let mut type_def_pair: Option<Pair<Rule>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::type_def => type_def_pair = Some(p),
            _ => {}
        }
    }

    let def = type_def_pair.ok_or("Missing type_def in type_decl")?;
    let inner = def.into_inner().next().ok_or("Empty type_def")?;

    match inner.as_rule() {
        Rule::class_type => walk_class_type(inner, &name, span),
        Rule::class_helper_type => walk_class_helper_type(inner, &name, span),
        Rule::record_type => walk_record_type(inner, &name, span),
        Rule::record_helper_type => walk_record_helper_type(inner, &name, span),
        Rule::interface_type => walk_interface_type(inner, &name, span),
        Rule::enum_type => walk_enum_type(inner, &name, span),
        Rule::array_type => {
            // Type alias for array: type TMyArray = array[0..9] of Integer;
            // Emit as a VarDecl with type hint
            Ok(Statement::with_span(
                StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: Some(type_ref_to_string(&inner)),
                        init: None,
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                },
                span,
            ))
        }
        Rule::pointer_type => {
            // Type alias for pointer
            let target = inner
                .into_inner()
                .next()
                .map(|p| type_ref_to_string(&p))
                .unwrap_or_default();
            Ok(Statement::with_span(
                StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: Some(format!("^{}", target)),
                        init: None,
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                },
                span,
            ))
        }
        Rule::subrange_type => Ok(Statement::with_span(
            StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name),
                    type_hint: Some(type_ref_to_string(&inner)),
                    init: None,
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Const,
            },
            span,
        )),
        Rule::type_alias => {
            // Simple type alias: type TMyInt = Integer;
            let aliased = inner
                .into_inner()
                .next()
                .map(|p| type_ref_to_string(&p))
                .unwrap_or_default();
            Ok(Statement::with_span(
                StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident(name),
                        type_hint: Some(aliased),
                        init: None,
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Const,
                },
                span,
            ))
        }
        other => Err(format!("Unexpected type_def inner: {:?}", other)),
    }
}

// ── Class type ─────────────────────────────────────────────────────────────

fn walk_class_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::class_heritage => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        parents.push(id.as_str().to_string());
                    }
                }
            }
            Rule::class_body => members = walk_class_body_members(p)?,
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents,
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: vec![],
        },
        span,
    ))
}

fn walk_class_helper_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut target = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_ref => target = type_ref_to_string(&p),
            Rule::class_body => members = walk_class_body_members(p)?,
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents: vec![format!("{}{}", PASCAL_HELPER_TARGET_PREFIX, target)],
            interfaces: Vec::new(),
            members,
            modifiers: ClassModifiers::default(),
            decorators: vec![],
        },
        span,
    ))
}

fn walk_class_body_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for m in pair.into_inner() {
        match m.as_rule() {
            Rule::field_decl => members.extend(walk_field_decl_members(m)?),
            Rule::class_constructor => members.push(walk_class_constructor_sig(m)?),
            Rule::class_destructor => members.push(walk_class_method_sig(m, true)?),
            Rule::class_procedure | Rule::class_function => {
                members.push(walk_class_method_sig(m, false)?);
            }
            Rule::class_class_member => members.push(walk_class_class_member(m)?),
            Rule::class_property_decl => members.push(walk_class_property_decl(m)?),
            _ => {}
        }
    }
    Ok(members)
}

// ── Record type ────────────────────────────────────────────────────────────

fn walk_record_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut members = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::record_body {
            members = walk_record_body_members(p)?;
        }
    }
    Ok(Statement::with_span(
        StmtKind::StructDecl {
            name: name.to_string(),
            interfaces: Vec::new(),
            members,
            visibility: Visibility::Public,
            decorators: vec![],
        },
        span,
    ))
}

fn walk_record_helper_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut target = String::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::type_ref => target = type_ref_to_string(&p),
            Rule::record_body => members = walk_record_body_members(p)?,
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::StructDecl {
            name: name.to_string(),
            interfaces: vec![format!("{}{}", PASCAL_HELPER_TARGET_PREFIX, target)],
            members,
            visibility: Visibility::Public,
            decorators: vec![],
        },
        span,
    ))
}

fn walk_record_body_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();
    for m in pair.into_inner() {
        match m.as_rule() {
            Rule::field_decl => members.extend(walk_field_decl_members(m)?),
            Rule::variant_part => members.extend(walk_variant_part_members(m)?),
            Rule::record_method_sig => members.push(walk_record_method_sig(m)?),
            Rule::record_class_method => members.push(walk_record_class_method(m)?),
            Rule::record_operator_method => members.push(walk_record_operator_method(m)?),
            _ => {}
        }
    }
    Ok(members)
}

// ── Interface type ─────────────────────────────────────────────────────────

fn walk_interface_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut parents = Vec::new();
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::interface_heritage => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        parents.push(id.as_str().to_string());
                    }
                }
            }
            Rule::interface_body => {
                for m in p.into_inner() {
                    match m.as_rule() {
                        Rule::interface_procedure => {
                            members.push(walk_interface_method(m, true)?);
                        }
                        Rule::interface_function => {
                            members.push(walk_interface_method(m, false)?);
                        }
                        Rule::interface_property_decl => {
                            members.push(walk_interface_property(m)?);
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::with_span(
        StmtKind::InterfaceDecl {
            name: name.to_string(),
            parents,
            members,
            decorators: vec![],
        },
        span,
    ))
}

fn walk_interface_method(pair: Pair<Rule>, is_sub: bool) -> Result<InterfaceMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(InterfaceMember::Method {
        name,
        params,
        return_type,
        is_sub,
        signature_source: None,
    })
}

fn walk_interface_property(pair: Pair<Rule>) -> Result<InterfaceMember, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut has_read = false;
    let mut has_write = false;

    for p in pair.into_inner() {
        if p.as_rule() == Rule::property_def {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                    Rule::property_specifiers => {
                        for spec in sp.into_inner() {
                            match spec.as_rule() {
                                Rule::property_read => has_read = true,
                                Rule::property_write => has_write = true,
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(InterfaceMember::Property {
        name,
        type_hint,
        is_readonly: has_read && !has_write,
        is_writeonly: !has_read && has_write,
    })
}

// ── Enum type ──────────────────────────────────────────────────────────────

fn walk_enum_type(pair: Pair<Rule>, name: &str, span: Span) -> Result<Statement, String> {
    let mut members = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::enum_value {
            let mut ename = String::new();
            let mut value: Option<Expression> = None;
            for ep in p.into_inner() {
                match ep.as_rule() {
                    Rule::identifier => ename = ep.as_str().to_string(),
                    Rule::expression => value = Some(walk_expression(ep)?),
                    _ => {}
                }
            }
            members.push(EnumMember {
                name: ename,
                value,
                constructor_args: Vec::new(),
            });
        }
    }

    Ok(Statement::with_span(
        StmtKind::EnumDecl {
            name: name.to_string(),
            members,
            visibility: Visibility::Public,
            is_flags: false,
            backing_type: None,
            interfaces: Vec::new(),
            body_members: Vec::new(),
            decorators: vec![],
        },
        span,
    ))
}

// ── Field declarations (in class/record bodies) ───────────────────────────

fn walk_field_decl_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            _ => {}
        }
    }

    Ok(names
        .into_iter()
        .map(|n| ClassMember::Field {
            name: n,
            type_hint: type_hint.clone(),
            init: None,
            modifiers: Modifiers::default(),
            with_events: false,
            array_bounds: None,
        })
        .collect())
}

fn variant_field_member(name: String, type_hint: Option<String>) -> ClassMember {
    let mut modifiers = Modifiers::default();
    modifiers
        .decorators
        .push(Expression::ident(PASCAL_VARIANT_FIELD_MARKER));
    ClassMember::Field {
        name,
        type_hint,
        init: None,
        modifiers,
        with_events: false,
        array_bounds: None,
    }
}

fn walk_variant_part_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::variant_selector => {
                if let Some((name, type_hint)) = walk_variant_selector(p) {
                    members.push(variant_field_member(name, Some(type_hint)));
                }
            }
            Rule::variant_arm => {
                members.extend(walk_variant_arm_members(p)?);
            }
            _ => {}
        }
    }

    Ok(members)
}

fn walk_variant_selector(pair: Pair<Rule>) -> Option<(String, String)> {
    let mut name: Option<String> = None;
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => name = Some(p.as_str().to_string()),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            _ => {}
        }
    }

    name.zip(type_hint)
}

fn walk_variant_arm_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut members = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::variant_field_list {
            for field in p.into_inner() {
                if field.as_rule() == Rule::variant_field_decl {
                    members.extend(walk_variant_field_decl_members(field)?);
                }
            }
        }
    }

    Ok(members)
}

fn walk_variant_field_decl_members(pair: Pair<Rule>) -> Result<Vec<ClassMember>, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            _ => {}
        }
    }

    Ok(names
        .into_iter()
        .map(|name| variant_field_member(name, type_hint.clone()))
        .collect())
}

// ── Class member signatures ────────────────────────────────────────────────

fn walk_class_constructor_sig(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut params = Vec::new();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                if sp.as_rule() == Rule::param_clause {
                    params = walk_param_clause(sp)?;
                }
            }
        }
    }

    Ok(ClassMember::Constructor {
        params,
        body: Vec::new(), // Body comes from method_impl
        base_args: None,
        initializer_target: crate::ast::ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    })
}

fn walk_class_method_sig(pair: Pair<Rule>, is_destructor: bool) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    Rule::method_directives => {
                        walk_method_directives(sp, &mut modifiers);
                    }
                    _ => {}
                }
            }
        }
    }

    if is_destructor {
        name = "Destroy".to_string();
    }

    let is_sub = return_type.is_none();
    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body: Vec::new(), // Body comes from method_impl
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub,
        },
    ))))
}

fn walk_class_class_member(pair: Pair<Rule>) -> Result<ClassMember, String> {
    // class procedure / class function / class var
    let _inner_text = pair.as_str().to_lowercase();
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut type_hint: Option<String> = None;
    let mut is_field = false;
    let mut modifiers = Modifiers {
        is_static: true,
        ..Default::default()
    };

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => name = sp.as_str().to_string(),
                        Rule::param_clause => params = walk_param_clause(sp)?,
                        Rule::return_type_clause => {
                            for rt in sp.into_inner() {
                                if rt.as_rule() == Rule::type_ref {
                                    return_type = Some(type_ref_to_string(&rt));
                                }
                            }
                        }
                        Rule::method_directives => {
                            walk_method_directives(sp, &mut modifiers);
                        }
                        _ => {}
                    }
                }
            }
            Rule::field_decl => {
                is_field = true;
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier_list => {
                            for id in sp.into_inner() {
                                if id.as_rule() == Rule::identifier {
                                    name = id.as_str().to_string();
                                }
                            }
                        }
                        Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_field {
        Ok(ClassMember::Field {
            name,
            type_hint,
            init: None,
            modifiers,
            with_events: false,
            array_bounds: None,
        })
    } else {
        let is_sub = return_type.is_none();
        Ok(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body: Vec::new(),
                modifiers,
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub,
            },
        ))))
    }
}

fn walk_class_property_decl(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut type_hint: Option<String> = None;
    let mut getter: Option<Vec<Statement>> = None;
    let mut setter: Option<PropertySetter> = None;
    let modifiers = Modifiers::default();

    for p in pair.into_inner() {
        if p.as_rule() == Rule::property_def {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => {
                        if name.is_empty() {
                            name = sp.as_str().to_string();
                        }
                    }
                    Rule::type_ref => type_hint = Some(type_ref_to_string(&sp)),
                    Rule::property_specifiers => {
                        for spec in sp.into_inner() {
                            match spec.as_rule() {
                                Rule::property_read => {
                                    // read GetFoo → getter delegates to method
                                    let getter_name = spec
                                        .into_inner()
                                        .find(|p| p.as_rule() == Rule::identifier)
                                        .map(|p| p.as_str().to_string())
                                        .unwrap_or_default();
                                    getter = Some(vec![Statement::new(StmtKind::Return(Some(
                                        Expression::new(ExprKind::Call {
                                            callee: Box::new(Expression::new(ExprKind::Member {
                                                object: Box::new(Expression::new(ExprKind::This)),
                                                field: getter_name,
                                                null_safe: false,
                                            })),
                                            args: Vec::new(),
                                            optional: false,
                                        }),
                                    )))]);
                                }
                                Rule::property_write => {
                                    let setter_name = spec
                                        .into_inner()
                                        .find(|p| p.as_rule() == Rule::identifier)
                                        .map(|p| p.as_str().to_string())
                                        .unwrap_or_default();
                                    let param = Param {
                                        name: "value".to_string(),
                                        type_hint: type_hint.clone(),
                                        default: None,
                                        pass_by: PassBy::Value,
                                        is_rest: false,
                                        is_kwargs: false,
                                        is_optional: false,
                                        is_nullable: false,
                                    };
                                    setter = Some(PropertySetter {
                                        param,
                                        body: vec![Statement::new(StmtKind::Expr(
                                            Expression::new(ExprKind::Call {
                                                callee: Box::new(Expression::new(
                                                    ExprKind::Member {
                                                        object: Box::new(Expression::new(
                                                            ExprKind::This,
                                                        )),
                                                        field: setter_name,
                                                        null_safe: false,
                                                    },
                                                )),
                                                args: vec![Argument::positional(
                                                    Expression::ident("value"),
                                                )],
                                                optional: false,
                                            }),
                                        ))],
                                    });
                                }
                                _ => {} // default, stored, nodefault
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    let is_auto = getter.is_none() && setter.is_none();
    Ok(ClassMember::Property {
        name,
        type_hint,
        getter,
        setter,
        is_auto,
        modifiers,
    })
}

fn walk_record_method_sig(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers::default();
    let mut method_kind = "";

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_kind_keyword => method_kind = p.as_str(),
            Rule::method_sig_body => {
                for sp in p.into_inner() {
                    match sp.as_rule() {
                        Rule::identifier => name = sp.as_str().to_string(),
                        Rule::param_clause => params = walk_param_clause(sp)?,
                        Rule::return_type_clause => {
                            for rt in sp.into_inner() {
                                if rt.as_rule() == Rule::type_ref {
                                    return_type = Some(type_ref_to_string(&rt));
                                }
                            }
                        }
                        Rule::method_directives => {
                            walk_method_directives(sp, &mut modifiers);
                        }
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    let kind_lower = method_kind.to_lowercase();
    if kind_lower == "constructor" {
        Ok(ClassMember::Constructor {
            params,
            body: Vec::new(),
            base_args: None,
            initializer_target: crate::ast::ConstructorInitializerTarget::Base,
            visibility: Visibility::Public,
        })
    } else {
        let is_sub =
            return_type.is_none() || kind_lower == "destructor" || kind_lower == "procedure";
        Ok(ClassMember::Method(Box::new(Statement::new(
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body: Vec::new(),
                modifiers,
                handles: Vec::new(),
                is_async: false,
                is_generator: false,
                is_sub,
            },
        ))))
    }
}

fn walk_record_class_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut modifiers = Modifiers {
        is_static: true,
        ..Default::default()
    };

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = sp.as_str().to_string(),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    Rule::method_directives => {
                        walk_method_directives(sp, &mut modifiers);
                    }
                    _ => {}
                }
            }
        }
    }

    let is_sub = return_type.is_none();
    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body: Vec::new(),
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub,
        },
    ))))
}

fn walk_record_operator_method(pair: Pair<Rule>) -> Result<ClassMember, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let modifiers = Modifiers {
        is_static: true,
        ..Default::default()
    };

    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_sig_body {
            for sp in p.into_inner() {
                match sp.as_rule() {
                    Rule::identifier => name = format!("operator_{}", sp.as_str()),
                    Rule::param_clause => params = walk_param_clause(sp)?,
                    Rule::return_type_clause => {
                        for rt in sp.into_inner() {
                            if rt.as_rule() == Rule::type_ref {
                                return_type = Some(type_ref_to_string(&rt));
                            }
                        }
                    }
                    _ => {}
                }
            }
        }
    }

    Ok(ClassMember::Method(Box::new(Statement::new(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body: Vec::new(),
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator: false,
            is_sub: false,
        },
    ))))
}

fn walk_method_directives(pair: Pair<Rule>, modifiers: &mut Modifiers) {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_directive {
            let kw = p.as_str().to_lowercase();
            match kw.as_str() {
                "virtual" => modifiers.is_virtual = true,
                "override" => modifiers.is_override = true,
                "abstract" => modifiers.is_abstract = true,
                "overload" => modifiers.is_overloads = true,
                _ => {} // reintroduce, inline, cdecl, stdcall, register, dynamic
            }
        }
    }
}

// ════════════════════════════════════════════════════════════════════════════
// Procedure / Function declarations and method implementations
// ════════════════════════════════════════════════════════════════════════════

fn walk_procedure_decl_or_method(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_impl_proc => return walk_method_impl_proc(p, span),
            Rule::standalone_procedure => return walk_standalone_procedure(p, span),
            _ => {}
        }
    }
    Err("procedure_decl_or_method: no inner match".into())
}

fn walk_function_decl_or_method(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::method_impl_func => return walk_method_impl_func(p, span),
            Rule::standalone_function => return walk_standalone_function(p, span),
            _ => {}
        }
    }
    Err("function_decl_or_method: no inner match".into())
}

fn walk_method_impl_proc(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 {
                    class_name = p.as_str().to_string();
                } else {
                    method_name = p.as_str().to_string();
                }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name: format!("{}.{}", class_name, method_name),
            params,
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: true,
        },
        span,
    ))
}

fn walk_method_impl_func(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 {
                    class_name = p.as_str().to_string();
                } else {
                    method_name = p.as_str().to_string();
                }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => return_type = Some(type_ref_to_string(&p)),
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name: format!("{}.{}", class_name, method_name),
            params,
            return_type,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: false,
        },
        span,
    ))
}

fn walk_constructor_method_impl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    // constructor ClassName.Create(...)
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_impl_body_proc {
            return walk_method_impl_body_proc(p, span, true);
        }
    }
    Err("constructor_method_impl: missing body".into())
}

fn walk_destructor_method_impl(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    for p in pair.into_inner() {
        if p.as_rule() == Rule::method_impl_body_proc {
            return walk_method_impl_body_proc(p, span, false);
        }
    }
    Err("destructor_method_impl: missing body".into())
}

fn walk_method_impl_body_proc(
    pair: Pair<Rule>,
    span: Span,
    _is_constructor: bool,
) -> Result<Statement, String> {
    let mut class_name = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut id_count = 0;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => {
                if id_count == 0 {
                    class_name = p.as_str().to_string();
                } else {
                    method_name = p.as_str().to_string();
                }
                id_count += 1;
            }
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::decl_section => walk_decl_section(p, &mut body)?,
            Rule::compound_statement => body.extend(walk_compound_statement(p)?),
            _ => {}
        }
    }

    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name: format!("{}.{}", class_name, method_name),
            params,
            return_type: None,
            body,
            modifiers: Modifiers::default(),
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: true,
        },
        span,
    ))
}

fn walk_standalone_procedure(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body = Vec::new();
    let mut is_forward = false;
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::method_directives => walk_method_directives(p, &mut modifiers),
            Rule::forward_directive => is_forward = true,
            Rule::procedure_body => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::decl_section => walk_decl_section(bp, &mut body)?,
                        Rule::compound_statement => body.extend(walk_compound_statement(bp)?),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_forward {
        // Forward declarations emit an empty function
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type: None,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: true,
        },
        span,
    ))
}

fn walk_standalone_function(pair: Pair<Rule>, span: Span) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut body = Vec::new();
    let mut is_forward = false;
    let mut modifiers = Modifiers::default();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => return_type = Some(type_ref_to_string(&p)),
            Rule::method_directives => walk_method_directives(p, &mut modifiers),
            Rule::forward_directive => is_forward = true,
            Rule::function_body => {
                for bp in p.into_inner() {
                    match bp.as_rule() {
                        Rule::decl_section => walk_decl_section(bp, &mut body)?,
                        Rule::compound_statement => body.extend(walk_compound_statement(bp)?),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
    }

    if is_forward {
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    let is_generator = body_has_yield(&body);

    Ok(Statement::with_span(
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles: Vec::new(),
            is_async: false,
            is_generator,
            is_sub: false,
        },
        span,
    ))
}

// ════════════════════════════════════════════════════════════════════════════
// Parameters
// ════════════════════════════════════════════════════════════════════════════

fn walk_param_clause(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param_list {
            return walk_param_list(p);
        }
    }
    Ok(Vec::new())
}

fn walk_param_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::param {
            params.extend(walk_param(p)?);
        }
    }
    Ok(params)
}

fn walk_param(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut pass_by = PassBy::Value;
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut default: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_mode => {
                let mode = p.as_str().to_lowercase();
                pass_by = match mode.as_str() {
                    "var" => PassBy::Ref,
                    "const" => PassBy::Const,
                    "out" => PassBy::Out,
                    _ => PassBy::Value,
                };
            }
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::param_default => {
                for dp in p.into_inner() {
                    if dp.as_rule() == Rule::expression {
                        default = Some(walk_expression(dp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(names
        .into_iter()
        .map(|n| Param {
            name: n,
            type_hint: type_hint.clone(),
            default: default.clone(),
            pass_by,
            is_rest: false,
            is_kwargs: false,
            is_optional: default.is_some(),
            is_nullable: false,
        })
        .collect())
}

// ════════════════════════════════════════════════════════════════════════════
// Statements
// ════════════════════════════════════════════════════════════════════════════

fn walk_compound_statement(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for p in pair.into_inner() {
        if p.as_rule() == Rule::statement_list {
            return walk_statement_list(p);
        }
    }
    Ok(Vec::new())
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for p in pair.into_inner() {
        if p.as_rule() == Rule::statement {
            let stmt = walk_statement(p)?;
            if !matches!(stmt.kind, StmtKind::Empty) {
                stmts.push(stmt);
            }
        }
    }
    Ok(stmts)
}

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let inner = pair.into_inner().next();
    let inner = match inner {
        Some(p) => p,
        None => return Ok(Statement::with_span(StmtKind::Empty, span)),
    };

    let kind = match inner.as_rule() {
        Rule::compound_statement => StmtKind::Block(walk_compound_statement(inner)?),
        Rule::inline_var_statement => walk_inline_var_statement(inner)?,
        Rule::if_statement => walk_if_statement(inner)?,
        Rule::for_statement => walk_for_statement(inner)?,
        Rule::for_in_statement => walk_for_in_statement(inner)?,
        Rule::while_statement => walk_while_statement(inner)?,
        Rule::repeat_statement => walk_repeat_statement(inner)?,
        Rule::case_statement => walk_case_statement(inner)?,
        Rule::with_statement => walk_with_statement(inner)?,
        Rule::try_statement => walk_try_statement(inner)?,
        Rule::raise_statement => walk_raise_statement(inner)?,
        Rule::yield_statement => {
            let value = inner
                .into_inner()
                .next()
                .map(walk_expression)
                .transpose()?
                .map(Box::new);
            StmtKind::Expr(Expression::new(ExprKind::Yield(value)))
        }
        Rule::exit_statement => walk_exit_statement(inner)?,
        Rule::halt_statement => walk_halt_statement(inner)?,
        Rule::break_statement => StmtKind::Break(BreakTarget::Implicit),
        Rule::continue_statement => StmtKind::Continue(ContinueTarget::Implicit),
        Rule::inherited_statement => walk_inherited_statement(inner)?,
        Rule::assign_or_call_statement => walk_assign_or_call(inner)?,
        Rule::empty_statement => StmtKind::Empty,
        other => return Err(format!("Unexpected statement rule: {:?}", other)),
    };

    Ok(Statement::with_span(kind, span))
}

fn walk_inline_var_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut array_bounds: Option<Vec<Expression>> = None;
    let mut init: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier_list => {
                for id in p.into_inner() {
                    if id.as_rule() == Rule::identifier {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::type_ref => {
                array_bounds = extract_array_bounds(&p)?;
                type_hint = Some(type_ref_to_string(&p));
            }
            Rule::inline_var_init => {
                for inner in p.into_inner() {
                    if inner.as_rule() == Rule::expression {
                        init = Some(walk_expression(inner)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::VarDecl {
        declarations: build_var_declarators(names, type_hint, init, array_bounds),
        kind: VarDeclKind::Dim,
    })
}

// ── If ─────────────────────────────────────────────────────────────────────

fn walk_if_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    // First is always expression (condition)
    let cond = walk_expression(parts.remove(0))?;
    // Second is the then statement
    let then_stmt = walk_statement(parts.remove(0))?;
    let then_body = flatten_stmt(then_stmt);

    let else_body = if !parts.is_empty() {
        // else_clause
        let else_clause = parts.remove(0);
        let else_stmt = else_clause
            .into_inner()
            .next()
            .map(|p| walk_statement(p))
            .transpose()?;
        else_stmt.map(|s| flatten_stmt(s))
    } else {
        None
    };

    Ok(StmtKind::If {
        cond,
        then_body,
        elifs: Vec::new(),
        else_body,
    })
}

// ── For ────────────────────────────────────────────────────────────────────

fn walk_for_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let src = pair.as_str().to_lowercase();
    let is_downto = src.contains(" downto ");

    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let binding = parts.remove(0);
    let (var_name, type_hint, start_expr) = walk_for_binding(binding)?;
    let end_expr = walk_expression(parts.remove(0))?; // end expression
    let body_stmt = walk_statement(parts.remove(0))?; // body statement
    let use_char_ordinal_loop = type_hint
        .as_deref()
        .is_some_and(|hint| hint.eq_ignore_ascii_case("char"))
        || pascal_expr_is_char_like(&start_expr)
        || pascal_expr_is_char_like(&end_expr);

    let init = if let Some(type_hint) = type_hint {
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(var_name.clone()),
                type_hint: Some(type_hint),
                init: Some(start_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Dim,
        })
    } else {
        Statement::new(StmtKind::Assign {
            targets: vec![Expression::ident(&var_name)],
            value: start_expr,
        })
    };

    let cond = Expression::new(ExprKind::Binary {
        op: if is_downto { BinOp::GtEq } else { BinOp::LtEq },
        left: Box::new(if use_char_ordinal_loop {
            pascal_ord_call(Expression::ident(&var_name))
        } else {
            Expression::ident(&var_name)
        }),
        right: Box::new(if use_char_ordinal_loop {
            pascal_ord_call(end_expr.clone())
        } else {
            end_expr.clone()
        }),
    });

    let update_assign = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&var_name)),
        value: Box::new(if use_char_ordinal_loop {
            pascal_chr_call(Expression::new(ExprKind::Binary {
                op: if is_downto { BinOp::Sub } else { BinOp::Add },
                left: Box::new(pascal_ord_call(Expression::ident(&var_name))),
                right: Box::new(Expression::int(1)),
            }))
        } else {
            Expression::new(ExprKind::Binary {
                op: if is_downto { BinOp::Sub } else { BinOp::Add },
                left: Box::new(Expression::ident(&var_name)),
                right: Box::new(Expression::int(1)),
            })
        }),
    });

    fn pascal_expr_is_char_like(expr: &Expression) -> bool {
        match &expr.kind {
            ExprKind::Lit(Literal::Char(_)) => true,
            ExprKind::Lit(Literal::Str(value)) => value.chars().count() == 1,
            _ => false,
        }
    }

    fn pascal_ord_call(expr: Expression) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Ord")),
            args: vec![Argument::positional(expr)],
            optional: false,
        })
    }

    fn pascal_chr_call(expr: Expression) -> Expression {
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("Chr")),
            args: vec![Argument::positional(expr)],
            optional: false,
        })
    }
    Ok(StmtKind::For {
        init: Some(Box::new(init)),
        cond: Some(cond),
        update: Some(update_assign),
        body: flatten_stmt(body_stmt),
    })
}

fn walk_for_binding(pair: Pair<Rule>) -> Result<(String, Option<String>, Expression), String> {
    let mut name = String::new();
    let mut type_hint = None;
    let mut start_expr = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier if name.is_empty() => name = p.as_str().to_string(),
            Rule::type_ref => type_hint = Some(type_ref_to_string(&p)),
            Rule::expression => start_expr = Some(walk_expression(p)?),
            _ => {}
        }
    }

    let start_expr = start_expr.ok_or("for binding missing start expression")?;
    Ok((name, type_hint, start_expr))
}

// ── For-in ─────────────────────────────────────────────────────────────────

fn walk_for_in_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let var_name = parts.remove(0).as_str().to_string();
    let iter_expr = walk_expression(parts.remove(0))?;
    let body_stmt = walk_statement(parts.remove(0))?;

    Ok(StmtKind::ForIn {
        var: var_name,
        key: None,
        iter: iter_expr,
        body: flatten_stmt(body_stmt),
        of: true, // Pascal for-in iterates values, like JS for...of
        else_body: None,
        is_async: false,
    })
}

// ── While ──────────────────────────────────────────────────────────────────

fn walk_while_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    let cond = walk_expression(parts.remove(0))?;
    let body_stmt = walk_statement(parts.remove(0))?;

    Ok(StmtKind::While {
        cond,
        body: flatten_stmt(body_stmt),
        else_body: None,
    })
}

// ── Repeat/Until ───────────────────────────────────────────────────────────

fn walk_repeat_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut cond: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_list => body = walk_statement_list(p)?,
            Rule::expression => cond = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(StmtKind::DoWhile {
        body,
        cond: cond.unwrap_or_else(|| Expression::bool(true)),
        until: true,
    })
}

// ── Case ───────────────────────────────────────────────────────────────────

fn walk_case_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr: Option<Expression> = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => {
                if expr.is_none() {
                    expr = Some(walk_expression(p)?);
                }
            }
            Rule::case_arm => {
                cases.push(walk_case_arm(p)?);
            }
            Rule::case_else => {
                for cp in p.into_inner() {
                    if cp.as_rule() == Rule::statement_list {
                        default = Some(walk_statement_list(cp)?);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(StmtKind::Switch {
        expr: expr.unwrap_or_else(|| Expression::null()),
        cases,
        default,
    })
}

fn walk_case_arm(pair: Pair<Rule>) -> Result<SwitchCase, String> {
    let mut conditions = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::case_value_list => {
                for cv in p.into_inner() {
                    if cv.as_rule() == Rule::case_value {
                        conditions.push(walk_case_value(cv)?);
                    }
                }
            }
            Rule::case_arm_body => {
                let inner = cv_first(p)?;
                match inner.as_rule() {
                    Rule::compound_statement => body = walk_compound_statement(inner)?,
                    Rule::statement => body = flatten_stmt(walk_statement(inner)?),
                    _ => {}
                }
            }
            _ => {}
        }
    }

    Ok(SwitchCase { conditions, body })
}

fn walk_case_value(pair: Pair<Rule>) -> Result<CaseCondition, String> {
    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();
    if parts.len() == 2 {
        // Range: expr..expr
        let from = walk_expression(parts[0].clone())?;
        let to = walk_expression(parts[1].clone())?;
        Ok(CaseCondition::Range { from, to })
    } else if parts.len() == 1 {
        let val = walk_expression(parts[0].clone())?;
        Ok(CaseCondition::Value(val))
    } else {
        Err("Empty case_value".into())
    }
}

// ── With ───────────────────────────────────────────────────────────────────

fn walk_with_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut items = Vec::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::expression => {
                items.push(WithItem {
                    expr: walk_expression(p)?,
                    var: None,
                });
            }
            Rule::statement => {
                body = flatten_stmt(walk_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(StmtKind::With {
        items,
        body,
        is_async: false,
    })
}

// ── Try ────────────────────────────────────────────────────────────────────

fn walk_try_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut body = Vec::new();
    let mut catches = Vec::new();
    let mut finally: Option<Vec<Statement>> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::statement_list => {
                body = walk_statement_list(p)?;
            }
            Rule::try_handler => {
                for hp in p.into_inner() {
                    match hp.as_rule() {
                        Rule::except_handler => {
                            for ep in hp.into_inner() {
                                if ep.as_rule() == Rule::except_body {
                                    catches = walk_except_body(ep)?;
                                }
                            }
                        }
                        Rule::finally_handler => {
                            for fp in hp.into_inner() {
                                if fp.as_rule() == Rule::statement_list {
                                    finally = Some(walk_statement_list(fp)?);
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

    Ok(StmtKind::Try {
        body,
        catches,
        else_body: None,
        finally,
    })
}

fn walk_except_body(pair: Pair<Rule>) -> Result<Vec<CatchClause>, String> {
    let mut clauses = Vec::new();
    let mut has_on_clauses = false;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::on_clause => {
                has_on_clauses = true;
                clauses.push(walk_on_clause(p)?);
            }
            Rule::except_else => {
                // else clause in except → catch-all
                for ep in p.into_inner() {
                    if ep.as_rule() == Rule::statement_list {
                        clauses.push(CatchClause {
                            types: Vec::new(),
                            var_name: None,
                            stack_var: None,
                            body: walk_statement_list(ep)?,
                            when_clause: None,
                        });
                    }
                }
            }
            Rule::statement_list => {
                if !has_on_clauses {
                    // Bare except with just a statement list → catch-all
                    clauses.push(CatchClause {
                        types: Vec::new(),
                        var_name: None,
                        stack_var: None,
                        body: walk_statement_list(p)?,
                        when_clause: None,
                    });
                }
            }
            _ => {}
        }
    }

    Ok(clauses)
}

fn walk_on_clause(pair: Pair<Rule>) -> Result<CatchClause, String> {
    let mut var_name: Option<String> = None;
    let mut type_name = String::new();
    let mut body = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::on_var_binding => {
                for vp in p.into_inner() {
                    if vp.as_rule() == Rule::identifier {
                        var_name = Some(vp.as_str().to_string());
                    }
                }
            }
            Rule::identifier => type_name = p.as_str().to_string(),
            Rule::statement => body = flatten_stmt(walk_statement(p)?),
            _ => {}
        }
    }

    Ok(CatchClause {
        types: vec![type_name],
        var_name,
        stack_var: None,
        body,
        when_clause: None,
    })
}

// ── Raise ──────────────────────────────────────────────────────────────────

fn walk_raise_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    Ok(StmtKind::Throw { expr, cause: None })
}

// ── Exit ───────────────────────────────────────────────────────────────────

fn walk_exit_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    Ok(StmtKind::Return(expr))
}

// ── Halt ───────────────────────────────────────────────────────────────────

fn walk_halt_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|p| p.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?;

    // Halt maps to a special call; emit as Expr(Call(Halt, [code]))
    let args = expr.into_iter().map(Argument::positional).collect();
    Ok(StmtKind::Expr(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Halt")),
        args,
        optional: false,
    })))
}

// ── Inherited statement ────────────────────────────────────────────────────

fn walk_inherited_statement(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut method: Option<String> = None;
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => method = Some(p.as_str().to_string()),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    Ok(StmtKind::Expr(Expression::new(ExprKind::SuperCall {
        method,
        args,
    })))
}

// ── Assign or call ─────────────────────────────────────────────────────────

fn walk_assign_or_call(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let src = pair.as_str();
    let parts: Vec<Pair<Rule>> = pair.into_inner().collect();

    if parts.len() == 1 {
        // Pure expression used as statement (procedure call, etc.)
        let expr = walk_expression(parts.into_iter().next().unwrap())?;

        // Pascal `FreeAndNil(x)` is sugar for `x := nil` — we have GC, so the
        // free is a no-op but the variable still needs to be cleared so that
        // `Assigned(x)` returns false afterwards. Rewrite at the walker.
        if let ExprKind::Call { callee, args, .. } = &expr.kind {
            if let ExprKind::Ident(name) = &callee.kind {
                if name.eq_ignore_ascii_case("FreeAndNil") && args.len() == 1 {
                    return Ok(StmtKind::Assign {
                        targets: vec![args[0].value.clone()],
                        value: Expression::null(),
                    });
                }
            }
        }

        // Pascal allows zero-arg procedure calls without parens: `Hello;` means
        // `Hello();`. At statement level, a bare identifier or member access
        // that isn't already a Call is implicitly a zero-arg invocation.
        let expr = match expr.kind {
            ExprKind::Call { .. }
            | ExprKind::New { .. }
            | ExprKind::Assign { .. }
            | ExprKind::Lit(_) => expr,
            ExprKind::Ident(_) | ExprKind::Member { .. } => Expression::with_span(
                ExprKind::Call {
                    callee: Box::new(expr.clone()),
                    args: Vec::new(),
                    optional: false,
                },
                expr.span,
            ),
            _ => expr,
        };
        return Ok(StmtKind::Expr(expr));
    }

    if parts.len() >= 2 {
        let target_pair = parts[0].clone();
        let target = walk_expression(target_pair)?;

        // Check for compound assignment operators
        // The grammar captures expression ~ (":=" | "+=" | "-=" | "*=" | "/=") ~ expression
        // After the first expression, remaining pairs are: potentially just the value expression
        // But the operator is part of the rule text, not a separate pair.
        // We need to detect the operator from the source text.

        let value_pair = parts.last().unwrap().clone();
        let value = walk_expression(value_pair)?;

        if src.contains(":=") {
            return Ok(StmtKind::Assign {
                targets: vec![target],
                value,
            });
        } else if src.contains("+=") {
            return Ok(StmtKind::CompoundAssign {
                target,
                op: CompoundOp::Add,
                value,
            });
        } else if src.contains("-=") {
            return Ok(StmtKind::CompoundAssign {
                target,
                op: CompoundOp::Sub,
                value,
            });
        } else if src.contains("*=") {
            return Ok(StmtKind::CompoundAssign {
                target,
                op: CompoundOp::Mul,
                value,
            });
        } else if src.contains("/=") {
            return Ok(StmtKind::CompoundAssign {
                target,
                op: CompoundOp::Div,
                value,
            });
        }

        // Fallback: assignment
        return Ok(StmtKind::Assign {
            targets: vec![target],
            value,
        });
    }

    Ok(StmtKind::Empty)
}

// ════════════════════════════════════════════════════════════════════════════
// Expressions
// ════════════════════════════════════════════════════════════════════════════

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let span = to_span(&pair);
    let kind = walk_expr_kind(pair)?;
    Ok(Expression::with_span(kind, span))
}

fn walk_expr_kind(pair: Pair<Rule>) -> Result<ExprKind, String> {
    match pair.as_rule() {
        Rule::expression => {
            // expression = { is_as_expression }
            let inner = pair.into_inner().next().ok_or("Empty expression")?;
            walk_expr_kind(inner)
        }

        Rule::is_as_expression => {
            // is_as_expression = { relational ~ (is_as_op ~ identifier)* }
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.len() == 1 {
                return walk_expr_kind(inner.remove(0));
            }

            let mut left = walk_expression(inner.remove(0))?;
            let mut i = 0;
            while i + 1 < inner.len() {
                let op_pair = &inner[i];
                let op_str = op_pair.as_str().to_lowercase();
                let type_name = inner[i + 1].as_str().to_string();

                if op_str == "is" {
                    left = Expression::new(ExprKind::IsType {
                        expr: Box::new(left),
                        type_name,
                    });
                } else {
                    // as
                    left = Expression::new(ExprKind::Cast {
                        expr: Box::new(left),
                        type_name,
                    });
                }
                i += 2;
            }
            Ok(left.kind)
        }

        Rule::relational => walk_binary_chain(pair, |op_str| match op_str {
            "<>" => BinOp::NotEq,
            "<=" => BinOp::LtEq,
            ">=" => BinOp::GtEq,
            "<" => BinOp::Lt,
            ">" => BinOp::Gt,
            "=" => BinOp::Eq,
            s if s.starts_with("in") => BinOp::In,
            _ => BinOp::Eq,
        }),

        Rule::additive => walk_binary_chain(pair, |op_str| match op_str {
            "+" => BinOp::Add,
            "-" => BinOp::Sub,
            s if s.starts_with("or") => BinOp::Or,
            s if s.starts_with("xor") => BinOp::BitXor,
            _ => BinOp::Add,
        }),

        Rule::multiplicative => walk_binary_chain(pair, |op_str| match op_str {
            "*" => BinOp::Mul,
            "/" => BinOp::Div,
            s if s.starts_with("div") => BinOp::IDiv,
            s if s.starts_with("mod") => BinOp::Mod,
            s if s.starts_with("and") => BinOp::And,
            s if s.starts_with("shl") => BinOp::Shl,
            s if s.starts_with("shr") => BinOp::Shr,
            _ => BinOp::Mul,
        }),

        Rule::unary => {
            // Pest does not include literal token matches (like "-", "@") as inner
            // pairs — they're consumed silently. Inspect the source text to decide
            // whether this unary node carries a prefix operator.
            let src = pair.as_str().trim_start();
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            // Always exactly one inner pair: either the inner `unary` (when there's
            // a prefix) or the `postfix` (no prefix).
            let operand_pair = inner.pop().ok_or("Empty unary")?;
            let operand = walk_expression(operand_pair)?;

            if src.starts_with('-') {
                Ok(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(operand),
                })
            } else if src.len() >= 3
                && src[..3].eq_ignore_ascii_case("not")
                && !src
                    .chars()
                    .nth(3)
                    .map_or(false, |c| c.is_alphanumeric() || c == '_')
            {
                Ok(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(operand),
                })
            } else if src.starts_with('@') {
                Ok(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(operand),
                })
            } else {
                Ok(operand.kind)
            }
        }

        Rule::postfix => {
            // postfix = { primary ~ postfix_op* }
            let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
            if inner.is_empty() {
                return Err("Empty postfix".into());
            }

            let mut expr = walk_expression(inner.remove(0))?;

            for op in inner {
                if op.as_rule() != Rule::postfix_op {
                    continue;
                }
                expr = walk_postfix_op(expr, op)?;
            }

            Ok(expr.kind)
        }

        Rule::primary => walk_primary(pair),

        // Passthrough for operator pairs that appear in binary chains
        Rule::relational_op | Rule::additive_op | Rule::multiplicative_op | Rule::is_as_op => {
            Err(format!(
                "Operator {:?} should not be walked as expression",
                pair.as_rule()
            ))
        }

        // Literals and identifiers that might appear directly
        Rule::int_literal => {
            let s = pair.as_str();
            if s.starts_with('$') {
                Ok(ExprKind::Lit(Literal::Int(
                    i64::from_str_radix(&s[1..], 16).unwrap_or(0),
                )))
            } else {
                Ok(ExprKind::Lit(Literal::Int(s.parse().unwrap_or(0))))
            }
        }
        Rule::real_literal => Ok(ExprKind::Lit(Literal::Float(
            pair.as_str().parse().unwrap_or(0.0),
        ))),
        Rule::string_literal => {
            let raw = pair.as_str();
            // Strip surrounding quotes and unescape ''
            let inner = &raw[1..raw.len() - 1];
            Ok(ExprKind::Lit(Literal::Str(
                inner.replace("''", "'").to_string(),
            )))
        }
        Rule::char_literal => {
            // #65 → 'A'
            let s = pair.as_str();
            let code: u32 = s[1..].parse().unwrap_or(0);
            Ok(ExprKind::Lit(Literal::Char(
                char::from_u32(code).unwrap_or('\0'),
            )))
        }
        Rule::identifier => Ok(ExprKind::Ident(pair.as_str().to_string())),

        other => Err(format!("Unexpected expression rule: {:?}", other)),
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<ExprKind, String> {
    // Primary can be a keyword literal or have inner pairs
    let src = pair.as_str().trim();
    let src_lower = src.to_lowercase();

    // Check for keyword literals that pest may not produce inner pairs for
    match src_lower.as_str() {
        "true" => return Ok(ExprKind::Lit(Literal::Bool(true))),
        "false" => return Ok(ExprKind::Lit(Literal::Bool(false))),
        "nil" => return Ok(ExprKind::Lit(Literal::Null)),
        "result" => return Ok(ExprKind::Ident("Result".to_string())),
        _ => {}
    }

    let inner = pair.into_inner().next();
    let inner = match inner {
        Some(p) => p,
        None => {
            // If no inner pair, the whole primary text is an identifier or keyword
            // (pest sometimes doesn't create inner pairs for case-insensitive keyword matches)
            return Ok(ExprKind::Ident(src.to_string()));
        }
    };

    match inner.as_rule() {
        Rule::int_literal => walk_expr_kind(inner),
        Rule::real_literal => walk_expr_kind(inner),
        Rule::string_literal => walk_expr_kind(inner),
        Rule::char_literal => walk_expr_kind(inner),
        Rule::identifier => Ok(ExprKind::Ident(inner.as_str().to_string())),
        Rule::set_literal => walk_set_literal(inner),
        Rule::tuple_array_literal => walk_tuple_array_literal(inner),
        Rule::new_expression => walk_new_expression(inner),
        Rule::lambda_procedure => walk_lambda_procedure(inner),
        Rule::lambda_function => walk_lambda_function(inner),
        Rule::inherited_expression => walk_inherited_expression(inner),
        Rule::type_cast_builtin => walk_type_cast_builtin(inner),
        Rule::true_keyword => Ok(ExprKind::Lit(Literal::Bool(true))),
        Rule::false_keyword => Ok(ExprKind::Lit(Literal::Bool(false))),
        Rule::nil_keyword => Ok(ExprKind::Lit(Literal::Null)),
        Rule::result_keyword => Ok(ExprKind::Ident("Result".to_string())),
        Rule::expression => {
            // Parenthesized expression: "(" ~ expression ~ ")"
            walk_expr_kind(inner)
        }
        other => Err(format!("Unexpected primary inner: {:?}", other)),
    }
}

// ── Postfix operations ─────────────────────────────────────────────────────

fn walk_postfix_op(expr: Expression, op: Pair<Rule>) -> Result<Expression, String> {
    let op_src = op.as_str();
    let parts: Vec<Pair<Rule>> = op.into_inner().collect();

    if op_src == "^" {
        // Dereference: ptr^
        return Ok(Expression::new(ExprKind::RefLoad(Box::new(expr))));
    }

    if op_src.starts_with('.') {
        // Field access or method call: obj.Field or obj.Method(args)
        // Grammar: "." ~ identifier ~ arg_list  |  "." ~ identifier
        let mut ident = String::new();
        let mut arg_list: Option<Pair<Rule>> = None;

        for p in &parts {
            match p.as_rule() {
                Rule::identifier => ident = p.as_str().to_string(),
                Rule::arg_list => arg_list = Some(p.clone()),
                _ => {}
            }
        }

        // Canonicalize property-style access for builtins (e.g. arr.Length →
        // __len__(arr)) so the compiler dispatches via compiler_common::canonical.
        // Only when there are no parens — `obj.Length(...)` is a real method call.
        if arg_list.is_none() {
            if let Some(canonical) = canonicalize_pascal_member(&ident) {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(canonical)),
                    args: vec![Argument::positional(expr)],
                    optional: false,
                }));
            }
        }

        let member = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: ident,
            null_safe: false,
        });

        if let Some(al) = arg_list {
            let args = walk_arg_list(al)?;
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args,
                optional: false,
            }));
        } else if matches!(&member.kind, ExprKind::Member { field, .. } if field.eq_ignore_ascii_case("Length"))
        {
            return Ok(Expression::new(ExprKind::Call {
                callee: Box::new(member),
                args: Vec::new(),
                optional: false,
            }));
        } else {
            return Ok(member);
        }
    }

    if op_src.starts_with('[') {
        // Index access: arr[i] or arr[i, j]
        let index_exprs = parts
            .into_iter()
            .flat_map(|p| {
                if p.as_rule() == Rule::array_index_list {
                    p.into_inner().collect::<Vec<_>>()
                } else {
                    vec![p]
                }
            })
            .filter(|p| p.as_rule() == Rule::expression)
            .map(walk_expression)
            .collect::<Result<Vec<_>, _>>()?;

        let mut indexed = expr;
        for index_expr in index_exprs {
            indexed = Expression::new(ExprKind::Index {
                object: Box::new(indexed),
                index: Box::new(index_expr),
                null_safe: false,
            });
        }
        return Ok(indexed);
    }

    if op_src.starts_with('<') && parts.iter().all(|p| p.as_rule() != Rule::arg_list) {
        return Ok(expr);
    }

    if op_src.starts_with('(') || op_src.starts_with('<') {
        // Function call: `F(args)` or generic call `F<T>(args)`. The
        // grammar collapses `generic_args` (silent rule) into the
        // postfix_op, so we just look for the `arg_list` child.
        // Generic type args are captured-and-discarded — the dynamic
        // VM is type-erased.
        let args = parts
            .into_iter()
            .find(|p| p.as_rule() == Rule::arg_list)
            .map(walk_arg_list)
            .transpose()?
            .unwrap_or_default();

        if let ExprKind::Ident(name) = &expr.kind {
            if name.eq_ignore_ascii_case("Format") && args.len() == 2 {
                if let ExprKind::Array(elements) = &args[1].value.kind {
                    let mut expanded_args = vec![args[0].clone()];
                    expanded_args.extend(
                        elements
                            .iter()
                            .map(|element| Argument::positional(element.value.clone())),
                    );
                    return Ok(Expression::new(ExprKind::Call {
                        callee: Box::new(expr),
                        args: expanded_args,
                        optional: false,
                    }));
                }
            }
            if name.eq_ignore_ascii_case("IntToStr") && args.len() == 1 {
                let value = args[0].value.clone();
                return Ok(Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(value),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Str(
                        String::new(),
                    )))),
                }));
            }
        }

        // Canonicalize Pascal's function-style builtins to canonical names so the
        // compiler can dispatch them via compiler_common::canonical regardless of
        // source language. Pascal is case-insensitive.
        if let ExprKind::Ident(name) = &expr.kind {
            if let Some(canonical) = canonicalize_pascal_builtin(name, args.len()) {
                return Ok(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident(canonical)),
                    args,
                    optional: false,
                }));
            }
        }

        // Check if the callee is an identifier that looks like a type cast
        // (e.g. Integer(x), String(x)) — the grammar handles builtin type casts via
        // type_cast_builtin, but identifier-based type casts (e.g. TMyType(x))
        // are just calls — let the compiler handle semantics.
        return Ok(Expression::new(ExprKind::Call {
            callee: Box::new(expr),
            args,
            optional: false,
        }));
    }

    // Fallback
    Ok(expr)
}

// ── Canonical builtin normalization ────────────────────────────────────────

/// Normalize Pascal's function-style builtins to canonical names so the compiler
/// can dispatch them through `compiler_common::canonical`. This keeps language
/// surface syntax in the walker; the compiler stays language-agnostic.
fn canonicalize_pascal_builtin(name: &str, argc: usize) -> Option<&'static str> {
    match (name.to_lowercase().as_str(), argc) {
        ("length", 1) => Some("__len__"),
        _ => None,
    }
}

/// Pascal property-style member access canonicalization (case-insensitive).
fn canonicalize_pascal_member(name: &str) -> Option<&'static str> {
    match name.to_lowercase().as_str() {
        _ => None,
    }
}

// ── Argument list ──────────────────────────────────────────────────────────

fn walk_arg_list(pair: Pair<Rule>) -> Result<Vec<Argument>, String> {
    pair.into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            let value = walk_expression(p)?;
            Ok(Argument::positional(value))
        })
        .collect()
}

// ── Set literal ────────────────────────────────────────────────────────────

fn walk_set_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements: Vec<ArrayElement> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            let value = walk_expression(p)?;
            Ok(ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<_, String>>()?;
    Ok(ExprKind::Array(elements))
}

fn walk_tuple_array_literal(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let elements: Vec<ArrayElement> = pair
        .into_inner()
        .filter(|p| p.as_rule() == Rule::expression)
        .map(|p| {
            let value = walk_expression(p)?;
            Ok(ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
        })
        .collect::<Result<_, String>>()?;
    Ok(ExprKind::Array(elements))
}

// ── New expression ─────────────────────────────────────────────────────────

fn walk_new_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut class_name = String::new();
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => class_name = p.as_str().to_string(),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    Ok(ExprKind::New {
        class: Box::new(Expression::ident(&class_name)),
        args,
    })
}

// ── Lambda expressions ─────────────────────────────────────────────────────

fn walk_lambda_procedure(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body = LambdaBody::Block(Vec::new());

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::compound_statement => {
                body = LambdaBody::Block(walk_compound_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: Vec::new(),
    })
}

fn walk_lambda_function(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut params = Vec::new();
    let mut body = LambdaBody::Block(Vec::new());

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::param_clause => params = walk_param_clause(p)?,
            Rule::type_ref => { /* return type hint — ignored for lambda body */ }
            Rule::compound_statement => {
                body = LambdaBody::Block(walk_compound_statement(p)?);
            }
            _ => {}
        }
    }

    Ok(ExprKind::Lambda {
        params,
        body,
        is_async: false,
        captures: Vec::new(),
    })
}

// ── Inherited expression ───────────────────────────────────────────────────

fn walk_inherited_expression(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut method: Option<String> = None;
    let mut args = Vec::new();

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::identifier => method = Some(p.as_str().to_string()),
            Rule::arg_list => args = walk_arg_list(p)?,
            _ => {}
        }
    }

    if method.is_some() || !args.is_empty() {
        Ok(ExprKind::SuperCall { method, args })
    } else {
        // Bare `inherited` → Super
        Ok(ExprKind::Super)
    }
}

// ── Type cast with builtin type ────────────────────────────────────────────

fn walk_type_cast_builtin(pair: Pair<Rule>) -> Result<ExprKind, String> {
    let mut type_name = String::new();
    let mut expr: Option<Expression> = None;

    for p in pair.into_inner() {
        match p.as_rule() {
            Rule::builtin_type => type_name = p.as_str().to_string(),
            Rule::expression => expr = Some(walk_expression(p)?),
            _ => {}
        }
    }

    Ok(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::null())),
        type_name,
    })
}

// ── Binary chain helper ────────────────────────────────────────────────────

fn walk_binary_chain<F>(pair: Pair<Rule>, map_op: F) -> Result<ExprKind, String>
where
    F: Fn(&str) -> BinOp,
{
    let mut inner: Vec<Pair<Rule>> = pair.into_inner().collect();
    if inner.len() == 1 {
        return walk_expr_kind(inner.remove(0));
    }

    // First operand
    let mut left = walk_expression(inner.remove(0))?;

    // Remaining: (op, operand) pairs
    let mut i = 0;
    while i + 1 < inner.len() {
        let op_str = inner[i].as_str().trim().to_lowercase();
        let right = walk_expression(inner[i + 1].clone())?;
        let bin_op = map_op(&op_str);

        left = Expression::new(ExprKind::Binary {
            op: bin_op,
            left: Box::new(left),
            right: Box::new(right),
        });
        i += 2;
    }

    Ok(left.kind)
}

// ════════════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════════════

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

fn type_ref_to_string(pair: &Pair<Rule>) -> String {
    pair.as_str().trim().to_string()
}

fn extract_array_bounds(pair: &Pair<Rule>) -> Result<Option<Vec<Expression>>, String> {
    for child in pair.clone().into_inner() {
        if matches!(child.as_rule(), Rule::array_type_ref | Rule::array_type) {
            let mut bounds = Vec::new();
            for inner in child.into_inner() {
                if inner.as_rule() == Rule::array_dimension {
                    for expr in inner.into_inner() {
                        if expr.as_rule() == Rule::expression {
                            bounds.push(walk_expression(expr)?);
                        }
                    }
                }
            }
            if !bounds.is_empty() {
                return Ok(Some(bounds));
            }
        }
    }
    Ok(None)
}

/// Flatten a single statement into a Vec — if it's a Block, unwrap it.
fn flatten_stmt(stmt: Statement) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::Block(stmts) => stmts,
        _ => vec![stmt],
    }
}

/// Get first inner pair from a compound pair.
fn cv_first(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    pair.into_inner()
        .next()
        .ok_or_else(|| "Expected inner pair".to_string())
}
